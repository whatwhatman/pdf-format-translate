# -*- coding: utf-8 -*-
"""LLM 翻译层：批量 + 并发 + 缓存 + 术语去重（少量注释）。

速度设计
--------
* 公式已被替换成 {{n}} 占位符，**不进上下文**——这比让模型重打公式快一个数量级，也避免乱码。
* 先跑首批构建「已括注术语表」，再并行处理剩余批次，既保持术语一致，又不会重复注释。
* 结果按原文内容哈希缓存，重跑同一份 PDF 秒级返回。
"""
import asyncio, hashlib, json, os, re, time
from typing import Callable, Dict, List

import httpx

CACHE_DIR = os.path.expanduser('~/.pdftool')
# v2：缓存键加入「公式/图片指纹」，避免不同插图但文本相同的条目互相串味
CACHE_FILE = os.path.join(CACHE_DIR, 'cache_v2.json')
os.makedirs(CACHE_DIR, exist_ok=True)

MAX_TERMS = 120

# ── 项目符号：源 PDF 列表项的行首标志 ──
BULLETS = '•●○▪▫■□◦‣∙⋅·'
# 分隔判定：项目符号前若以这些字符收尾（忽略空格），才认为一项已结束
_ITEM_END = '.,;:!?。，、；：！？）】》」』”’…'


_MARK_VAR = re.compile(r"\[\[\s*['‘’\"“”]?\s*(/?)\s*([bi])\s*['‘’\"“”]?\s*\]\]")
_MARK_JUNK = re.compile(r'\[\[[^\]]{0,12}\]\]')


def fix_markers(s: str) -> str:
    """把模型写坏的强调标记还原成规范形式。

    模型经常把 [[b]] 写成 [['b']]、[[b']]、[[' /b']] 之类，直接渲染会漏出
    字面量（用户看到满篇 [[b']]）。这里统一规范化，并清掉无法还原的 [[...]] 垃圾。
    """
    if not s or '[[' not in s:
        return s
    s = _MARK_VAR.sub(lambda m: '[[%s%s]]' % (m.group(1), m.group(2)), s)
    # 清掉仍然不合法的 [[...]]（占位符 {{n}} 不受影响）
    s = _MARK_JUNK.sub(lambda m: m.group(0) if m.group(0) in
                       ('[[b]]', '[[/b]]', '[[i]]', '[[/i]]') else '', s)
    return s


def tidy(s: str) -> str:
    """清掉译文里的空括号、错位括号等「视觉垃圾」，保留原始排版标记。

    只对普通文本片段生效：{{n}} 公式占位符与 [[b]] 强调标记原样放过。
    """
    if not s:
        return s
    chunks = re.split(r'(\{\{\d+\}\}|\[\[/?[bi]\]\])', s)
    out = []
    for ch in chunks:
        if re.fullmatch(r'\{\{\d+\}\}|\[\[/?[bi]\]\]', ch or ''):
            out.append(ch)
            continue
        t = ch or ''
        if not t:
            continue
        # 假括号：空内容 / 只有空白与标点 → 删
        t = re.sub(r'[（(]\s*[)）]', '', t)
        t = re.sub(r'[（(][\s,，.。、;；:：·・]{0,6}[)）]', '', t)
        # 夹在中文里的半角括号统一成全角——必须整对转换，否则会出现「（E)」这种
        # 一半全角一半半角的畸形括号
        def _pair(m):
            i, nxt = m.start() - 1, m.end()
            while i >= 0 and t[i] == ' ':
                i -= 1
            while nxt < len(t) and t[nxt] == ' ':
                nxt += 1
            prev_ok = i >= 0 and re.match(r'[\u4e00-\u9fff]', t[i])
            next_ok = nxt < len(t) and re.match(r'[\u4e00-\u9fff]', t[nxt])
            return '（%s）' % m.group(1) if (prev_ok or next_ok) else m.group(0)
        t = re.sub(r'\(([^()]{0,40})\)', _pair, t)
        # 同向重复括号：（（x））→（x）
        for _ in range(3):
            t2 = re.sub(r'（\s*（([^（）]*)）\s*）', r'（\1）', t)
            if t2 == t:
                break
            t = t2
        out.append(t)
    s = ''.join(out)
    s = re.sub(r'[ \t]{2,}', ' ', s)
    s = re.sub(r'\s+([，。；：、！？])', r'\1', s)
    s = re.sub(r'([（(])\s+', r'\1', s)
    s = re.sub(r'\s+([）)])', r'\1', s)
    return s.strip()


def strip_bullet(s: str):
    """去掉句首的项目符号（渲染时改用真正的列表符号）。→ (文本, 是否带符号)"""
    if not s:
        return s, False
    t = s.lstrip()
    if t[:1] in BULLETS:
        return t[1:].lstrip(), True
    return t, False


def _bullet_cuts(t: str):
    """找出模板中应另起一项的下标（跳过 {{n}} 占位符内部）。"""
    idx, pos = [], 0
    while pos < len(t):
        if t.startswith('{{', pos):
            j = t.find('}}', pos)
            pos = (j + 2) if j > 0 else len(t)
            continue
        if t[pos] in BULLETS:
            k = pos - 1
            while k >= 0 and t[k] in ' \t':
                k -= 1
            if k < 0 or t[k] in _ITEM_END or t[k] in BULLETS:
                idx.append(pos)
        pos += 1
    return idx


def split_items(blocks) -> int:
    """把「被打平成一段」的列表还原成多个块，恢复原文的换行位置。

    源 PDF 的 itemize 常被行距判定合并进同一个段落，于是出现「…，• …」
    这种内联项目符号。这里按符号切开，渲染器再排成真正的列表。
    （refs 整表复制给新块，未被引用到的条目自然不会被输出。）
    """
    from .extract import Block
    made = 0

    def rec(bs):
        nonlocal made
        out = []
        for b in bs:
            if b.t == 'box':
                b.blocks = rec(b.blocks)
                out.append(b)
            elif b.t == 'para':
                t = b.template or ''
                cs = _bullet_cuts(t)
                if not cs:
                    out.append(b)
                    continue
                segs, prev = [], 0
                for c in cs:
                    segs.append(t[prev:c])
                    prev = c
                segs.append(t[prev:])
                lead = segs[0]
                rest = [s.strip() for s in segs[1:] if s.strip()]
                if lead.strip() or not rest:
                    b.template = lead.rstrip()
                    out.append(b)
                for s in rest:
                    nb = Block(t='para', role='list', level=b.level, size=b.size,
                               bold=b.bold, italic=b.italic, color=b.color,
                               align='left', page=b.page, template=s,
                               refs=b.refs)          # 公式图共享，不做深拷贝
                    out.append(nb)
                    made += 1
            else:
                out.append(b)
        return out

    blocks[:] = rec(blocks)
    return made


def _load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            return json.load(open(CACHE_FILE, encoding='utf-8'))
        except Exception:
            return {}
    return {}


def _save_cache(c):
    try:
        json.dump(c, open(CACHE_FILE, 'w', encoding='utf-8'))
    except Exception:
        pass


ANNOT_RULE = (
    "5. 术语注释：专业术语首次出现时，用半角括号紧附原文（如 “拉普拉斯变换（transformée de Laplace）”、"
    "“模（module）”、“辐角（argument）”）。只注专业术语，每个术语全文仅注一次；"
    "若列表里已出现过该术语则不要再注。日常用词、普通动词不要括注。"
)


def build_prompt(items: Dict[str, str], src: str, tgt: str, extra_terms: List[str],
                 strict: bool = False) -> str:
    extra = ''
    if extra_terms:
        extra = ('\n以下术语在本文档的前文中已经括注过，**不要再重复括注**：'
                 + '、'.join(extra_terms[:MAX_TERMS]) + '\n')
    if strict:
        extra += ('\n【严格要求】下列条目上一轮被原样退回，必须逐条译出：'
                  '即使只是一个单词或只是一个标题，也要给出%s译文，不得保留任何%s原词。\n'
                  % (tgt, src))
        extra += '\n再次强调：每个编号只译它自己那一条，不要与前后条目连起来译，不要输出换行。\n'
    head = (
        f"你是精通 {src} 和 {tgt} 的学术技术文档译者。请将下列编号条目逐条翻译成{tgt}。\n"
        "严格要求：\n"
        "1. 原样保留 {{数字}} 占位符，个数与相对位置不得改变，不得改写其中的数字。\n"
        "1a. 每条独立成译：严禁把相邻条目合并、严禁在译文里换行（需要断句一律用空格）。\n"
        "2. 原样保留 [[b]]…[[/b]] 与 [[i]]…[[/i]] 强调标记的个数与包裹范围。\n"
        "3. 忠于原文语义，不要增译、不要写解释性旁白、不要添加原文没有的标题。\n"
        "4. 数学符号、单位、编号、人名保持原样；专有名词保留原文或采用通行中译。\n"
        "4b. 章节标题、小节标题、图表题注一律要翻译，条目中不得残留未译的外文单词。\n"
        "4c. 条目里的 • 是列表项标记：必须原样保留，位置与个数不得改变，"
        "不要改写成逗号、顿号、破折号或任何别的符号，也不要合并或删除。\n"
        + ANNOT_RULE + '\n'
        "6. 只输出 JSON 对象，形如 {\"0\": \"译文\", \"1\": \"译文\"}，不要输出任何其它文本。\n"
        + extra
    )
    body = '\n'.join('%s\t%s' % (k, v.replace('\n', ' ')) for k, v in items.items())
    return head + '\n----\n' + body


def repair(src: str, out: str) -> str:
    """把译文里的 {{n}} 占位符与 [[b]] 标记校正到与原文一致。

    模型偶尔会重排、漏掉或多吐占位符，而占位符直接对应公式插图，错一个就张冠李戴。
    这里按出现顺序重映射；数量不足时把缺失的补到末尾，多余的直接丢弃。
    """
    if not src:
        return out
    S = re.findall(r'\{\{\d+\}\}', src)
    T = re.findall(r'\{\{\d+\}\}', out)
    if T != S:
        if not T:                                  # 全丢了：把公式插在末尾
            out = out.rstrip() + ' ' + ' '.join(S)
        elif len(T) >= len(S):
            for i, old in enumerate(T):            # 多余的在末尾删掉
                if i < len(S):
                    out = out.replace(old, S[i], 1)
                else:
                    out = out.replace(old, '', 1)
        else:
            for i, old in enumerate(T):
                out = out.replace(old, S[i], 1)
            out = out.rstrip() + ' ' + ' '.join(S[len(T):])
    # 强调标记：保证 [[b]] 与 [[/b]] 成对
    nb, nbe = out.count('[[b]]'), out.count('[[/b]]')
    if nb > nbe:
        out += '[[/b]]' * (nb - nbe)
    elif nbe > nb:
        out = '[[b]]' * (nbe - nb) + out
    return out


class Translator:
    def __init__(self, base_url='https://api.deepseek.com/v1', api_key='', model='deepseek-chat',
                 concurrency=8, batch_chars=1400, temperature=0.2, proxy=None):
        self.base_url = base_url.rstrip('/')
        self.api_key = api_key
        self.model = model
        self.concurrency = concurrency
        self.batch_chars = batch_chars
        self.temperature = temperature
        self.proxy = proxy or None
        self.cache = _load_cache()

    # ---------- 单个批次 ----------
    async def _call(self, client, items, src, tgt, terms, strict=False):
        payload = {
            'model': self.model,
            'messages': [{'role': 'user', 'content': build_prompt(items, src, tgt, terms, strict)}],
            'temperature': self.temperature,
        }
        try:
            payload['response_format'] = {'type': 'json_object'}
        except Exception:
            pass
        r = await client.post(self.base_url + '/chat/completions', json=payload, timeout=180)
        r.raise_for_status()
        data = r.json()
        txt = data['choices'][0]['message']['content']
        obj = None
        try:
            obj = json.loads(txt)
        except Exception:
            m = re.search(r'\{.*\}', txt, re.S)
            if m:
                try:
                    obj = json.loads(m.group(0))
                except Exception:
                    obj = None
        if not isinstance(obj, dict):
            raise ValueError('模型未返回合法 JSON')
        return obj

    async def _batch(self, client, sem, items, src, tgt, terms, log, strict=False):
        """译一批。

        关键：模型几乎总是忽略我们给的键、自己从 0 重新编号输出，
        所以这里把条目改成本批的 0..n-1 再发，回包按位对号（并兼容按 key 命中）。
        """
        keys = list(items)
        local = {str(i): items[k] for i, k in enumerate(keys)}
        async with sem:
            for attempt in range(2):
                try:
                    # 硬截止：SSE 流偶尔会一直滴心跳而不吐正文，read-timeout 永远
                    # 不触发（实测卡了 27 分钟）。用整体墙钟超时兜住，宁可失败保留原文。
                    obj = await asyncio.wait_for(
                        self._call(client, local, src, tgt, terms, strict),
                        timeout=getattr(self, 'hard_timeout', 150))
                    vals = [obj.get(str(i)) for i in range(len(keys))]
                    if not any(v for v in vals):          # 键形制不认识 → 按出现顺序兜底
                        vals = list(obj.values())[:len(keys)]
                    return {k: self._clean_val(vals[i] if i < len(vals) else '',
                                               items[k])
                            for i, k in enumerate(keys)}
                except Exception as e:
                    if attempt == 1:
                        log('批次失败，保留原文：%s' % str(e)[:90])
                        return dict(items)
                    await asyncio.sleep(1.5 * (attempt + 1))
            return dict(items)

    @staticmethod
    def _clean_val(v, fallback):
        v = '' if v is None else str(v)
        v = v.replace('\r', ' ').replace('\n', ' ').strip()
        return v or fallback

    @staticmethod
    def _find_merged(todo, texts, out):
        """找出被模型「连条」译出的条目下标。"""
        def plain(s):
            return re.sub(r'\{\{\d+\}\}|\[\[/?[bi]\]\]|\s', '', s or '')
        bad, seen = [], {}
        for i in todo:
            v = out[i]
            if not v:
                continue
            p = plain(v)
            # ① 不同原文却拿到完全一样的译文
            if p in seen and plain(texts[i]) != plain(texts[seen[p]]):
                bad += [i, seen[p]]
            else:
                seen.setdefault(p, i)
            # ② 译文比原文长出一大截（通常是吞掉了后一条）
            src = plain(texts[i])
            if len(src) > 6 and len(p) > len(src) * 1.8 + 12:
                bad.append(i)
        return sorted(set(bad))


    # ---------- 主入口 ----------
    async def run(self, texts: List[str], src='法语', tgt='简体中文',
                  progress: Callable = None, log: Callable = None,
                  sigs: List[str] = None) -> List[str]:
        log = log or (lambda m: None)
        progress = progress or (lambda f, m: None)
        sigs = sigs or [''] * len(texts)
        # 缓存键必须带上插图指纹：文本相同而公式不同的两个条目不能共用译文
        keys = [hashlib.md5(('%s\x01%s' % (t, sigs[i] if i < len(sigs) else ''))
                            .encode('utf-8')).hexdigest()
                for i, t in enumerate(texts)]
        out = []
        todo = []
        for i, t in enumerate(texts):
            k = keys[i]
            if has_text(t) and k in self.cache:
                out.append(self.cache[k])
            else:
                out.append(None)
                if has_text(t):
                    todo.append(i)
                else:
                    out[i] = t          # 没有文字可译，原样返回
        progress(0.05, '缓存命中 %d / %d' % (len(texts) - len(todo), len(texts)))

        if not todo:
            return [o if o is not None else t for t, o in zip(texts, out)]

        # 分批
        batches, cur, cur_sz = [], [], 0
        for idx in todo:
            cur.append(idx)
            cur_sz += len(texts[idx]) + 8
            if cur_sz > self.batch_chars or len(cur) >= 40:
                batches.append(cur); cur, cur_sz = [], 0
        if cur:
            batches.append(cur)
        log('共 %d 段待译，分为 %d 个批次' % (len(todo), len(batches)))

        sem = asyncio.Semaphore(self.concurrency)
        async with httpx.AsyncClient(trust_env=True, proxy=self.proxy) as client:
            # 先译首批，抽取已括注术语，供后续批次去重
            def items_of(bidxs):
                return {str(j): texts[j] for j in bidxs}

            res0 = await self._batch(client, sem, items_of(batches[0]), src, tgt, [], log)
            for j, v in res0.items():
                out[int(j)] = v
            terms = []
            for v in res0.values():
                for m in re.finditer(r'（([^（）]{2,36})）', v):
                    t = m.group(1).strip()
                    if re.search(r'[A-Za-zÀ-ÿ]', t) and t not in terms:
                        terms.append(t)
            progress(0.15, '首批完成，已收集 %d 个术语' % len(terms))

            done = 1
            if len(batches) > 1:
                tasks = [self._batch(client, sem, items_of(b), src, tgt, terms, log)
                         for b in batches[1:]]
                for coro in asyncio.as_completed(tasks):
                    r = await coro
                    for j, v in r.items():
                        out[int(j)] = v
                    # 术语表持续累积，后续批次重复注解越来越少
                    for m in re.finditer(r'（([^（）]{2,36})）', ' '.join(r.values())):
                        t = m.group(1).strip()
                        if re.search(r'[A-Za-zÀ-ÿ]', t) and t not in terms:
                            terms.append(t)
                    done += 1
                    progress(0.15 + 0.7 * done / len(batches),
                             '已译 %d/%d 段' % (done, len(batches)))

            # 自检：模型偶尔会把相邻条目连起来译（表现为两条拿到同一份译文、
            # 或译文字数明显超出原文）。发现就逐条单独重译，成本极低但很关键。
            bad = self._find_merged(todo, texts, out)
            if bad:
                bad = bad[:40]
                progress(0.88, '重译 %d 段疑似串条的译文' % len(bad))
                log('自检发现 %d 段疑似与相邻条目串译，逐条重译' % len(bad))
                tasks = [self._batch(client, sem, {str(i): texts[i]},
                                     src, tgt, terms, log, strict=True) for i in bad]
                for coro in asyncio.as_completed(tasks):
                    r = await coro
                    for k, v in r.items():
                        if v:
                            out[int(k)] = v

            # 补译：模型原样退回的条目（多为短标题）再走一遍严格指令
            stuck = [i for i, (t, v) in enumerate(zip(texts, out))
                     if v is not None and t.strip() and v.strip() == t.strip()
                     and len(re.sub(r'[^A-Za-zÀ-ÿ]', '', t)) > 3]
            if stuck:
                progress(0.9, '补译 %d 段未译条目' % len(stuck))
                log('补译 %d 段原样退回的条目' % len(stuck))
                r = await self._batch(client, sem, {str(i): texts[i] for i in stuck},
                                      src, tgt, terms, log, strict=True)
                for k, v in r.items():
                    out[int(k)] = v

        out = [tidy(repair(fix_markers(t), fix_markers(v))) if v is not None
               else tidy(fix_markers(t)) for t, v in zip(texts, out)]
        for i, v in enumerate(out):
            if v != texts[i]:
                self.cache[keys[i]] = v
        _save_cache(self.cache)
        progress(0.9, '翻译完成')
        return out


# 测速探针用的固定短文本（与真实任务同源的句子，长度接近，便于横向比较）
PROBE_TEXTS = [
    'La somme de deux nombres complexes est commutative et associative.',
    'Le module du produit de deux nombres complexes est le produit des modules.',
    'Pour tout nombre complexe non nul z, il existe un inverse égal à un sur z.',
]


class CloudTranslator(Translator):
    """WorkBuddy 云服务的免密钥大模型（用户无需自备 API Key）。

    仅支持流式（SSE），因此要自己把 delta 拼起来；思考型模型会先吐
    reasoning_content，这里只取 content。
    """

    def __init__(self, endpoint='', publishable_key='', model='default',
                 concurrency=8, batch_chars=1400, temperature=0.2,
                 request_timeout=100, **kw):
        super().__init__(base_url=endpoint, api_key='', model=model,
                         concurrency=concurrency, batch_chars=batch_chars,
                         temperature=temperature)
        self.endpoint = endpoint.rstrip('/')
        self.publishable_key = publishable_key
        self.header = 'x-wb-webapp-access-key'
        self.request_timeout = request_timeout

    def models(self):
        import urllib.request
        r = urllib.request.Request(self.endpoint + '/.cloud/llm/models',
                                   headers={self.header: self.publishable_key})
        with urllib.request.urlopen(r, timeout=30) as f:
            data = json.loads(f.read().decode('utf-8'))
        return [m for m in data if m.get('enabled') is not False]

    async def probe(self, client, timeout=45):
        """测速探针：发一小批固定文本，返回 {ok, sec, ttft, note}。

        用同一个提示词跑每个模型，比较「总耗时」与「首字延迟」；顺便校验
        它是不是真的输出中文（有些模型会原样返回或输出英文）。
        """
        import time as _t
        items = {str(i): t for i, t in enumerate(PROBE_TEXTS)}
        t0 = _t.time()
        box = []
        obj = await asyncio.wait_for(
            self._call(client, items, '法语', '简体中文', [], False,
                       first_cb=lambda: box.append(_t.time() - t0)),
            timeout=timeout)
        sec = _t.time() - t0
        vals = [str(obj.get(str(i), '')) for i in range(len(items))]
        zh = sum(1 for v in vals if any('\u4e00' <= c <= '\u9fff' for c in v))
        return {'ok': zh >= max(1, len(items) // 2), 'sec': round(sec, 2),
                'ttft': round(box[0], 2) if box else None,
                'zh': zh, 'n': len(items)}

    async def _call(self, client, items, src, tgt, terms, strict=False,
                    first_cb=None):
        system = ('你是精通 %s 和 %s 的学术技术文档译者，译文准确、通顺、术语一致。'
                  % (src, tgt))
        payload = {
            'model': self.model,
            'stream': True,
            'messages': [{'role': 'system', 'content': system},
                         {'role': 'user', 'content': build_prompt(items, src, tgt, terms, strict)}],
            'temperature': self.temperature,
        }
        url = self.endpoint + '/.cloud/llm/chat/completions'
        headers = {self.header: self.publishable_key, 'Accept': 'text/event-stream'}
        txt = ''
        # 单批读超时：个别批次可能长时间不吐字，宁可失败重试也不要卡死整条流水线
        async with client.stream('POST', url, json=payload, headers=headers,
                                 timeout=self.request_timeout) as r:
            r.raise_for_status()
            async for line in r.aiter_lines():
                if not line.startswith('data:'):
                    continue
                d = line[5:].strip()
                if not d or d == '[DONE]':
                    continue
                try:
                    ch = json.loads(d)
                except Exception:
                    continue
                chs = ch.get('choices') or [{}]
                delta = chs[0].get('delta') or {}
                if delta.get('content'):
                    if first_cb is not None:
                        try:
                            first_cb()
                        except Exception:
                            pass
                        first_cb = None
                    txt += delta['content']
        obj = None
        try:
            obj = json.loads(txt)
        except Exception:
            m = re.search(r'\{.*\}', txt, re.S)
            if m:
                try:
                    obj = json.loads(m.group(0))
                except Exception:
                    obj = None
        if not isinstance(obj, dict):
            raise ValueError('模型未返回合法 JSON：%s' % txt[:120])
        return obj


_MARKERS = re.compile(r'\{\{\d+\}\}|\[\[/?[bi]\]\]')


def has_text(t: str) -> bool:
    """该模板是否真的有需要翻译的文字（纯公式图段落可直接跳过）。

    这类段落的模板一律是 "{{0}}"，若拿去翻译不仅浪费 token，
    还会因为文本完全相同而共享同一份缓存译文 → 公式串到别处去。
    """
    if not t:
        return False
    return bool(re.search(r'\w', _MARKERS.sub('', t)))


def unit_sig(b) -> str:
    """公式/插图指纹：文本相同但图不同的两个条目，不能共用一份译文。"""
    parts = []
    for i in sorted(getattr(b, 'refs', {}) or {}):
        r = b.refs[i]
        if getattr(r, 'kind', '') == 'math':
            parts.append('m:' + str(getattr(r, 'latex', '')))
        else:
            png = getattr(r, 'png', b'') or b''
            parts.append('i:%d:%s' % (len(png),
                                      hashlib.md5(png[:8192]).hexdigest()[:12]))
    return '|'.join(parts)


def collect_templates(blocks):
    """深度优先收集所有需要翻译的模板串（含盒内内容与页眉页脚）。"""
    out = []

    def rec(bs):
        for b in bs:
            if b.t == 'box':
                rec(b.blocks)
            elif b.t == 'para' and has_text(b.template):
                out.append(b)
    rec(blocks)
    return out


def assign(blocks, values):
    """把译文按收集顺序写回各 Block.template。"""
    ii = 0
    def rec(bs):
        nonlocal ii
        for b in bs:
            if b.t == 'box':
                rec(b.blocks)
            elif b.t == 'para' and has_text(b.template):
                b.template = values[ii] if ii < len(values) else b.template
                ii += 1
    rec(blocks)
    return ii
