"""扫描件（无文字层 PDF）处理：整页交给视觉模型翻译。

为什么不是 OCR → 现有文本管道：
数学讲义经 Tesseract 后 θ→0、∀θ∈]0,π/2[→"vo e]0, Zb"，分数、上下标、根号等
结构全部丢失——而公式保真正是本工具的核心价值。视觉模型直接读整页图，
能还原 re^{iφ}、∀θ∈]0,π/2] 这类结构（实测远端 shelf，见 scan 对照报告）。
OCR（含 scan-ocr-benchmark 技能的方法）退居：
  ① 评测视觉模型译文的准确率基线；② 无视觉模型时的兜底。
"""
import asyncio
import re


def scan_ratio(pdf_path: str, sample: int = 6) -> float:
    """返回「无文字层页面占比」：1.0 = 完全扫描件，0.0 = 完全有文字层。"""
    import pymupdf
    d = pymupdf.open(pdf_path)
    n = len(d)
    idxs = range(n) if n <= sample else [int(i * n / sample) for i in range(sample)]
    blank = 0
    total = 0
    for i in idxs:
        p = d[i]
        total += 1
        txt = p.get_text().strip()
        # 文字极少（或许是页码/印章）且含至少一张大位图 ⇒ 这一页是影印图
        imgs = p.get_images(full=True)
        big = any(im[2] >= 300 and im[3] >= 300 for im in imgs)
        if len(txt) < 20 and big:
            blank += 1
    d.close()
    return (blank / total) if total else 0.0


PAGE_PROMPT = """把这一页%src%讲义页面翻译成%tgt%。要求：
1. 输出 Markdown：标题用 #/##，列表用 - ，强调用 **粗体**。
2. 数学公式一律用 LaTeX：行内 $...$，独立成行的用 $$...$$。必须保留原公式含义，
   上下标、分数、根号、求和号等结构写全，不要把 θ 写成 0。
3. 数字、小数位数必须照抄原文，一个都不要多写或少写（例如原文只给到
   3,141 592 653 589 793…，就绝不能自己补更多位）。
4. 图片/示意图用一行 > 图：<简短中文说明> 表示。
5. 不要保留原文；不要加解释、不要前后寒暄；只输出本页译文。
6. 若本页几乎空白或只有页码，输出单个词：BLANK"""


async def translate_pages(pdf_path, tr, client, src='法语', tgt='简体中文',
                          dpi=170, model=None, log=None, prog=None,
                          concurrency=4):
    """逐页渲染成图，交给视觉模型翻译。返回 [(page_index, markdown)]。"""
    import pymupdf
    log = log or (lambda m: None)
    prog = prog or (lambda f, m: None)
    d = pymupdf.open(pdf_path)
    total = len(d)

    async def one(i):
        p = d[i]
        pix = p.get_pixmap(dpi=dpi)
        md = ''
        for attempt in (1, 2):
            try:
                md = await tr.vision_page(
                    client, pix.tobytes('png'),
                    PAGE_PROMPT.replace('%src%', src).replace('%tgt%', tgt),
                    model=model)
                break
            except Exception as e:
                if attempt == 2:
                    log('第 %d 页视觉翻译失败：%s' % (i + 1, str(e)[:90]))
                    md = ''
        return (i, clean_md(md))

    sem = asyncio.Semaphore(concurrency)
    done = 0

    async def guarded(i):
        nonlocal done
        async with sem:
            r = await one(i)
        done += 1
        prog(min(0.95, done / max(1, total)),
             '视觉翻译 %d/%d 页' % (done, total))
        return r

    res = await asyncio.gather(*[guarded(i) for i in range(total)])
    d.close()
    return sorted(res, key=lambda t: t[0])


def clean_md(t: str) -> str:
    """清掉模型常见的包裹式芜杂（代码块围栏、前后寒暄）。"""
    t = (t or '').strip()
    if not t:
        return ''
    if t.strip().upper() in ('BLANK', 'BLANK.'):
        return ''
    m = re.match(r'^```(?:markdown|md)?\s*\n(.*?)\n```\s*$', t, re.S)
    if m:
        t = m.group(1).strip()
    # 去掉首行可能的「以下是译文：」之类
    t = re.sub(r'^(以下是|下面是)?[^#\n]{0,30}译文[：:]\s*\n', '', t)
    return t.strip()
