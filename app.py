# -*- coding: utf-8 -*-
"""pdft —— 上传 PDF，得到「原格式的中文译本」。

用法
----
  Web:  python app.py            → 打开 http://127.0.0.1:8765
  CLI:  python app.py in.pdf -o out.docx
"""
import asyncio, json, os, re, shutil, sys, threading, time, traceback, uuid
from typing import Dict

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from pdftrans import extract as X, translate as T, build as B

APP_DIR = os.path.dirname(os.path.abspath(__file__))
HOME = os.environ.get('PDFTOOL_HOME', os.path.expanduser('~/.pdftool'))
os.makedirs(HOME, exist_ok=True)
CONFIG = os.path.join(HOME, 'config.json')
JOBS: Dict[str, dict] = {}

# ── 公共服务限制 ──
MAX_UPLOAD = 30 * 1024 * 1024          # 单文件上限 30 MB
MAX_PAGES = 60                          # 单文档页数上限
MAX_CONCURRENT = 2                      # 同时处理的任务数
JOB_TTL = 2 * 3600                      # 任务产物保留 2 小时
_SEM = threading.Semaphore(MAX_CONCURRENT)

# ── 内置免密钥通道（WorkBuddy 云服务，已为本应用开通）──
# 密钥**绝不写进代码**：只从环境变量 WB_PUBLISHABLE_KEY 读取，
# 或放在同目录的 .env 文件里（.env 已列入 .gitignore，不会进仓库）。
# 部署时通过启动命令注入：WB_PUBLISHABLE_KEY=xxx python app.py
def _load_dotenv():
    p = os.path.join(APP_DIR, '.env')
    if not os.path.exists(p):
        return
    try:
        for line in open(p, encoding='utf-8'):
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            k, v = line.split('=', 1)
            k, v = k.strip(), v.strip().strip('"').strip("'")
            if k and k not in os.environ:          # 真实环境变量优先
                os.environ[k] = v
    except Exception:
        pass


_load_dotenv()

BUILTIN = {
    'endpoint': os.environ.get('WB_ENDPOINT',
                               'https://pdf-translator.app.workbuddy.host'),
    'publishable_key': os.environ.get('WB_PUBLISHABLE_KEY', ''),
}
BUILTIN_ENABLED = bool(BUILTIN['publishable_key'])

DEFAULT_CFG = {
    'engine': 'builtin',        # builtin | custom | none
    'model': 'deepseek-v4-flash',
    # 自定义接口（选填）
    'base_url': 'https://api.deepseek.com/v1',
    'api_key': '',
    'custom_model': 'deepseek-chat',
    'concurrency': 8,
    'src': '法语',
    'tgt': '简体中文',
    'mode': 'faithful',     # faithful=公式保真成图 | editable=尝试重建可编辑公式
    'dpi': 300,
    'formats': ['docx', 'pdf'],
}

app = FastAPI(title='pdft')

# 允许来自 GitHub Pages 等静态前端的跨域调用（前端部署在别处时用）。
# 可用 CORS_ORIGINS 环境变量覆盖，逗号分隔；默认放行两类来源：
#   - 本服务自身（同源，用于本地 static 页面）
#   - https://*.github.io（GitHub Pages 前端）
_CORS = os.environ.get('CORS_ORIGINS', '').strip()
_ORIGINS = [o.strip() for o in _CORS.split(',') if o.strip()] or [
    'https://pdf-translator.app.workbuddy.host',
    'https://whatwhatman.github.io',
]
from fastapi.middleware.cors import CORSMiddleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=_ORIGINS,
    allow_origin_regex=r'https://[a-z0-9-]+\.github\.io',
    allow_methods=['*'],
    allow_headers=['*'],
)


def load_cfg():
    c = dict(DEFAULT_CFG)
    if os.path.exists(CONFIG):
        try:
            c.update(json.load(open(CONFIG, encoding='utf-8')))
        except Exception:
            pass
    # 一次性迁移：早期默认模型是 default，实测比 deepseek-v4-flash 慢约 10 倍
    # （同负载 9 批×16 条：38~129s vs 5s）。老配置里的 default 自动升级。
    if int(c.get('_cfg_version', 1)) < 2:
        if str(c.get('model', '')).strip() in ('', 'default'):
            c['model'] = DEFAULT_CFG['model']
        c['_cfg_version'] = 2
        try:
            save_cfg(c)
        except Exception:
            pass
    return c


def save_cfg(c):
    json.dump(c, open(CONFIG, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)


# ─────────────────────────── 核心流水线 ───────────────────────────
def make_translator(cfg, log):
    engine = cfg.get('engine', 'builtin')
    if engine == 'builtin':
        if not BUILTIN_ENABLED:
            raise RuntimeError(
                '内置免密钥通道未配置：请设置环境变量 WB_PUBLISHABLE_KEY'
                '（或在同目录建 .env 写入该变量），或改用「自定义接口」填自己的 API Key。')
        log('使用内置免密钥通道（WorkBuddy 云服务），模型：%s' % cfg.get('model'))
        return T.CloudTranslator(endpoint=BUILTIN['endpoint'],
                                 publishable_key=BUILTIN['publishable_key'],
                                 model=cfg.get('model') or 'deepseek-v4-flash',
                                 concurrency=int(cfg.get('concurrency', 8)))
    if engine == 'custom':
        if not cfg.get('api_key'):
            return None
        log('使用自定义接口：%s / %s' % (cfg.get('base_url'), cfg.get('custom_model')))
        return T.Translator(base_url=cfg.get('base_url'), api_key=cfg.get('api_key'),
                            model=cfg.get('custom_model'),
                            concurrency=int(cfg.get('concurrency', 8)))
    return None


def run_pipeline(job: dict, pdf_path: str, cfg: dict):
    t0 = time.time()
    logs = job['logs']

    def log(m):
        logs.append('%.1fs  %s' % (time.time() - t0, m))

    def prog(f, m):
        job['progress'] = round(min(0.99, f) * 100)
        job['message'] = m

    try:
        prog(0.01, '正在抽取版式…')
        doc = X.extract(pdf_path, mode=cfg.get('mode', 'faithful'),
                        dpi=int(cfg.get('dpi', 300) or 300), progress=prog)
        if doc.pages > MAX_PAGES:
            raise RuntimeError('文档 %d 页，超过公共服务的 %d 页上限，请拆分后上传'
                               % (doc.pages, MAX_PAGES))
        log('抽取完成：%d 页，%d 个内容块' % (doc.pages, len(doc.blocks)))

        body_units = T.collect_templates(doc.blocks)
        hf_units = T.collect_templates(doc.header + doc.footer)
        units = body_units + hf_units
        log('待译文本 %d 段（正文 %d + 页眉页脚 %d）'
            % (len(units), len(body_units), len(hf_units)))

        tr = make_translator(cfg, log)
        if tr is None:
            log('未启用翻译引擎，仅做版式还原')
        else:
            texts = [b.template for b in units]
            vals = asyncio.run(tr.run(texts, src=cfg.get('src', '法语'),
                                      tgt=cfg.get('tgt', '简体中文'),
                                      progress=prog, log=log,
                                      sigs=[T.unit_sig(b) for b in units]))
            T.assign(doc.blocks, vals[:len(body_units)])
            T.assign(doc.header + doc.footer, vals[len(body_units):])

        # 后处理：恢复被挤成一行的列表、清掉空括号等视觉垃圾
        n1 = T.split_items(doc.blocks)
        n2 = T.split_items(doc.header) + T.split_items(doc.footer)

        def _clean(bs):
            for b in bs:
                if b.t == 'para':
                    b.template = T.tidy(b.template or '')
                elif b.t == 'box':
                    _clean(b.blocks)

        for bs in (doc.blocks, doc.header, doc.footer):
            _clean(bs)
        if n1 or n2:
            log('还原列表项 %d 个（正文 %d / 页眉页脚 %d）' % (n1 + n2, n1, n2))

        base = os.path.join(job['dir'], os.path.splitext(os.path.basename(pdf_path))[0])
        prog(0.9, '正在生成文件…')
        job['files'] = []
        fmts = cfg.get('formats') or ['docx']
        html_path = None
        if 'pdf' in fmts or 'html' in fmts:
            html_path = base + '.译本.html'
            B.render_html(doc, html_path, progress=prog)
        if 'docx' in fmts:
            p = base + '.译本.docx'
            B.render_docx(doc, p, progress=prog)
            job['files'].append({'name': os.path.basename(p), 'size': os.path.getsize(p)})
        if 'pdf' in fmts:
            p = base + '.译本.pdf'
            log('正在用浏览器打印 PDF…')
            try:
                B.render_pdf(doc, p, html_path=html_path)
                # ── 自动适配：原 PDF 每页 97% 满时，译文常会溢出到下一页，
                # 出现大量半空白页。这里量一下页数，超出就等比缩放重排一次，
                # 让「译文第 N 页 ≈ 原文第 N 页」，最多重排一次（幂等、可控）。
                try:
                    import pymupdf as _fz
                    src_pages = sum(1 for b in doc.blocks if b.t == 'pagebreak') + 1
                    got = len(_fz.open(p))
                    # ── 迭代重排：一轮缩放可能仍差一两页（强制分页把溢出内容
                    # 推到下一页、留下半空页）。先按页数比估一档，仍不齐就再
                    # 降 4%，直到页数对齐、页数不再下降或触到 80% 下限。
                    z = min(0.97, max(0.80, ((src_pages + 0.4) / max(got, 1)) ** 0.5)) \
                        if src_pages else 1.0
                    prev_z = None
                    for _ in range(4):
                        if not (src_pages and got > src_pages) or z < 0.795:
                            break
                        if prev_z is not None:
                            est = min(0.97, max(0.80, ((src_pages + 0.4) / max(got, 1)) ** 0.5))
                            z = est if est < prev_z - 0.005 else max(0.80, prev_z - 0.04)
                        prev_z = z
                        B.render_html(doc, html_path, zoom=z)
                        B.render_pdf(doc, p, html_path=html_path)
                        new_got = len(_fz.open(p))
                        log('已按 %.0f%% 缩放重排：%d 页 → %d 页（原文件 %d 页）'
                            % (z * 100, got, new_got, src_pages))
                        if new_got >= got:
                            z -= 0.04          # 没起色 → 更激进地缩
                        got = new_got
                except Exception as e:
                    log('页数适配跳过：%s' % str(e)[:80])
                job['files'].append({'name': os.path.basename(p),
                                     'size': os.path.getsize(p)})
            except Exception as e:
                log('PDF 生成失败：%s' % str(e)[:120])
        if 'html' in fmts:
            job['files'].append({'name': os.path.basename(html_path),
                                 'size': os.path.getsize(html_path)})
        job['done'] = True
        job['progress'] = 100
        job['message'] = '完成，用时 %.1f 秒' % (time.time() - t0)
        log(job['message'])
    except Exception as e:
        job['error'] = str(e)
        job['message'] = '出错：%s' % e
        job['done'] = True
        log(traceback.format_exc()[-600:])


# ─────────────────────────── API ───────────────────────────
@app.get('/', response_class=HTMLResponse)
def index():
    return open(os.path.join(APP_DIR, 'static', 'index.html'), encoding='utf-8').read()


app.mount('/static', StaticFiles(directory=os.path.join(APP_DIR, 'static')), name='static')


@app.get('/api/config')
def get_cfg():
    c = load_cfg()
    c['api_key'] = '***已保存***' if c.get('api_key') else ''
    return c


@app.post('/api/config')
def post_cfg(c: dict):
    old = load_cfg()
    if c.get('api_key') in ('', '***已保存***'):
        c['api_key'] = old.get('api_key', '')
    if 'formats' not in c:
        c['formats'] = old.get('formats', DEFAULT_CFG['formats'])
    old.update({k: v for k, v in c.items()
                if k in DEFAULT_CFG and str(v).strip() != ''})
    if not str(old.get('concurrency', '')).isdigit():
        old['concurrency'] = DEFAULT_CFG['concurrency']
    save_cfg(old)
    return {'ok': True}


@app.get('/api/models')
def get_models():
    """内置免密钥通道的可用模型列表。"""
    if not BUILTIN_ENABLED:
        return {'models': [], 'error': '内置通道未配置密钥（WB_PUBLISHABLE_KEY）'}
    try:
        tr = T.CloudTranslator(endpoint=BUILTIN['endpoint'],
                               publishable_key=BUILTIN['publishable_key'])
        ms = tr.models()
        out = []
        for m in ms:
            if m.get('id') in ('hunyuan-image-v3.0',):
                continue
            out.append({'id': m['id'], 'name': m.get('name') or m['id'],
                        'desc': m.get('descriptionZh') or m.get('descriptionEn') or ''})
        return {'models': out}
    except Exception as e:
        return JSONResponse({'models': [], 'error': str(e)[:200]}, status_code=200)


# ─────────────────── 模型测速（给用户推荐最快的那一个） ───────────────────
SPEED_FILE = os.path.join(HOME, 'speed.json')
SPEED_TTL = 15 * 60            # 测速结果 15 分钟内复用，不必反复烧额度
SPEED = {'running': False, 'ts': 0, 'results': [], 'done': 0, 'total': 0,
         'error': '', 'note': ''}


def _load_speed():
    try:
        d = json.load(open(SPEED_FILE, encoding='utf-8'))
        SPEED.update({k: d.get(k, SPEED[k]) for k in
                      ('ts', 'results', 'done', 'total')})
    except Exception:
        pass


def _save_speed():
    try:
        json.dump({'ts': SPEED['ts'], 'results': SPEED['results'],
                   'done': SPEED['done'], 'total': SPEED['total']},
                  open(SPEED_FILE, 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1)
    except Exception:
        pass


async def _bench_models(models, log):
    """并发给每个模型发一小批探针文本，记录耗时/首字延迟/是否真出中文。"""
    import httpx
    sem = asyncio.Semaphore(6)
    SPEED['results'] = []

    async with httpx.AsyncClient(trust_env=True, timeout=60) as client:
        async def one(mid):
            tr = T.CloudTranslator(endpoint=BUILTIN['endpoint'],
                                   publishable_key=BUILTIN['publishable_key'],
                                   model=mid, request_timeout=45)
            row = {'id': mid, 'ok': False, 'sec': None, 'ttft': None,
                   'note': ''}
            async with sem:
                try:
                    r = await tr.probe(client, timeout=45)      # 并发 3 批，测吞吐
                    row.update(r)
                    if not r['ok'] and not row.get('note'):
                        row['note'] = '探针未通过'
                except asyncio.TimeoutError:
                    row['note'] = '超时（>45s）'
                except Exception as e:
                    row['note'] = str(e)[:60]
            SPEED['results'].append(row)
            SPEED['done'] += 1
            return row

        rows = await asyncio.gather(*[one(m) for m in models])
    ok = [r for r in rows if r['ok'] and r['sec']]
    ok.sort(key=lambda r: r['sec'])
    SPEED['results'] = sorted(rows, key=lambda r: (not r['ok'],
                                                  r['sec'] if r['sec'] else 9e9))
    SPEED['ts'] = time.time()
    if ok:
        best = ok[0]['id']
        SPEED['note'] = '最快：%s（%.1fs）' % (best, ok[0]['sec'])
        log('模型测速完成，最快 %s（%.1fs）' % (best, ok[0]['sec']))
    else:
        SPEED['note'] = '全部模型均未通过探针'
    _save_speed()
    return SPEED['results']


def _start_bench(force=False):
    if SPEED['running']:
        return False
    SPEED['results'] = []
    SPEED['done'] = 0
    SPEED['total'] = 0
    SPEED['error'] = ''
    SPEED['running'] = True

    def work():
        try:
            if not BUILTIN_ENABLED:
                SPEED['error'] = ('内置通道未配置密钥（WB_PUBLISHABLE_KEY），'
                                  '无法测速；可改用「自定义接口」填自己的 API Key。')
                return
            tr = T.CloudTranslator(endpoint=BUILTIN['endpoint'],
                                   publishable_key=BUILTIN['publishable_key'])
            ms = [m['id'] for m in tr.models()
                  if m.get('id') not in ('hunyuan-image-v3.0',)]
            SPEED['total'] = len(ms)
            asyncio.run(_bench_models(ms, lambda m: None))
        except Exception as e:
            SPEED['error'] = str(e)[:200]
        finally:
            SPEED['running'] = False

    threading.Thread(target=work, daemon=True).start()
    return True


@app.get('/api/speed')
def get_speed(refresh: int = 0):
    """模型测速状态与结果。refresh=1 强制重测（否则 15 分钟内复用缓存）。"""
    if not SPEED['running']:          # 测速中不要用磁盘上的旧结果覆盖内存进度
        _load_speed()
    stale = (time.time() - SPEED.get('ts', 0)) > SPEED_TTL
    if refresh or (stale and not SPEED['results'] and not SPEED['running']):
        _start_bench(force=True)
    best = next((r['id'] for r in SPEED['results'] if r['ok']), '')
    return {'running': SPEED['running'], 'done': SPEED['done'],
            'total': SPEED['total'], 'ts': SPEED['ts'], 'note': SPEED['note'],
            'error': SPEED['error'], 'best': best,
            'recommended': DEFAULT_CFG['model'],
            'results': SPEED['results']}


def _cleanup_jobs():
    """清理过期任务与僵尸上传。"""
    now = time.time()
    for jid in list(JOBS):
        j = JOBS[jid]
        if j.get('done') and now - j.get('ts', now) > JOB_TTL:
            shutil.rmtree(j['dir'], ignore_errors=True)
            JOBS.pop(jid, None)


@app.post('/api/upload')
async def upload(file: UploadFile = File(...), cfg_json: str = Form(None)):
    _cleanup_jobs()
    # 并发与存量控制
    running = sum(1 for j in JOBS.values() if not j.get('done'))
    if running >= MAX_CONCURRENT or len(JOBS) > 60:
        raise HTTPException(429, '当前任务较多，请稍后再试')
    fname = os.path.basename(file.filename or 'input.pdf')
    if not fname.lower().endswith('.pdf'):
        raise HTTPException(400, '只支持 PDF 文件')
    jid = uuid.uuid4().hex[:10]
    d = os.path.join(HOME, 'jobs', jid)
    os.makedirs(d, exist_ok=True)
    pdf = os.path.join(d, 'input.pdf')
    size = 0
    with open(pdf, 'wb') as f:
        while True:
            chunk = await file.read(1 << 20)
            if not chunk:
                break
            size += len(chunk)
            if size > MAX_UPLOAD:
                f.close()
                shutil.rmtree(d, ignore_errors=True)
                raise HTTPException(413, '文件超过 30 MB 上限')
            f.write(chunk)
    JOBS[jid] = {'dir': d, 'progress': 0, 'message': '排队中', 'logs': [],
                 'files': [], 'done': False, 'error': None, 'pdf': pdf,
                 'ts': time.time()}
    job_cfg = load_cfg()
    # 可选：随请求携带本次任务的设置（供 GitHub Pages 等外部静态前端使用）。
    # 只作用于这一个任务，不改写服务器上的全局配置，避免多访客互相覆盖。
    if cfg_json:
        try:
            sent = json.loads(cfg_json)
            allowed = set(DEFAULT_CFG) | {'formats'}
            job_cfg.update({k: v for k, v in sent.items()
                            if k in allowed and v is not None})
        except Exception as e:
            JOBS[jid]['logs'].append('忽略无效的任务设置：%s' % str(e)[:80])
    th = threading.Thread(target=_run_with_sem, args=(JOBS[jid], pdf, job_cfg), daemon=True)
    th.start()
    return {'job': jid}


def _run_with_sem(job, pdf, cfg):
    with _SEM:
        job['ts'] = time.time()
        run_pipeline(job, pdf, cfg)


@app.get('/api/env')
def get_env():
    """服务器能力探测：PDF 输出是否可用等。"""
    return {
        'browser': bool(B.find_browser()),
        'story': hasattr(__import__('pymupdf'), 'Story'),
        'max_pages': MAX_PAGES,
        'max_upload_mb': MAX_UPLOAD // (1024 * 1024),
    }


@app.get('/api/job/{jid}')
def get_job(jid: str):
    j = JOBS.get(jid)
    if not j:
        raise HTTPException(404, 'no job')
    return {'progress': j['progress'], 'message': j['message'], 'logs': j['logs'][-40:],
            'files': j['files'], 'done': j['done'], 'error': j['error'],
            'nlogs': len(j['logs'])}


@app.get('/api/file/{jid}/{name}')
def get_file(jid: str, name: str):
    j = JOBS.get(jid)
    if not j:
        raise HTTPException(404, 'no job')
    name = os.path.basename(name)          # 防路径穿越
    p = os.path.join(j['dir'], name)
    if not os.path.exists(p):
        raise HTTPException(404, 'no file')
    from urllib.parse import quote
    ct = 'text/html; charset=utf-8'
    if name.endswith('.docx'):
        ct = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    elif name.endswith('.pdf'):
        ct = 'application/pdf'
    return FileResponse(p, media_type=ct,
                        headers={'Content-Disposition':
                                 "attachment; filename*=UTF-8''%s" % quote(name)})


# ─────────────────────────── CLI ───────────────────────────
def cli(argv):
    import argparse
    ap = argparse.ArgumentParser(description='PDF → 原格式译本')
    ap.add_argument('pdf')
    ap.add_argument('-o', '--out', default='')
    ap.add_argument('--mode', default=None, choices=['faithful', 'editable'])
    ap.add_argument('--engine', default=None, choices=['builtin', 'custom', 'none'])
    ap.add_argument('--model', default=None, help='内置通道的模型 id')
    ap.add_argument('--tgt', default=None)
    ap.add_argument('--src', default=None)
    ap.add_argument('--formats', default=None, help='逗号分隔，如 docx,pdf,html')
    ap.add_argument('--open', action='store_true')
    a = ap.parse_args(argv)
    cfg = load_cfg()
    for k in ('mode', 'tgt', 'src', 'engine', 'model'):
        v = getattr(a, k)
        if v:
            cfg[k] = v
    if a.formats:
        cfg['formats'] = [x.strip() for x in a.formats.split(',')]
    jid = 'cli'
    d = os.path.join(HOME, 'jobs', jid)
    os.makedirs(d, exist_ok=True)
    j = {'dir': d, 'progress': 0, 'message': '', 'logs': [], 'files': [], 'done': False,
         'error': None}
    JOBS[jid] = j
    run_pipeline(j, a.pdf, cfg)
    for m in j['logs']:
        print(' ', m)
    if j['error']:
        print(j['error']); return 1
    for f in j['files']:
        print('→', os.path.join(d, f['name']))
    if a.out and j['files']:
        # 优先拷 PDF，其次 docx，最后取第一个（原先固定取 docx，
        # 未勾选 docx 时会 IndexError 崩掉）
        pick = ([f for f in j['files'] if f['name'].endswith('.pdf')]
                or [f for f in j['files'] if f['name'].endswith('.docx')]
                or j['files'])[0]
        src = os.path.join(d, pick['name'])
        shutil.copy(src, a.out)
        print('copied →', a.out)
    if a.open and j['files']:
        os.system('open "%s"' % os.path.join(d, j['files'][0]['name']))
    return 0


if __name__ == '__main__':
    if len(sys.argv) > 1 and not sys.argv[1].startswith('-') and os.path.exists(sys.argv[1]):
        sys.exit(cli(sys.argv[1:]))
    import uvicorn
    port = int(os.environ.get('PORT', '8765'))
    host = os.environ.get('HOST', '0.0.0.0' if os.environ.get('PORT') else '127.0.0.1')
    print('pdft → http://%s:%d' % (host, port))
    uvicorn.run(app, host=host, port=port, log_level='warning')
