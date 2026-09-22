# -*- coding: utf-8 -*-
"""PDF → 结构化中间表示（IR）。

思路（今天手工流程的自动化版本）
------------------------------
1. PyMuPDF 取 dict 级文本：每个 span 带 font / size / color / bbox。
2. 判定数学：出现最多的字体族视为正文字体；凡「数学字体族且非正文字体族」或含数学 Unicode 的
   span 标记为数学片段。
3. **公式不重新打字**（提效的关键）：
   - faithful 模式：按 bbox 高清裁切原页面当作公式图，与原文像素级一致、零出错；
   - editable 模式：由 glyph + 几何（相对基线的偏移 → 上/下标）重建 LaTeX，重建失败降级为图。
4. 还原色块：由 drawings 里的填充矩形反推 tcolorbox 的边框色 / 底色 / 是否为左侧色条。
5. 连续同类行按间距合并成段落；字号聚类出标题；跨页重复文本识别为页眉页脚。
"""
import re, warnings
from dataclasses import dataclass, field
from typing import List, Optional

import pymupdf

warnings.filterwarnings('ignore')

_MATH_FAMILY = re.compile(
    r'^(CMMI|CMSY|CMEX|MSAM|MSBM|CMU|CMR|CMTI|CMBX|CMSS|CMFinite|LMMath|'
    r'Cambria|STIX|STIXTwo|Asana|XITS|Euler|MnSymbol|MathJax|Symbol)', re.I)

_MATH_UNICODE = ((0x0370, 0x03FF), (0x1D400, 0x1D7FF), (0x2190, 0x21FF),
                 (0x2200, 0x22FF), (0x27C0, 0x27EF), (0x2A00, 0x2AFF),
                 (0x2100, 0x214F))

_MATH_CHARS = set('≤≥≠≈≡≃≅∼∝∞∂∇±×÷·∈∉⊂⊆⊃⊇∪∩∅∀∃∑∏∫⟨⟩∥⊥∠⩽⩾→←↦⇒⇐⇔↔↑↓')

# 项目符号（itemize 行首）。行首出现它们时必须另起一段，否则列表会被压成一行
BULLETS = '•●○▪▫■□◦‣∙⋅·'

# 公式连接词：出现在公式内部的罗马字词（Re、Im、et、où…）。
# 段落的普通文本若全部由这些词构成，则整段按一块高清图输出，保住多行公式的对齐。
_FORMULA_WORDS = set(('re im et où ou avec si pour donc arg cos sin tan ln log exp '
                      'mod min max sup inf sh ch th').split())

# 页码：1 / 12、- 3 -、3
_PAGE_NUM = re.compile(r'^\s*[-–—.]?\s*\d{1,4}\s*(?:[/|·]\s*\d{1,4}\s*)?[-–—.]?\s*$')

def _formula_only(text: str, words=None) -> bool:
    """去掉公式连接词（Re、Im、et…）后什么都不剩 ⇒ 这串文字其实是公式的一部分。

    LaTeX 排版里 Re/Im/et 用的是正文字体，不能仅凭字体把它们排除在公式外，
    否则含 Re(z) 的公式行会被误判为「文字行」而遭逐段裁切。
    words 可指定更窄的连接词表：并入行内公式区时只认 Re/Im/et 这类真正的
    记号词，'donc'『avec』这类普通词要留给翻译。
    """
    pool = _FORMULA_WORDS if words is None else words
    s = re.sub(r'[^a-zà-ÿ]', '', text.lower())
    # ⚠️ 必须固定替换顺序（最长词优先）：pool 若是 set，迭代顺序随
    # PYTHONHASHSEED 变化——'lim' 先被替换则并入数学区、先被 'im' 吃掉
    # 则剩 'l' 判为文字，同一份 PDF 每次跑结果都可能不同！
    words_sorted = sorted(pool, key=len, reverse=True)
    prev = None
    while prev != s:
        prev = s
        for w in words_sorted:
            s = s.replace(w, '')
    return not s

# 行内公式区允许吞并的连接词/函数名（保持公式完整、不拆行）；
# 排除 donc/avec/si/pour/ou/où 这类真正的句子词，它们要留给翻译
_ZONE_WORDS = set('re im et mod arg cos sin tan cot ln log exp sup inf min max '
                  'ch sh th '
                  # lim/arccos… 也是正体排版的公式记号词：不并入数学区的话，
                  # lim 的下标（θ→0）会被拆成孤立小图，「lim」留在文字里变成
                  # 「因此 θlim [图]」这种断裂排版。
                  'lim arccos arcsin arctan'.split())

# 「值得裁图」的数学运算符：一个孤立的区若不含这些、也不含字母数字，
# 就只是标点/分隔符（如页脚的 ·），应当当文字处理，否则会变成小块乱码
_MATH_OPERATORS = set('=+−-×÷±∓≤≥≠≈≡∼≃≅∈∉⊂⊆⊃⊇∪∩∅∀∃∑∏∫⟨⟩∥⊥∠'
                      '→←↦⇒⇐⇔↔↑↓√∞∂∇%^_')

# 句子标点：只有这些可以被"还给文字流"（随之进图的标点会在公式图边上变成
# 孤立小符号）；括号/竖线/斜杠等可能是公式结构，不能动
_SENTENCE_PUNCT = set(',.;:!?·、。，；：！？…\'"“”‘’')

def _line_kind(spans, body_root):
    """一行（视觉行）的分类 → (has_math, pure, bullet, math_x0, math_x1)。

    pure = 整行都是公式（文字几乎没有，或只剩 Re/Im/et 这类公式连接词）；
    bullet = 行首是项目符号/编号（这类行的公式必须留在行内，不进成片合并）。
    """
    has_math = False
    t = []
    mx0, mx1 = 1e9, 0.0
    for s in spans:
        fam = _family(s['font'])
        if any(_is_math_char(c) for c in s['text']) or \
           (_MATH_FAMILY.match(fam) and fam != body_root):
            has_math = True
            mx0 = min(mx0, s['bbox'][0])
            mx1 = max(mx1, s['bbox'][2])
        else:
            t.append(s['text'])
    lead = ''.join(s['text'] for s in spans).lstrip()
    bullet = lead[:1] in BULLETS or bool(_ENUM.match(lead))
    rest = ''.join(t).strip().lstrip(BULLETS).strip()
    pure = has_math and (len(rest) < 3 or _formula_only(rest))
    return has_math, pure, bullet, mx0, mx1

def _is_math_char(ch: str) -> bool:
    if ch in _MATH_CHARS:
        return True
    cp = ord(ch)
    return any(a <= cp <= b for a, b in _MATH_UNICODE)

def _family(fontname: str) -> str:
    f = fontname.split('+')[-1]
    m = re.match(r'^([A-Za-z]+?)(?=\d|-|$)', f)
    return m.group(1) if m else f

# ───────────────────────── LaTeX 轻量重建 ─────────────────────────
_GREEK = {
    'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma', 'δ': r'\delta', 'ε': r'\varepsilon',
    'ζ': r'\zeta', 'η': r'\eta', 'θ': r'\theta', 'ι': r'\iota', 'κ': r'\kappa',
    'λ': r'\lambda', 'μ': r'\mu', 'ν': r'\nu', 'ξ': r'\xi', 'π': r'\pi',
    'ρ': r'\rho', 'σ': r'\sigma', 'τ': r'\tau', 'υ': r'\upsilon', 'φ': r'\varphi',
    'χ': r'\chi', 'ψ': r'\psi', 'ω': r'\omega',
    'Γ': r'\Gamma', 'Δ': r'\Delta', 'Θ': r'\Theta', 'Λ': r'\Lambda', 'Ξ': r'\Xi',
    'Π': r'\Pi', 'Σ': r'\Sigma', 'Φ': r'\Phi', 'Ψ': r'\Psi', 'Ω': r'\Omega',
    'ϑ': r'\vartheta', 'ϕ': r'\phi', 'ϖ': r'\varpi', 'ϱ': r'\varrho',
}
_SYM = {
    '≤': r'\leqslant', '⩽': r'\leqslant', '≥': r'\geqslant', '⩾': r'\geqslant',
    '≠': r'\neq', '≈': r'\approx', '≡': r'\equiv', '∼': r'\sim', '≃': r'\simeq',
    '≅': r'\cong', '∝': r'\propto', '∞': r'\infty', '∂': r'\partial', '∇': r'\nabla',
    '±': r'\pm', '∓': r'\mp', '×': r'\times', '÷': r'\div', '·': r'\cdot',
    '∈': r'\in', '∉': r'\notin', '⊂': r'\subset', '⊆': r'\subseteq', '⊃': r'\supset',
    '⊇': r'\supseteq', '∪': r'\cup', '∩': r'\cap', '∅': r'\varnothing',
    '∀': r'\forall', '∃': r'\exists', '∑': r'\sum', '∏': r'\prod', '∫': r'\int',
    '⟨': r'\langle', '⟩': r'\rangle', '∥': r'\parallel', '⊥': r'\perp', '∠': r'\angle',
    '→': r'\to', '⟶': r'\longrightarrow', '↦': r'\mapsto', '⇒': r'\Rightarrow',
    '⇐': r'\Leftarrow', '⇔': r'\iff', '↔': r'\leftrightarrow', '↑': r'\uparrow',
    '↓': r'\downarrow', 'ℂ': r'\mathbb{C}', 'ℕ': r'\mathbb{N}', 'ℝ': r'\mathbb{R}',
    'ℤ': r'\mathbb{Z}', 'ℚ': r'\mathbb{Q}', 'ℙ': r'\mathbb{P}', '−': '-',
}
_PLAIN = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
             " ()[]{}+-*/=,.;:!?'\"<>|~@")
_UNSAFE = set('√∛∜')                    # 需要根号横线等版式信息 → 一律走图片

def _char_latex(ch: str) -> Optional[str]:
    if ch in _UNSAFE or ch in '̄':
        return None
    if ch in _GREEK:
        return _GREEK[ch]
    if ch in _SYM:
        return _SYM[ch]
    if ch in _PLAIN:
        return '\\' + ch if ch in '{}$&#_^%' else ch
    if ord(ch) < 128:
        return ch
    return None

def build_math_latex(metas: List[dict]) -> Optional[str]:
    """由同一行内相邻数学 span 重建 LaTeX；无法保证语义时返回 None。"""
    if not metas:
        return None
    main = max(m['size'] for m in metas)
    baseline = max(m['y1'] for m in metas if m['size'] == main)
    height = max(m['y1'] for m in metas) - min(m['y0'] for m in metas)
    if height > main * 2.0:
        return None

    def level(m):
        if abs(m['y1'] - baseline) <= main * 0.30 and m['size'] >= main * 0.97:
            return 'b'
        if m['y1'] < baseline - main * 0.15:
            return 'u'
        if m['y1'] > baseline + main * 0.15:
            return 'd'
        return 'b'

    parts, buf, cur = [], [], 'b'

    def flush():
        if not buf:
            return ''
        body = re.sub(r'\s+', ' ', ''.join(buf)).strip()
        if not body:
            return ''
        return body if cur == 'b' else ('^{%s}' if cur == 'u' else '_{%s}') % body

    for m in metas:
        piece = []
        for ch in m['text']:
            if ch.isspace():
                piece.append(' '); continue
            c = _char_latex(ch)
            if c is None:
                return None
            piece.append(c)
        lv = level(m)
        if lv != cur:
            parts.append(flush()); buf = []; cur = lv
        buf.extend(piece)
    parts.append(''.join(buf) if False else flush())
    res = ''.join(parts)
    return res or None

# ─────────────────────────── IR ───────────────────────────
@dataclass
class Run:
    kind: str = 'text'      # text | math | image
    text: str = ''
    latex: str = ''
    png: bytes = b''
    w: float = 0.0
    h: float = 0.0

@dataclass
class Block:
    t: str = 'para'         # para | box | pagebreak
    role: str = 'body'      # body | heading | list | display
    level: int = 0
    align: str = 'left'
    size: float = 0.0
    bold: bool = False
    italic: bool = False
    color: str = '000000'
    display: bool = False   # 是否独立居中陈列
    template: str = ''      # 含 {{i}} 公式占位符与 [[b]] 强调标记
    refs: dict = field(default_factory=dict)
    border: str = ''
    fill: str = ''
    style: str = 'frame'
    blocks: List['Block'] = field(default_factory=list)
    page: int = 0

@dataclass
class Doc:
    width: float = 595.0
    height: float = 842.0
    margin_l: float = 62.0
    margin_r: float = 62.0
    margin_t: float = 60.0
    margin_b: float = 45.0
    body_size: float = 11.0
    body_font: str = ''
    pages: int = 1
    has_pagenum: bool = False
    header: List[Block] = field(default_factory=list)
    footer: List[Block] = field(default_factory=list)
    blocks: List[Block] = field(default_factory=list)

# ─────────────────────────── 工具 ───────────────────────────
def _col(c) -> str:
    if isinstance(c, int):
        return '%02X%02X%02X' % ((c >> 16) & 255, (c >> 8) & 255, c & 255)
    if not c:
        return '000000'
    return '%02X%02X%02X' % tuple(int(round(v * 255)) for v in c)

_SPAN_CACHE = {}


def _page_spans(page):
    """该页所有字符 span 的 (bbox, id, 向上扩展量)，供裁图时做「邻行字符遮挡」。

    向上扩展量：组合用重音符（如 U+20D7 向量箭头、U+0304 上划线）的墨迹画在
    字母**上方**，但 span 的 bbox 通常只包住字母本身——不加这个扩展，邻行的
    重音墨迹会漏进裁图（用户看到的"公式上方的奇怪小符号"）。
    """
    key = id(page)
    hit = _SPAN_CACHE.get(key)
    if hit is None:
        out = []
        for bb, spans in _lines(page):
            for s in spans:
                up = 0.0
                if any(0x300 <= ord(c) <= 0x36F or 0x20D0 <= ord(c) <= 0x20F0
                       for c in s['text']):
                    up = 0.6 * s['size']
                out.append((tuple(s['bbox']), id(s), up))
        _SPAN_CACHE.clear()            # 只留当前页，避免累积
        _SPAN_CACHE[key] = out
        hit = out
    return hit


_DRAW_CACHE = {}


def _page_rules(page):
    """该页所有「细线段」（框线、分隔线、分数线、根号横杆）的 (bbox, 是否细线)。

    与 _page_spans 同理做缓存——_cropped 每裁一张图都要用，不能每次重新解析。
    """
    key = id(page)
    hit = _DRAW_CACHE.get(key)
    if hit is None:
        out = []
        try:
            for dr in page.get_drawings():
                for it in dr['items']:
                    if it[0] != 'l':
                        continue
                    p0, p1 = it[1], it[2]
                    x0, x1 = sorted((p0.x, p1.x))
                    y0, y1 = sorted((p0.y, p1.y))
                    # 水平/垂直线段的 bbox 有一个维度为 0，PyMuPDF 视其为「空矩形」，
                    # intersects() 会直接返回 False（这正是框线一直漏进裁图的原因）。
                    # 补 0.5pt 的“厚度”让相交判断可用。
                    if y1 - y0 < 0.6:
                        y0 -= 0.5
                        y1 += 0.5
                    if x1 - x0 < 0.6:
                        x0 -= 0.5
                        x1 += 0.5
                    out.append((x0, y0, x1, y1))
        except Exception:
            out = []
        _DRAW_CACHE.clear()
        _DRAW_CACHE[key] = out
        hit = out
    return hit


def _cropped(page, bbox, dpi, own_ids=None):
    """裁出 bbox 区域的高清 PNG。

    own_ids：本图包含的 span id 集合。给定时，落在裁图矩形内、但**不属于本图**的
    字符会被涂白——PDF 的行在纵向上常互相咬合（分数的上下标、cases 与 bullet 交错），
    纯矩形裁图必然把邻行的字符切进公式图里（用户看到公式上下的"乱码残片"）。
    """
    clip = pymupdf.Rect(bbox[0] - 1.2, bbox[1] - 1.2, bbox[2] + 1.2, bbox[3] + 1.2)
    pix = page.get_pixmap(clip=clip, dpi=dpi)
    if own_ids is None:
        return pix.tobytes('png')
    raw = pix.tobytes('png')
    try:
        import io
        from PIL import Image, ImageDraw
        img = Image.open(io.BytesIO(raw)).convert('RGB')
        k = dpi / 72.0
        W, H = img.size

        def _px(r, pad=0.0):
            x0 = (r.x0 - clip.x0) * k
            y0 = (r.y0 - clip.y0) * k
            x1 = (r.x1 - clip.x0) * k
            y1 = (r.y1 - clip.y0) * k
            x0, x1 = max(0.0, min(x0, x1) - pad), min(float(W), max(x0, x1) + pad)
            y0, y1 = max(0.0, min(y0, y1) - pad), min(float(H), max(y0, y1) + pad)
            return (x0, y0, x1, y1)

        # ── 遮挡：把不属于本图的字符涂白，且绝不涂到公式本身 ──
        # 外扩 0.8pt 很关键：字形墨迹常比声明 bbox 略大（斜体出挑、组合重音、
        # 向量箭头），不外扩就会在公式图上留下"半截笔画/奇怪小符号"。
        own_rect = {}
        for sbb, sid, up in _page_spans(page):
            if sid not in own_ids:
                continue
            r = pymupdf.Rect(sbb)
            if up:
                r = pymupdf.Rect(r.x0, r.y0 - up, r.x1, r.y1)
            own_rect[sid] = r
        core = None
        for r in own_rect.values():
            core = r if core is None else (core | r)
        if core is not None:
            core = pymupdf.Rect(core.x0 - 2.5, core.y0 - 2.5,
                                core.x1 + 2.5, core.y1 + 2.5)

        mask = Image.new('L', img.size, 0)
        md = ImageDraw.Draw(mask)
        foreign = 0
        for sbb, sid, up in _page_spans(page):
            if sid in own_ids:
                continue
            r = pymupdf.Rect(sbb)
            if up:
                r = pymupdf.Rect(r.x0, r.y0 - up, r.x1, r.y1)
            r = r & clip
            if r.is_empty or r.width < 0.2 or r.height < 0.2:
                continue
            x0, y0, x1, y1 = _px(r, pad=0.8)
            if x1 - x0 < 0.5 or y1 - y0 < 0.5:
                continue
            md.rectangle((x0, y0, x1, y1), fill=255)
            foreign += 1
        # 自己的字符区回填（外扩 1.2pt，压过上面外来的 0.8pt）
        for r in own_rect.values():
            rr = r & clip
            if rr.is_empty:
                continue
            md.rectangle(_px(rr, pad=1.2), fill=0)

        # 版式线条（定义框边框、分隔线）会被矩形裁图框进来，表现为公式上方/下方
        # 一条多余的横线或竖线。做法：对「线条带」做像素级清理——只涂白彩色像素
        # （框线是有色的）与浅色像素，黑色字迹照旧保留，即便框线穿过字也不会切字。
        rules = 0
        for (sx0, sy0, sx1, sy1) in _page_rules(page):
            seg = pymupdf.Rect(sx0, sy0, sx1, sy1)
            if min(seg.width, seg.height) > 4.0:      # 不是细线（填充块）→ 不动
                continue
            if not seg.intersects(clip):
                continue
            if core is not None and core.contains(seg):
                continue                              # 属于本图内容（分数线等）
            bx0, by0, bx1, by1 = _px(seg & clip, pad=1.4)
            if bx1 - bx0 < 1 or by1 - by0 < 1:
                continue
            box = (int(bx0), int(by0), int(bx1), int(by1))
            try:
                from PIL import ImageChops
                band = img.crop(box)
                r_, g_, b_ = band.convert('RGB').split()
                mx = ImageChops.lighter(ImageChops.lighter(r_, g_), b_)
                mn = ImageChops.darker(ImageChops.darker(r_, g_), b_)
                sat = ImageChops.subtract(mx, mn)              # 彩度
                lum = band.convert('L')
                chrome = Image.eval(sat, lambda v: 255 if v > 40 else 0)
                light = Image.eval(lum, lambda v: 255 if v > 170 else 0)
                cmask = ImageChops.lighter(chrome, light)
                band.paste(Image.new('RGB', band.size, (255, 255, 255)),
                           (0, 0), cmask)
                img.paste(band, box)
            except Exception:
                md.rectangle((bx0, by0, bx1, by1), fill=255)
            rules += 1
        if not foreign and not rules:
            return raw
        white = Image.new('RGB', img.size, (255, 255, 255))
        out = io.BytesIO()
        Image.composite(white, img, mask).save(out, 'PNG')
        return out.getvalue()
    except Exception as e:
        try:
            log_fn = getattr(page, '_pdftool_log', None)
            if log_fn:
                log_fn('遮挡失败（保留原图）：%s' % str(e)[:80])
        except Exception:
            pass
        return raw

def _norm_hf(txt):
    """页眉/页脚比对用的归一化文本：数字→#，空白→空。

    这样「… 1 / 7 …」「… 2 / 7 …」会被视作同一行，页脚才能被正确识别。
    """
    return re.sub(r'\s+', '', re.sub(r'\d+', '#', txt or ''))


def _split_cols(spans, gap=12.0):
    """把一行 span 按横向大间隙切成最多 3 组——页眉/页脚常是「左/中/右」三栏。

    不切开的话，三栏会被当成一整行：页脚会变成
    「BIFAST 2A · 2026/2027-S1」+「2 / 7」+「Seb Godillon」首尾相接的怪字符串。
    """
    if not spans:
        return [spans]
    ss = sorted(spans, key=lambda s: s['bbox'][0])
    groups, cur = [], [ss[0]]
    for s in ss[1:]:
        if s['bbox'][0] - cur[-1]['bbox'][2] > gap:
            groups.append(cur)
            cur = [s]
        else:
            cur.append(s)
    groups.append(cur)
    if len(groups) <= 3:
        return groups
    n = len(groups)
    out = [[], [], []]
    for i, g in enumerate(groups):
        out[min(2, i * 3 // n)].extend(g)
    return [g for g in out if g]


def _lines(page):
    out = []
    for b in page.get_text('dict')['blocks']:
        if b.get('type', 0) != 0:
            continue
        for l in b['lines']:
            if not l['spans']:
                continue
            out.append((tuple(l['bbox']), list(l['spans'])))
    out.sort(key=lambda t: (t[0][1], t[0][0]))
    # 同一视觉基线带合并：PDF 会把一个分数拆成「分子行 / 分母行 / 根号行」
    # 等多个伪行，不并回真实行就会被分别裁图（公式碎成一堆小片）。
    # 判据：候选行的纵向中心落在带内最高行的范围内（允许 15% 外扩）——
    # 分子/分母/上下标必然落在主行范围内，而下一真实行的中心在线距之外。
    bands = []          # [bbox_union, spans, tall_bbox]
    for bb, spans in out:
        c = (bb[1] + bb[3]) / 2
        h = bb[3] - bb[1]
        if bands:
            ub, usp, tb = bands[-1]
            th = tb[3] - tb[1]
            if tb[1] - 0.15 * th <= c <= tb[3] + 0.15 * th:
                nb = (min(ub[0], bb[0]), min(ub[1], bb[1]),
                      max(ub[2], bb[2]), max(ub[3], bb[3]))
                usp.extend(spans)
                usp.sort(key=lambda s: (s['bbox'][0], s['bbox'][1]))
                bands[-1] = [nb, usp, bb if h > th else tb]
                continue
        bands.append([bb, sorted(spans, key=lambda s: (s['bbox'][0], s['bbox'][1])), bb])
    # 后处理：带是按 y 顺序生长的，先遇到的矮行（根号、上标）会自立门户，
    # 后遇到的高行（主行）应把它吞并回来；相邻带互查，直到收敛。
    changed = True
    while changed:
        changed = False
        for k in range(len(bands) - 1):
            A, B = bands[k], bands[k + 1]
            ta, tb = A[2], B[2]
            ha, hb = ta[3] - ta[1], tb[3] - tb[1]
            ca = (A[0][1] + A[0][3]) / 2
            cb = (B[0][1] + B[0][3]) / 2
            if (ta[1] - 0.15 * ha <= cb <= ta[3] + 0.15 * ha) or \
               (tb[1] - 0.15 * hb <= ca <= tb[3] + 0.15 * hb):
                nb = (min(A[0][0], B[0][0]), min(A[0][1], B[0][1]),
                      max(A[0][2], B[0][2]), max(A[0][3], B[0][3]))
                A[1].extend(B[1])
                A[1].sort(key=lambda s: (s['bbox'][0], s['bbox'][1]))
                bands[k] = [nb, A[1], ta if ha >= hb else tb]
                del bands[k + 1]
                changed = True
                break
    merged = [(b[0], b[1]) for b in bands]
    merged.sort(key=lambda t: (round(t[0][1], 1), round(t[0][0], 1)))
    return merged

# 编号列表行首（1. / 2) / 3、）。这些行同样必须各自成段
_ENUM = re.compile(r'^\d{1,2}[.)、]\s')


def _split_bullets(line_group, doc, page, mode, dpi):
    """一个物理段落里若某行以项目符号/编号开头 ⇒ 该行起另建一段。

    行距判定（gap > 5.5pt）对列表常常失效：itemize 的行距往往比段间距还小，
    于是「• 甲」「• 乙」被合并成一行，译出来就是「…，• …」。这里先看行首符号，
    再交给 _build_para。编号项保留自身编号（role 不变），符号项交给渲染器
    换成真正的列表符号。
    """
    segs, cur = [], []
    for it in line_group:
        txt = ''.join(s['text'] for s in it[1]).lstrip()
        new_item = txt[:1] in BULLETS or bool(_ENUM.match(txt))
        if new_item and cur:
            segs.append(cur)
            cur = [it]
        else:
            cur.append(it)
    if cur:
        segs.append(cur)
    out = []
    for g in segs:
        b = _build_para(g, doc, page, mode, dpi)
        if b:
            if (''.join(s['text'] for s in g[0][1]).lstrip()[:1]) in BULLETS:
                b.role = 'list'
                b.display = False
                b.align = 'left'
            b._vy = (min(it[0][1] for it in g) + max(it[0][3] for it in g)) / 2
            out.append(b)
    return out


def _find_boxes(page, text_w):
    rects = []          # (rect, fill_hex, is_white)
    for dr in page.get_drawings():
        f = dr['fill']
        if not f:
            continue
        c = _col(f)
        white = c == 'FFFFFF' or (isinstance(f, tuple) and min(f) > 0.97)
        r = dr['rect']
        if r.width < text_w * 0.30 or r.height < 10:
            continue
        rects.append((r, c, white))
    boxes = []
    for r, c, white in rects:
        inner = None
        for r2, c2, w2 in rects:
            if r2 == r:
                continue
            if (r2.y0 >= r.y0 - 2 and r2.y1 <= r.y1 + 2 and
                    r2.x0 >= r.x0 - 1 and r2.x1 <= r.x1 + 1 and
                    r2.width * r2.height > 0.6 * r.width * r.height):
                if inner is None or r2.width * r2.height > inner[0].width * inner[0].height:
                    inner = (r2, c2)
        if inner is None:
            continue
        fill = inner[1]
        # 内嵌矩形左上角明显内缩 ⇒ tcolorbox 的左侧色条样式
        stripe = (inner[0].x0 - r.x0) > 2.0
        boxes.append((r, c, fill, 'stripe' if stripe else 'frame'))
    uniq = []
    for it in boxes:
        if all(abs(it[0].y0 - o[0].y0) > 3 or abs(it[0].y1 - o[0].y1) > 3 for o in uniq):
            uniq.append(it)
    return uniq

def _build_para(line_group, doc, page, mode, dpi):
    """line_group: [(bbox, spans)] → Block（失败返回 None）"""
    sizes, colors, bold_n, italic_n, tot = [], {}, 0, 0, 0
    for _, spans in line_group:
        for s in spans:
            n = len(s['text']) or 1
            sizes += [s['size']] * n
            colors[s['color']] = colors.get(s['color'], 0) + n
            if 'Bold' in s['font'] or 'black' in s['font'].lower():
                bold_n += n
            if 'Italic' in s['font'] or 'Oblique' in s['font']:
                italic_n += n
            tot += n
    if tot == 0:
        return None
    main_size = max(set(sizes), key=sizes.count)
    color = _col(max(colors, key=colors.get))
    x0 = min(g[0][0] for g in line_group); x1 = max(g[0][2] for g in line_group)
    y0 = min(g[0][1] for g in line_group); y1 = max(g[0][3] for g in line_group)

    body_font_root = _family(doc.body_font)
    runs, template, idx = [], [], 0

    def add_text(txt, bold=False, italic=False):
        nonlocal template
        if not txt:
            return
        if bold:
            template.append('[[b]]' + txt + '[[/b]]')
        elif italic:
            template.append('[[i]]' + txt + '[[/i]]')
        else:
            template.append(txt)
        runs.append(Run(kind='text', text=txt))

    def add_ref(run):
        nonlocal idx, template
        template.append('{{%d}}' % idx)
        runs.append(run)
        idx += 1

    # ── 跨行数学结构检测（cases/矩阵/多行行内公式）：
    #    大括号延伸件(CMEX)与上下标会被逐行裁成孤立小图，视觉上像一堆奇怪符号。
    #    判定命中 ⇒ 该段落所有数学内容合并裁成一整块图，文字照常保留翻译。──
    math_spans = []
    for _, ss in line_group:
        for sp in ss:
            fam = _family(sp['font'])
            if any(_is_math_char(c) for c in sp['text']) or \
               (_MATH_FAMILY.match(fam) and fam != body_font_root):
                math_spans.append(sp)
    merge_math = False
    mbox_all = None
    if len(line_group) >= 2 and math_spans:
        # ⚠️ 真·跨行结构（cases/矩阵）的每一行**只有公式**；若是「每个伪行
        # 都跟着正文文字」的列表项（- Si q1q2<0 les deux… / - Si q1q2>0 …），
        # 合并会把两个列表项的公式缝成一张假 cases 图、文字被拦腰截断。
        _word_re = re.compile(r'[A-Za-zÀ-ÿ]{2,}')
        lines_have_prose = 0
        for _, ss in line_group:
            txt = ''.join(sp['text'] for sp in ss
                          if not (any(_is_math_char(c) for c in sp['text']) or
                                  (_MATH_FAMILY.match(_family(sp['font'])) and
                                   _family(sp['font']) != body_font_root)))
            words = [w for w in _word_re.findall(txt) if w.lower() not in _ZONE_WORDS]
            if len(' '.join(words)) >= 6:
                lines_have_prose += 1
        my0 = min(s['bbox'][1] for s in math_spans)
        my1 = max(s['bbox'][3] for s in math_spans)
        h_line = max((g[0][3] - g[0][1]) for g in line_group) or 1
        if lines_have_prose < len(line_group) and my1 - my0 > 1.7 * h_line:
            bx0 = min(s['bbox'][0] for s in math_spans)
            bx1 = max(s['bbox'][2] for s in math_spans)
            area = max(1e-6, (bx1 - bx0) * (my1 - my0))
            cover = 0.0
            math_ids = set(id(s) for s in math_spans)
            for _, ss in line_group:
                for sp in ss:
                    if id(sp) in math_ids:
                        continue
                    ix = min(sp['bbox'][2], bx1) - max(sp['bbox'][0], bx0)
                    iy = min(sp['bbox'][3], my1) - max(sp['bbox'][1], my0)
                    if ix > 0 and iy > 0:
                        cover += ix * iy
            if cover / area < 0.12:      # 文字大面积压在数学区里时不合并
                merge_math = True
                mbox_all = (bx0, my0, bx1, my1)

    merged_done = False
    for li, (bb, spans) in enumerate(line_group):
        # ── 逐行把 span 划成「文字区」与「数学区」 ──
        # 数学区 = 连续的一段：数学 span，或虽在正文字体但属于公式的碎片
        # （括号、逗号、空格，以及 Re/Im/et 这类公式连接词）。整个区只裁一张
        # 图——一行至多一两个占位符。碎片太多（十个八个 {{n}}）时模型会把
        # 占位符放错位置，译文里就出现「，² ²，」这种乱序残片。
        zones = []           # [is_math, [span, ...]]
        for sp in spans:
            fam = _family(sp['font'])
            ism = bool(any(_is_math_char(c) for c in sp['text']) or
                       (_MATH_FAMILY.match(fam) and fam != body_font_root))
            if ism:
                joins = True
            else:
                st = sp['text'].strip()
                if st and all(c in BULLETS for c in st):
                    joins = False                     # 项目符号始终算文字
                else:
                    joins = (not re.search(r'[A-Za-zÀ-ÿ]', sp['text'])) \
                        or _formula_only(sp['text'], _ZONE_WORDS)
            if joins and zones and zones[-1][0]:
                zones[-1][1].append(sp)
            elif (not joins) and zones and not zones[-1][0]:
                zones[-1][1].append(sp)
            else:
                zones.append([joins, [sp]])
        # 区内若一个真数学 span 都没有（只是标点/空格）⇒ 并回文字区
        reduced = []
        for ism, zsp in zones:
            real = any(any(_is_math_char(c) for c in s['text']) or
                       (_MATH_FAMILY.match(_family(s['font'])) and
                        _family(s['font']) != body_font_root) for s in zsp)
            ism = ism and real
            if reduced and reduced[-1][0] == ism:
                reduced[-1][1].extend(zsp)
            else:
                reduced.append([ism, zsp])

        # 极小的孤立"数学区"= 版式残件（上划线、延伸段、孤零零的点/逗号/分隔符）。
        # 保留会变成公式图上面/下面的一小块乱码。装饰件直接丢，标点并回文字。
        final = []
        for ism, zsp in reduced:
            if ism:
                w = max(s['bbox'][2] for s in zsp) - min(s['bbox'][0] for s in zsp)
                txt = ''.join(s['text'] for s in zsp).strip()
                has_op = any(c in _MATH_OPERATORS for c in txt)
                if (not has_op) and len(txt) <= 2 and \
                        (w < 7.0 or not any(c.isalnum() for c in txt)):
                    if any((0xF8F0 <= ord(c) <= 0xF8FF) or c in '¯ˉ\u0304\u0305˙'
                           for c in txt):
                        continue                    # 装饰残件：丢弃
                    ism = False                     # 孤立标点/分隔符：当文字
            if final and final[-1][0] == ism:
                final[-1][1].extend(zsp)
            else:
                final.append([ism, zsp])
        reduced = final

        # 把数学区首尾的纯标点（逗号、句号、空格）还给文字流：它们属于句子，
        # 跟着公式进图会在公式图片边上变成"奇怪的小符号"（尤其是大图被提升为
        # 居中块之后，逗号就孤立地贴在图的四周）。
        def _punct_only(sp):
            """只认真正的句子标点（, . ; : · 等）——括号、竖线 | 、斜杠都可能是
            公式结构的一部分，绝不能移出图外（否则公式会缺一半竖线）。"""
            t = sp['text']
            if any(_is_math_char(c) for c in t):
                return False
            return all(c.isspace() or c in _SENTENCE_PUNCT for c in t)

        split_zones = []
        for ism, zsp in reduced:
            if not ism:
                split_zones.append([False, zsp])
                continue
            core, head, tail = list(zsp), [], []
            while core and _punct_only(core[0]):
                head.append(core.pop(0))
            while core and _punct_only(core[-1]):
                tail.insert(0, core.pop())
            if not core:
                split_zones.append([False, head + tail])
                continue
            if head:
                split_zones.append([False, head])
            split_zones.append([True, core])
            if tail:
                split_zones.append([False, tail])
        merged_zones = []
        for ism, zsp in split_zones:
            if merged_zones and merged_zones[-1][0] == ism:
                merged_zones[-1][1].extend(zsp)
            else:
                merged_zones.append([ism, zsp])
        reduced = merged_zones

        for ism, zsp in reduced:
            if not ism:
                buf, b_b, b_i = [], None, None
                prev_x1 = None
                for s in zsp:
                    sb = 'Bold' in s['font'] or 'black' in s['font'].lower()
                    si = 'Italic' in s['font'] or 'Oblique' in s['font']
                    # 同一基线带里两个 span 横向间隙明显（> 0.3 倍字高）时补一个
                    # 空格——band 合并把「节号 2」这类独立伪行并进主行后，span
                    # 之间没有空格字符，直接 join 会得到「2Fonctions」这种粘连。
                    if buf and prev_x1 is not None and s['text'] \
                            and not s['text'][0].isspace() \
                            and buf and buf[-1] and not buf[-1][-1].isspace():
                        size = s.get('size') or 10.0
                        if s['bbox'][0] - prev_x1 > 0.3 * max(6.0, size):
                            buf.append(' ')
                    if sb != b_b or si != b_i:
                        if buf:
                            add_text(''.join(buf), b_b or False, b_i or False)
                        buf, b_b, b_i = [], sb, si
                    buf.append(s['text'])
                    prev_x1 = s['bbox'][2]
                if buf:
                    add_text(''.join(buf), b_b or False, b_i or False)
                continue
            if merge_math:
                # 跨行数学结构：只插一次整块大图，其余数学片段全部并入
                if not merged_done:
                    merged_done = True
                    add_ref(Run(kind='image',
                                png=_cropped(page, mbox_all, dpi,
                                             own_ids={id(x) for x in math_spans}),
                                w=mbox_all[2] - mbox_all[0],
                                h=mbox_all[3] - mbox_all[1]))
                continue
            metas = [{'text': s['text'], 'size': s['size'], 'y0': s['bbox'][1],
                      'y1': s['bbox'][3], 'x0': s['bbox'][0], 'x1': s['bbox'][2],
                      'bold': 'Bold' in s['font'], 'italic': 'Italic' in s['font'],
                      'font': s['font']} for s in zsp]
            mbox = (min(m['x0'] for m in metas), min(m['y0'] for m in metas),
                    max(m['x1'] for m in metas), max(m['y1'] for m in metas))
            # 邻行残留清理：基线带合并偶尔把下一行的孤立字符（如 √t 的 √）并进来。
            # 若区内 span 较多，且有 span 孤零零吊在主行下方（下方只有它自己），
            # 判定为下一行的字符 → 不纳入本图（交给遮挡逻辑涂白）。
            own = {id(x) for x in zsp}
            if len(zsp) >= 5:
                ctr = sorted((s['bbox'][1] + s['bbox'][3]) / 2 for s in zsp)
                mc = ctr[len(ctr) // 2]
                below = [s for s in zsp if (s['bbox'][1] + s['bbox'][3]) / 2 > mc + 6.5]
                if len(below) == 1:
                    own.discard(id(below[0]))
                    keep = [s for s in zsp if id(s) in own]
                    if keep:
                        mbox = (min(s['bbox'][0] for s in keep),
                                min(s['bbox'][1] for s in keep),
                                max(s['bbox'][2] for s in keep),
                                max(s['bbox'][3] for s in keep))
            latex = build_math_latex(metas) if mode == 'editable' else None
            if latex:
                add_ref(Run(kind='math', latex=latex))
            else:
                add_ref(Run(kind='image',
                            png=_cropped(page, mbox, dpi, own_ids=own),
                            w=mbox[2] - mbox[0], h=mbox[3] - mbox[1]))
        if li < len(line_group) - 1:
            add_text(' ')

    if not runs:
        return None

    tmpl = ''.join(template)
    plain = re.sub(r'\{\{\d+\}\}', '', tmpl).replace('[[b]]', '').replace('[[/b]]', '') \
        .replace('[[i]]', '').replace('[[/i]]', '').strip()

    # 整段成图时，行首的项目符号不要画进图里——渲染层会用原生列表符号，
    # 否则会看到两个符号（图里一个、列表一个）
    def _crop_rect():
        _x0 = x0
        left = None
        for _, ss in line_group:
            for sp in ss:
                if left is None or sp['bbox'][0] < left['bbox'][0]:
                    left = sp
        if left is not None:
            st = left['text'].strip()
            if st and all(c in BULLETS for c in st):
                _x0 = min(x1 - 1.0, left['bbox'][2] + 1.0)
        return (_x0, y0, x1, y1)

    # 纯公式且非可编辑模式：整段按一块高分辨率图输出（多行的对齐关系不会丢）
    only_math = not plain
    b = Block(t='para', size=main_size, color=color,
              bold=bold_n > tot * 0.6, italic=italic_n > tot * 0.6)
    if only_math and mode != 'editable':
        box = _crop_rect()
        b = Block(t='para', role='display', size=main_size, color=color, display=True)
        b.align = 'center'
        _all_ids = {id(sp) for _, ss in line_group for sp in ss}
        b.refs = {0: Run(kind='image', png=_cropped(page, box, dpi, own_ids=_all_ids),
                         w=box[2] - box[0], h=box[3] - box[1])}
        b.template = '{{0}}'
        return b

    b.refs = {}
    for r in runs:                       # add_ref 已按出现顺序用同一计数器编号
        if r.kind != 'text':
            b.refs[len(b.refs)] = r
    b.template = tmpl

    # 公式主导：普通文本只剩公式连接词（Re、Im、et…）⇒ 整段裁成一块高清图，
    # 保住多行公式（cases/aligned）的对齐结构
    if b.refs and len(plain) < 220:
        s = re.sub(r'[^a-zà-ÿ]', '', plain.lower())
        prev = None
        while prev != s:
            prev = s
            for w in _FORMULA_WORDS:
                s = s.replace(w, '')
        if not s:
            box = _crop_rect()
            b.role = 'display'; b.display = True; b.align = 'center'
            _all_ids = {id(sp) for _, ss in line_group for sp in ss}
            b.refs = {0: Run(kind='image', png=_cropped(page, box, dpi, own_ids=_all_ids),
                             w=box[2] - box[0], h=box[3] - box[1])}
            b.template = '{{0}}'
            return b

    if not plain:
        b.role = 'display'
        b.display = True
        b.align = 'center'
    if plain.lstrip()[:1] in '•·▪‣◦–—' or plain.lstrip()[:2] in ('- ', '* '):
        b.role = 'list'
    elif b.size >= doc.body_size + 2.0 and b.bold and len(plain) < 120:
        b.role = 'heading'
        b.level = 1 if b.size >= doc.body_size + 7 else 2
    tw = doc.width - doc.margin_l - doc.margin_r
    if x0 - doc.margin_l > 14 and (tw - (x1 - doc.margin_l)) > 14 and (x1 - x0) < 0.8 * tw:
        b.align = 'center'
    return b

# ─────────────────────────── 主流程 ───────────────────────────
def extract(path: str, mode: str = 'faithful', dpi: int = 300, progress=None) -> Doc:
    pdf = pymupdf.open(path)
    n = len(pdf)
    fam_cnt, size_cnt = {}, {}
    minx, maxx, miny, maxy = 1e9, 0, 1e9, 0
    for p in pdf:
        for b in p.get_text('dict')['blocks']:
            if b.get('type', 0) != 0:
                continue
            for l in b['lines']:
                for s in l['spans']:
                    t = s['text'].strip()
                    if not t:
                        continue
                    fam_cnt[_family(s['font'])] = fam_cnt.get(_family(s['font']), 0) + len(t)
                    k = round(s['size'], 1)
                    size_cnt[k] = size_cnt.get(k, 0) + len(t)
                    minx = min(minx, l['bbox'][0]); maxx = max(maxx, l['bbox'][2])
                    miny = min(miny, l['bbox'][1]); maxy = max(maxy, l['bbox'][3])
    doc = Doc(width=float(pdf[0].rect.width), height=float(pdf[0].rect.height),
              margin_l=float(minx), margin_r=float(pdf[0].rect.width - maxx),
              margin_t=float(miny), margin_b=float(pdf[0].rect.height - maxy),
              body_size=max(size_cnt, key=size_cnt.get) if size_cnt else 11.0,
              body_font=max(fam_cnt, key=fam_cnt.get) if fam_cnt else '', pages=n)

    # 页眉页脚：跨页重复且在页边缘
    sig = {}
    for i, p in enumerate(pdf):
        for bb, spans in _lines(p):
            txt = ''.join(s['text'] for s in spans).strip()
            if not txt or len(txt) > 140:
                continue
            if bb[1] < doc.height * 0.14 or bb[1] > doc.height * 0.86:
                # 页码逐页变化，整行文本不重复——把数字归一成 # 再比对，
                # 否则「BIFAST 2A · 2026/2027-S1  2 / 7  Seb Godillon」这类页脚
                # 会被当成正文，三栏首尾相接成怪字符串（用户报的"字符丢失"）。
                sig.setdefault(_norm_hf(txt), []).append(i)
    rep = {k for k, v in sig.items() if len(v) >= max(2, int(n * 0.6))}
    doc.has_pagenum = False

    for pi, page in enumerate(pdf):
        if progress:
            progress(pi / max(n, 1), '正在解析第 %d/%d 页' % (pi + 1, n))
        lines = _lines(page)
        # 登记本页 span（对象级），供 _cropped 的邻行遮挡使用：
        # 必须与主流程共用同一批 span 对象，否则 id 对不上会把公式本身涂白
        _SPAN_CACHE.clear()
        _SPAN_CACHE[id(page)] = [
            (tuple(s['bbox']), id(s),
             0.6 * s['size'] if any(0x300 <= ord(c) <= 0x36F or
                                    0x20D0 <= ord(c) <= 0x20F0 for c in s['text']) else 0.0)
            for _, ss in lines for s in ss]
        _DRAW_CACHE.clear()          # 版式线条缓存也只留当前页
        boxes = _find_boxes(page, maxx - minx)

        def container(bb):
            # 必须返回 boxes 列表里的同一个元组对象（key=id 才稳定）
            best = None
            for item in boxes:
                r = item[0]
                if r.y0 - 6 <= bb[1] and bb[3] <= r.y1 + 6 and r.x0 - 8 <= bb[0] and bb[2] <= r.x1 + 8:
                    if best is None or r.width * r.height < best[1]:
                        best = (item, r.width * r.height)
            return best[0] if best else None

        # ── 行级「纯数学行」成片合并 ──
        # cases/矩阵/分数等多行公式：每一行都是纯数学行。逐行裁图会把大括号
        # 延伸件、分数线上下两侧切成一堆孤立碎片（看起来像奇怪的小符号）。
        # 这里把纵向相连、横向交叠的连续纯数学行合并裁成一整块 display 图。
        # 安全边界（宁可不合并，绝不错切）：
        #   · 带项目符号/编号的行不参与——每个列表项的公式留在该项行内；
        #   · 含文字的行不参与——文字必须留下翻译，只整体搬走纯数学行。
        body_root = _family(doc.body_font)
        kinds = [_line_kind(spans, body_root) for _, spans in lines]

        # 续行并入（先于成片挖取）：cases/分数的组成部分常落在所属行（带符号
        # 或文字的混排行）的上一带或下一带，单独裁出就像公式被腰斩/腰斩。
        # 判据：纵向间隙≤12pt（含部分重叠），且横向重叠超过纯行自身宽度
        # 的一半（防止把下一行的独立公式误并进只含一小段行内公式的文字行）。
        def _attachable(t):
            # 目标行：带项目符号/编号的行（cases 行、分数子行依附于 bullet 项）。
            # 普通文字行不吸收独立公式——否则居中大公式会按 x 序插进文字中间
            # （定义 3.1 的 |z|=√(a²+b²) 就是这么被搅乱的）。
            return kinds[t][0] and kinds[t][2]

        def _merge_into(src, dst):
            sb, ssp = lines[src]
            db, dsp = lines[dst]
            nb = (min(db[0], sb[0]), min(db[1], sb[1]),
                  max(db[2], sb[2]), max(db[3], sb[3]))
            merged = dsp + ssp
            merged.sort(key=lambda s: (s['bbox'][0], s['bbox'][1]))
            lines[dst] = (nb, merged)
            kinds[dst] = _line_kind(merged, body_root)
            del lines[src]
            del kinds[src]

        i = 0
        while i < len(lines):
            bb, _ = lines[i]
            if kinds[i][1] and not kinds[i][2]:      # 纯数学行且不带符号
                target = None
                # 优先并入上一行：cases 的行属于它上面的条目；只有上一行不可
                # 并入时才向下找（有的框把符号行垂直居中在 cases 两行之间）。
                for t in (i - 1, i + 1):
                    if not (0 <= t < len(lines)) or not kinds[t][0]:
                        continue
                    tb = lines[t][0]
                    yov = min(bb[3], tb[3]) - max(bb[1], tb[1])
                    ygap = max(bb[1] - tb[3], tb[1] - bb[3], -5.0)
                    xov = min(bb[2], kinds[t][4]) - max(bb[0], kinds[t][3])
                    # 纯行必须基本「被包含」在目标行的数学区里（cases 行、分数
                    # 子行）；超出太多的说明是独立居中大公式，不能并入（定义 3.1）
                    if xov <= 0.8 * (bb[2] - bb[0]):
                        continue
                    if kinds[t][2]:
                        if ygap > 12:                # 符号行：允许小间隙
                            continue
                    elif yov > 1.0:                  # 普通行：必须纵向重叠
                        pass                         # （分数分子/分母就叠在主行上）
                    else:
                        continue
                    target = t
                    break
                if target is not None:
                    _merge_into(i, target)
                    i = max(0, i - 1)              # 合并后重查相邻行（链式并入）
                    continue
            i += 1

        cand = [i for i in range(len(lines))
                if kinds[i][1] and not kinds[i][2]]
        dug_boxes, zone_lines = [], set()
        if len(cand) >= 2:
            parent = {i: i for i in cand}

            def find(i):
                while parent[i] != i:
                    parent[i] = parent[parent[i]]
                    i = parent[i]
                return i

            for ai in range(len(cand)):
                for bi in range(ai + 1, len(cand)):
                    i, j = cand[ai], cand[bi]
                    if find(i) == find(j):
                        continue
                    a, b = lines[i][0], lines[j][0]
                    yov = min(a[3], b[3]) - max(a[1], b[1])
                    xov = min(a[2], b[2]) - max(a[0], b[0])
                    ygap = max(a[1], b[1]) - min(a[3], b[3])
                    xgap = max(a[0], b[0]) - min(a[2], b[2])
                    minw = min(a[2] - a[0], b[2] - b[0])
                    if (yov > 0 and xgap < 18) or \
                       (ygap < 12 and xov > 0.25 * minw) or \
                       (ygap < 6 and xgap < 18):
                        parent[find(i)] = find(j)
            comps = {}
            for i in cand:
                comps.setdefault(find(i), []).append(i)
            for idxs in comps.values():
                if len(idxs) < 2:
                    continue
                _ids = {id(sp) for i in idxs for sp in lines[i][1]}
                dug_boxes.append((min(lines[i][0][0] for i in idxs),
                                  min(lines[i][0][1] for i in idxs),
                                  max(lines[i][0][2] for i in idxs),
                                  max(lines[i][0][3] for i in idxs), _ids))
                zone_lines.update(idxs)

        # ── 嵌入位图（照片、示意图）：此前完全被忽略，译本会整图丢失。
        # 每张图当成一个「挖块」走同一管线（容器归属、排序、整块裁剪渲染），
        # 图内的标注文字（中心落在图内的行）跟随图片，不再进正文。 ──
        try:
            seen_xref = set()
            page_area = doc.width * doc.height
            for _img in page.get_images(full=True):
                xref = _img[0]
                if xref in seen_xref:
                    continue
                seen_xref.add(xref)
                for r in page.get_image_rects(xref):
                    w, h = r.width, r.height
                    if w < 40 or h < 40 or w * h < 3000:
                        continue                       # 图标、装饰点
                    if w * h > 0.85 * page_area:
                        continue                       # 整页背景（幻灯片常见）
                    bbox = (r.x0, r.y0, r.x1, r.y1)
                    _ids, drop = set(), set()
                    for k, (lbb, lsp) in enumerate(lines):
                        cx = (lbb[0] + lbb[2]) / 2
                        cy = (lbb[1] + lbb[3]) / 2
                        if r.x0 - 2 <= cx <= r.x1 + 2 and r.y0 - 2 <= cy <= r.y1 + 2:
                            _ids.update(id(sp) for sp in lsp)
                            drop.add(k)                # 图内标注文字跟图走
                    dug_boxes.append((r.x0, r.y0, r.x1, r.y1, _ids))
                    zone_lines.update(drop)
        except Exception:
            pass

        if zone_lines:
            lines = [ln for k, ln in enumerate(lines) if k not in zone_lines]
            kinds = [k for q, k in enumerate(kinds) if q not in zone_lines]

        # 相邻挖块合并：多行 aligned 公式偶尔被拆成两个纵向紧邻的分量，
        # 合并后才能保住整体对齐（右边界也不会被截断）
        changed = True
        while changed and len(dug_boxes) > 1:
            changed = False
            for i in range(len(dug_boxes)):
                for j in range(i + 1, len(dug_boxes)):
                    a, b = dug_boxes[i], dug_boxes[j]
                    yov = min(a[3], b[3]) - max(a[1], b[1])
                    ygap = max(a[1], b[1]) - min(a[3], b[3])
                    xov = min(a[2], b[2]) - max(a[0], b[0])
                    xgap = max(a[0], b[0]) - min(a[2], b[2])
                    if (yov > 0 and xgap < 12) or (ygap < 8 and xov > 0.4 * min(a[2] - a[0], b[2] - b[0])):
                        dug_boxes[i] = (min(a[0], b[0]), min(a[1], b[1]),
                                        max(a[2], b[2]), max(a[3], b[3]),
                                        a[4] | b[4])
                        dug_boxes.pop(j)
                        changed = True
                        break
                if changed:
                    break

        # 挖块 → 独立 display 块，按容器归属待插入
        dug_by_key = {}
        for (bx0, by0, bx1, by1, _own) in dug_boxes:
            blk = Block(role='display', display=True, align='center', page=pi,
                        template='{{0}}',
                        refs={0: Run(kind='image',
                                     png=_cropped(page, (bx0, by0, bx1, by1), dpi,
                                                  own_ids=_own),
                                     w=bx1 - bx0, h=by1 - by0)})
            blk._vy = (by0 + by1) / 2
            blk._y0, blk._y1 = by0, by1
            cont2 = container((bx0, by0, bx1, by1))
            dug_by_key.setdefault(id(cont2) if cont2 else None, []).append(blk)

        groups, order = {}, []
        for li, (bb, spans) in enumerate(lines):
            txt = ''.join(s['text'] for s in spans).strip()
            if not txt:
                continue
            if _norm_hf(txt) in rep and (bb[1] < doc.height * 0.14 or bb[1] > doc.height * 0.86):
                # 页眉/页脚常是「左 / 中 / 右」三栏：按横向大间隙切开，各成一栏，
                # 否则三栏会被拼成一串（页脚出现 “S12 / 7Seb” 这种怪字符）
                target = doc.header if bb[1] < doc.height * 0.5 else doc.footer
                for grp in _split_cols(spans):
                    # 页码那一栏不留文本（否则每页的「1 / 7」「2 / 7」都会堆进来），
                    # 只记一个标记，由渲染层生成 Word 域 / 自动页码
                    gtxt = ''.join(s['text'] for s in grp).strip()
                    if _PAGE_NUM.match(gtxt):
                        doc.has_pagenum = True
                        continue
                    gb = (min(s['bbox'][0] for s in grp), bb[1],
                          max(s['bbox'][2] for s in grp), bb[3])
                    blk = _build_para([(gb, grp)], doc, page, mode, dpi)
                    if blk:
                        target.append(blk)
                continue
            # 页边缘的纯页码：不进正文，改由 Word 的 PAGE/NUMPAGES 域生成
            if _PAGE_NUM.match(txt) and (bb[3] > doc.height * 0.90 or bb[1] < doc.height * 0.10):
                doc.has_pagenum = True
                continue
            cont = container(bb)
            key = id(cont) if cont else None
            if key not in groups:
                groups[key] = []
                order.append((cont, key, bb[1]))
            groups[key].append((bb, spans, kinds[li][1]))

        order.sort(key=lambda t: t[2])
        page_blocks = []
        for cont, key, first_y in order:
            items = groups[key]
            paras = []
            grp, prev_bottom, grp_pure = [], None, None
            for bb, spans, isp in items:
                # 纯数学行与文字行不共段：否则段内跨行公式会触发整段合并裁图，
                # 把文字一起圈进大图、还丢掉翻译（例 2.3 的教训）
                boundary = False
                if prev_bottom is not None and (bb[1] - prev_bottom) > 5.5:
                    # 同一行左右并排时 gap 为负 → 保持同一段（章节号 + 标题）
                    boundary = True
                if grp_pure is not None and isp != grp_pure:
                    boundary = True
                if boundary and grp:
                    paras.append(grp)
                    grp = []
                    grp_pure = None
                if grp_pure is None:
                    grp_pure = isp
                grp.append((bb, spans))
                prev_bottom = max(prev_bottom or bb[3], bb[3])
            if grp:
                paras.append(grp)
            built = []
            for lg in paras:
                built.extend(_split_bullets(lg, doc, page, mode, dpi))
            dugs = dug_by_key.pop(key, [])
            if dugs:
                built = sorted(built + dugs, key=lambda b: getattr(b, '_vy', first_y))
                # 标签居中贴在 cases 旁时，挖块的 _vy 可能略小于标签行而被排到
                # 前面（图上标签下）。若列表/文字段的中心落在挖块纵向范围内，
                # 它应当排在挖块之前。
                for di in range(1, len(built)):
                    a, b = built[di - 1], built[di]
                    if a.role == 'display' and b.role != 'display' and \
                            hasattr(a, '_y0') and \
                            a._y0 - 2 <= getattr(b, '_vy', 0) <= a._y1 + 2:
                        built[di - 1], built[di] = b, a
            if not built:
                continue
            if cont is not None:
                r, c, f, s = cont
                bx = Block(t='box', border=c, fill=f, style=s, blocks=built, page=pi)
                bx._vy = first_y
                page_blocks.append(bx)
            else:
                page_blocks.extend(built)
        for blks in dug_by_key.values():        # 未被任何容器认领的挖块
            page_blocks.extend(blks)
        page_blocks.sort(key=lambda b: getattr(b, '_vy', 0))
        doc.blocks.extend(page_blocks)
        if pi < n - 1:
            doc.blocks.append(Block(t='pagebreak', page=pi))
    # 页眉页脚按内容去重，只保留一次
    def uniq(bs):
        seen, out = set(), []
        for b in bs:
            k = b.template.strip()
            if k and k not in seen:
                seen.add(k); out.append(b)
        return out
    doc.header = uniq(doc.header)
    doc.footer = uniq(doc.footer)
    pdf.close()
    return doc
