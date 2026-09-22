"""把视觉模型输出的 Markdown（含 LaTeX）渲染成 HTML / DOCX / PDF。

走视觉路径时得到的是「正文 + LaTeX 公式」的 Markdown，跟原版式版？
 preserves 的版式还原不一样：这里按文档结构（标题/正文/列表/图注）重排，
 公式交给 MathJax 渲染——牺牲了原页面特有的方框、色块，但保住了公式，
 且扫描件本来也没有可提取的矢量版式。
"""
import html as H
import os
import re


def _inline(s: str) -> str:
    """Markdown 行内标记 → HTML（保持 LaTeX 原样，交给 MathJax）。"""
    s = H.escape(s, quote=False)
    s = s.replace(r'\(', '\u0001').replace(r'\)', '\u0002')
    s = s.replace(r'\[', '\u0003').replace(r'\]', '\u0004')
    s = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', s)
    s = re.sub(r'`(.+?)`', r'<code>\1</code>', s)
    s = (s.replace('\u0001', r'\(').replace('\u0002', r'\)')
          .replace('\u0003', r'\[').replace('\u0004', r'\]'))
    # $...$ → \(...\)；$$...$$ → \[...\]
    s = re.sub(r'\$\$(.+?)\$\$', r'\\[\1\\]', s, flags=re.S)
    s = re.sub(r'(?<!\\)\$([^$\n]+?)\$', r'\\(\1\\)', s)
    return s


def md_to_html_body(md: str) -> str:
    """整段 Markdown（可能多页拼接）→ HTML 片段。"""
    out, buf_line, in_list = [], [], False
    lines = (md or '').split('\n')

    def flush_para():
        if not buf_line:
            return
        txt = ' '.join(x.strip() for x in buf_line if x.strip())
        buf_line.clear()
        if txt:
            out.append('<p>%s</p>' % _inline(txt))

    for ln in lines:
        s = ln.rstrip()
        if not s.strip():
            flush_para()
            if in_list:
                out.append('</ul>')
                in_list = False
            continue
        m = re.match(r'^(#{1,4})\s+(.*)$', s)
        if m:
            flush_para()
            if in_list:
                out.append('</ul>')
                in_list = False
            lv = len(m.group(1))
            out.append('<h%d>%s</h%d>' % (lv, _inline(m.group(2).strip()), lv))
            continue
        if re.match(r'^\s*[-*]\s+', s):
            flush_para()
            if not in_list:
                out.append('<ul>')
                in_list = True
            item = re.sub(r'^\s*[-*]\s+', '', s)
            out.append('<li>%s</li>' % _inline(item))
            continue
        if s.lstrip().startswith('>'):            # > 图：说明
            flush_para()
            if in_list:
                out.append('</ul>')
                in_list = False
            out.append('<p class="figcap">%s</p>'
                       % _inline(s.lstrip()[1:].strip()))
            continue
        if s.startswith('$$') or s.startswith('\\['):
            flush_para()
            body = s.strip().strip('$').replace('\\[', '').replace('\\]', '')
            out.append('<div class="disp">\\[%s\\]</div>' % body)
            continue
        buf_line.append(s)
    flush_para()
    if in_list:
        out.append('</ul>')
    return '\n'.join(out)


_PAGE_CSS = """
body{font-family:-apple-system,"PingFang SC","Microsoft YaHei",serif;
 font-size:15px;line-height:1.85;color:#1a1a1a;max-width:760px;margin:0 auto;
 padding:36px 30px}
@media print{body{padding:0;font-size:11pt}
 @page{size:A4;margin:2.0cm 2.1cm 2.0cm}}
h1{font-size:1.55em;margin:1.4em 0 .7em;line-height:1.4}
h2{font-size:1.28em;margin:1.25em 0 .6em}
h3{font-size:1.1em;margin:1.1em 0 .5em}
h4{font-size:1em;margin:1em 0 .45em}
p{margin:.62em 0;text-align:justify}
ul{margin:.5em 0 .8em 1.3em;padding:0}
li{margin:.28em 0}
code{background:#f4f4f4;padding:1px 4px;border-radius:3px;font-size:.92em}
.figcap{color:#555;font-size:.92em;text-align:center;margin:.5em 0 1em;
 font-style:italic}
.disp{margin:.9em 0;text-align:center;overflow-x:auto}
.pb{page-break-after:always;height:0}
.note{color:#666;font-size:.85em;border-left:3px solid #ddd;padding-left:10px;
 margin:1.2em 0}
"""

_MATHJAX = ('<script>window.MathJax={tex:{inlineMath:[["\\\\(","\\\\)"]],'
            'displayMath:[["\\\\[","\\\\]"]]},options:{skipHtmlTags:'
            '["script","noscript","style","textarea","pre","code"]}};</script>\n'
            '<script id="MathJax-script" async '
            'src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js">'
            '</script>')


def render_html_pages(md_pages, out_path: str, title='译文', note=None):
    """md_pages: [(name, markdown)] → 单文件 HTML。"""
    body = []
    if note:
        body.append('<p class="note">%s</p>' % H.escape(note))
    for i, (_, md) in enumerate(md_pages):
        if md:
            body.append(md_to_html_body(md))
        if i < len(md_pages) - 1:
            body.append('<div class="pb"></div>')
    html = ('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1">'
            '<title>%s</title><style>%s</style>%s</head><body>%s</body></html>'
            % (H.escape(title), _PAGE_CSS, _MATHJAX, '\n'.join(body)))
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return out_path


def render_docx_pages(md_pages, out_path: str):
    """Markdown → DOCX（公式以 LaTeX 源码落为斜体文本，Word 内可读可改）。"""
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(11)
    SMALL = re.compile(r'[\u4e00-\u9fff]')

    def setfont(run, size=11, italic=False, bold=False):
        run.font.size = Pt(size)
        run.italic = italic
        run.bold = bold
        # 中西文混排：中文用宋体，西文保持 Times
        if SMALL.search(run.text or ''):
            try:
                run.font.name = '宋体'
                run._element.rPr.rFonts.set(
                    __import__('docx.oxml.ns', fromlist=['qn']).qn('w:eastAsia'),
                    '宋体')
            except Exception:
                pass

    for _, md in md_pages:
        if not md:
            continue
        lines = md.split('\n')
        i = 0
        while i < len(lines):
            s = lines[i].rstrip()
            i += 1
            if not s.strip():
                continue
            m = re.match(r'^(#{1,4})\s+(.*)$', s)
            if m:
                lv = min(len(m.group(1)), 4)
                p = doc.add_heading(level=lv)
                r = p.add_run(strip_md(m.group(2)))
                setfont(r, size=18 - 2 * lv, bold=True)
                continue
            if re.match(r'^\s*[-*]\s+', s):
                items = []
                while i <= len(lines) and re.match(r'^\s*[-*]\s+', lines[i - 1] if i <= len(lines) else ''):
                    cur = lines[i - 1]
                    if not re.match(r'^\s*[-*]\s+', cur):
                        break
                    items.append(re.sub(r'^\s*[-*]\s+', '', cur).strip())
                    i += 1
                for it in items:
                    p = doc.add_paragraph(style='List Bullet')
                    r = p.add_run(strip_md(it))
                    setfont(r)
                continue
            if s.lstrip().startswith('>'):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(strip_md(s.lstrip()[1:].strip()))
                setfont(r, size=9.5, italic=True)
                continue
            if s.startswith('$$') or s.startswith('\\['):
                body = s.strip().strip('$').replace('\\[', '').replace('\\]', '')
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(body)
                setfont(r, size=10.5, italic=True)
                continue
            # 普通段落：合并续行
            para = [s.strip()]
            while i < len(lines):
                nxt = lines[i].rstrip()
                if (not nxt.strip() or re.match(r'^(#{1,4})\s+', nxt)
                        or re.match(r'^\s*[-*]\s+', nxt)
                        or nxt.startswith('$$') or nxt.lstrip().startswith('>')):
                    break
                para.append(nxt.strip())
                i += 1
            txt = ' '.join(x for x in para if x)
            if txt:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(4)
                r = p.add_run(strip_md(txt))
                setfont(r)
    doc.save(out_path)
    return out_path


def strip_md(s: str) -> str:
    """去掉 **、` 与公式定界符，保留 LaTeX 正文（Word 里以源码呈现）。"""
    s = re.sub(r'\*\*(.+?)\*\*', r'\1', s)
    s = re.sub(r'`(.+?)`', r'\1', s)
    s = s.replace('$$', '').replace('\\[', '').replace('\\]', '')
    s = s.replace('\\(', '').replace('\\)', '')
    return re.sub(r'\s{2,}', ' ', s).strip()
