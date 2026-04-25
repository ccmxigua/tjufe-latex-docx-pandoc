#!/opt/homebrew/bin/python3
from __future__ import annotations

import copy
import hashlib
import re
import sys
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
XML_NS = 'http://www.w3.org/XML/1998/namespace'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

ET.register_namespace('w', W_NS)
ET.register_namespace('r', R_NS)
ET.register_namespace('', CT_NS)

PAGE_W = '10431'
PAGE_H = '14740'
TOP = '1134'
BOTTOM = '850'
LEFT = '1134'
RIGHT = '1134'
HEADER = '850'
FOOTER = '992'
TOC_PLACEHOLDER = '__TJUFE_TOC_PLACEHOLDER__'
FOOTNOTE_NUMFMT = 'decimalEnclosedCircleChinese'
UNIT_PREFIXES = ('单位：', '单位:', '计量单位：', '计量单位:', '数据单位：', '数据单位:')
SOURCE_PREFIXES = ('资料来源：', '资料来源:', '数据来源：', '数据来源:', '来源：', '来源:')
BODY_ATLEAST_STYLE_IDS = {
    'Normal',
    'BodyText',
    'FirstParagraph',
    'Bibliography',
    'AbstractBodyCN',
    'AbstractBodyEN',
    'KeywordsLineCN',
    'KeywordsLineEN',
    'AcknowledgementsBody',
    'StatementBody',
}
XREF_RE = re.compile(r'\[\[\[TJUFE_XREF:([^|\]]+)\|([^\]]+)\]\]\]')
EQLABEL_RE = re.compile(r'^\[\[\[TJUFE_EQLABEL:([^\]]+)\]\]\]$')
LABEL_RE = re.compile(r'TJUFE_LABEL__(.+?)__')


def sanitize_bookmark_name(label: str) -> str:
    name = re.sub(r'[^0-9A-Za-z_]', '_', label)
    if not name:
        name = 'eqref'
    if name[0].isdigit():
        name = f'bm_{name}'
    if len(name) > 24:
        digest = hashlib.sha1(name.encode('utf-8')).hexdigest()[:8]
        head = name[:15].rstrip('_')
        name = f'{head}_{digest}'
    return f'TJUFE_{name}'


def qn(ns: str, local: str) -> str:
    mapping = {'w': W_NS, 'm': M_NS, 'r': R_NS, 'rel': REL_NS, 'ct': CT_NS, 'wp': WP_NS, 'a': A_NS, 'pic': PIC_NS}
    return f'{{{mapping[ns]}}}{local}'


def paragraph_text(p: ET.Element) -> str:
    texts = []
    for t in p.findall('.//' + qn('w', 't')):
        if t.text:
            texts.append(t.text)
    return ''.join(texts).strip()


def paragraph_style(p: ET.Element) -> str | None:
    ppr = p.find(qn('w', 'pPr'))
    if ppr is None:
        return None
    pstyle = ppr.find(qn('w', 'pStyle'))
    if pstyle is None:
        return None
    return pstyle.get(qn('w', 'val'))


def set_paragraph_style(p: ET.Element, style_id: str) -> None:
    ppr = p.find(qn('w', 'pPr'))
    if ppr is None:
        ppr = ET.SubElement(p, qn('w', 'pPr'))
    pstyle = ppr.find(qn('w', 'pStyle'))
    if pstyle is None:
        pstyle = ET.SubElement(ppr, qn('w', 'pStyle'))
    pstyle.set(qn('w', 'val'), style_id)


def has_run_style(p: ET.Element, style_id: str) -> bool:
    for rstyle in p.findall('.//' + qn('w', 'rStyle')):
        if rstyle.get(qn('w', 'val')) == style_id:
            return True
    return False


def paragraph_has_math(p: ET.Element) -> bool:
    return p.find('.//' + qn('m', 'oMathPara')) is not None or p.find('.//' + qn('m', 'oMath')) is not None


def paragraph_has_math_para(p: ET.Element) -> bool:
    return p.find('.//' + qn('m', 'oMathPara')) is not None


def paragraph_has_numpr(p: ET.Element) -> bool:
    ppr = p.find(qn('w', 'pPr'))
    if ppr is None:
        return False
    return ppr.find(qn('w', 'numPr')) is not None


def paragraph_has_drawing(p: ET.Element) -> bool:
    return p.find('.//' + qn('w', 'drawing')) is not None


def ensure_text_run(p: ET.Element) -> ET.Element:
    run = p.find(qn('w', 'r'))
    if run is None:
        run = ET.SubElement(p, qn('w', 'r'))
    text = run.find(qn('w', 't'))
    if text is None:
        text = ET.SubElement(run, qn('w', 't'))
    return text


def set_paragraph_text(p: ET.Element, text: str) -> None:
    texts = p.findall('.//' + qn('w', 't'))
    if texts:
        texts[0].text = text
        for extra in texts[1:]:
            extra.text = ''
    else:
        ensure_text_run(p).text = text


def clear_paragraph_content(p: ET.Element) -> None:
    for child in list(p):
        if child.tag != qn('w', 'pPr'):
            p.remove(child)


def append_plain_run(p: ET.Element, text: str) -> None:
    if not text:
        return
    r = ET.SubElement(p, qn('w', 'r'))
    t = ET.SubElement(r, qn('w', 't'))
    t.set(f'{{{XML_NS}}}space', 'preserve')
    t.text = text


def append_internal_hyperlink(p: ET.Element, anchor: str, text: str) -> None:
    hl = ET.SubElement(p, qn('w', 'hyperlink'))
    hl.set(qn('w', 'anchor'), anchor)
    hl.set(qn('w', 'history'), '1')
    r = ET.SubElement(hl, qn('w', 'r'))
    t = ET.SubElement(r, qn('w', 't'))
    t.set(f'{{{XML_NS}}}space', 'preserve')
    t.text = text


def is_text_only_run(el: ET.Element) -> bool:
    if el.tag != qn('w', 'r'):
        return False
    saw_text = False
    for child in list(el):
        if child.tag == qn('w', 'rPr'):
            continue
        if child.tag == qn('w', 't'):
            saw_text = True
            continue
        return False
    return saw_text


def run_text(el: ET.Element) -> str:
    texts: list[str] = []
    for t in el.findall(qn('w', 't')):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)


def make_text_run(text: str, template_run: ET.Element | None = None) -> ET.Element:
    r = ET.Element(qn('w', 'r'))
    if template_run is not None:
        rpr = template_run.find(qn('w', 'rPr'))
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
    t = ET.SubElement(r, qn('w', 't'))
    t.set(f'{{{XML_NS}}}space', 'preserve')
    t.text = text
    return r


def make_internal_hyperlink(anchor: str, text: str, template_run: ET.Element | None = None) -> ET.Element:
    hl = ET.Element(qn('w', 'hyperlink'))
    hl.set(qn('w', 'anchor'), anchor)
    hl.set(qn('w', 'history'), '1')
    r = make_text_run(text, template_run)
    hl.append(r)
    return hl


def localize_xref_display(text: str) -> str:
    s = text.strip()
    patterns: list[tuple[re.Pattern[str], str]] = [
        (re.compile(r'^Figure\s+(.+)$', re.I), '图{}'),
        (re.compile(r'^Table\s+(.+)$', re.I), '表{}'),
        (re.compile(r'^Theorem\s+(.+)$', re.I), '定理{}'),
        (re.compile(r'^Lemma\s+(.+)$', re.I), '引理{}'),
        (re.compile(r'^Proposition\s+(.+)$', re.I), '命题{}'),
        (re.compile(r'^Corollary\s+(.+)$', re.I), '推论{}'),
        (re.compile(r'^Appendix\s+(.+)$', re.I), '附录{}'),
        (re.compile(r'^Section\s+(.+)$', re.I), '第{}节'),
    ]
    for pat, fmt in patterns:
        m = pat.match(s)
        if m:
            return fmt.format(m.group(1).strip())
    return text


def xref_display_from_bookmark_paragraph(style: str | None, text: str) -> str | None:
    s = text.strip()
    if not s:
        return None
    m = re.match(r'^(第\d+章)', s)
    if m:
        return m.group(1)
    m = re.match(r'^(附录[A-Z])', s)
    if m:
        return m.group(1)
    m = re.match(r'^((?:\d+\.\d+(?:\.\d+)?|[A-Z]\.\d+(?:\.\d+)?))\s+', s)
    if m and style in {'Heading2', 'Heading3'}:
        return f'第{m.group(1)}节'
    m = re.match(r'^(续?图[^\s　]+)', s)
    if m:
        return m.group(1)
    m = re.match(r'^(续?表[^\s　]+)', s)
    if m:
        return m.group(1)
    # Do not treat a generic trailing parenthetical like “（右）” in normal prose as
    # the display text of a cross-reference. Restrict this fallback to formula-like
    # blocks only, so figure references remain 图X-Y / 表X-Y.
    if style in {'EquationBlock', 'Formula'}:
        m = re.search(r'(（[^）]+）)\s*$', s)
        if m:
            return m.group(1)
    m = re.match(r'^(命题|推论|引理|定理)\s*([A-Z]?\d+(?:\.\d+)*)', s)
    if m:
        return f'{m.group(1)}{m.group(2)}'
    if s.startswith('证明'):
        return '证明'
    return None


def next_semantic_display(paragraphs: list[ET.Element], idx: int) -> str | None:
    probe = idx + 1
    while probe < len(paragraphs):
        p = paragraphs[probe]
        s = paragraph_text(p).strip()
        if s:
            return xref_display_from_bookmark_paragraph(paragraph_style(p), s)
        probe += 1
    return None


def replace_xref_placeholders_in_paragraph(p: ET.Element, bookmark_map: dict[str, str], eq_display_map: dict[str, str]) -> None:
    text = ''.join((t.text or '') for t in p.findall('.//' + qn('w', 't')))
    if '[[[TJUFE_XREF:' not in text:
        return
    children = list(p)
    rebuilt: list[ET.Element] = []
    i = 0
    if children and children[0].tag == qn('w', 'pPr'):
        rebuilt.append(children[0])
        i = 1

    while i < len(children):
        child = children[i]
        if not is_text_only_run(child):
            rebuilt.append(child)
            i += 1
            continue

        buffer_runs: list[ET.Element] = []
        buffer_texts: list[str] = []
        while i < len(children) and is_text_only_run(children[i]):
            buffer_runs.append(children[i])
            buffer_texts.append(run_text(children[i]))
            i += 1

        combined = ''.join(buffer_texts)
        matches = list(XREF_RE.finditer(combined))
        if not matches:
            rebuilt.extend(buffer_runs)
            continue

        template_run = next((r for r in buffer_runs if run_text(r)), buffer_runs[0] if buffer_runs else None)
        pos = 0
        for m in matches:
            plain = combined[pos:m.start()]
            if plain:
                rebuilt.append(make_text_run(plain, template_run))
            label = m.group(1).strip()
            shown = localize_xref_display(eq_display_map.get(label, m.group(2)))
            anchor = bookmark_map.get(label)
            if anchor:
                rebuilt.append(make_internal_hyperlink(anchor, shown, template_run))
            else:
                rebuilt.append(make_text_run(shown, template_run))
            pos = m.end()
        tail = combined[pos:]
        if tail:
            rebuilt.append(make_text_run(tail, template_run))

    for child in list(p):
        p.remove(child)
    for child in rebuilt:
        p.append(child)


def extract_eq_label_marker(text: str) -> str | None:
    m = EQLABEL_RE.match(text.strip())
    if not m:
        return None
    return m.group(1).strip()


def extract_generic_label_markers(text: str) -> list[str]:
    return [m.group(1).strip() for m in LABEL_RE.finditer(text) if m.group(1).strip()]


def strip_label_markers_in_paragraph(p: ET.Element) -> None:
    children = list(p)
    rebuilt: list[ET.Element] = []
    i = 0
    if children and children[0].tag == qn('w', 'pPr'):
        rebuilt.append(children[0])
        i = 1

    while i < len(children):
        child = children[i]
        if not is_text_only_run(child):
            rebuilt.append(child)
            i += 1
            continue

        buffer_runs: list[ET.Element] = []
        buffer_texts: list[str] = []
        while i < len(children) and is_text_only_run(children[i]):
            buffer_runs.append(children[i])
            buffer_texts.append(run_text(children[i]))
            i += 1

        combined = ''.join(buffer_texts)
        cleaned = LABEL_RE.sub('', combined)
        if not cleaned:
            continue
        template_run = next((r for r in buffer_runs if run_text(r)), buffer_runs[0] if buffer_runs else None)
        rebuilt.append(make_text_run(cleaned, template_run))

    for child in list(p):
        p.remove(child)
    for child in rebuilt:
        p.append(child)


def strip_label_markers_in_metadata(document_root: ET.Element) -> None:
    for elem in document_root.iter():
        if elem.tag == qn('w', 'tblCaption'):
            val = elem.get(qn('w', 'val'))
            if val:
                elem.set(qn('w', 'val'), LABEL_RE.sub('', val).strip())


def extract_trailing_equation_label(p: ET.Element) -> str | None:
    text = paragraph_text(p)
    m = re.search(r'(（[^）]+）)\s*$', text)
    return m.group(1) if m else None


def attach_bookmarks_to_paragraph(p: ET.Element, labels: list[str], bookmark_map: dict[str, str], next_bookmark_id: int) -> int:
    if not labels:
        return next_bookmark_id
    insert_at = 1 if len(p) > 0 and p[0].tag == qn('w', 'pPr') else 0
    has_visible_content = any(child.tag != qn('w', 'pPr') for child in list(p))
    for label in labels:
        if label in bookmark_map:
            continue
        bookmark_name = sanitize_bookmark_name(label)
        bookmark_map[label] = bookmark_name
        bm_start = ET.Element(qn('w', 'bookmarkStart'))
        bm_start.set(qn('w', 'id'), str(next_bookmark_id))
        bm_start.set(qn('w', 'name'), bookmark_name)
        bm_end = ET.Element(qn('w', 'bookmarkEnd'))
        bm_end.set(qn('w', 'id'), str(next_bookmark_id))
        p.insert(insert_at, bm_start)
        if has_visible_content:
            p.append(bm_end)
            insert_at += 1
        else:
            p.insert(insert_at + 1, bm_end)
            insert_at += 2
        next_bookmark_id += 1
    return next_bookmark_id


def rebuild_heading_with_tab(p: ET.Element, prefix: str, title: str) -> None:
    ppr = p.find(qn('w', 'pPr'))
    for child in list(p):
        if child.tag != qn('w', 'pPr'):
            p.remove(child)

    r_prefix = ET.SubElement(p, qn('w', 'r'))
    t_prefix = ET.SubElement(r_prefix, qn('w', 't'))
    t_prefix.set(f'{{{XML_NS}}}space', 'preserve')
    t_prefix.text = prefix

    r_tab = ET.SubElement(p, qn('w', 'r'))
    ET.SubElement(r_tab, qn('w', 'tab'))

    r_title = ET.SubElement(p, qn('w', 'r'))
    t_title = ET.SubElement(r_title, qn('w', 't'))
    t_title.set(f'{{{XML_NS}}}space', 'preserve')
    t_title.text = title


def normalize_heading_separator(p: ET.Element, style: str | None, text: str) -> str:
    match = None
    if style == 'Heading2':
        match = re.match(r'^((?:\d+\.\d+|[A-Z]\.\d+))\s+(.+)$', text)
    elif style == 'Heading3':
        match = re.match(r'^((?:\d+\.\d+\.\d+|[A-Z]\.\d+\.\d+))\s+(.+)$', text)
    if not match:
        return text
    prefix, title = match.groups()
    normalized = f'{prefix}  {title}'
    set_paragraph_text(p, normalized)
    set_paragraph_style(p, style)
    return normalized


def is_body_heading1(text: str) -> bool:
    return bool(re.match(r'^第\d+章\s+.+', text))


def is_appendix_heading1(text: str) -> bool:
    return bool(re.match(r'^附录([A-Z])\s+.+', text))


def chapter_label_from_heading(text: str) -> str | None:
    m = re.match(r'^第(\d+)章', text)
    if m:
        return m.group(1)
    m = re.match(r'^附录([A-Z])', text)
    if m:
        return m.group(1)
    return None


def is_body_heading2(text: str) -> bool:
    return bool(re.match(r'^(\d+\.\d+|[A-Z]\.\d+)\s+.+', text))


def is_body_heading3(text: str) -> bool:
    return bool(re.match(r'^(\d+\.\d+\.\d+|[A-Z]\.\d+\.\d+)\s+.+', text))


def is_research_outputs_heading(text: str) -> bool:
    return text == '在学期间发表的学术论文与研究成果'


def heading_scope_kind(style: str | None, text: str) -> str | None:
    if style == 'AppendixHeading' or is_appendix_heading1(text):
        return 'appendix'
    if style == 'ResearchOutputsHeading' or is_research_outputs_heading(text):
        return 'research_outputs'
    if style == 'AcknowledgementsHeading' or text in {'后 记', '后记'}:
        return 'ack'
    return None


def should_demote_heading_styled_paragraph(scope_kind: str, style: str | None, text: str, p: ET.Element) -> bool:
    if style is None:
        return False
    if scope_kind == 'appendix' and style == 'AppendixHeading':
        if is_appendix_heading1(text):
            return False
        if paragraph_has_numpr(p) or paragraph_has_drawing(p):
            return False
        return True
    if scope_kind == 'research_outputs' and style == 'ResearchOutputsHeading':
        return text != '在学期间发表的学术论文与研究成果'
    if scope_kind == 'ack' and style == 'AcknowledgementsHeading':
        return text not in {'后 记', '后记'}
    return False


def body_style_for_demoted_heading(scope_kind: str, p: ET.Element) -> str:
    if paragraph_has_math_para(p):
        return 'EquationBlock'
    if scope_kind == 'research_outputs':
        return 'Bibliography'
    if scope_kind == 'ack':
        return 'AcknowledgementsBody'
    return 'Normal'


def append_equation_number(p: ET.Element, label: str, bookmark_name: str | None = None, bookmark_id: int | None = None) -> None:
    ppr = ensure_p_pr(p)
    jc = ppr.find(qn('w', 'jc'))
    if jc is not None:
        ppr.remove(jc)

    tabs = ppr.find(qn('w', 'tabs'))
    if tabs is None:
        tabs = ET.SubElement(ppr, qn('w', 'tabs'))
    for child in list(tabs):
        if child.tag == qn('w', 'tab'):
            tabs.remove(child)

    tab_center = ET.SubElement(tabs, qn('w', 'tab'))
    tab_center.set(qn('w', 'val'), 'center')
    tab_center.set(qn('w', 'pos'), '4500')

    tab_right = ET.SubElement(tabs, qn('w', 'tab'))
    tab_right.set(qn('w', 'val'), 'right')
    tab_right.set(qn('w', 'pos'), '9000')

    children = list(p)
    insert_at = 1 if children and children[0].tag == qn('w', 'pPr') else 0
    if not any(child.tag == qn('w', 'r') and child.find(qn('w', 'tab')) is not None for child in children[:insert_at+2]):
        r_center = ET.Element(qn('w', 'r'))
        ET.SubElement(r_center, qn('w', 'tab'))
        p.insert(insert_at, r_center)

    r_tab = ET.SubElement(p, qn('w', 'r'))
    ET.SubElement(r_tab, qn('w', 'tab'))
    if bookmark_name and bookmark_id is not None:
        bm_start = ET.SubElement(p, qn('w', 'bookmarkStart'))
        bm_start.set(qn('w', 'id'), str(bookmark_id))
        bm_start.set(qn('w', 'name'), bookmark_name)
    r_num = ET.SubElement(p, qn('w', 'r'))
    rpr = ET.SubElement(r_num, qn('w', 'rPr'))
    fonts = ET.SubElement(rpr, qn('w', 'rFonts'))
    fonts.set(qn('w', 'ascii'), 'Times New Roman')
    fonts.set(qn('w', 'hAnsi'), 'Times New Roman')
    fonts.set(qn('w', 'eastAsia'), 'SimSun')
    sz = ET.SubElement(rpr, qn('w', 'sz'))
    sz.set(qn('w', 'val'), '24')
    szcs = ET.SubElement(rpr, qn('w', 'szCs'))
    szcs.set(qn('w', 'val'), '24')
    t = ET.SubElement(r_num, qn('w', 't'))
    t.text = label
    if bookmark_name and bookmark_id is not None:
        bm_end = ET.SubElement(p, qn('w', 'bookmarkEnd'))
        bm_end.set(qn('w', 'id'), str(bookmark_id))


def ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(tag)
    if child is None:
        child = ET.SubElement(parent, tag)
    return child


def remove_children(parent: ET.Element, tags: set[str]) -> None:
    for child in list(parent):
        if child.tag in tags:
            parent.remove(child)


def configure_footnote_pr(footnote_pr: ET.Element) -> None:
    remove_children(footnote_pr, {qn('w', 'numRestart'), qn('w', 'numFmt')})
    num_restart = ET.SubElement(footnote_pr, qn('w', 'numRestart'))
    num_restart.set(qn('w', 'val'), 'eachPage')
    num_fmt = ET.SubElement(footnote_pr, qn('w', 'numFmt'))
    num_fmt.set(qn('w', 'val'), FOOTNOTE_NUMFMT)


def configure_sectpr(sectpr: ET.Element, *, page_fmt: str, page_start: int | None, header_rid: str | None, footer_rid: str | None) -> None:
    remove_children(sectpr, {qn('w', 'headerReference'), qn('w', 'footerReference'), qn('w', 'pgSz'), qn('w', 'pgMar'), qn('w', 'pgNumType'), qn('w', 'footnotePr'), qn('w', 'docGrid')})

    if header_rid:
        hr = ET.SubElement(sectpr, qn('w', 'headerReference'))
        hr.set(qn('w', 'type'), 'default')
        hr.set(qn('r', 'id'), header_rid)
    if footer_rid:
        fr = ET.SubElement(sectpr, qn('w', 'footerReference'))
        fr.set(qn('w', 'type'), 'default')
        fr.set(qn('r', 'id'), footer_rid)

    pgsz = ET.SubElement(sectpr, qn('w', 'pgSz'))
    pgsz.set(qn('w', 'w'), PAGE_W)
    pgsz.set(qn('w', 'h'), PAGE_H)

    pgmar = ET.SubElement(sectpr, qn('w', 'pgMar'))
    pgmar.set(qn('w', 'top'), TOP)
    pgmar.set(qn('w', 'right'), RIGHT)
    pgmar.set(qn('w', 'bottom'), BOTTOM)
    pgmar.set(qn('w', 'left'), LEFT)
    pgmar.set(qn('w', 'header'), HEADER)
    pgmar.set(qn('w', 'footer'), FOOTER)
    pgmar.set(qn('w', 'gutter'), '0')

    pgnum = ET.SubElement(sectpr, qn('w', 'pgNumType'))
    pgnum.set(qn('w', 'fmt'), page_fmt)
    if page_start is not None:
        pgnum.set(qn('w', 'start'), str(page_start))

    footnote_pr = ET.SubElement(sectpr, qn('w', 'footnotePr'))
    configure_footnote_pr(footnote_pr)

    doc_grid = ET.SubElement(sectpr, qn('w', 'docGrid'))
    doc_grid.set(qn('w', 'linePitch'), '326')
    doc_grid.set(qn('w', 'charSpace'), '0')


def next_rel_id(rels_root: ET.Element) -> str:
    nums = []
    for rel in rels_root.findall(qn('rel', 'Relationship')):
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try:
                nums.append(int(rid[3:]))
            except ValueError:
                pass
    return f'rId{max(nums, default=0) + 1}'


def ensure_relationship(rels_root: ET.Element, target: str, rel_type: str) -> str:
    for rel in rels_root.findall(qn('rel', 'Relationship')):
        if rel.get('Target') == target and rel.get('Type') == rel_type:
            return rel.get('Id')
    rid = next_rel_id(rels_root)
    rel = ET.SubElement(rels_root, qn('rel', 'Relationship'))
    rel.set('Id', rid)
    rel.set('Type', rel_type)
    rel.set('Target', target)
    return rid


def ensure_content_type(ct_root: ET.Element, part_name: str, content_type: str) -> None:
    for ov in ct_root.findall(qn('ct', 'Override')):
        if ov.get('PartName') == part_name:
            ov.set('ContentType', content_type)
            return
    ov = ET.SubElement(ct_root, qn('ct', 'Override'))
    ov.set('PartName', part_name)
    ov.set('ContentType', content_type)


def is_fixed_header_heading(style: str | None, text: str) -> bool:
    if style in {'ReferencesHeading', 'AppendixHeading', 'ResearchOutputsHeading', 'AcknowledgementsHeading'}:
        return True
    return style == 'Heading1' and (
        is_body_heading1(text)
        or is_appendix_heading1(text)
        or text in {'参考文献', '后 记', '后记'}
        or is_research_outputs_heading(text)
    )


def collect_fixed_header_sections(body: ET.Element) -> list[tuple[int, str]]:
    sections: list[tuple[int, str]] = []
    for idx, child in enumerate(list(body)):
        if child.tag != qn('w', 'p'):
            continue
        style = paragraph_style(child)
        text = paragraph_text(child)
        if is_fixed_header_heading(style, text):
            sections.append((idx, text))
    return sections


def header_xml(title: str) -> bytes:
    hdr = ET.Element(qn('w', 'hdr'))
    p = ET.SubElement(hdr, qn('w', 'p'))
    ppr = ET.SubElement(p, qn('w', 'pPr'))
    pstyle = ET.SubElement(ppr, qn('w', 'pStyle'))
    pstyle.set(qn('w', 'val'), 'Header')
    jc = ET.SubElement(ppr, qn('w', 'jc'))
    jc.set(qn('w', 'val'), 'center')
    pbdr = ET.SubElement(ppr, qn('w', 'pBdr'))
    bottom = ET.SubElement(pbdr, qn('w', 'bottom'))
    bottom.set(qn('w', 'val'), 'thinThickSmallGap')
    bottom.set(qn('w', 'sz'), '12')
    bottom.set(qn('w', 'space'), '1')
    bottom.set(qn('w', 'color'), 'auto')
    run = ET.SubElement(p, qn('w', 'r'))
    text = ET.SubElement(run, qn('w', 't'))
    text.text = title
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(hdr, encoding='utf-8')


def write_header_part(unzip_dir: Path, rels_root: ET.Element, ct_root: ET.Element, *, index: int, title: str) -> str:
    target = f'header-fixed-{index}.xml'
    rid = ensure_relationship(rels_root, target, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header')
    ensure_content_type(ct_root, f'/word/{target}', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
    (unzip_dir / 'word' / target).write_bytes(header_xml(title))
    return rid


def footer_xml() -> bytes:
    xml = fr'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W_NS}" xmlns:r="{R_NS}">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:fldSimple w:instr=" PAGE \\* MERGEFORMAT ">
      <w:r><w:t>1</w:t></w:r>
    </w:fldSimple>
  </w:p>
</w:ftr>
'''
    return xml.encode('utf-8')


def make_toc_paragraph() -> ET.Element:
    p = ET.Element(qn('w', 'p'))
    ppr = ET.SubElement(p, qn('w', 'pPr'))
    pstyle = ET.SubElement(ppr, qn('w', 'pStyle'))
    pstyle.set(qn('w', 'val'), 'Normal')

    r1 = ET.SubElement(p, qn('w', 'r'))
    fld1 = ET.SubElement(r1, qn('w', 'fldChar'))
    fld1.set(qn('w', 'fldCharType'), 'begin')

    r2 = ET.SubElement(p, qn('w', 'r'))
    instr = ET.SubElement(r2, qn('w', 'instrText'))
    instr.set(f'{{{XML_NS}}}space', 'preserve')
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '

    r3 = ET.SubElement(p, qn('w', 'r'))
    fld3 = ET.SubElement(r3, qn('w', 'fldChar'))
    fld3.set(qn('w', 'fldCharType'), 'separate')

    r4 = ET.SubElement(p, qn('w', 'r'))
    t = ET.SubElement(r4, qn('w', 't'))
    t.text = ''

    r5 = ET.SubElement(p, qn('w', 'r'))
    fld5 = ET.SubElement(r5, qn('w', 'fldChar'))
    fld5.set(qn('w', 'fldCharType'), 'end')
    return p


def update_settings(settings_root: ET.Element) -> None:
    update_fields = settings_root.find(qn('w', 'updateFields'))
    if update_fields is None:
        update_fields = ET.SubElement(settings_root, qn('w', 'updateFields'))
    update_fields.set(qn('w', 'val'), 'true')

    footnote_pr = settings_root.find(qn('w', 'footnotePr'))
    if footnote_pr is None:
        footnote_pr = ET.SubElement(settings_root, qn('w', 'footnotePr'))
    configure_footnote_pr(footnote_pr)


def normalize_body_line_spacing(styles_root: ET.Element) -> None:
    for style in styles_root.findall(qn('w', 'style')):
        style_id = style.get(qn('w', 'styleId'))
        if style_id not in BODY_ATLEAST_STYLE_IDS:
            continue
        ppr = style.find(qn('w', 'pPr'))
        if ppr is None:
            ppr = ET.SubElement(style, qn('w', 'pPr'))
        spacing = ppr.find(qn('w', 'spacing'))
        if spacing is None:
            spacing = ET.SubElement(ppr, qn('w', 'spacing'))
        spacing.set(qn('w', 'line'), '400')
        spacing.set(qn('w', 'lineRule'), 'atLeast')


def clean_caption_title(text: str, kind: str) -> tuple[bool, str]:
    text = re.sub(r'\s+', ' ', text.strip())
    explicit_continued = text.startswith(f'续{kind}')
    text = re.sub(rf'^续?{kind}\s*(?:[A-Z]?\d+(?:-\d+)?)\s*', '', text)
    return explicit_continued, text.strip()


def format_caption_text(kind: str, label: str, title: str, continued: bool) -> str:
    prefix = f'续{kind}{label}' if continued else f'{kind}{label}'
    return f'{prefix}    {title}'.strip()


def make_caption_label(chapter_label: str, seq: int) -> str:
    if chapter_label.isdigit():
        return f'{chapter_label}-{seq}'
    return f'{chapter_label}{seq}'


def is_blank_paragraph(elem: ET.Element) -> bool:
    return elem.tag == qn('w', 'p') and paragraph_text(elem) == ''


def make_blank_paragraph(style_id: str = 'Normal') -> ET.Element:
    p = ET.Element(qn('w', 'p'))
    ppr = ET.SubElement(p, qn('w', 'pPr'))
    pstyle = ET.SubElement(ppr, qn('w', 'pStyle'))
    pstyle.set(qn('w', 'val'), style_id)
    return p


def remove_adjacent_blank_paragraphs(body: ET.Element, idx: int, *, before: bool = False, after: bool = False) -> None:
    children = list(body)
    if before:
        cursor = idx - 1
        while cursor >= 0:
            child = children[cursor]
            if child.tag == qn('w', 'p') and is_blank_paragraph(child):
                body.remove(child)
                cursor -= 1
                continue
            break
    if after:
        children = list(body)
        if idx >= len(children):
            return
        cursor = idx + 1
        while cursor < len(children):
            child = children[cursor]
            if child.tag == qn('w', 'p') and is_blank_paragraph(child):
                body.remove(child)
                children = list(body)
                continue
            break


def ensure_blank_paragraph_before(body: ET.Element, idx: int, style_id: str = 'Normal') -> int:
    if idx <= 0:
        return idx
    children = list(body)
    prev_child = children[idx - 1]
    if prev_child.tag == qn('w', 'p') and is_blank_paragraph(prev_child):
        return idx - 1
    body.insert(idx, make_blank_paragraph(style_id))
    return idx


def ensure_blank_paragraph_after(body: ET.Element, idx: int, style_id: str = 'Normal') -> int:
    children = list(body)
    insert_at = idx + 1
    if insert_at >= len(children):
        return insert_at
    next_child = children[insert_at]
    if next_child.tag == qn('w', 'p') and is_blank_paragraph(next_child):
        return insert_at
    body.insert(insert_at, make_blank_paragraph(style_id))
    return insert_at


def next_nonblank_index(children: list[ET.Element], start: int) -> int | None:
    idx = start
    while idx < len(children):
        child = children[idx]
        if child.tag == qn('w', 'p') and is_blank_paragraph(child):
            idx += 1
            continue
        if child.tag not in {qn('w', 'p'), qn('w', 'tbl')}:
            idx += 1
            continue
        return idx
    return None


def prev_nonblank_index(children: list[ET.Element], start: int) -> int | None:
    idx = start
    while idx >= 0:
        child = children[idx]
        if child.tag == qn('w', 'p') and is_blank_paragraph(child):
            idx -= 1
            continue
        if child.tag not in {qn('w', 'p'), qn('w', 'tbl')}:
            idx -= 1
            continue
        return idx
    return None


def table_style(tbl: ET.Element) -> str | None:
    tbl_pr = tbl.find(qn('w', 'tblPr'))
    if tbl_pr is None:
        return None
    tbl_style = tbl_pr.find(qn('w', 'tblStyle'))
    if tbl_style is None:
        return None
    return tbl_style.get(qn('w', 'val'))


def ensure_tbl_pr(tbl: ET.Element) -> ET.Element:
    tbl_pr = tbl.find(qn('w', 'tblPr'))
    if tbl_pr is None:
        tbl_pr = ET.Element(qn('w', 'tblPr'))
        tbl.insert(0, tbl_pr)
    return tbl_pr


def ensure_tc_pr(tc: ET.Element) -> ET.Element:
    tc_pr = tc.find(qn('w', 'tcPr'))
    if tc_pr is None:
        tc_pr = ET.Element(qn('w', 'tcPr'))
        tc.insert(0, tc_pr)
    return tc_pr


def ensure_tr_pr(tr: ET.Element) -> ET.Element:
    tr_pr = tr.find(qn('w', 'trPr'))
    if tr_pr is None:
        tr_pr = ET.Element(qn('w', 'trPr'))
        tr.insert(0, tr_pr)
    return tr_pr


def ensure_p_pr(p: ET.Element) -> ET.Element:
    ppr = p.find(qn('w', 'pPr'))
    if ppr is None:
        ppr = ET.Element(qn('w', 'pPr'))
        p.insert(0, ppr)
    return ppr


def set_keep_next(p: ET.Element, enabled: bool = True) -> None:
    ppr = ensure_p_pr(p)
    keep_next = ppr.find(qn('w', 'keepNext'))
    if enabled:
        if keep_next is None:
            keep_next = ET.SubElement(ppr, qn('w', 'keepNext'))
        keep_next.set(qn('w', 'val'), '1')
    elif keep_next is not None:
        ppr.remove(keep_next)


def set_paragraph_alignment(p: ET.Element, align: str = 'center') -> None:
    ppr = ensure_p_pr(p)
    jc = ppr.find(qn('w', 'jc'))
    if jc is None:
        jc = ET.SubElement(ppr, qn('w', 'jc'))
    jc.set(qn('w', 'val'), align)


def set_paragraph_spacing_auto(p: ET.Element) -> None:
    """Prevent inline pictures from being clipped by inherited exact line spacing.

    Body text uses an exact line height in the reference DOCX.  If picture-only
    paragraphs merely delete their direct w:line/w:lineRule attributes, Word
    keeps inheriting the exact height from Normal and clips tall inline images.
    Therefore picture paragraphs must explicitly opt back into automatic line
    height and must not snap to the document grid.
    """
    ppr = ensure_p_pr(p)
    snap = ppr.find(qn('w', 'snapToGrid'))
    if snap is None:
        snap = ET.SubElement(ppr, qn('w', 'snapToGrid'))
    snap.set(qn('w', 'val'), '0')

    spacing = ppr.find(qn('w', 'spacing'))
    if spacing is None:
        spacing = ET.SubElement(ppr, qn('w', 'spacing'))
    spacing.set(qn('w', 'before'), '0')
    spacing.set(qn('w', 'after'), '0')
    spacing.set(qn('w', 'line'), '240')
    spacing.set(qn('w', 'lineRule'), 'auto')


def normalize_drawing_paragraph(p: ET.Element) -> None:
    if not paragraph_has_drawing(p):
        return
    set_paragraph_style(p, 'Normal')
    set_paragraph_alignment(p, 'center')
    set_paragraph_spacing_auto(p)


def set_border(elem: ET.Element, edge: str, val: str, *, sz: str | None = None, space: str = '0', color: str = 'auto') -> None:
    border = elem.find(qn('w', edge))
    if border is None:
        border = ET.SubElement(elem, qn('w', edge))
    border.attrib.clear()
    border.set(qn('w', 'val'), val)
    if val != 'nil':
        if sz is not None:
            border.set(qn('w', 'sz'), sz)
        border.set(qn('w', 'space'), space)
        border.set(qn('w', 'color'), color)


def apply_three_line_table_style(tbl: ET.Element) -> None:
    tbl_pr = ensure_tbl_pr(tbl)

    jc = tbl_pr.find(qn('w', 'jc'))
    if jc is None:
        jc = ET.SubElement(tbl_pr, qn('w', 'jc'))
    jc.set(qn('w', 'val'), 'center')

    tbl_borders = tbl_pr.find(qn('w', 'tblBorders'))
    if tbl_borders is None:
        tbl_borders = ET.SubElement(tbl_pr, qn('w', 'tblBorders'))
    for child in list(tbl_borders):
        tbl_borders.remove(child)
    set_border(tbl_borders, 'top', 'single', sz='12')
    set_border(tbl_borders, 'left', 'nil')
    set_border(tbl_borders, 'bottom', 'single', sz='12')
    set_border(tbl_borders, 'right', 'nil')
    set_border(tbl_borders, 'insideH', 'nil')
    set_border(tbl_borders, 'insideV', 'nil')

    rows = tbl.findall(qn('w', 'tr'))
    for row_idx, tr in enumerate(rows):
        for tc in tr.findall(qn('w', 'tc')):
            tc_pr = ensure_tc_pr(tc)
            tc_borders = tc_pr.find(qn('w', 'tcBorders'))
            if tc_borders is not None:
                tc_pr.remove(tc_borders)
            if row_idx == 0:
                tc_borders = ET.SubElement(tc_pr, qn('w', 'tcBorders'))
                set_border(tc_borders, 'bottom', 'single', sz='8')
            for p in tc.findall(qn('w', 'p')):
                set_paragraph_alignment(p, 'center')



def set_repeat_table_header(tbl: ET.Element, enabled: bool = True) -> None:
    rows = tbl.findall(qn('w', 'tr'))
    if not rows:
        return
    tr_pr = ensure_tr_pr(rows[0])
    tbl_header = tr_pr.find(qn('w', 'tblHeader'))
    if enabled:
        if tbl_header is None:
            tbl_header = ET.SubElement(tr_pr, qn('w', 'tblHeader'))
        tbl_header.set(qn('w', 'val'), '1')
    elif tbl_header is not None:
        tr_pr.remove(tbl_header)



def is_unit_line(text: str) -> bool:
    return text.startswith(UNIT_PREFIXES)


def is_source_line(text: str) -> bool:
    return text.startswith(SOURCE_PREFIXES)


def table_plain_text(tbl: ET.Element) -> str:
    texts = []
    for t in tbl.findall('.//' + qn('w', 't')):
        if t.text:
            texts.append(t.text)
    return ''.join(texts).strip()


def is_stylable_body_table(children: list[ET.Element], idx: int, content_start_idx: int | None) -> bool:
    if content_start_idx is None or idx < content_start_idx:
        return False
    tbl = children[idx]
    if tbl.tag != qn('w', 'tbl'):
        return False
    if table_style(tbl) == 'FigureTable':
        return False
    if table_plain_text(tbl) == '':
        return False
    return True


def inline_to_anchor(inline: ET.Element) -> ET.Element:
    anchor = ET.Element(qn('wp', 'anchor'))
    anchor.set('distT', '0')
    anchor.set('distB', '0')
    anchor.set('distL', '0')
    anchor.set('distR', '0')
    anchor.set('simplePos', '0')
    anchor.set('relativeHeight', '251659264')
    anchor.set('behindDoc', '0')
    anchor.set('locked', '0')
    anchor.set('layoutInCell', '1')
    anchor.set('allowOverlap', '0')

    simple_pos = ET.SubElement(anchor, qn('wp', 'simplePos'))
    simple_pos.set('x', '0')
    simple_pos.set('y', '0')

    position_h = ET.SubElement(anchor, qn('wp', 'positionH'))
    position_h.set('relativeFrom', 'column')
    align_h = ET.SubElement(position_h, qn('wp', 'align'))
    align_h.text = 'center'

    position_v = ET.SubElement(anchor, qn('wp', 'positionV'))
    position_v.set('relativeFrom', 'paragraph')
    pos_offset = ET.SubElement(position_v, qn('wp', 'posOffset'))
    pos_offset.text = '0'

    extent = inline.find(qn('wp', 'extent'))
    if extent is not None:
        anchor.append(copy.deepcopy(extent))
    else:
        fallback_extent = ET.SubElement(anchor, qn('wp', 'extent'))
        fallback_extent.set('cx', '0')
        fallback_extent.set('cy', '0')

    effect_extent = inline.find(qn('wp', 'effectExtent'))
    if effect_extent is not None:
        anchor.append(copy.deepcopy(effect_extent))

    ET.SubElement(anchor, qn('wp', 'wrapTopAndBottom'))

    doc_pr = inline.find(qn('wp', 'docPr'))
    if doc_pr is not None:
        anchor.append(copy.deepcopy(doc_pr))
    else:
        fallback_docpr = ET.SubElement(anchor, qn('wp', 'docPr'))
        fallback_docpr.set('id', '1')
        fallback_docpr.set('name', 'Picture')

    c_nv = inline.find(qn('wp', 'cNvGraphicFramePr'))
    if c_nv is not None:
        anchor.append(copy.deepcopy(c_nv))
    else:
        c_nv = ET.SubElement(anchor, qn('wp', 'cNvGraphicFramePr'))
        locks = ET.SubElement(c_nv, qn('a', 'graphicFrameLocks'))
        locks.set('noChangeAspect', '1')

    graphic = inline.find(qn('a', 'graphic'))
    if graphic is not None:
        anchor.append(copy.deepcopy(graphic))

    return anchor


def convert_paragraph_drawings_to_anchor(p: ET.Element) -> bool:
    changed = False
    for drawing in p.findall('.//' + qn('w', 'drawing')):
        for child in list(drawing):
            if child.tag == qn('wp', 'inline'):
                drawing.remove(child)
                drawing.append(inline_to_anchor(child))
                changed = True
    if changed:
        normalize_drawing_paragraph(p)
    return changed


def extract_figure_table_paragraphs(tbl: ET.Element) -> list[ET.Element]:
    paras: list[ET.Element] = []
    for p in tbl.findall('.//' + qn('w', 'p')):
        if not p.findall('.//' + qn('w', 'drawing')):
            continue
        new_p = copy.deepcopy(p)
        normalize_drawing_paragraph(new_p)
        set_keep_next(new_p, True)
        paras.append(new_p)
    return paras


def unwrap_figure_tables(body: ET.Element) -> None:
    idx = 0
    while idx < len(body):
        child = body[idx]
        if child.tag == qn('w', 'tbl') and table_style(child) == 'FigureTable':
            paras = extract_figure_table_paragraphs(child)
            body.remove(child)
            insert_at = idx
            for p in paras:
                body.insert(insert_at, p)
                insert_at += 1
            idx = insert_at
            continue
        idx += 1


def strip_leading_tab_runs(p: ET.Element) -> None:
    removable = []
    started = False
    for child in list(p):
        if child.tag == qn('w', 'pPr'):
            continue
        if child.tag != qn('w', 'r'):
            break
        has_tab = child.find(qn('w', 'tab')) is not None
        texts = [t.text or '' for t in child.findall('.//' + qn('w', 't'))]
        if not started and has_tab and not ''.join(texts).strip():
            removable.append(child)
            continue
        started = True
        break
    for child in removable:
        p.remove(child)


def tighten_list_indentation(numbering_root: ET.Element) -> None:
    for lvl in numbering_root.findall('.//' + qn('w', 'lvl')):
        ppr = lvl.find(qn('w', 'pPr'))
        if ppr is None:
            ppr = ET.SubElement(lvl, qn('w', 'pPr'))
        ind = ppr.find(qn('w', 'ind'))
        if ind is None:
            ind = ET.SubElement(ppr, qn('w', 'ind'))
        ilvl = int(lvl.get(qn('w', 'ilvl'), '0'))
        ind.set(qn('w', 'left'), str(420 + ilvl * 420))
        ind.set(qn('w', 'hanging'), '180')
        suff = lvl.find(qn('w', 'suff'))
        if suff is None:
            suff = ET.SubElement(lvl, qn('w', 'suff'))
        suff.set(qn('w', 'val'), 'space')


def replace_math_token_sequence(p: ET.Element, old_tokens: list[str], new_tokens: list[str]) -> bool:
    mts = p.findall('.//' + qn('m', 't'))
    values = [mt.text or '' for mt in mts]
    changed = False
    i = 0
    while i <= len(values) - len(old_tokens):
        if values[i:i + len(old_tokens)] == old_tokens:
            old_nodes = mts[i:i + len(old_tokens)]
            parent_map = {child: parent for parent in p.iter() for child in parent}
            run_nodes = [parent_map.get(node) for node in old_nodes]
            if not run_nodes or any(r is None for r in run_nodes):
                i += 1
                continue
            container = parent_map.get(run_nodes[0])
            if container is None or any(parent_map.get(r) is not container for r in run_nodes):
                i += 1
                continue
            container_children = list(container)
            insert_at = container_children.index(run_nodes[0])
            for run in run_nodes:
                if run in list(container):
                    container.remove(run)
            for token in new_tokens:
                mr = ET.Element(qn('m', 'r'))
                mt = ET.SubElement(mr, qn('m', 't'))
                mt.text = token
                container.insert(insert_at, mr)
                insert_at += 1
            changed = True
            mts = p.findall('.//' + qn('m', 't'))
            values = [mt.text or '' for mt in mts]
            i += len(new_tokens)
            continue
        i += 1
    return changed


def apply_term_corrections(body: ET.Element) -> None:
    """Apply terminology corrections and English residue replacements.
    
    Rules:
    - "进程" → "过程" (in stochastic process context)
    - "具有内存和持久性的系统" → "具有记忆性和持续性的系统"
    - "Fractional CIR process" → "分数 CIR 过程"
    - "has a unique solution." → "具有唯一解。"
    - Lemma/Theorem citation format: "引理 4.1 Mehrdoust and Fallah 2020" → "引理 4.1（Mehrdoust and Fallah, 2020）"
    """
    # Pattern for lemma/theorem/proposition/corollary citation format
    # Matches: 引理 4.1 Mehrdoust and Fallah 2020 → 引理 4.1（Mehrdoust and Fallah, 2020）
    lemma_thm_pattern = re.compile(
        r'^(引理|定理|命题|推论)\s*(\d+(?:\.\d+)?)\s+([A-Za-z][A-Za-z\s]+?)\s+(\d{4})$'
    )
    
    for p in body.iter(qn('w', 'p')):
        text = paragraph_text(p)
        if not text:
            continue
        
        # Rule 1: "进程" → "过程" (only in specific contexts to avoid over-correction)
        # Apply IN PLACE to preserve OMML math objects
        if '随机进程' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '随机进程' in w_t.text:
                    w_t.text = w_t.text.replace('随机进程', '随机过程')
        if '高斯进程' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '高斯进程' in w_t.text:
                    w_t.text = w_t.text.replace('高斯进程', '高斯过程')
        if '布朗进程' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '布朗进程' in w_t.text:
                    w_t.text = w_t.text.replace('布朗进程', '布朗运动')
        
        # Rule 2: "具有内存和持久性的系统" → "具有记忆性和持续性的系统"
        if '具有内存和持久性的系统' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '具有内存和持久性的系统' in w_t.text:
                    w_t.text = w_t.text.replace('具有内存和持久性的系统', '具有记忆性和持续性的系统')
        if '此类进程用于描述具有记忆性和持续性的系统' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '此类进程用于描述具有记忆性和持续性的系统' in w_t.text:
                    w_t.text = w_t.text.replace('此类进程用于描述具有记忆性和持续性的系统', '此类过程用于描述具有记忆性和持续性的系统')
        
        # Rule 3: "Fractional CIR process" → "分数 CIR 过程"
        if 'Fractional CIR process' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and 'Fractional CIR process' in w_t.text:
                    w_t.text = w_t.text.replace('Fractional CIR process', '分数 CIR 过程')
        
        # Rule 4: "has a unique solution." → "具有唯一解。"
        if 'has a unique solution.' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and 'has a unique solution.' in w_t.text:
                    w_t.text = w_t.text.replace('has a unique solution.', '具有唯一解。')
        
        # Rule 5: Lemma/Theorem citation format normalization
        # "引理 4.1 Mehrdoust and Fallah 2020" → "引理 4.1（Mehrdoust and Fallah, 2020）"
        m = lemma_thm_pattern.match(text.strip())
        if m:
            type_word = m.group(1)
            number = m.group(2)
            authors = m.group(3).strip()
            year = m.group(4)
            new_text = f'{type_word}{number}（{authors}, {year}）'
            # Replace entire paragraph text for this case since it's a full-line pattern
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text:
                    w_t.text = w_t.text.replace(text.strip(), new_text)
                    break  # Only replace first occurrence
        
        # Rule 6: P617 - Fix H-value classification paragraph variable loss and fraction errors
        # Bad: '当时，fBm 表示标准布朗运动。如果是，则' or 'H∈(12,1)'
        # Good: '当 H=1/2 时，fBm 表示标准布朗运动。如果是 H∈(1/2,1)，则'
        # IMPORTANT: Apply replacements IN PLACE to preserve OMML math objects!
        # Do NOT use set_paragraph_text() as it destroys math nodes!
        if '当时，fBm 表示标准布朗运动' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '当时，fBm 表示标准布朗运动' in w_t.text:
                    w_t.text = w_t.text.replace('当时，fBm 表示标准布朗运动', '当 H=1/2 时，fBm 表示标准布朗运动')
        if '如果是，则' in text and 'fBm' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '如果是，则' in w_t.text:
                    w_t.text = w_t.text.replace('如果是，则', '如果是 H∈(1/2,1)，则')
        if '当时，fBm 的增量呈负相关' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '当时，fBm 的增量呈负相关' in w_t.text:
                    w_t.text = w_t.text.replace('当时，fBm 的增量呈负相关', '当 H∈(0,1/2) 时，fBm 的增量呈负相关')
        # Rule 7: Fix fraction errors like H∈(12,1) -> H∈(1/2,1) and H∈[12,1) -> H∈[1/2,1)
        if 'H∈(12,1)' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and 'H∈(12,1)' in w_t.text:
                    w_t.text = w_t.text.replace('H∈(12,1)', 'H∈(1/2,1)')
        if 'H∈[12,1)' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and 'H∈[12,1)' in w_t.text:
                    w_t.text = w_t.text.replace('H∈[12,1)', 'H∈[1/2,1)')
        # Rule 8: Unified terminology - option names (回望期权 / 回顾期权 / 回溯看跌期权 → 回望期权)
        if '回溯看跌期权' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '回溯看跌期权' in w_t.text:
                    w_t.text = w_t.text.replace('回溯看跌期权', '回望看跌期权')
        if '回顾看跌期权' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '回顾看跌期权' in w_t.text:
                    w_t.text = w_t.text.replace('回顾看跌期权', '回望看跌期权')
        if '回顾看涨期权' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '回顾看涨期权' in w_t.text:
                    w_t.text = w_t.text.replace('回顾看涨期权', '回望看涨期权')
        if '回顾期权' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '回顾期权' in w_t.text:
                    w_t.text = w_t.text.replace('回顾期权', '回望期权')
        # Rule 9: Unified terminology - 进程 → 过程 (stochastic process context)
        if '此类进程' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '此类进程' in w_t.text:
                    w_t.text = w_t.text.replace('此类进程', '此类过程')
        if '随机进程' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and '随机进程' in w_t.text:
                    w_t.text = w_t.text.replace('随机进程', '随机过程')
        # Rule 10: Fix isolated citation-conversion residuals
        if 'Reference Wolfram 2023 Leastsquares 2023' in text:
            for w_t in p.findall('.//' + qn('w', 't')):
                if w_t.text and 'Reference Wolfram 2023 Leastsquares 2023' in w_t.text:
                    w_t.text = w_t.text.replace(' Reference Wolfram 2023 Leastsquares 2023', '')
                    w_t.text = w_t.text.replace('Reference Wolfram 2023 Leastsquares 2023', '')
        if 'RF 方法约需 540 秒 ' in text and '。作为比较' in text:
            text_nodes = p.findall('.//' + qn('w', 't'))
            for i, w_t in enumerate(text_nodes):
                if not w_t.text or '540 秒 ' not in w_t.text:
                    continue
                next_text = ''
                for j in range(i + 1, len(text_nodes)):
                    if text_nodes[j].text:
                        next_text = text_nodes[j].text
                        break
                if next_text.startswith('。作为比较'):
                    w_t.text = w_t.text.replace('540 秒 ', '540 秒')
        # Rule 11: Final targeted paragraph rewrite for fraction-loss paragraphs
        # These two prose paragraphs contain only lightweight inline math; full paragraph rewrite is acceptable.
        if '我们重点关注' in text and '简化为标准布朗运动' in text:
            new_text = '在本节中，我们将在具有比例交易成本的分数双 Heston 模型下开发浮动行使价回望看跌期权的定价模型。 我们重点关注 H∈[1/2,1) 的情况，其中 H=1/2 简化为标准布朗运动。 fBm 的积分表示由下式给出：'
            clear_paragraph_content(p)
            append_plain_run(p, new_text)
            text = new_text
        if '根据赫斯特参数的不同值对过程系列进行分类' in text:
            new_text = text
            new_text = new_text.replace('当时，fBm表示标准布朗运动。', '当H=1/2时，fBm表示标准布朗运动。')
            new_text = new_text.replace('如果是 ，则', '如果是 H∈(1/2,1)，则')
            new_text = new_text.replace('当时，fBm的增量呈负相关', '当H∈(0,1/2)时，fBm的增量呈负相关')
            new_text = new_text.replace('H∈(12,1)', 'H∈(1/2,1)')
            new_text = new_text.replace('H∈(0,12)', 'H∈(0,1/2)')
            new_text = new_text.replace('此类进程', '此类过程')
            new_text = new_text.replace('具有内存和持久性的系统', '具有记忆性和持续性的系统')
            if new_text != text:
                clear_paragraph_content(p)
                append_plain_run(p, new_text)
                text = new_text


def normalize_body(body: ET.Element) -> None:
    unwrap_figure_tables(body)

    current_chapter_label: str | None = None
    equation_no = 0
    table_no = 0
    figure_no = 0
    table_labels: dict[str, str] = {}
    figure_labels: dict[str, str] = {}
    first_abstract_idx = None
    pending_eq_label: str | None = None
    pending_generic_labels: list[str] = []
    bookmark_map: dict[str, str] = {}
    eq_display_map: dict[str, str] = {}
    next_bookmark_id = 1

    paragraphs = list(body.iter(qn('w', 'p')))
    for idx, p in enumerate(paragraphs):
        style = paragraph_style(p)
        text = paragraph_text(p)
        marker_label = extract_eq_label_marker(text)
        if marker_label:
            pending_eq_label = marker_label
            if p in list(body):
                body.remove(p)
            continue

        generic_labels = extract_generic_label_markers(text)
        marker_only_text = LABEL_RE.sub('', text).strip()
        if generic_labels and marker_only_text == '':
            prev_idx = prev_nonblank_index(paragraphs, idx - 1)
            attached = False
            if prev_idx is not None:
                prev_p = paragraphs[prev_idx]
                if prev_p in list(body):
                    next_bookmark_id = attach_bookmarks_to_paragraph(prev_p, generic_labels, bookmark_map, next_bookmark_id)
                    display = xref_display_from_bookmark_paragraph(paragraph_style(prev_p), paragraph_text(prev_p))
                    if display:
                        for label in generic_labels:
                            eq_display_map.setdefault(label, display)
                    if paragraph_has_math(prev_p):
                        eq_text = extract_trailing_equation_label(prev_p)
                        if eq_text:
                            for label in generic_labels:
                                eq_display_map.setdefault(label, eq_text)
                    attached = True
            if not attached:
                pending_generic_labels.extend(generic_labels)
            if p in list(body):
                body.remove(p)
            continue
        elif generic_labels:
            strip_label_markers_in_paragraph(p)
            text = paragraph_text(p)
            next_bookmark_id = attach_bookmarks_to_paragraph(p, generic_labels, bookmark_map, next_bookmark_id)
            display = xref_display_from_bookmark_paragraph(paragraph_style(p), text) or next_semantic_display(paragraphs, idx)
            if display:
                for label in generic_labels:
                    eq_display_map.setdefault(label, display)

        text = normalize_heading_separator(p, style, text)
        if first_abstract_idx is None and style == 'Heading1' and text in {'摘 要', '摘要', 'Abstract'}:
            first_abstract_idx = idx
        if style == 'Heading1':
            if text in {'摘 要', '摘要'}:
                set_paragraph_style(p, 'AbstractTitleCN')
            elif text == 'Abstract':
                set_paragraph_style(p, 'AbstractTitleEN')
            elif text in {'参考文献'}:
                set_paragraph_style(p, 'ReferencesHeading')
            elif text in {'后 记', '后记'}:
                set_paragraph_style(p, 'AcknowledgementsHeading')
            elif is_research_outputs_heading(text):
                set_paragraph_style(p, 'ResearchOutputsHeading')
            elif is_appendix_heading1(text):
                set_paragraph_style(p, 'AppendixHeading')
            elif is_body_heading1(text):
                pass
            else:
                set_paragraph_style(p, 'Normal')
        elif style == 'Heading2':
            if not is_body_heading2(text):
                set_paragraph_style(p, 'Normal')
        elif style == 'Heading3':
            if not is_body_heading3(text):
                set_paragraph_style(p, 'Normal')

        if pending_generic_labels and p in list(body):
            next_bookmark_id = attach_bookmarks_to_paragraph(p, pending_generic_labels, bookmark_map, next_bookmark_id)
            pending_generic_labels = []

        if paragraph_has_math(p):
            if paragraph_has_math_para(p):
                set_paragraph_style(p, 'EquationBlock')
            else:
                set_paragraph_style(p, 'Normal')

        chapter_label = chapter_label_from_heading(text)
        if chapter_label:
            current_chapter_label = chapter_label
            equation_no = 0
            table_no = 0
            figure_no = 0
            table_labels = {}
            figure_labels = {}
        elif current_chapter_label and paragraph_has_math_para(p):
            if pending_eq_label:
                equation_no += 1
                if current_chapter_label.isdigit():
                    eq_label = f'（{current_chapter_label}.{equation_no}）'
                else:
                    eq_label = f'（{current_chapter_label}{equation_no}）'
                bookmark_name = sanitize_bookmark_name(pending_eq_label)
                bookmark_map[pending_eq_label] = bookmark_name
                eq_display_map[pending_eq_label] = eq_label
                bookmark_id = next_bookmark_id
                next_bookmark_id += 1
                append_equation_number(p, eq_label, bookmark_name, bookmark_id)
                pending_eq_label = None
        elif current_chapter_label and re.match(r'^（[A-Z]\.\d+）$', text):
            set_paragraph_text(p, re.sub(r'^（([A-Z])\.(\d+)）$', r'（）', text))
        elif current_chapter_label and style == 'TableCaption' and text:
            explicit_continued, title = clean_caption_title(text, '表')
            if explicit_continued or title in table_labels:
                label = table_labels.get(title, make_caption_label(current_chapter_label, max(table_no, 1)))
                set_paragraph_text(p, format_caption_text('表', label, title, True))
            else:
                table_no += 1
                label = make_caption_label(current_chapter_label, table_no)
                table_labels[title] = label
                set_paragraph_text(p, format_caption_text('表', label, title, False))
        elif current_chapter_label and style == 'ImageCaption' and text:
            explicit_continued, title = clean_caption_title(text, '图')
            if explicit_continued or title in figure_labels:
                label = figure_labels.get(title, make_caption_label(current_chapter_label, max(figure_no, 1)))
                set_paragraph_text(p, format_caption_text('图', label, title, True))
            else:
                figure_no += 1
                label = make_caption_label(current_chapter_label, figure_no)
                figure_labels[title] = label
                set_paragraph_text(p, format_caption_text('图', label, title, False))

    if bookmark_map:
        anchor_to_label = {anchor: label for label, anchor in bookmark_map.items()}
        for p in body.iter(qn('w', 'p')):
            style = paragraph_style(p)
            text = paragraph_text(p).strip()
            display = xref_display_from_bookmark_paragraph(style, text)
            if not display:
                continue
            for b in p.findall(qn('w', 'bookmarkStart')):
                name = b.get(qn('w', 'name'))
                label = anchor_to_label.get(name)
                if label and label not in eq_display_map:
                    eq_display_map[label] = display

    children = list(body)
    content_start_idx = 0
    for probe_idx, probe_child in enumerate(children):
        if probe_child.tag != qn('w', 'p'):
            continue
        probe_style = paragraph_style(probe_child)
        if probe_style in {'AbstractTitleCN', 'AbstractTitleEN'}:
            content_start_idx = probe_idx
            break

    for p in body.iter(qn('w', 'p')):
        replace_xref_placeholders_in_paragraph(p, bookmark_map, eq_display_map)

    idx = 0
    current_scope_kind: str | None = None
    while idx < len(children):
        child = children[idx]
        if child.tag == qn('w', 'tbl') and is_stylable_body_table(children, idx, content_start_idx):
            idx = ensure_blank_paragraph_before(body, idx)
            children = list(body)
            idx = children.index(child)
            apply_three_line_table_style(child)
            set_repeat_table_header(child)
            ensure_blank_paragraph_after(body, idx)
            children = list(body)
            idx = children.index(child) + 1
            continue
        if child.tag != qn('w', 'p'):
            idx += 1
            continue

        style = paragraph_style(child)
        text = paragraph_text(child)
        new_scope_kind = heading_scope_kind(style, text)
        if new_scope_kind is not None:
            current_scope_kind = new_scope_kind

        if current_scope_kind and should_demote_heading_styled_paragraph(current_scope_kind, style, text, child):
            set_paragraph_style(child, body_style_for_demoted_heading(current_scope_kind, child))
            style = paragraph_style(child)

        if style == 'ReferencesHeading':
            current_scope_kind = None
            remove_adjacent_blank_paragraphs(body, idx, before=True, after=False)
            children = list(body)
            idx = children.index(child)
        elif style in {'AppendixHeading', 'ResearchOutputsHeading', 'AcknowledgementsHeading'}:
            remove_adjacent_blank_paragraphs(body, idx, before=True, after=False)
            ensure_blank_paragraph_after(body, idx)
            children = list(body)
            idx = children.index(child)
        elif style == 'Heading1' and is_body_heading1(text):
            current_scope_kind = None
        if style == 'TableCaption':
            ensure_blank_paragraph_before(body, idx)
            set_keep_next(child, True)
            children = list(body)
            idx = children.index(child)

            prev_idx = prev_nonblank_index(children, idx - 1)
            if prev_idx is not None and children[prev_idx].tag == qn('w', 'p') and is_unit_line(paragraph_text(children[prev_idx])):
                unit_para = children[prev_idx]
                set_paragraph_style(unit_para, 'TableMetaLine')
                set_keep_next(unit_para, True)
                body.remove(unit_para)
                insert_pos = list(body).index(child) + 1
                body.insert(insert_pos, unit_para)
                children = list(body)
                idx = children.index(child)

            last_block_idx = idx
            next_idx = next_nonblank_index(children, idx + 1)
            if next_idx is not None and children[next_idx].tag == qn('w', 'p'):
                text = paragraph_text(children[next_idx])
                if is_unit_line(text):
                    set_paragraph_style(children[next_idx], 'TableMetaLine')
                    set_keep_next(children[next_idx], True)
                    last_block_idx = next_idx
                    next_idx = next_nonblank_index(children, next_idx + 1)
            if next_idx is not None and children[next_idx].tag == qn('w', 'tbl'):
                apply_three_line_table_style(children[next_idx])
                set_repeat_table_header(children[next_idx])
                last_block_idx = next_idx
                source_idx = next_nonblank_index(children, next_idx + 1)
                if source_idx is not None and children[source_idx].tag == qn('w', 'p') and is_source_line(paragraph_text(children[source_idx])):
                    set_paragraph_style(children[source_idx], 'SourceNote')
                    last_block_idx = source_idx

            children = list(body)
            ensure_blank_paragraph_after(body, last_block_idx)
            children = list(body)
            idx = children.index(child) + 1
            continue

        if style == 'ImageCaption':
            ensure_blank_paragraph_before(body, idx)
            set_keep_next(child, True)
            children = list(body)
            idx = children.index(child)

            last_block_idx = idx
            next_idx = next_nonblank_index(children, idx + 1)
            while next_idx is not None and children[next_idx].tag == qn('w', 'p') and paragraph_has_drawing(children[next_idx]):
                last_block_idx = next_idx
                next_idx = next_nonblank_index(children, next_idx + 1)
            if next_idx is not None and children[next_idx].tag == qn('w', 'p') and is_source_line(paragraph_text(children[next_idx])):
                set_paragraph_style(children[next_idx], 'SourceNote')
                last_block_idx = next_idx

            children = list(body)
            ensure_blank_paragraph_after(body, last_block_idx)
            children = list(body)
            idx = children.index(child) + 1
            continue

        if idx >= content_start_idx and paragraph_has_drawing(child):
            normalize_drawing_paragraph(child)
            prev_idx = prev_nonblank_index(children, idx - 1)
            if prev_idx is None or children[prev_idx].tag != qn('w', 'p') or paragraph_style(children[prev_idx]) != 'ImageCaption':
                ensure_blank_paragraph_before(body, idx)
                children = list(body)
                idx = children.index(child)
                last_block_idx = idx
                next_idx = next_nonblank_index(children, idx + 1)
                while next_idx is not None and children[next_idx].tag == qn('w', 'p') and paragraph_has_drawing(children[next_idx]):
                    last_block_idx = next_idx
                    next_idx = next_nonblank_index(children, next_idx + 1)
                if next_idx is not None and children[next_idx].tag == qn('w', 'p') and is_source_line(paragraph_text(children[next_idx])):
                    set_paragraph_style(children[next_idx], 'SourceNote')
                    last_block_idx = next_idx
                children = list(body)
                ensure_blank_paragraph_after(body, last_block_idx)
                children = list(body)
                idx = children.index(child) + 1
                continue

        idx += 1


def process_docx(input_path: Path, output_path: Path) -> None:
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        unzip_dir = td_path / 'unzipped'
        unzip_dir.mkdir()
        with zipfile.ZipFile(input_path) as zf:
            zf.extractall(unzip_dir)

        document_path = unzip_dir / 'word' / 'document.xml'
        styles_path = unzip_dir / 'word' / 'styles.xml'
        settings_path = unzip_dir / 'word' / 'settings.xml'
        rels_path = unzip_dir / 'word' / '_rels' / 'document.xml.rels'
        ct_path = unzip_dir / '[Content_Types].xml'

        document_tree = ET.parse(document_path)
        document_root = document_tree.getroot()
        styles_tree = ET.parse(styles_path)
        styles_root = styles_tree.getroot()
        settings_tree = ET.parse(settings_path)
        settings_root = settings_tree.getroot()
        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        ct_tree = ET.parse(ct_path)
        ct_root = ct_tree.getroot()

        # create footer part and relationships
        footer_rid = ensure_relationship(rels_root, 'footer1.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer')
        ensure_content_type(ct_root, '/word/footer1.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
        (unzip_dir / 'word' / 'footer1.xml').write_bytes(footer_xml())

        update_settings(settings_root)
        normalize_body_line_spacing(styles_root)

        body = document_root.find(qn('w', 'body'))
        if body is None:
            raise RuntimeError('document body not found')

        # remove pandoc auto title-block residue at document top
        removable = {'Author', 'Date'}
        while len(body) > 0:
            first = body[0]
            if first.tag != qn('w', 'p'):
                break
            style = paragraph_style(first)
            if style in removable:
                body.remove(first)
                continue
            break

        # replace TOC placeholder paragraphs
        children = list(body)
        for idx, child in enumerate(children):
            if child.tag == qn('w', 'p') and paragraph_text(child) == TOC_PLACEHOLDER:
                body.remove(child)
                body.insert(idx, make_toc_paragraph())
                break

        normalize_body(body)
        apply_term_corrections(body)
        strip_label_markers_in_metadata(document_root)

        # locate abstract start and fixed-header sections
        body_children = list(body)
        abstract_body_idx = None
        for idx, child in enumerate(body_children):
            if child.tag != qn('w', 'p'):
                continue
            style = paragraph_style(child)
            if style == 'AbstractTitleCN':
                abstract_body_idx = idx
                break
            if style == 'AbstractTitleEN' and abstract_body_idx is None:
                abstract_body_idx = idx
                break

        # cover/title section ends at first abstract page
        if abstract_body_idx is not None:
            title_break_p = ET.Element(qn('w', 'p'))
            ppr = ET.SubElement(title_break_p, qn('w', 'pPr'))
            sectpr = ET.SubElement(ppr, qn('w', 'sectPr'))
            configure_sectpr(sectpr, page_fmt='decimal', page_start=1, header_rid=None, footer_rid=None)
            body.insert(abstract_body_idx, title_break_p)

        section_starts = collect_fixed_header_sections(body)
        header_rids = [
            write_header_part(unzip_dir, rels_root, ct_root, index=i + 1, title=title)
            for i, (_, title) in enumerate(section_starts)
        ]

        # final section properties = last fixed-header section (or headerless fallback)
        final_sectpr = body.find(qn('w', 'sectPr'))
        if final_sectpr is None:
            final_sectpr = ET.SubElement(body, qn('w', 'sectPr'))
        if section_starts:
            final_page_start = 1 if len(section_starts) == 1 else None
            configure_sectpr(final_sectpr, page_fmt='decimal', page_start=final_page_start, header_rid=header_rids[-1], footer_rid=footer_rid)
        else:
            configure_sectpr(final_sectpr, page_fmt='decimal', page_start=1, header_rid=None, footer_rid=footer_rid)

        # end each completed body section right before the next heading and bind it to a fixed header
        for section_idx in range(len(section_starts) - 1, 0, -1):
            insert_at, _ = section_starts[section_idx]
            prev_header_rid = header_rids[section_idx - 1]
            page_start = 1 if section_idx - 1 == 0 else None
            body_break_p = ET.Element(qn('w', 'p'))
            ppr = ET.SubElement(body_break_p, qn('w', 'pPr'))
            sectpr = ET.SubElement(ppr, qn('w', 'sectPr'))
            configure_sectpr(sectpr, page_fmt='decimal', page_start=page_start, header_rid=prev_header_rid, footer_rid=footer_rid)
            body.insert(insert_at, body_break_p)

        # abstract/toc section ends right before first fixed-header body section
        if section_starts:
            first_heading_idx = section_starts[0][0]
            front_break_p = ET.Element(qn('w', 'p'))
            ppr = ET.SubElement(front_break_p, qn('w', 'pPr'))
            sectpr = ET.SubElement(ppr, qn('w', 'sectPr'))
            configure_sectpr(sectpr, page_fmt='upperRoman', page_start=1, header_rid=None, footer_rid=footer_rid)
            body.insert(first_heading_idx, front_break_p)

        document_tree.write(document_path, encoding='utf-8', xml_declaration=True)
        styles_tree.write(styles_path, encoding='utf-8', xml_declaration=True)
        settings_tree.write(settings_path, encoding='utf-8', xml_declaration=True)
        rels_tree.write(rels_path, encoding='utf-8', xml_declaration=True)
        ct_tree.write(ct_path, encoding='utf-8', xml_declaration=True)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for path in sorted(unzip_dir.rglob('*')):
                if path.is_file():
                    zf.write(path, path.relative_to(unzip_dir))


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print('usage: postprocess_tjufe_docx.py <input.docx> <output.docx>', file=sys.stderr)
        return 2
    input_path = Path(argv[1]).resolve()
    output_path = Path(argv[2]).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    process_docx(input_path, output_path)
    print(str(output_path))
    return 0


if __name__ == '__main__':
    raise SystemExit(main(sys.argv))
