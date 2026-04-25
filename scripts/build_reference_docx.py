#!/opt/homebrew/bin/python3
from __future__ import annotations

import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_NS = 'http://www.w3.org/XML/1998/namespace'

ET.register_namespace('w', W_NS)

ROOT = Path(__file__).resolve().parent.parent
OUT_DOCX = ROOT / 'reference.docx'
PANDOC = shutil.which('pandoc') or '/opt/homebrew/bin/pandoc'

PAGE_W = '10431'   # 16K 184mm
PAGE_H = '14740'   # 16K 260mm
TOP = '1134'       # 20mm
BOTTOM = '850'     # 15mm
LEFT = '1134'      # 20mm
RIGHT = '1134'     # 20mm
HEADER = '850'     # 1.5cm
FOOTER = '992'     # 1.75cm
DOCGRID_LINE_PITCH = '326'


def qn(tag: str) -> str:
    prefix, local = tag.split(':', 1)
    if prefix != 'w':
        raise ValueError(tag)
    return f'{{{W_NS}}}{local}'


def find_style(styles_root: ET.Element, style_id: str) -> ET.Element | None:
    for st in styles_root.findall(qn('w:style')):
        if st.get(qn('w:styleId')) == style_id:
            return st
    return None


def remove_children(parent: ET.Element, tags: list[str]) -> None:
    names = {qn(tag) for tag in tags}
    for child in list(parent):
        if child.tag in names:
            parent.remove(child)


def ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(qn(tag))
    if child is None:
        child = ET.SubElement(parent, qn(tag))
    return child


def apply_run_style(rpr: ET.Element, spec: dict) -> None:
    remove_children(rpr, ['w:rFonts', 'w:b', 'w:bCs', 'w:sz', 'w:szCs', 'w:lang', 'w:vertAlign', 'w:color', 'w:i', 'w:iCs', 'w:spacing'])
    fonts = ET.SubElement(rpr, qn('w:rFonts'))
    fonts.set(qn('w:ascii'), spec.get('ascii', 'Times New Roman'))
    fonts.set(qn('w:hAnsi'), spec.get('hAnsi', spec.get('ascii', 'Times New Roman')))
    fonts.set(qn('w:eastAsia'), spec.get('eastAsia', 'SimSun'))
    fonts.set(qn('w:cs'), spec.get('cs', spec.get('ascii', 'Times New Roman')))
    if spec.get('bold'):
        ET.SubElement(rpr, qn('w:b'))
        ET.SubElement(rpr, qn('w:bCs'))
    sz = ET.SubElement(rpr, qn('w:sz'))
    sz.set(qn('w:val'), str(spec.get('size', 24)))
    szcs = ET.SubElement(rpr, qn('w:szCs'))
    szcs.set(qn('w:val'), str(spec.get('size', 24)))
    if spec.get('vertAlign'):
        vert_align = ET.SubElement(rpr, qn('w:vertAlign'))
        vert_align.set(qn('w:val'), spec['vertAlign'])
    lang = ET.SubElement(rpr, qn('w:lang'))
    lang.set(qn('w:val'), spec.get('lang', 'en-US'))
    lang.set(qn('w:eastAsia'), spec.get('lang_east', 'zh-CN'))


def apply_para_style(ppr: ET.Element, spec: dict) -> None:
    remove_children(
        ppr,
        ['w:jc', 'w:spacing', 'w:ind', 'w:keepNext', 'w:keepLines', 'w:outlineLvl', 'w:numPr', 'w:contextualSpacing', 'w:tabs']
    )
    if spec.get('contextualSpacing'):
        ET.SubElement(ppr, qn('w:contextualSpacing'))
    jc = ET.SubElement(ppr, qn('w:jc'))
    jc.set(qn('w:val'), spec.get('align', 'both'))
    spacing = ET.SubElement(ppr, qn('w:spacing'))
    spacing.set(qn('w:before'), str(spec.get('before', 0)))
    spacing.set(qn('w:after'), str(spec.get('after', 0)))
    spacing.set(qn('w:line'), str(spec.get('line', 360)))
    spacing.set(qn('w:lineRule'), spec.get('lineRule', 'auto'))
    if 'outlineLvl' in spec:
        outline = ET.SubElement(ppr, qn('w:outlineLvl'))
        outline.set(qn('w:val'), str(spec['outlineLvl']))
    ind_kwargs = {}
    if 'left' in spec:
        ind_kwargs[qn('w:left')] = str(spec['left'])
    if 'firstLineChars' in spec:
        ind_kwargs[qn('w:firstLineChars')] = str(spec['firstLineChars'])
    if 'leftChars' in spec:
        ind_kwargs[qn('w:leftChars')] = str(spec['leftChars'])
    if 'rightChars' in spec:
        ind_kwargs[qn('w:rightChars')] = str(spec['rightChars'])
    if 'firstLine' in spec:
        ind_kwargs[qn('w:firstLine')] = str(spec['firstLine'])
    if 'hanging' in spec:
        ind_kwargs[qn('w:hanging')] = str(spec['hanging'])
    if 'hangingChars' in spec:
        ind_kwargs[qn('w:hangingChars')] = str(spec['hangingChars'])
    if ind_kwargs:
        ind = ET.SubElement(ppr, qn('w:ind'))
        for k, v in ind_kwargs.items():
            ind.set(k, v)
    if spec.get('tabs'):
        tabs = ET.SubElement(ppr, qn('w:tabs'))
        for tab_spec in spec['tabs']:
            tab = ET.SubElement(tabs, qn('w:tab'))
            tab.set(qn('w:val'), str(tab_spec.get('val', 'left')))
            tab.set(qn('w:pos'), str(tab_spec['pos']))


def upsert_style(styles_root: ET.Element, style_id: str, style_type: str, name: str, p_spec: dict, r_spec: dict, based_on: str = 'Normal', custom: bool = False) -> None:
    style = find_style(styles_root, style_id)
    if style is None:
        style = ET.SubElement(styles_root, qn('w:style'))
        style.set(qn('w:type'), style_type)
        style.set(qn('w:styleId'), style_id)
        if custom:
            style.set(qn('w:customStyle'), '1')
        name_el = ET.SubElement(style, qn('w:name'))
        name_el.set(qn('w:val'), name)
        based = ET.SubElement(style, qn('w:basedOn'))
        based.set(qn('w:val'), based_on)
        ET.SubElement(style, qn('w:qFormat'))
    ppr = ensure_child(style, 'w:pPr')
    rpr = ensure_child(style, 'w:rPr')
    apply_para_style(ppr, p_spec)
    apply_run_style(rpr, r_spec)


def upsert_character_style(styles_root: ET.Element, style_id: str, name: str, r_spec: dict, based_on: str = 'DefaultParagraphFont', custom: bool = False) -> None:
    style = find_style(styles_root, style_id)
    if style is None:
        style = ET.SubElement(styles_root, qn('w:style'))
        style.set(qn('w:type'), 'character')
        style.set(qn('w:styleId'), style_id)
    if custom:
        style.set(qn('w:customStyle'), '1')
    else:
        style.attrib.pop(qn('w:customStyle'), None)
    name_el = ensure_child(style, 'w:name')
    name_el.set(qn('w:val'), name)
    based = ensure_child(style, 'w:basedOn')
    based.set(qn('w:val'), based_on)
    if style.find(qn('w:qFormat')) is None:
        ET.SubElement(style, qn('w:qFormat'))
    ppr = style.find(qn('w:pPr'))
    if ppr is not None:
        style.remove(ppr)
    rpr = ensure_child(style, 'w:rPr')
    apply_run_style(rpr, r_spec)


def normalize_paragraph_indentation(styles_root: ET.Element, preserve_firstline: set[str]) -> list[str]:
    patched: list[str] = []
    for style in styles_root.findall(qn('w:style')):
        if style.get(qn('w:type')) != 'paragraph':
            continue
        style_id = style.get(qn('w:styleId')) or ''
        if style_id in preserve_firstline:
            continue
        based = style.find(qn('w:basedOn'))
        if based is None or based.get(qn('w:val')) != 'Normal':
            continue
        ppr = ensure_child(style, 'w:pPr')
        ind = ppr.find(qn('w:ind'))
        if ind is None:
            ind = ET.SubElement(ppr, qn('w:ind'))
        if qn('w:firstLineChars') in ind.attrib or qn('w:firstLine') in ind.attrib:
            continue
        if qn('w:hanging') in ind.attrib or qn('w:hangingChars') in ind.attrib:
            continue
        ind.set(qn('w:firstLineChars'), '0')
        patched.append(style_id)
    return patched


def ensure_document_section(document_root: ET.Element) -> None:
    body = document_root.find(qn('w:body'))
    sect = body.find(qn('w:sectPr')) if body is not None else None
    if body is None:
        raise RuntimeError('document body not found')
    if sect is None:
        sect = ET.SubElement(body, qn('w:sectPr'))
    for child in list(sect):
        if child.tag in {qn('w:pgSz'), qn('w:pgMar'), qn('w:docGrid')}:
            sect.remove(child)
    pg_sz = ET.SubElement(sect, qn('w:pgSz'))
    pg_sz.set(qn('w:w'), PAGE_W)
    pg_sz.set(qn('w:h'), PAGE_H)
    pg_mar = ET.SubElement(sect, qn('w:pgMar'))
    pg_mar.set(qn('w:top'), TOP)
    pg_mar.set(qn('w:right'), RIGHT)
    pg_mar.set(qn('w:bottom'), BOTTOM)
    pg_mar.set(qn('w:left'), LEFT)
    pg_mar.set(qn('w:header'), HEADER)
    pg_mar.set(qn('w:footer'), FOOTER)
    pg_mar.set(qn('w:gutter'), '0')
    doc_grid = ET.SubElement(sect, qn('w:docGrid'))
    doc_grid.set(qn('w:linePitch'), DOCGRID_LINE_PITCH)
    doc_grid.set(qn('w:charSpace'), '0')


def build_seed_docx(tmpdir: Path) -> Path:
    seed_md = tmpdir / 'seed.md'
    seed_md.write_text('# Seed\n\nReference template seed.\n', encoding='utf-8')
    out = tmpdir / 'seed.docx'
    subprocess.run([PANDOC, str(seed_md), '-o', str(out)], check=True)
    return out


def main() -> None:
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        seed_docx = build_seed_docx(td_path)
        unzip_dir = td_path / 'unzipped'
        unzip_dir.mkdir()
        with zipfile.ZipFile(seed_docx) as zf:
            zf.extractall(unzip_dir)

        styles_path = unzip_dir / 'word' / 'styles.xml'
        document_path = unzip_dir / 'word' / 'document.xml'

        styles_tree = ET.parse(styles_path)
        styles_root = styles_tree.getroot()

        body_p = {'align': 'both', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 200}
        noindent_body_p = {'align': 'both', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 0}
        # Cover/title-page content uses single spacing in the school spec.
        # Using exact 240 twips clips larger fonts in Word.
        cover_meta_p = {'align': 'center', 'before': 0, 'after': 0, 'line': 240, 'lineRule': 'auto', 'firstLineChars': 0}
        cover_label_p = {'align': 'distribute', 'before': 0, 'after': 0, 'line': 240, 'lineRule': 'auto', 'firstLineChars': 0}
        heading1_center_p = {'align': 'center', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 0}
        heading_left_p = {'align': 'left', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 0}
        toc_heading_p = {'align': 'center', 'before': 0, 'after': 0, 'line': 360, 'lineRule': 'auto', 'firstLineChars': 0}
        toc_body_p = {'align': 'both', 'before': 0, 'after': 0, 'line': 360, 'lineRule': 'auto', 'firstLineChars': 0}
        signature_p = {'align': 'right', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 0}
        table_meta_p = {'align': 'right', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 0}
        source_note_p = {'align': 'both', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 0}
        equation_block_p = {'align': 'center', 'before': 120, 'after': 120, 'line': 360, 'lineRule': 'auto', 'firstLineChars': 0}

        body_r = {'ascii': 'Times New Roman', 'eastAsia': 'SimSun', 'size': 24}
        body_r_b = {'ascii': 'Times New Roman', 'eastAsia': 'SimSun', 'size': 24, 'bold': True}
        tnr_r = {'ascii': 'Times New Roman', 'eastAsia': 'Times New Roman', 'size': 24}
        black_r_16 = {'ascii': 'Times New Roman', 'eastAsia': 'SimHei', 'size': 32, 'bold': True}
        black_r_15 = {'ascii': 'Times New Roman', 'eastAsia': 'SimHei', 'size': 30}
        black_r_14 = {'ascii': 'Times New Roman', 'eastAsia': 'SimHei', 'size': 28}
        black_r_12 = {'ascii': 'Times New Roman', 'eastAsia': 'SimHei', 'size': 24}
        black_r_10_5 = {'ascii': 'Times New Roman', 'eastAsia': 'SimHei', 'size': 21, 'bold': True}
        song_r_22 = {'ascii': 'Times New Roman', 'eastAsia': 'SimSun', 'size': 44, 'bold': True}
        song_r_18 = {'ascii': 'Times New Roman', 'eastAsia': 'SimSun', 'size': 36, 'bold': True}
        song_r_16 = {'ascii': 'Times New Roman', 'eastAsia': 'SimSun', 'size': 32}
        song_r_14 = {'ascii': 'Times New Roman', 'eastAsia': 'SimSun', 'size': 28}
        song_r_14_b = {'ascii': 'Times New Roman', 'eastAsia': 'SimSun', 'size': 28, 'bold': True}
        song_r_10_5 = {'ascii': 'Times New Roman', 'eastAsia': 'SimSun', 'size': 21}
        song_r_9 = {'ascii': 'Times New Roman', 'eastAsia': 'SimSun', 'size': 18}
        english_15_b = {'ascii': 'Times New Roman', 'eastAsia': 'Times New Roman', 'size': 30, 'bold': True}
        english_14_b = {'ascii': 'Times New Roman', 'eastAsia': 'Times New Roman', 'size': 28, 'bold': True}
        kaiti_12 = {'ascii': 'Times New Roman', 'eastAsia': 'KaiTi', 'size': 24}
        kaiti_9 = {'ascii': 'Times New Roman', 'eastAsia': 'KaiTi', 'size': 18}
        footer_r = {'ascii': 'Times New Roman', 'eastAsia': 'Times New Roman', 'size': 15}
        header_r = {'ascii': 'Times New Roman', 'eastAsia': 'SimHei', 'size': 21}
        footnote_ref_r = {'ascii': 'Times New Roman', 'eastAsia': 'Times New Roman', 'size': 16, 'vertAlign': 'superscript'}
        keyword_label_cn_r = {'ascii': 'Times New Roman', 'eastAsia': 'SimHei', 'size': 24, 'bold': True}
        keyword_label_en_r = {'ascii': 'Times New Roman', 'eastAsia': 'Times New Roman', 'size': 24, 'bold': True}
        signature_r = {'ascii': 'Times New Roman', 'eastAsia': 'SimHei', 'size': 24}

        references_heading_p = {'align': 'center', 'before': 800, 'after': 800, 'line': 400, 'lineRule': 'atLeast', 'outlineLvl': 0, 'firstLineChars': 0}
        postpage_heading_p = {'align': 'center', 'before': 0, 'after': 800, 'line': 400, 'lineRule': 'atLeast', 'outlineLvl': 0, 'firstLineChars': 0}

        style_specs = [
            ('Normal', 'paragraph', 'Normal', body_p, body_r, 'Normal', False),
            ('BodyText', 'paragraph', 'Body Text', noindent_body_p, body_r, 'Normal', False),
            ('FirstParagraph', 'paragraph', 'First Paragraph', noindent_body_p, body_r, 'Normal', False),
            ('Title', 'paragraph', 'Title', {'align': 'center', 'before': 0, 'after': 0, 'line': 240, 'lineRule': 'auto', 'firstLineChars': 0}, song_r_22, 'Normal', False),
            ('Subtitle', 'paragraph', 'Subtitle', {'align': 'center', 'before': 0, 'after': 0, 'line': 240, 'lineRule': 'auto', 'firstLineChars': 0}, song_r_18, 'Normal', False),
            ('Heading1', 'paragraph', 'heading 1', {**heading1_center_p, 'outlineLvl': 0}, black_r_15, 'Normal', False),
            ('Heading2', 'paragraph', 'heading 2', {**heading_left_p, 'outlineLvl': 1, 'left': 0, 'firstLineChars': 0}, black_r_14, 'Normal', False),
            ('Heading3', 'paragraph', 'heading 3', {**heading_left_p, 'outlineLvl': 2, 'left': 0, 'firstLineChars': 0}, black_r_12, 'Normal', False),
            ('AbstractTitleCN', 'paragraph', 'Abstract Title CN', {'align': 'center', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'outlineLvl': 0, 'firstLineChars': 0}, black_r_16, 'Normal', True),
            ('AbstractTitleEN', 'paragraph', 'Abstract Title EN', {'align': 'center', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'outlineLvl': 0, 'firstLineChars': 0}, {'ascii': 'Times New Roman', 'eastAsia': 'Times New Roman', 'size': 32, 'bold': True}, 'Normal', True),
            ('AbstractBodyCN', 'paragraph', 'Abstract Body CN', body_p, body_r, 'Normal', True),
            ('AbstractBodyEN', 'paragraph', 'Abstract Body EN', body_p, tnr_r, 'Normal', True),
            ('KeywordsLineCN', 'paragraph', 'Keywords Line CN', noindent_body_p, body_r, 'Normal', True),
            ('KeywordsLineEN', 'paragraph', 'Keywords Line EN', noindent_body_p, tnr_r, 'Normal', True),
            ('TOCHeading', 'paragraph', 'TOC Heading', toc_heading_p, black_r_16, 'Normal', True),
            ('TOC1', 'paragraph', 'toc 1', toc_body_p, song_r_14, 'Normal', False),
            ('TOC2', 'paragraph', 'toc 2', {'align': 'both', 'before': 0, 'after': 0, 'line': 360, 'lineRule': 'auto', 'leftChars': 200, 'firstLineChars': 0}, body_r, 'Normal', False),
            ('TOC3', 'paragraph', 'toc 3', {'align': 'both', 'before': 0, 'after': 0, 'line': 360, 'lineRule': 'auto', 'leftChars': 400, 'firstLineChars': 0}, body_r, 'Normal', False),
            ('Bibliography', 'paragraph', 'Bibliography', {'align': 'both', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 0}, kaiti_12, 'Normal', False),
            ('AcknowledgementsBody', 'paragraph', 'Acknowledgements Body', body_p, kaiti_12, 'Normal', True),
            ('StatementBody', 'paragraph', 'Statement Body', body_p, body_r, 'Normal', True),
            ('SignatureLine', 'paragraph', 'Signature Line', signature_p, signature_r, 'Normal', True),
            ('Caption', 'paragraph', 'Caption', {'align': 'center', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 0}, black_r_10_5, 'Normal', False),
            ('TableCaption', 'paragraph', 'Table Caption', {'align': 'center', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 0}, black_r_10_5, 'Normal', True),
            ('ImageCaption', 'paragraph', 'Image Caption', {'align': 'center', 'before': 0, 'after': 0, 'line': 400, 'lineRule': 'atLeast', 'firstLineChars': 0}, black_r_10_5, 'Normal', True),
            ('EquationBlock', 'paragraph', 'Equation Block', equation_block_p, body_r, 'Normal', True),
            ('TableMetaLine', 'paragraph', 'Table Meta Line', table_meta_p, black_r_10_5, 'Normal', True),
            ('SourceNote', 'paragraph', 'Source Note', source_note_p, song_r_9, 'Normal', True),
            ('FootnoteText', 'paragraph', 'Footnote Text', {'align': 'left', 'before': 0, 'after': 0, 'line': 240, 'lineRule': 'auto', 'firstLineChars': 0}, {'ascii': 'Times New Roman', 'eastAsia': 'SimSun', 'size': 18}, 'Normal', False),
            ('CoverTopLine', 'paragraph', 'Cover Top Line', cover_meta_p, song_r_16, 'Normal', True),
            ('CoverInfoLabel', 'paragraph', 'Cover Info Label', cover_label_p, song_r_14, 'Normal', True),
            ('CoverInfoValue', 'paragraph', 'Cover Info Value', cover_meta_p, song_r_14, 'Normal', True),
            ('CoverDate', 'paragraph', 'Cover Date', cover_meta_p, song_r_14, 'Normal', True),
            ('TitlePageMeta', 'paragraph', 'Title Page Meta', {'align': 'both', 'before': 0, 'after': 0, 'line': 240, 'lineRule': 'auto', 'firstLineChars': 0}, song_r_14_b, 'Normal', True),
            ('TitlePageDegree', 'paragraph', 'Title Page Degree', cover_meta_p, song_r_14_b, 'Normal', True),
            ('EnglishTitle', 'paragraph', 'English Title', cover_meta_p, english_15_b, 'Normal', True),
            ('EnglishSubtitle', 'paragraph', 'English Subtitle', cover_meta_p, english_14_b, 'Normal', True),
            ('TitlePageInfoLabel', 'paragraph', 'Title Page Info Label', cover_label_p, song_r_14_b, 'Normal', True),
            ('TitlePageInfoValue', 'paragraph', 'Title Page Info Value', cover_meta_p, song_r_14_b, 'Normal', True),
            ('ReferencesHeading', 'paragraph', 'References Heading', references_heading_p, black_r_15, 'Normal', True),
            ('AppendixHeading', 'paragraph', 'Appendix Heading', postpage_heading_p, black_r_15, 'Normal', True),
            ('ResearchOutputsHeading', 'paragraph', 'Research Outputs Heading', postpage_heading_p, black_r_15, 'Normal', True),
            ('AcknowledgementsHeading', 'paragraph', 'Acknowledgements Heading', postpage_heading_p, black_r_15, 'Normal', True),
            ('Header', 'paragraph', 'Header', {'align': 'center', 'before': 0, 'after': 0, 'line': 240, 'lineRule': 'exact', 'firstLineChars': 0}, header_r, 'Normal', True),
            ('Footer', 'paragraph', 'Footer', {'align': 'center', 'before': 0, 'after': 0, 'line': 240, 'lineRule': 'exact', 'firstLineChars': 0}, footer_r, 'Normal', False),
        ]

        for spec in style_specs:
            upsert_style(styles_root, *spec)

        normalize_paragraph_indentation(
            styles_root,
            preserve_firstline={
                'Normal',
                'AbstractBodyCN',
                'AbstractBodyEN',
                'AcknowledgementsBody',
                'StatementBody',
            },
        )

        upsert_character_style(styles_root, 'FootnoteReference', 'Footnote Reference', footnote_ref_r)
        upsert_character_style(styles_root, 'KeywordLabelCN', 'Keyword Label CN', keyword_label_cn_r, custom=True)
        upsert_character_style(styles_root, 'KeywordLabelEN', 'Keyword Label EN', keyword_label_en_r, custom=True)

        styles_tree.write(styles_path, encoding='utf-8', xml_declaration=True)

        document_tree = ET.parse(document_path)
        document_root = document_tree.getroot()
        ensure_document_section(document_root)
        document_tree.write(document_path, encoding='utf-8', xml_declaration=True)

        if OUT_DOCX.exists():
            OUT_DOCX.unlink()
        with zipfile.ZipFile(OUT_DOCX, 'w', zipfile.ZIP_DEFLATED) as zf:
            for path in sorted(unzip_dir.rglob('*')):
                if path.is_file():
                    zf.write(path, path.relative_to(unzip_dir))

    print(str(OUT_DOCX))


if __name__ == '__main__':
    main()


