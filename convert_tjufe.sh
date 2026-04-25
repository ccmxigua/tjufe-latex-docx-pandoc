#!/bin/bash
set -euo pipefail

SCAN_ONLY=0
SCAN_DOCX_PATH=""

if [ "${1:-}" = "--scan-docx" ]; then
  if [ "$#" -lt 2 ]; then
    echo "usage: ./convert_tjufe.sh --scan-docx <docx> [metadata.yaml]" >&2
    exit 2
  fi
  SCAN_ONLY=1
  SCAN_DOCX_PATH="$2"
  INPUT="$SCAN_DOCX_PATH"
  OUTPUT="$SCAN_DOCX_PATH"
  shift 2
else
  if [ "$#" -lt 2 ]; then
    echo "usage: ./convert_tjufe.sh <input.tex> <output.docx> [metadata.yaml] [extra pandoc args...]" >&2
    exit 2
  fi

  INPUT="$1"
  OUTPUT="$2"
  shift 2
fi

METADATA=""
if [ "$#" -gt 0 ]; then
  case "$1" in
    *.yaml|*.yml)
      METADATA="$1"
      shift
      ;;
  esac
fi

if [ -z "$METADATA" ]; then
  INPUT_DIR="$(cd "$(dirname "$INPUT")" && pwd)"
  for cand in \
    "$INPUT_DIR/metadata.yaml" \
    "$INPUT_DIR/metadata.yml" \
    "$INPUT_DIR/$(basename "${INPUT%.*}").yaml" \
    "$INPUT_DIR/$(basename "${INPUT%.*}").yml"
  do
    if [ -f "$cand" ]; then
      METADATA="$cand"
      break
    fi
  done
fi

EXTRA_ARGS=("$@")
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
FILTER="$SCRIPT_DIR/filters/tjufe-thesis.lua"
REFDOC="$SCRIPT_DIR/reference.docx"
BUILD_REF="$SCRIPT_DIR/scripts/build_reference_docx.py"
PREPROCESS_TEX="$SCRIPT_DIR/scripts/preprocess_tjufe_tex.py"
POSTPROCESS="$SCRIPT_DIR/scripts/postprocess_tjufe_docx.py"
TMP_DIR="$(mktemp -d)"
TMP_TEX="$TMP_DIR/preprocessed.tex"
TMP_DOCX="$TMP_DIR/raw.docx"

cleanup() {
  rm -rf "$TMP_DIR"
}
trap cleanup EXIT

resolve_bibliography_path() {
  if [ -z "$METADATA" ] || [ ! -f "$METADATA" ]; then
    return 0
  fi
  local raw
  raw="$(awk -F':' '/^bibliography:[[:space:]]*/{sub(/^[^:]+:[[:space:]]*/, ""); print; exit}' "$METADATA" | tr -d '"' | tr -d "'")"
  raw="${raw## }"
  raw="${raw%% }"
  if [ -z "$raw" ]; then
    return 0
  fi
  if [[ "$raw" = /* ]]; then
    printf '%s\n' "$raw"
  else
    local meta_dir
    meta_dir="$(cd "$(dirname "$METADATA")" && pwd)"
    printf '%s/%s\n' "$meta_dir" "$raw"
  fi
}

scan_bibliography_or_fail() {
  local bib_path="$1"
  local metadata_path="${2:-}"
  if [ -z "$bib_path" ] || [ ! -f "$bib_path" ]; then
    return 0
  fi
  /opt/homebrew/bin/python3 - "$bib_path" "$metadata_path" <<'PY'
from pathlib import Path
import datetime as dt
import re
import sys

path = Path(sys.argv[1])
metadata_path = Path(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2] else None
text = path.read_text(encoding='utf-8', errors='ignore')
starts = list(re.finditer(r'@\w+\s*\{\s*([^,\s]+)\s*,', text, flags=re.I))


def load_simple_yaml(path):
    if path is None or not path.is_file():
        return {}
    out = {}
    for line in path.read_text(encoding='utf-8', errors='ignore').splitlines():
        if not line.strip() or line.lstrip().startswith('#'):
            continue
        if line.startswith((' ', '\t', '- ')):
            continue
        if ':' not in line:
            continue
        key, value = line.split(':', 1)
        out[key.strip()] = value.strip().strip('"').strip("'")
    return out


def extract_field(entry, name):
    m = re.search(r'\\b' + re.escape(name) + r'\\s*=\\s*([{"\'])', entry, flags=re.I)
    if not m:
        return None
    opener = m.group(1)
    start = m.end(1)
    if opener == '{':
        depth = 1
        i = start
        out = []
        while i < len(entry):
            ch = entry[i]
            if ch == '{':
                depth += 1
                out.append(ch)
            elif ch == '}':
                depth -= 1
                if depth == 0:
                    break
                out.append(ch)
            else:
                out.append(ch)
            i += 1
        return ''.join(out).strip()
    end = entry.find(opener, start)
    if end == -1:
        return None
    return entry[start:end].strip()


def contains_cjk(s):
    return bool(re.search(r'[\u4e00-\u9fff]', s or ''))


def ascii_ratio(s):
    if not s:
        return 0.0
    chars = [ch for ch in s if not ch.isspace()]
    if not chars:
        return 0.0
    return sum(1 for ch in chars if ord(ch) < 128) / len(chars)


def split_authors(author):
    if not author:
        return []
    return [a.strip() for a in re.split(r'\s+and\s+', author, flags=re.I) if a.strip()]


def looks_foreign_entry(fields):
    lang = ' '.join(filter(None, [fields.get('language'), fields.get('langid')])).lower()
    if any(token in lang for token in ['english', 'en', 'eng']):
        return True
    author = fields.get('author', '')
    title = fields.get('title', '')
    journal = fields.get('journal', '') or fields.get('booktitle', '') or fields.get('publisher', '')
    merged = ' '.join(filter(None, [author, title, journal]))
    if merged and not contains_cjk(merged) and ascii_ratio(merged) > 0.85:
        return True
    return False


def infer_degree_mode(meta):
    hint = ' '.join(part for part in [
        meta.get('degree_type', ''),
        meta.get('cover_degree_line', ''),
        meta.get('doctoral_apply_degree', ''),
    ] if part)
    if '博士' in hint:
        return 'doctor'
    if '专业硕士' in hint or '专业学位' in hint:
        return 'professional_master'
    if '硕士' in hint:
        return 'master'
    return 'unknown'


meta = load_simple_yaml(metadata_path)
degree_mode = infer_degree_mode(meta)
if degree_mode == 'doctor':
    min_total = 100
elif degree_mode == 'professional_master':
    min_total = 30
elif degree_mode == 'master':
    min_total = 40
else:
    min_total = None

bad = []
warn = []
entries = []
current_year = dt.datetime.now().year
recent_cutoff = current_year - 4

for i, m in enumerate(starts):
    key = m.group(1).strip()
    start = m.start()
    end = starts[i + 1].start() if i + 1 < len(starts) else len(text)
    entry = text[start:end]
    reasons = []
    if 'Auto-generated placeholder during bibliography merge' in entry:
        reasons.append('placeholder-note')
    if 'Original local .bib entry was not found in this project' in entry:
        reasons.append('missing-local-bib-note')
    if 'complete metadata before formal bibliography compilation' in entry:
        reasons.append('incomplete-metadata-note')
    if re.search(r'\byear\s*=\s*\{\s*XXXX\s*\}', entry, flags=re.I):
        reasons.append('year-XXXX')
    m_title = re.search(r'\btitle\s*=\s*\{\s*([^{}]+?)\s*\}', entry, flags=re.I | re.S)
    if m_title and m_title.group(1).strip().lower() == key.lower():
        reasons.append('title-equals-citekey')

    fields = {name: extract_field(entry, name) for name in ['author', 'title', 'year', 'language', 'langid', 'journal', 'booktitle', 'publisher']}
    if not fields.get('author'):
        warn.append((key, 'missing-author'))
    if not fields.get('title'):
        warn.append((key, 'missing-title'))
    if not fields.get('year'):
        warn.append((key, 'missing-year'))

    title = (fields.get('title') or '').strip()
    if title and looks_foreign_entry(fields):
        m_word = re.search(r'[A-Za-z][A-Za-z0-9\-]*', title)
        if m_word:
            w = m_word.group(0)
            if w[:1].islower():
                warn.append((key, 'english-title-first-word-not-capitalized'))

    authors = split_authors(fields.get('author'))
    if len(authors) >= 4 and 'et al.' not in entry and '，等' not in entry and ', et al.' not in entry and ' et al' not in entry:
        warn.append((key, 'author-count-ge-4-without-et-al'))

    year_text = (fields.get('year') or '').strip()
    year_num = None
    m_year = re.search(r'(19|20)\d{2}', year_text)
    if m_year:
        year_num = int(m_year.group(0))

    entries.append({
        'key': key,
        'year': year_num,
        'is_foreign': looks_foreign_entry(fields),
    })

    if reasons:
        bad.append((key, reasons))

if bad:
    print(f'BIBLIOGRAPHY_SCAN_FAILED: {path}')
    for key, reasons in bad[:30]:
        print(f'  - {key}: ' + ','.join(reasons))
    if len(bad) > 30:
        print(f'  ... and {len(bad)-30} more bad entries')
    sys.exit(10)

total = len(entries)
recent = sum(1 for e in entries if e['year'] is not None and e['year'] >= recent_cutoff)
foreign = sum(1 for e in entries if e['is_foreign'])

if min_total is not None and total < min_total:
    print(f'BIBLIOGRAPHY_SCAN_WARN: total_too_few:{total}:{degree_mode}:{min_total}', file=sys.stderr)
if total > 0 and recent * 4 < total:
    print(f'BIBLIOGRAPHY_SCAN_WARN: recent_ratio_too_low:{recent}/{total}:cutoff={recent_cutoff}', file=sys.stderr)
if total > 0 and foreign * 5 < total:
    print(f'BIBLIOGRAPHY_SCAN_WARN: foreign_ratio_too_low:{foreign}/{total}', file=sys.stderr)
for key, reason in warn:
    print(f'BIBLIOGRAPHY_SCAN_WARN: {key}:{reason}', file=sys.stderr)
PY
}

scan_docx_or_fail() {
  local docx_path="$1"
  local metadata_path="${2:-}"
  if [ ! -f "$docx_path" ]; then
    echo "DOCX_SCAN_FAILED: file not found: $docx_path" >&2
    return 11
  fi
  /opt/homebrew/bin/python3 - "$docx_path" "$metadata_path" <<'PY'
from pathlib import Path
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

path = Path(sys.argv[1])
metadata_path = Path(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2] else None
W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
needles = [
    'Auto-generated placeholder during bibliography merge',
    'Original local .bib entry was not found in this project',
    'complete metadata before formal bibliography compilation',
    'XXXX',
    '__TJUFE_TOC_PLACEHOLDER__',
]
key_labels = ['U_BS', 'U_fBm', 'U_Model', 'N_BS', 'N_fBm', 'N_Model']


def qn(local: str) -> str:
    return W_NS + local


def style_map(styles_root):
    out = {}
    for style in styles_root.findall(qn('style')):
        style_id = style.get(qn('styleId'))
        if style_id:
            out[style_id] = style
    return out


def get_style_spacing(style):
    if style is None:
        return None
    ppr = style.find(qn('pPr'))
    if ppr is None:
        return None
    return ppr.find(qn('spacing'))


def get_style_rpr(style):
    if style is None:
        return None
    return style.find(qn('rPr'))


def get_east_asia_font(style):
    rpr = get_style_rpr(style)
    if rpr is None:
        return None
    fonts = rpr.find(qn('rFonts'))
    if fonts is None:
        return None
    return fonts.get(qn('eastAsia'))


def get_font_size(style):
    rpr = get_style_rpr(style)
    if rpr is None:
        return None
    sz = rpr.find(qn('sz'))
    if sz is None:
        return None
    return sz.get(qn('val'))


def style_is_bold(style):
    rpr = get_style_rpr(style)
    if rpr is None:
        return False
    return rpr.find(qn('b')) is not None or rpr.find(qn('bCs')) is not None


def paragraph_style(p):
    ppr = p.find(qn('pPr'))
    if ppr is None:
        return None
    pstyle = ppr.find(qn('pStyle'))
    if pstyle is None:
        return None
    return pstyle.get(qn('val'))


def paragraph_text(node):
    return ''.join(t.text or '' for t in node.iter(qn('t'))).strip()


def looks_reference_line(text):
    return bool(re.match(r'^\[\d+\]', text))


def is_reference_line_ending_ok(text):
    return bool(re.search(r'[.。．]$', text))


def is_backmatter_start(style):
    return style in {
        'ReferencesHeading', 'AppendixHeading', 'ResearchOutputsHeading', 'AcknowledgementsHeading'
    }


def chinese_char_count(text):
    return len(re.findall(r'[\u4e00-\u9fff]', text))


def english_word_count(text):
    return len(re.findall(r"[A-Za-z]+(?:[-'][A-Za-z]+)*", text))


def load_simple_yaml(path):
    if path is None or not path.is_file():
        return {}
    out = {}
    for line in path.read_text(encoding='utf-8', errors='ignore').splitlines():
        if not line.strip() or line.lstrip().startswith('#'):
            continue
        if line.startswith((' ', '\t', '- ')):
            continue
        if ':' not in line:
            continue
        key, value = line.split(':', 1)
        out[key.strip()] = value.strip().strip('"').strip("'")
    return out


def has_bad_cn_keyword_punct(raw):
    return any(ch in raw for ch in [',', ';', '；', '、'])


def has_bad_en_keyword_punct(raw):
    return any(ch in raw for ch in ['，', '；', '、', '。']) or (',' in raw)


def has_bad_en_keyword_spacing(raw):
    if ';' not in raw:
        return False
    if re.search(r';(?=[^\s])', raw):
        return True
    if re.search(r'\s{2,}', raw):
        return True
    return False


def has_terminal_keyword_punct(raw):
    return bool(re.search(r'[;；，,。.]\s*$', raw))


def has_bad_en_abstract_punct(raw):
    return any(ch in raw for ch in ['，', '；', '、', '。', '：'])


def has_bad_en_abstract_spacing(raw):
    if re.search(r'\b[A-Za-z]+[，；：]\b', raw):
        return True
    if re.search(r'\s{2,}', raw):
        return True
    return False


def first_row_has_tbl_header(tbl):
    rows = tbl.findall(qn('tr'))
    if not rows:
        return False
    tr_pr = rows[0].find(qn('trPr'))
    if tr_pr is None:
        return False
    return tr_pr.find(qn('tblHeader')) is not None


def is_blank_paragraph(child):
    return child.tag == qn('p') and paragraph_text(child) == ''


def paragraph_has_drawing(p):
    return p.find('.//' + qn('drawing')) is not None


def paragraph_has_manual_page_break(p):
    return any(br.get(qn('type')) == 'page' for br in p.iter(qn('br')))


def paragraph_has_page_break_before(p):
    ppr = p.find(qn('pPr'))
    return ppr is not None and ppr.find(qn('pageBreakBefore')) is not None


def is_structural_page_break_style(style):
    return style in {
        'Heading1', 'AbstractTitleCN', 'AbstractTitleEN', 'TOCHeading',
        'ReferencesHeading', 'AppendixHeading', 'ResearchOutputsHeading',
        'AcknowledgementsHeading', 'StatementHeading', 'Title',
    }


def parse_numbered_caption(text, kind):
    text = re.sub(r'\s+', ' ', (text or '').strip())
    pattern = rf'^(续?{kind})\s*([A-Z]?\d+(?:-\d+)?)\s*(.*)$'
    m = re.match(pattern, text)
    if not m:
        return None
    prefix, label, title = m.groups()
    return {
        'continued': prefix.startswith('续'),
        'label': label,
        'scope': caption_scope_from_label(label),
        'title': title.strip(),
    }


def caption_scope_from_label(label):
    if '-' in label:
        return label.split('-', 1)[0]
    m = re.match(r'^([A-Z])\d+', label)
    if m:
        return m.group(1)
    return label


def has_blank_before(children, idx):
    return idx > 0 and is_blank_paragraph(children[idx - 1])


def has_blank_after(children, idx):
    return idx + 1 < len(children) and is_blank_paragraph(children[idx + 1])


def next_nonblank_child_index(children, start):
    idx = start
    while idx < len(children):
        child = children[idx]
        if is_blank_paragraph(child):
            idx += 1
            continue
        if child.tag not in {qn('p'), qn('tbl')}:
            idx += 1
            continue
        return idx
    return None


def prev_nonblank_child_index(children, start):
    idx = start
    while idx >= 0:
        child = children[idx]
        if is_blank_paragraph(child):
            idx -= 1
            continue
        if child.tag not in {qn('p'), qn('tbl')}:
            idx -= 1
            continue
        return idx
    return None


def table_cluster_end_index(children, caption_idx):
    last_idx = caption_idx
    next_idx = next_nonblank_child_index(children, caption_idx + 1)
    if next_idx is not None and children[next_idx].tag == qn('p') and paragraph_style(children[next_idx]) == 'TableMetaLine':
        last_idx = next_idx
        next_idx = next_nonblank_child_index(children, next_idx + 1)
    if next_idx is not None and children[next_idx].tag == qn('tbl'):
        last_idx = next_idx
        source_idx = next_nonblank_child_index(children, next_idx + 1)
        if source_idx is not None and children[source_idx].tag == qn('p') and paragraph_style(children[source_idx]) == 'SourceNote':
            last_idx = source_idx
    return last_idx


def image_caption_cluster_end_index(children, caption_idx):
    last_idx = caption_idx
    next_idx = next_nonblank_child_index(children, caption_idx + 1)
    while next_idx is not None and children[next_idx].tag == qn('p') and paragraph_has_drawing(children[next_idx]):
        last_idx = next_idx
        next_idx = next_nonblank_child_index(children, next_idx + 1)
    if next_idx is not None and children[next_idx].tag == qn('p') and paragraph_style(children[next_idx]) == 'SourceNote':
        last_idx = next_idx
    return last_idx


def drawing_cluster_end_index(children, drawing_idx):
    last_idx = drawing_idx
    next_idx = next_nonblank_child_index(children, drawing_idx + 1)
    while next_idx is not None and children[next_idx].tag == qn('p') and paragraph_has_drawing(children[next_idx]):
        last_idx = next_idx
        next_idx = next_nonblank_child_index(children, next_idx + 1)
    if next_idx is not None and children[next_idx].tag == qn('p') and paragraph_style(children[next_idx]) == 'SourceNote':
        last_idx = next_idx
    return last_idx


def image_caption_drawing_indices(children, caption_idx):
    indices = []
    idx = caption_idx - 1
    while idx >= 0:
        child = children[idx]
        if child.tag == qn('p') and paragraph_has_drawing(child):
            indices.append(idx)
            idx -= 1
            continue
        if is_blank_paragraph(child):
            idx -= 1
            continue
        break
    idx = caption_idx + 1
    while idx < len(children):
        child = children[idx]
        if child.tag == qn('p') and paragraph_has_drawing(child):
            indices.append(idx)
            idx += 1
            continue
        if is_blank_paragraph(child):
            idx += 1
            continue
        break
    return sorted(set(indices))


def has_subfigure_label(text):
    return bool(re.search(r'[（(]\s*[a-zA-Z]\s*[）)]', text or ''))


def text_has_unit_marker(text):
    raw = text or ''
    if re.search(r'(单位|计量单位|数据单位)\s*[:：]', raw, flags=re.I):
        return True
    if re.search(r'[（(][^）)]{0,18}(?:%|％|元|万元|亿元|人|个|次|件|项|年|月|日|天|小时|分钟|秒|kg|g|t|m|km|cm|mm|℃|K|Pa|MPa|GB|MB)[^）)]{0,18}[）)]', raw, flags=re.I):
        return True
    return False


def text_suggests_measured_figure(text):
    return bool(re.search(r'(趋势|分布|统计|占比|比例|规模|数量|金额|均值|指数|得分|变化|增长|下降|产出|投入|变量|样本|回归|散点|柱状|折线|曲线|直方|箱线|热力|收入|利润|成本|资产|负债|人数|频率)', text or ''))


def nearby_text_for_indices(children, indices, window=3):
    if not indices:
        return ''
    start = max(0, min(indices) - window)
    end = min(len(children), max(indices) + window + 1)
    parts = []
    for child in children[start:end]:
        if child.tag == qn('p'):
            parts.append(paragraph_text(child))
    return ' '.join(p for p in parts if p)


def has_prior_figure_reference(children, start_idx, label):
    scanned = 0
    pat = re.compile(rf'图\s*{re.escape(label)}')
    idx = start_idx - 1
    while idx >= 0 and scanned < 8:
        child = children[idx]
        if child.tag != qn('p'):
            idx -= 1
            continue
        style = paragraph_style(child)
        text = paragraph_text(child)
        if style in {'Heading1', 'Heading2', 'Heading3', 'AbstractTitleCN', 'AbstractTitleEN', 'ReferencesHeading'}:
            break
        if not text or paragraph_has_drawing(child) or style in {'ImageCaption', 'TableCaption', 'SourceNote', 'TableMetaLine'}:
            idx -= 1
            continue
        scanned += 1
        compact = re.sub(r'\s+', '', text)
        if pat.search(compact) or ('如图' in compact and label in compact):
            return True
        idx -= 1
    return False


def research_output_has_detail_note(text):
    raw = text or ''
    bracket_parts = re.findall(r'[（(][^）)]{1,80}[）)]', raw)
    if not bracket_parts:
        return False
    joined = ' '.join(bracket_parts)
    return bool(re.search(r'(SCI|SSCI|CSSCI|EI|CSCD|北大核心|核心期刊|检索|收录|影响因子|IF\s*=|IF[:：]|项目|课题|基金|专利|奖项|获奖|会议|录用|发表|出版)', joined, flags=re.I))


def looks_like_research_output_item(text):
    raw = (text or '').strip()
    if not raw:
        return False
    if re.match(r'^(\[?\d+\]?|[一二三四五六七八九十]+[、.．])', raw):
        return True
    return bool(re.search(r'(\[J\]|\[C\]|\[D\]|\[M\]|\[P\]|论文|期刊|会议|项目|课题|专利|奖项|获奖|发表|录用|出版)', raw, flags=re.I))


def research_output_item_has_type_marker(text):
    return bool(re.search(r'(\[J\]|\[C\]|\[D\]|\[M\]|\[P\]|论文|期刊|会议|项目|课题|专利|奖项|获奖)', text or '', flags=re.I))


meta = load_simple_yaml(metadata_path)
degree_hint = ' '.join(
    part for part in [
        meta.get('degree_type', ''),
        meta.get('cover_degree_line', ''),
        meta.get('doctoral_apply_degree', ''),
    ] if part
)
if '博士' in degree_hint:
    degree_mode = 'doctor'
    cn_abstract_min = 2000
    cn_abstract_max = 3500
    en_abstract_min_words = 250
    body_char_min = 100000
elif '硕士' in degree_hint:
    degree_mode = 'master'
    cn_abstract_min = 800
    cn_abstract_max = 1800
    en_abstract_min_words = 120
    body_char_min = 30000
else:
    degree_mode = 'unknown'
    cn_abstract_min = 800
    cn_abstract_max = 3000
    en_abstract_min_words = 120
    body_char_min = None

with zipfile.ZipFile(path) as zf:
    data = zf.read('word/document.xml')
    styles_data = zf.read('word/styles.xml')
root = ET.fromstring(data)
styles_root = ET.fromstring(styles_data)
styles = style_map(styles_root)
text = ''.join(t.text or '' for t in root.iter(qn('t')))
hits = [needle for needle in needles if needle in text]
extra_failures = []
warnings = []
if '“”' in text:
    extra_failures.append('empty-quotes')
if '在哪里' in text:
    extra_failures.append('formula-where-mistranslation')
if '基金价值路径和增强价值路径' in text:
    missing = [label for label in key_labels if label not in text]
    if missing:
        extra_failures.append('missing-key-labels:' + ','.join(missing))

for style_id in ['Normal', 'BodyText', 'FirstParagraph', 'Bibliography', 'AbstractBodyCN', 'AbstractBodyEN', 'KeywordsLineCN', 'KeywordsLineEN', 'AcknowledgementsBody', 'StatementBody']:
    spacing = get_style_spacing(styles.get(style_id))
    if spacing is None:
        extra_failures.append(f'missing-style-spacing:{style_id}')
        continue
    if spacing.get(qn('line')) != '400' or spacing.get(qn('lineRule')) != 'exact':
        extra_failures.append(f'bad-style-spacing:{style_id}:{spacing.get(qn("line"))}:{spacing.get(qn("lineRule"))}')

spacing = get_style_spacing(styles.get('FootnoteText'))
if spacing is None or spacing.get(qn('line')) != '240' or spacing.get(qn('lineRule')) != 'auto':
    extra_failures.append('bad-style-spacing:FootnoteText')

spacing = get_style_spacing(styles.get('EquationBlock'))
if spacing is None or spacing.get(qn('before')) != '120' or spacing.get(qn('after')) != '120' or spacing.get(qn('line')) != '360' or spacing.get(qn('lineRule')) != 'auto':
    extra_failures.append('bad-style-spacing:EquationBlock')

if get_east_asia_font(styles.get('TableMetaLine')) != 'SimHei' or get_font_size(styles.get('TableMetaLine')) != '21':
    extra_failures.append('bad-style-font:TableMetaLine')
if get_east_asia_font(styles.get('SourceNote')) != 'SimSun' or get_font_size(styles.get('SourceNote')) != '18':
    extra_failures.append('bad-style-font:SourceNote')
if get_east_asia_font(styles.get('SignatureLine')) != 'SimHei' or get_font_size(styles.get('SignatureLine')) != '24':
    extra_failures.append('bad-style-font:SignatureLine')
if get_east_asia_font(styles.get('CoverDate')) != 'SimSun' or get_font_size(styles.get('CoverDate')) != '28' or style_is_bold(styles.get('CoverDate')):
    extra_failures.append('bad-style-font:CoverDate')

body = root.find(qn('body'))
children = list(body) if body is not None else []
cn_abstract_parts = []
en_abstract_parts = []
body_parts = []
cn_keywords = None
en_keywords = None
cn_title = None
research_outputs_found = False
research_output_items = []
in_cn_abstract = False
in_en_abstract = False
in_main_body = False
in_references = False
in_research_outputs = False
main_body_blank_run = 0
main_body_excessive_blank_reported = False

for child in children:
    if child.tag == qn('tbl'):
        if in_main_body:
            main_body_blank_run = 0
        if in_cn_abstract:
            extra_failures.append('abstract-contains-table:cn')
        if in_en_abstract:
            extra_failures.append('abstract-contains-table:en')
        continue
    if child.tag != qn('p'):
        continue

    p = child
    style = paragraph_style(p)
    ptext = paragraph_text(p)

    if in_main_body:
        if ptext == '' and not paragraph_has_drawing(p):
            main_body_blank_run += 1
            if main_body_blank_run >= 3 and not main_body_excessive_blank_reported:
                warnings.append(f'main_body_excessive_blank_paragraphs:{main_body_blank_run}')
                main_body_excessive_blank_reported = True
        else:
            main_body_blank_run = 0
        if ptext and not is_structural_page_break_style(style):
            if paragraph_has_manual_page_break(p):
                warnings.append('main_body_manual_page_break')
            if paragraph_has_page_break_before(p):
                warnings.append('main_body_page_break_before_nonheading')

    if cn_title is None and style == 'Title' and ptext:
        cn_title = ptext

    if style == 'AbstractTitleCN':
        in_cn_abstract = True
        in_en_abstract = False
        in_references = False
        continue
    if style == 'AbstractTitleEN':
        in_en_abstract = True
        in_cn_abstract = False
        in_references = False
        continue

    if style == 'ReferencesHeading' or ptext == '参考文献':
        in_references = True
        in_research_outputs = False
        in_main_body = False
        in_cn_abstract = False
        in_en_abstract = False
        continue

    if style == 'Heading1' and not in_main_body:
        in_main_body = True
    if is_backmatter_start(style):
        in_main_body = False
        in_cn_abstract = False
        in_en_abstract = False
        if style == 'ResearchOutputsHeading':
            research_outputs_found = True
            in_research_outputs = True
        else:
            in_research_outputs = False
        if style != 'ReferencesHeading':
            in_references = False

    if in_cn_abstract:
        if p.find('.//' + qn('drawing')) is not None:
            extra_failures.append('abstract-contains-drawing:cn')
        if p.find('.//' + M_NS + 'oMath') is not None or p.find('.//' + M_NS + 'oMathPara') is not None:
            extra_failures.append('abstract-contains-math:cn')
        if style == 'KeywordsLineCN':
            cn_keywords = ptext
            in_cn_abstract = False
        elif ptext:
            cn_abstract_parts.append(ptext)
        continue

    if in_en_abstract:
        if p.find('.//' + qn('drawing')) is not None:
            extra_failures.append('abstract-contains-drawing:en')
        if p.find('.//' + M_NS + 'oMath') is not None or p.find('.//' + M_NS + 'oMathPara') is not None:
            extra_failures.append('abstract-contains-math:en')
        if style == 'KeywordsLineEN':
            en_keywords = ptext
            in_en_abstract = False
        elif ptext:
            en_abstract_parts.append(ptext)
        continue

    if in_references and ptext:
        if style == 'Bibliography':
            if not looks_reference_line(ptext):
                warnings.append('references-bad-prefix')
            if not is_reference_line_ending_ok(ptext):
                warnings.append('references-bad-ending')
        elif style in {'AppendixHeading', 'ResearchOutputsHeading', 'AcknowledgementsHeading', 'Heading1'}:
            in_references = False

    if in_research_outputs and ptext:
        if style == 'ResearchOutputsHeading':
            pass
        elif style in {'AcknowledgementsHeading', 'AppendixHeading', 'ReferencesHeading', 'Heading1'}:
            in_research_outputs = False
        elif looks_like_research_output_item(ptext):
            research_output_items.append(ptext)

    if in_main_body and ptext:
        body_parts.append(ptext)

if cn_title:
    cn_title_len = len(re.sub(r'\s+', '', cn_title))
    if cn_title_len > 25:
        warnings.append(f'title_too_long:{cn_title_len}')

cn_abstract_text = ''.join(cn_abstract_parts)
en_abstract_text = ' '.join(en_abstract_parts)
cn_chars = chinese_char_count(cn_abstract_text)
if cn_chars and cn_chars < cn_abstract_min:
    warnings.append(f'cn_abstract_too_short:{cn_chars}:{degree_mode}')
if cn_chars and cn_chars > cn_abstract_max:
    warnings.append(f'cn_abstract_too_long:{cn_chars}:{degree_mode}')
if en_abstract_text:
    en_words = english_word_count(en_abstract_text)
    if en_words < en_abstract_min_words:
        warnings.append(f'en_abstract_too_short:{en_words}:{degree_mode}')
    if has_bad_en_abstract_punct(en_abstract_text):
        warnings.append('en_abstract_punct_bad')
    if has_bad_en_abstract_spacing(en_abstract_text):
        warnings.append('en_abstract_spacing_bad')

if cn_keywords:
    raw = re.sub(r'^关键词[:：]\s*', '', cn_keywords).strip()
    parts = [x.strip() for x in re.split(r'[，]+', raw) if x.strip()]
    if not (3 <= len(parts) <= 5):
        warnings.append(f'cn_keywords_count:{len(parts)}')
    if has_bad_cn_keyword_punct(raw):
        warnings.append('cn_keywords_punct_bad')
else:
    warnings.append('cn_keywords_missing')

if en_keywords:
    raw = re.sub(r'^(Key Words|Keywords)[:：]\s*', '', en_keywords, flags=re.I).strip()
    parts = [x.strip() for x in re.split(r';\s*', raw) if x.strip()]
    if not (3 <= len(parts) <= 5):
        warnings.append(f'en_keywords_count:{len(parts)}')
    if has_bad_en_keyword_punct(raw):
        warnings.append('en_keywords_punct_bad')
    if has_bad_en_keyword_spacing(raw):
        warnings.append('en_keywords_spacing_bad')
    if has_terminal_keyword_punct(raw):
        warnings.append('en_keywords_terminal_punct_bad')
else:
    warnings.append('en_keywords_missing')

if degree_mode == 'doctor':
    if not research_outputs_found:
        warnings.append('research_outputs_missing:doctor')
    elif not research_output_items:
        warnings.append('research_outputs_empty:doctor')

for idx, item in enumerate(research_output_items, 1):
    if not research_output_item_has_type_marker(item):
        warnings.append(f'research_output_type_marker_missing:{idx}')
    if not research_output_has_detail_note(item):
        warnings.append(f'research_output_detail_note_missing:{idx}')

for idx, child in enumerate(children):
    if child.tag != qn('p'):
        continue
    style = paragraph_style(child)
    if style == 'TableCaption':
        if not has_blank_before(children, idx):
            warnings.append('table_cluster_missing_blank_before')
        cluster_end = table_cluster_end_index(children, idx)
        if not has_blank_after(children, cluster_end):
            warnings.append('table_cluster_missing_blank_after')
    elif style == 'ImageCaption':
        if not has_blank_before(children, idx):
            warnings.append('image_cluster_missing_blank_before')
        cluster_end = image_caption_cluster_end_index(children, idx)
        if not has_blank_after(children, cluster_end):
            warnings.append('image_cluster_missing_blank_after')
        parsed = parse_numbered_caption(paragraph_text(child), '图')
        if parsed and not parsed['continued']:
            label = parsed['label']
            if not has_prior_figure_reference(children, idx, label):
                warnings.append(f'figure_text_reference_missing_before:{label}')
            drawing_indices = image_caption_drawing_indices(children, idx)
            nearby_text = nearby_text_for_indices(children, drawing_indices + [idx])
            if len(drawing_indices) >= 2 and not has_subfigure_label(nearby_text):
                warnings.append(f'subfigure_labels_missing:{label}')
            if text_suggests_measured_figure(parsed['title'] + ' ' + nearby_text) and not text_has_unit_marker(nearby_text):
                warnings.append(f'image_unit_unconfirmed:{label}')
    elif paragraph_has_drawing(child):
        prev_idx = prev_nonblank_child_index(children, idx - 1)
        if prev_idx is not None and children[prev_idx].tag == qn('p') and paragraph_style(children[prev_idx]) == 'ImageCaption':
            continue
        if not has_blank_before(children, idx):
            warnings.append('image_cluster_missing_blank_before')
        cluster_end = drawing_cluster_end_index(children, idx)
        if not has_blank_after(children, cluster_end):
            warnings.append('image_cluster_missing_blank_after')

seen_table_captions = set()
seen_image_captions = set()
for child in children:
    if child.tag != qn('p'):
        continue
    style = paragraph_style(child)
    if style == 'TableCaption':
        parsed = parse_numbered_caption(paragraph_text(child), '表')
        if not parsed or not parsed['title']:
            continue
        key = (parsed['scope'], parsed['title'])
        if key in seen_table_captions and not parsed['continued']:
            warnings.append('continued_table_caption_missing')
        seen_table_captions.add(key)
    elif style == 'ImageCaption':
        parsed = parse_numbered_caption(paragraph_text(child), '图')
        if not parsed or not parsed['title']:
            continue
        key = (parsed['scope'], parsed['title'])
        if key in seen_image_captions and not parsed['continued']:
            warnings.append('continued_image_caption_missing')
        seen_image_captions.add(key)

for idx, child in enumerate(children):
    if child.tag != qn('p'):
        continue
    if paragraph_style(child) != 'TableCaption':
        continue
    caption_text = paragraph_text(child).strip()
    if not caption_text.startswith('续表'):
        continue
    next_tbl = None
    for probe in children[idx + 1:]:
        if probe.tag == qn('tbl'):
            next_tbl = probe
            break
        if probe.tag == qn('p') and paragraph_text(probe).strip() != '':
            break
    if next_tbl is not None and not first_row_has_tbl_header(next_tbl):
        warnings.append('continued_table_header_missing')

if body_char_min is not None:
    body_chars = chinese_char_count(''.join(body_parts))
    if body_chars < body_char_min:
        warnings.append(f'body_char_count_too_short:{body_chars}:{degree_mode}')

if hits or extra_failures:
    print(f'DOCX_SCAN_FAILED: {path}')
    for needle in hits:
        print(f'  - matched: {needle}')
    for item in extra_failures:
        print(f'  - matched: {item}')
    sys.exit(12)

for item in warnings:
    print(f'DOCX_SCAN_WARN: {item}', file=sys.stderr)
PY
}

if [ "$SCAN_ONLY" = "1" ]; then
  scan_docx_or_fail "$SCAN_DOCX_PATH" "$METADATA"
  exit 0
fi

mkdir -p "$(dirname "$OUTPUT")"

if [ ! -f "$REFDOC" ] || [ "$BUILD_REF" -nt "$REFDOC" ]; then
  /opt/homebrew/bin/python3 "$BUILD_REF"
fi

BIB_PATH="$(resolve_bibliography_path || true)"
if [ -n "$BIB_PATH" ]; then
  scan_bibliography_or_fail "$BIB_PATH" "$METADATA"
fi

PREPROCESS_ARGS=("$INPUT" "$TMP_TEX")
if [ -n "$METADATA" ]; then
  PREPROCESS_ARGS+=(--metadata "$METADATA")
fi
/opt/homebrew/bin/python3 "$PREPROCESS_TEX" "${PREPROCESS_ARGS[@]}"

PANDOC_ARGS=(
  "$TMP_TEX"
  --from=latex+raw_tex
  --to=docx
  --standalone
  --reference-doc="$REFDOC"
  --lua-filter="$FILTER"
  --resource-path="$(dirname "$INPUT"):$PWD"
  -o "$TMP_DOCX"
)

if [ -n "$METADATA" ]; then
  PANDOC_ARGS+=(--metadata-file="$METADATA")
fi

if [ "${#EXTRA_ARGS[@]}" -gt 0 ]; then
  PANDOC_ARGS+=("${EXTRA_ARGS[@]}")
fi

/opt/homebrew/bin/pandoc "${PANDOC_ARGS[@]}"
/opt/homebrew/bin/python3 "$POSTPROCESS" "$TMP_DOCX" "$OUTPUT"
scan_docx_or_fail "$OUTPUT" "$METADATA"

echo "$OUTPUT"
