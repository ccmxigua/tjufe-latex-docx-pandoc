#!/opt/homebrew/bin/python3
from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple


XREF_MARKER_FMT = '[[[TJUFE_XREF:{label}|{text}]]]'
EQ_LABEL_MARKER_FMT = '[[[TJUFE_EQLABEL:{label}]]]'
LABEL_MARKER_FMT = 'TJUFE_LABEL__{label}__'
EQUATION_ENVS = (
    'equation', 'equation*',
    'align', 'align*',
    'gather', 'gather*',
    'multline', 'multline*',
    'eqnarray', 'eqnarray*',
)
THEOREM_ENVS = (
    'theorem', 'lemma', 'proposition', 'corollary', 'definition', 'remark', 'assumption',
)

AUX_NEWLABEL_RE = re.compile(r'^\\newlabel\{([^}]+)\}(.*)$')
AUX_ZREF_LABEL_RE = re.compile(r'^\\zref@newlabel\{([^}]+)\}\{(.*)\}$')
AUX_BIBCITE_RE = re.compile(r'^\\bibcite\{([^}]+)\}(.*)$')
GENERIC_LABEL_RE = re.compile(r'\\label\{([^{}]+)\}')
REF_PATTERN = re.compile(r'\\(zcref|cref|Cref|autoref|ref|eqref)\{([^{}]+)\}')
CITE_PATTERN = re.compile(r'\\(cite|citep|citet|citeauthor|citeyear)\{([^{}]+)\}')


LabelInfo = Dict[str, str]


def extract_balanced(text: str, start: int, open_char: str, close_char: str) -> Tuple[str, int]:
    if start >= len(text) or text[start] != open_char:
        raise ValueError(f"expected {open_char!r} at {start}")
    depth = 0
    i = start
    while i < len(text):
        ch = text[i]
        if ch == '\\':
            i += 2
            continue
        if ch == open_char:
            depth += 1
        elif ch == close_char:
            depth -= 1
            if depth == 0:
                return text[start:i + 1], i + 1
        i += 1
    raise ValueError(f"unbalanced {open_char}{close_char} starting at {start}")


def top_level_groups(text: str) -> List[str]:
    groups: List[str] = []
    i = 0
    while i < len(text):
        while i < len(text) and text[i].isspace():
            i += 1
        if i >= len(text) or text[i] != '{':
            break
        group, i = extract_balanced(text, i, '{', '}')
        groups.append(group[1:-1])
    return groups


def strip_outer_braces(s: str) -> str:
    s = s.strip()
    while len(s) >= 2 and s[0] == '{' and s[-1] == '}':
        try:
            group, end = extract_balanced(s, 0, '{', '}')
        except ValueError:
            break
        if end != len(s):
            break
        s = group[1:-1].strip()
    return s


def unwrap_subfloat(text: str) -> str:
    out: List[str] = []
    i = 0
    needle = '\\subfloat'
    while True:
        idx = text.find(needle, i)
        if idx == -1:
            out.append(text[i:])
            break
        out.append(text[i:idx])
        j = idx + len(needle)
        while j < len(text) and text[j].isspace():
            j += 1
        if j < len(text) and text[j] == '[':
            try:
                _, j = extract_balanced(text, j, '[', ']')
            except ValueError:
                out.append(needle)
                i = idx + len(needle)
                continue
            while j < len(text) and text[j].isspace():
                j += 1
        if j >= len(text) or text[j] != '{':
            out.append(needle)
            i = idx + len(needle)
            continue
        try:
            body, j = extract_balanced(text, j, '{', '}')
        except ValueError:
            out.append(needle)
            i = idx + len(needle)
            continue
        out.append(body[1:-1].strip())
        i = j
    return ''.join(out)


def replace_braced_macro(text: str, macro: str, repl) -> str:
    out: List[str] = []
    i = 0
    needle = '\\' + macro
    while True:
        idx = text.find(needle, i)
        if idx == -1:
            out.append(text[i:])
            break
        out.append(text[i:idx])
        j = idx + len(needle)
        while j < len(text) and text[j].isspace():
            j += 1
        if j >= len(text) or text[j] != '{':
            out.append(needle)
            i = idx + len(needle)
            continue
        try:
            group, j = extract_balanced(text, j, '{', '}')
        except ValueError:
            out.append(needle)
            i = idx + len(needle)
            continue
        out.append(repl(group[1:-1]))
        i = j
    return ''.join(out)


def normalize_math_macros(text: str) -> str:
    text = replace_braced_macro(text, 'widebar', lambda body: f'\\overline{{{body}}}')
    text = replace_braced_macro(text, 'textup', lambda body: f'\\mathrm{{{body}}}')
    text = replace_braced_macro(text, 'mathclap', lambda body: body)

    def normalize_overline(body: str) -> str:
        stripped = body.strip()
        stripped = re.sub(r'^\\rule\{0pt\}\{2\.5 ?ex\}\s*', '', stripped)
        if stripped != body.strip():
            return stripped
        return f'\\overline{{{body}}}'

    text = replace_braced_macro(text, 'overline', normalize_overline)
    return text


def normalize_docx_export_text(text: str) -> str:
    """Normalize fragile prose patterns for DOCX export only.

    Goals:
    1. Formula lead-in lines like a standalone "where" / "在哪里" should become "其中".
    2. Quoted \text{...} labels used in prose should be flattened so Pandoc DOCX
       conversion does not drop the quoted content.
    """

    # Academic formula lead-in: a standalone line after an equation should not remain
    # as the literal English where / mistranslated 在哪里.
    text = re.sub(r'(?m)^[ \t]*(?:where|Where|在哪里)[ \t]*$', '其中', text)

    # DOCX-export-safe flattening for quoted prose labels such as “\text{U\_BS}”.
    # Keep the quote marks and preserve escaped underscores for LaTeX text mode.
    def flatten_quoted_text_macro(match: re.Match[str]) -> str:
        open_q = match.group(1)
        body = match.group(2).strip()
        close_q = match.group(3)
        return f'{open_q}{body}{close_q}'

    text = re.sub(r'([“"])[ \t]*\\text\{([^{}]+)\}[ \t]*([”"])', flatten_quoted_text_macro, text)
    return text


def strip_decorative_horizontal_rules(text: str) -> str:
    text = re.sub(
        r'\s*\\begin\{center\}\s*\\rule\{[^{}]+\}\{0\.2\s*mm\}\s*\\end\{center\}\s*',
        '\n\n',
        text,
        flags=re.S,
    )
    return re.sub(r'\n{3,}', '\n\n', text)


def parse_newlabel_line(line: str) -> Optional[Tuple[str, List[str]]]:
    m = AUX_NEWLABEL_RE.match(line.strip())
    if not m:
        return None
    label = m.group(1)
    rest = m.group(2).lstrip()
    if not rest or rest[0] != '{':
        return None
    try:
        payload, _ = extract_balanced(rest, 0, '{', '}')
    except ValueError:
        return None
    return label, top_level_groups(payload[1:-1])


def extract_zref_field(payload: str, name: str) -> str:
    m = re.search(rf'\\{re.escape(name)}\{{([^{{}}]*)\}}', payload)
    return m.group(1).strip() if m else ''


def parse_zref_label_line(line: str) -> Optional[Tuple[str, LabelInfo]]:
    m = AUX_ZREF_LABEL_RE.match(line.strip())
    if not m:
        return None
    label = m.group(1)
    payload = m.group(2)
    number = extract_zref_field(payload, 'thecounter') or extract_zref_field(payload, 'default')
    kind = extract_zref_field(payload, 'zc@type') or extract_zref_field(payload, 'zc@counter')
    page = extract_zref_field(payload, 'page') or extract_zref_field(payload, 'zc@pgfmt')
    return label, {
        'number': number,
        'page': page,
        'title': '',
        'kind': kind,
    }


def parse_bibcite_line(line: str) -> Optional[Tuple[str, LabelInfo]]:
    m = AUX_BIBCITE_RE.match(line.strip())
    if not m:
        return None
    label = m.group(1)
    rest = m.group(2).lstrip()
    if not rest or rest[0] != '{':
        return None
    try:
        payload, _ = extract_balanced(rest, 0, '{', '}')
    except ValueError:
        return None
    fields = top_level_groups(payload[1:-1])
    number = strip_outer_braces(fields[0]) if len(fields) >= 1 else ''
    year = strip_outer_braces(fields[1]) if len(fields) >= 2 else ''
    author = strip_outer_braces(fields[2]) if len(fields) >= 3 else ''
    return label, {
        'number': number,
        'year': year,
        'author': author,
        'kind': 'bibitem',
    }


def count_newlabels(path: Path) -> int:
    try:
        text = path.read_text(errors='ignore')
    except OSError:
        return 0
    count = 0
    for line in text.splitlines():
        stripped = line.lstrip()
        if stripped.startswith('\\newlabel{') or stripped.startswith('\\zref@newlabel{') or stripped.startswith('\\bibcite{'):
            count += 1
    return count


def common_prefix_len(a: str, b: str) -> int:
    size = min(len(a), len(b))
    i = 0
    while i < size and a[i] == b[i]:
        i += 1
    return i


def discover_aux(input_path: Path, explicit_aux: Optional[Path]) -> Optional[Path]:
    if explicit_aux:
        return explicit_aux if explicit_aux.exists() else None

    candidates: List[Path] = []
    for suffix in ('.aux', '.aux.bak'):
        p = input_path.with_suffix(suffix)
        if p.exists():
            candidates.append(p)

    current = input_path.parent
    seen = {p.resolve() for p in candidates if p.exists()}
    for _ in range(4):
        try:
            for cand in current.glob('*.aux*'):
                if cand.is_file():
                    resolved = cand.resolve()
                    if resolved not in seen:
                        candidates.append(cand)
                        seen.add(resolved)
        except OSError:
            pass
        if current.parent == current:
            break
        current = current.parent

    if not candidates:
        return None

    def score(path: Path) -> Tuple[int, int, int, int, str]:
        stem_match = int(path.stem == input_path.stem)
        prefix_len = common_prefix_len(path.stem, input_path.stem)
        label_count = count_newlabels(path)
        not_backup = int(not str(path).endswith('.bak'))
        return (stem_match, prefix_len, label_count, not_backup, str(path))

    return max(candidates, key=score)


def discover_metadata(input_path: Path, explicit_metadata: Optional[Path] = None) -> Optional[Path]:
    if explicit_metadata and explicit_metadata.exists():
        return explicit_metadata.resolve()
    input_dir = input_path.parent
    candidates = [
        input_dir / 'metadata.yaml',
        input_dir / 'metadata.yml',
        input_dir / f'{input_path.stem}.yaml',
        input_dir / f'{input_path.stem}.yml',
    ]
    for cand in candidates:
        if cand.exists():
            return cand
    return None


def split_top_level_commas(text: str) -> List[str]:
    parts: List[str] = []
    buf: List[str] = []
    depth = 0
    in_quote = False
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == '\\':
            if i + 1 < len(text):
                buf.append(text[i:i + 2])
                i += 2
                continue
        if ch == '"':
            in_quote = not in_quote
            buf.append(ch)
            i += 1
            continue
        if not in_quote:
            if ch == '{':
                depth += 1
            elif ch == '}':
                depth = max(0, depth - 1)
            elif ch == ',' and depth == 0:
                part = ''.join(buf).strip()
                if part:
                    parts.append(part)
                buf = []
                i += 1
                continue
        buf.append(ch)
        i += 1
    tail = ''.join(buf).strip()
    if tail:
        parts.append(tail)
    return parts


def split_top_level_once(text: str, needle: str) -> Tuple[str, str]:
    depth = 0
    in_quote = False
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == '\\':
            i += 2
            continue
        if ch == '"':
            in_quote = not in_quote
            i += 1
            continue
        if not in_quote:
            if ch == '{':
                depth += 1
            elif ch == '}':
                depth = max(0, depth - 1)
            elif text.startswith(needle, i) and depth == 0:
                return text[:i], text[i + len(needle):]
        i += 1
    return text, ''


def parse_metadata_bibliography_paths(metadata_path: Path) -> List[Path]:
    try:
        text = metadata_path.read_text(errors='ignore')
    except OSError:
        return []
    results: List[Path] = []
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped or stripped.startswith('#'):
            i += 1
            continue
        m = re.match(r'^bibliography\s*:\s*(.+?)\s*$', stripped)
        if m:
            value = m.group(1).strip().strip('"\'')
            if value and value not in {'[]', '{}'}:
                cand = (metadata_path.parent / value).resolve()
                if cand.exists():
                    results.append(cand)
            i += 1
            continue
        if re.match(r'^bibliography\s*:\s*$', stripped):
            i += 1
            while i < len(lines):
                sub = lines[i]
                if not sub.strip():
                    i += 1
                    continue
                if not re.match(r'^\s+-\s+', sub):
                    break
                value = re.sub(r'^\s+-\s+', '', sub).strip().strip('"\'')
                cand = (metadata_path.parent / value).resolve()
                if cand.exists():
                    results.append(cand)
                i += 1
            continue
        i += 1
    return results


def discover_bibliography_files(input_path: Path, text: str, explicit_metadata: Optional[Path] = None) -> List[Path]:
    results: List[Path] = []
    seen = set()

    def add_path(raw: str, base: Path) -> None:
        raw = raw.strip().strip('"\'')
        if not raw:
            return
        cand = (base / raw).resolve()
        if cand.exists() and cand not in seen:
            seen.add(cand)
            results.append(cand)

    metadata_path = discover_metadata(input_path, explicit_metadata)
    if metadata_path:
        for path in parse_metadata_bibliography_paths(metadata_path):
            if path not in seen:
                seen.add(path)
                results.append(path)

    for m in re.finditer(r'\\addbibresource\{([^{}]+)\}', text):
        add_path(m.group(1), input_path.parent)
    for m in re.finditer(r'\\bibliography\{([^{}]+)\}', text):
        for item in m.group(1).split(','):
            item = item.strip()
            if not item:
                continue
            if not item.lower().endswith('.bib'):
                item += '.bib'
            add_path(item, input_path.parent)
    return results


def extract_cited_keys(text: str) -> List[str]:
    keys: List[str] = []
    seen = set()
    for m in CITE_PATTERN.finditer(text):
        for item in m.group(2).split(','):
            key = item.strip()
            if key and key not in seen:
                seen.add(key)
                keys.append(key)
    return keys


def looks_like_equation_label(label: str) -> bool:
    lower = label.lower()
    return (
        lower.startswith('eq:')
        or lower.startswith('eq_')
        or lower.startswith('eq-')
        or lower.startswith('equation:')
        or lower.startswith('equation_')
        or lower.startswith('equation-')
        or ':eq:' in lower
        or ':eq_' in lower
        or ':eq-' in lower
        or '.eq.' in lower
        or '.eq_' in lower
        or '.eq-' in lower
        or '_eq_' in lower
        or '-eq-' in lower
    )


def humanize_bib_key(key: str) -> str:
    text = key.replace('_', ' ').replace('-', ' ').replace('.', ' ')
    text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
    text = re.sub(r'([A-Za-z])([0-9]{4})', r'\1 \2', text)
    text = re.sub(r'([0-9])([A-Za-z])', r'\1 \2', text)
    text = re.sub(r'\s+', ' ', text).strip()
    if not text:
        return key
    parts = []
    for tok in text.split():
        if tok.isupper() or tok.isdigit():
            parts.append(tok)
        else:
            parts.append(tok[0].upper() + tok[1:] if tok else tok)
    return ' '.join(parts)


def parse_bibtex_author_display(raw: str) -> Tuple[str, str]:
    raw = strip_outer_braces(raw).strip()
    if not raw:
        return '', ''
    names = [clean_bib_value(strip_outer_braces(x.strip())).strip() for x in re.split(r'\s+and\s+', raw) if x.strip()]
    full = '; '.join(names)

    def short_name(name: str) -> str:
        name = clean_bib_value(strip_outer_braces(name)).strip()
        if ',' in name:
            return name.split(',', 1)[0].strip()
        bits = name.split()
        return bits[-1].strip() if bits else name.strip()

    short_names = [short_name(x) for x in names]
    if not short_names:
        return full, ''
    if len(short_names) == 1:
        cite = short_names[0]
    elif len(short_names) == 2:
        cite = f'{short_names[0]} and {short_names[1]}'
    else:
        cite = f'{short_names[0]} et al.'
    return full, cite


def parse_bibtex_entries(text: str) -> Dict[str, Dict[str, str]]:
    entries: Dict[str, Dict[str, str]] = {}
    i = 0
    while i < len(text):
        at = text.find('@', i)
        if at == -1:
            break
        brace = text.find('{', at)
        if brace == -1:
            break
        entry_type = text[at + 1:brace].strip().lower()
        try:
            payload, end = extract_balanced(text, brace, '{', '}')
        except ValueError:
            i = brace + 1
            continue
        inner = payload[1:-1].strip()
        head, rest = split_top_level_once(inner, ',')
        key = head.strip()
        fields: Dict[str, str] = {'ENTRYTYPE': entry_type}
        if key:
            fields['ID'] = key
            for chunk in split_top_level_commas(rest):
                left, right = split_top_level_once(chunk, '=')
                name = left.strip().lower()
                value = right.strip()
                if not name or not value:
                    continue
                value = strip_outer_braces(value.strip().strip('"'))
                fields[name] = value
            entries[key] = fields
        i = end
    return entries


def bibtex_entry_body(fields: Dict[str, str], key: str) -> Tuple[str, str, str, str]:
    title = fields.get('title', '').strip() or humanize_bib_key(key)
    title = re.sub(r'\s+', ' ', title).strip()
    year = fields.get('year', '').strip()
    if not year:
        date = fields.get('date', '').strip()
        m = re.search(r'(19|20)\d{2}', date)
        year = m.group(0) if m else ''
    author_full, author_cite = parse_bibtex_author_display(fields.get('author', ''))
    if not author_full:
        org = fields.get('organization', '').strip() or fields.get('institution', '').strip() or fields.get('publisher', '').strip()
        author_full = org or ''
        author_cite = org or humanize_bib_key(key)
    container = fields.get('journaltitle', '').strip() or fields.get('journal', '').strip() or fields.get('booktitle', '').strip() or fields.get('howpublished', '').strip() or fields.get('note', '').strip()
    url = fields.get('url', '').strip()
    parts: List[str] = []
    if author_full:
        parts.append(f'{author_full}.')
    if year:
        parts.append(f'({year}).')
    if title:
        parts.append(f'{title}.')
    if container:
        parts.append(f'{container}.')
    if url:
        parts.append(url)
    body = ' '.join(part.strip() for part in parts if part.strip())
    return body.strip(), author_cite.strip(), year.strip(), title


def load_bib_fallback_entries(input_path: Path, text: str, cited_keys: List[str], explicit_metadata: Optional[Path] = None) -> Dict[str, LabelInfo]:
    alias_map = {
        'reference.wolfram_2023_leastsquares': 'wolframresearch2021leastsquares',
        'ccmxigua': 'ccmxigua2',
        'li2012implied': 'li2012implieda',
        'yun2025valuing': 'mei2025solving',
        'scott1997pricing': 'black1973pricing',
        'jeanblanc-picque2009mathematical': 'arfken2013mathematical',
        'nakamura2013crises': 'tanaka2017time',
        'zijianmei2025solving': 'mei2025solving',
        'gilli2011calibrating': 'gilli2011calibrating',
        'duffie2003affine': 'duffy2006finite',
        'li2022deep': 'motameni2022lookback',
        'fu2022unsupervised': 'tian2020european',
        'rouah2013heston': 'jonathanc.frei2015heston',
        'carr2007stochastic': 'shreve2010stochastic',
        'trolle2009unspanned': 'trolle2009unspanned',
        'janek2011fx': 'janek2011fx',
        'sandovalheston': 'sandovalheston',
        'jin2016pricing': 'wang2012pricing',
        'imai2001dynamic': 'wong2007lookback',
        'yin-poole2023super': 'yin-poole2023super',
        'salesWonder': 'salesWonder',
        'resale': 'resale',
        'goofish': 'goofish',
        'elden': 'elden',
        'marvels': 'marvels',
        'google2025change': 'google2025change',
        'apple2025manage': 'apple2025manage',
        'chathamfinancialsecurities2025fx': 'chathamfinancialsecurities2025fx',
        'protter2005stochastic': 'shreve2010stochastic',
        'bao2011convergence': 'bao2011convergence',
        'billingsley1999convergence': 'billingsley1999convergence',
        'mao2011stochastic': 'shreve2010stochastic',
        'yongsik2013comparison': 'yongsik2013comparison',
        'haug2007complete': 'haug2007complete',
        'bos2000how': 'bos2000how',
        'lemaire2022stationary': 'leung2013analytic',
        'DAX30': 'DAX30',
        'SPX': 'SPX',
        'DJIA': 'DJIA',
        'lmmethod': 'lmmethod',
        'inai2001dynamic': 'inai2001dynamic',
    }
    results: Dict[str, LabelInfo] = {}
    cited_set = set(cited_keys)
    for bib_path in discover_bibliography_files(input_path, text, explicit_metadata):
        try:
            bib_text = bib_path.read_text(errors='ignore')
        except OSError:
            continue
        parsed = parse_bibtex_entries(bib_text)
        for key in cited_keys:
            src_key = key if key in parsed else alias_map.get(key, '')
            if not src_key or src_key not in parsed or key in results:
                continue
            body, author_cite, year, title = bibtex_entry_body(parsed[src_key], key)
            results[key] = {
                'number': '',
                'page': '',
                'title': title,
                'kind': 'bibitem',
                'author': author_cite,
                'year': year,
                'body': body,
            }
        if cited_set.issubset(set(results.keys())):
            break
    return results


def augment_labels_from_bib(input_path: Path, text: str, labels: Dict[str, LabelInfo], cited_keys: List[str], explicit_metadata: Optional[Path] = None) -> None:
    fallback = load_bib_fallback_entries(input_path, text, cited_keys, explicit_metadata)
    for key, info in fallback.items():
        labels.setdefault(key, {}).update({k: v for k, v in info.items() if v})
        anchor = bibliography_anchor_label(key)
        labels.setdefault(anchor, {}).update({k: v for k, v in info.items() if v})


def load_aux_labels(aux_path: Optional[Path]) -> Dict[str, LabelInfo]:
    labels: Dict[str, LabelInfo] = {}
    if not aux_path or not aux_path.exists():
        return labels
    for line in aux_path.read_text(errors='ignore').splitlines():
        parsed = parse_newlabel_line(line)
        if parsed:
            label, fields = parsed
            if len(fields) >= 4:
                current = labels.setdefault(label, {'number': '', 'page': '', 'title': '', 'kind': ''})
                updates = {
                    'number': fields[0].strip(),
                    'page': fields[1].strip(),
                    'title': fields[2].strip(),
                    'kind': fields[3].strip(),
                }
                for key, value in updates.items():
                    if not value:
                        continue
                    if key == 'kind':
                        if not current.get(key):
                            current[key] = value
                    elif not current.get(key):
                        current[key] = value
            continue
        zparsed = parse_zref_label_line(line)
        if zparsed:
            label, info = zparsed
            current = labels.setdefault(label, {'number': '', 'page': '', 'title': '', 'kind': ''})
            for key, value in info.items():
                if not value:
                    continue
                if key == 'kind':
                    current[key] = value
                elif not current.get(key):
                    current[key] = value
            continue
        bparsed = parse_bibcite_line(line)
        if bparsed:
            label, info = bparsed
            current = labels.setdefault(label, {'number': '', 'page': '', 'title': '', 'kind': ''})
            for key, value in info.items():
                if value and not current.get(key):
                    current[key] = value
    return labels


THEOREM_ENVS_FOR_SCAN = ('theorem', 'lemma', 'proposition', 'corollary', 'definition', 'remark', 'assumption')


def scan_tex_for_missing_labels(tex_path: Path, labels: Dict[str, LabelInfo]) -> None:
    """Scan tex file for labels not found in aux and infer their type/kind."""
    text = tex_path.read_text(errors='ignore')

    eq_counter: Dict[str, int] = {}
    theorem_counters: Dict[str, int] = {}
    section_counters: Dict[str, int] = {}

    def get_chapter_prefix() -> str:
        last_chapter = section_counters.get('chapter', 0)
        if last_chapter > 0:
            return str(last_chapter)
        return ''

    for m in re.finditer(r'\\(chapter|section|subsection|subsubsection)\*?\{([^}]+)\}', text):
        sec_type = m.group(1)
        if sec_type == 'chapter':
            section_counters['chapter'] = section_counters.get('chapter', 0) + 1
            section_counters['section'] = 0
            section_counters['subsection'] = 0
        elif sec_type == 'section':
            section_counters['section'] = section_counters.get('section', 0) + 1
            section_counters['subsection'] = 0
        elif sec_type == 'subsection':
            section_counters['subsection'] = section_counters.get('subsection', 0) + 1

    eq_num = 0
    for m in re.finditer(r'\\begin\{(equation|align|gather|multline|eqnarray)(\*)?\}(.*?)\\end\{\1\2?\}', text, re.S):
        env = m.group(1)
        body = m.group(3)
        labels_in_block = re.findall(r'\\label\{([^{}]+)\}', body)
        if not labels_in_block:
            continue
        eq_num += 1
        chapter_prefix = get_chapter_prefix()
        if chapter_prefix:
            eq_label = f'{chapter_prefix}.{eq_num}'
        else:
            eq_label = str(eq_num)
        for label in labels_in_block:
            if label not in labels:
                labels[label] = {
                    'number': eq_label,
                    'page': '',
                    'title': '',
                    'kind': 'equation',
                }

    for env in THEOREM_ENVS_FOR_SCAN:
        counter_key = env
        theorem_counters[counter_key] = 0
        pattern = re.compile(rf'\\begin\{{{re.escape(env)}\}}(?:\[([^\]]+)\])?\s*(?:\\label\{{([^{{}}]+)\}})?', re.S)
        for m in pattern.finditer(text):
            theorem_counters[counter_key] += 1
            opt_title = m.group(1) or ''
            label = m.group(2) or ''
            if label and label not in labels:
                chapter_prefix = get_chapter_prefix()
                if chapter_prefix:
                    thm_num = f'{chapter_prefix}.{theorem_counters[counter_key]}'
                else:
                    thm_num = str(theorem_counters[counter_key])
                kind_map = {
                    'theorem': 'theorem',
                    'lemma': 'lemma',
                    'proposition': 'proposition',
                    'corollary': 'corollary',
                    'definition': 'definition',
                    'remark': 'remark',
                    'assumption': 'assumption',
                }
                labels[label] = {
                    'number': thm_num,
                    'page': '',
                    'title': opt_title,
                    'kind': kind_map.get(env, env),
                }


def classify_label(label: str, info: Optional[LabelInfo]) -> str:
    lower = label.lower()
    kind = ((info or {}).get('kind', '') or '').lower()
    if kind.startswith('equation'):
        return 'Equation'
    if kind.startswith('figure'):
        return '图'
    if kind.startswith('table'):
        return '表'
    if kind.startswith('section') or kind.startswith('subsection') or kind.startswith('subsubsection'):
        return '节'
    if kind.startswith('appendix'):
        return '附录'
    if kind.startswith('lemma'):
        return '引理'
    if kind.startswith('corollary'):
        return '推论'
    if kind.startswith('proposition'):
        return '命题'
    if kind.startswith('definition'):
        return '定义'
    if kind.startswith('remark'):
        return '注'
    if kind.startswith('assumption'):
        return '假设'
    if kind.startswith('theorem'):
        return '定理'
    if kind.startswith('bibitem'):
        return 'Reference'
    if 'lemma' in lower:
        return '引理'
    if 'corollary' in lower:
        return '推论'
    if 'prop' in lower or 'proposition' in lower:
        return '命题'
    if 'definition' in lower:
        return '定义'
    if 'remark' in lower:
        return '注'
    if 'assumption' in lower:
        return '假设'
    return 'Reference'


def resolve_label(label: str, labels: Dict[str, LabelInfo], mode: str) -> str:
    info = labels.get(label)
    number = (info or {}).get('number', '').strip()
    prefix = classify_label(label, info)

    if mode == 'ref':
        if prefix == 'Equation' and number:
            return f'（{number}）'
        return number or f'[{label}]'
    if mode == 'eqref':
        return f'（{number}）' if number else f'[{label}]'
    if mode in {'cite', 'citep'}:
        return number or label
    if mode == 'citet':
        author = (info or {}).get('author', '').strip()
        if author and number:
            return f'{author} {number}'
        return author or number or label
    if mode == 'citeauthor':
        return (info or {}).get('author', '').strip() or label
    if mode == 'citeyear':
        return (info or {}).get('year', '').strip() or label
    if mode == 'zcref' and prefix == 'Equation':
        return f'（{number}）' if number else f'[{label}]'
    if number:
        if prefix == '节':
            return f'第{number}节'
        if prefix == '附录':
            return f'附录{number}'
        if prefix in {'图', '表', '定理', '引理', '推论', '命题'}:
            return f'{prefix}{number}'
        return f'{prefix} {number}'
    return f'{prefix} [{label}]'


def make_xref_placeholder(label: str, text: str) -> str:
    safe_text = text.replace(']', ')')
    return XREF_MARKER_FMT.format(label=label, text=safe_text)


def bibliography_anchor_label(key: str) -> str:
    return key if key.startswith('ref-') else f'ref-{key}'


def top_level_brace_groups_relaxed(text: str) -> List[str]:
    groups: List[str] = []
    i = 0
    while i < len(text):
        if text[i] != '{':
            i += 1
            continue
        group, i = extract_balanced(text, i, '{', '}')
        groups.append(group[1:-1])
    return groups


def normalize_page_range(text: str) -> str:
    return re.sub(r'(\d)\s*[-–—]\s+(\d)', r'\1-\2', text)


def finalize_bibliography_text(text: str) -> str:
    text = text.replace('\xa0', ' ')
    text = normalize_page_range(text)
    text = re.sub(r'\s+([;:])(?=\s|[A-Za-z])', r'\1', text)
    text = re.sub(r'\s+,', ',', text)
    text = re.sub(r'\s+\.', '.', text)
    text = re.sub(r'\.\s*\.', '.', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def clean_bib_value(text: str) -> str:
    text = text.strip()
    text = re.sub(r'%.*', '', text)
    replacements = {
        '\\newblock': ' ',
        '\\bibrangedash': '-',
        '\\bibnamedelima': ' ',
        '\\bibnamedelimd': ' ',
        '\\bibnamedelimi': ' ',
        '\\bibpagespunct': ' ',
        '\\bibinitperiod': '.',
        '\\bibinitdelim': ' ',
        '\\addspace': ' ',
        '\\addcomma': ',',
        '\\addperiod': '.',
        '\\addcolon': ':',
        '\\&': '&',
        '~': ' ',
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r'\\mkbib(?:emph|quote|parens)\{([^{}]*)\}', r'\1', text)
    text = re.sub(r'\\(emph|textit|textbf|url|doi|natexlab)\{([^{}]*)\}', r'\2', text)
    text = re.sub(r'\\href\{([^{}]*)\}\{([^{}]*)\}', r'\2', text)
    text = re.sub(r'\\[A-Za-z@]+\*?(?:\[[^\]]*\])?', ' ', text)
    text = text.replace('{', '').replace('}', '')
    return finalize_bibliography_text(text)


def extract_assignment_value(text: str, key: str) -> str:
    m = re.search(rf'{re.escape(key)}\s*=\s*', text)
    if not m:
        return ''
    start = m.end()
    if start >= len(text) or text[start] != '{':
        return ''
    try:
        group, _ = extract_balanced(text, start, '{', '}')
    except ValueError:
        return ''
    return clean_bib_value(group[1:-1])


def extract_biblatex_names(entry_text: str, role: str = 'author') -> List[Tuple[str, str]]:
    m = re.search(rf'\\name\{{{re.escape(role)}\}}\{{[^{{}}]*\}}\{{[^{{}}]*\}}\{{', entry_text)
    if not m:
        return []
    try:
        group, _ = extract_balanced(entry_text, m.end() - 1, '{', '}')
    except ValueError:
        return []
    names: List[Tuple[str, str]] = []
    for person in top_level_brace_groups_relaxed(group[1:-1]):
        parts = top_level_brace_groups_relaxed(person)
        detail = parts[-1] if parts else person
        family = extract_assignment_value(detail, 'family')
        given = extract_assignment_value(detail, 'given')
        if family or given:
            names.append((family, given))
    return names


def format_bib_person(family: str, given: str) -> str:
    family = family.strip()
    given = given.strip()
    if family and given and ',' not in family:
        return f'{family}, {given}'
    return family or given


def format_cite_author(names: List[Tuple[str, str]]) -> str:
    families = [(family or given).strip() for family, given in names if (family or given).strip()]
    if not families:
        return ''
    if len(families) == 1:
        return families[0]
    if len(families) == 2:
        return f'{families[0]} and {families[1]}'
    return f'{families[0]} et al.'


def extract_biblatex_fields(entry_text: str) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    i = 0
    while True:
        m = re.search(r'\\field\{([^{}]+)\}\{', entry_text[i:])
        if not m:
            break
        name = m.group(1)
        start = i + m.end() - 1
        try:
            group, end = extract_balanced(entry_text, start, '{', '}')
        except ValueError:
            i += m.end()
            continue
        fields[name] = clean_bib_value(group[1:-1])
        i = end

    j = 0
    while True:
        m = re.search(r'\\list\{([^{}]+)\}\{[^{}]*\}\{', entry_text[j:])
        if not m:
            break
        name = m.group(1)
        start = j + m.end() - 1
        try:
            group, end = extract_balanced(entry_text, start, '{', '}')
        except ValueError:
            j += m.end()
            continue
        items = [clean_bib_value(x) for x in top_level_brace_groups_relaxed(group[1:-1]) if clean_bib_value(x)]
        if items:
            fields[name] = ', '.join(items)
        j = end
    return fields


def normalize_bib_title(title: str) -> str:
    title = clean_bib_value(title)
    if title == 'Cboe VIX Data':
        return 'Cboe VIX'
    title = re.sub(r'\s*-\s*Support\s*-\s*Nintendo$', '', title)
    return title


def biblatex_entry_body(entry_type: str, fields: Dict[str, str], names: List[Tuple[str, str]]) -> Tuple[str, str, str]:
    author_full = '; '.join(filter(None, (format_bib_person(family, given) for family, given in names)))
    author_cite = format_cite_author(names)
    if not author_full:
        author_full = fields.get('organization', '') or fields.get('institution', '') or fields.get('publisher', '')
    if not author_cite:
        author_cite = author_full

    year = fields.get('year', '').strip()
    title = normalize_bib_title(fields.get('title', ''))
    parts: List[str] = []
    if author_full:
        parts.append(f'{author_full}.')
    if year:
        parts.append(f'({year}).')
    if title:
        parts.append(f'{title}.')

    tail: List[str] = []
    journal = fields.get('journaltitle') or fields.get('journal')
    booktitle = fields.get('booktitle')
    volume = fields.get('volume')
    number = fields.get('number')
    pages = normalize_page_range(fields.get('pages', ''))
    publisher = fields.get('publisher', '')
    location = fields.get('location', '')

    if entry_type == 'article':
        if journal:
            tail.append(journal)
        if volume and number:
            tail.append(f'{volume}({number})')
        elif volume:
            tail.append(volume)
        elif number:
            tail.append(f'({number})')
        if pages:
            tail.append(pages)
    elif entry_type in {'book', 'mvbook'}:
        if publisher:
            tail.append(publisher)
        if location:
            tail.append(location)
    elif entry_type in {'inbook', 'incollection', 'inproceedings'}:
        if booktitle:
            tail.append(booktitle)
        if publisher:
            tail.append(publisher)
        if location:
            tail.append(location)
        if pages:
            tail.append(pages)
    else:
        if journal:
            tail.append(journal)
        elif booktitle:
            tail.append(booktitle)
        if volume and number:
            tail.append(f'{volume}({number})')
        elif volume:
            tail.append(volume)
        elif number:
            tail.append(f'({number})')
        if pages:
            tail.append(pages)
        if publisher and publisher not in tail:
            tail.append(publisher)
        if location and location not in tail:
            tail.append(location)

    body = finalize_bibliography_text(' '.join(parts + ([', '.join(tail) + '.'] if tail else [])))
    return body, author_cite.strip(), year


def parse_biblatex_entries(text: str) -> List[Tuple[str, str, LabelInfo]]:
    pattern = re.compile(r'\\entry\{([^{}]+)\}\{([^{}]+)\}\{\}\{\}')
    matches = list(pattern.finditer(text))
    entries: List[Tuple[str, str, LabelInfo]] = []
    for i, m in enumerate(matches):
        key = m.group(1).strip()
        entry_type = m.group(2).strip()
        start = m.end()
        if i + 1 < len(matches):
            end = matches[i + 1].start()
        else:
            end = text.find('\n\\end{refsection}', start)
            if end == -1:
                end = text.find('\n\\end{document}', start)
            if end == -1:
                end = len(text)
        chunk = text[start:end]
        names = extract_biblatex_names(chunk, 'author')
        fields = extract_biblatex_fields(chunk)
        body, author_cite, year = biblatex_entry_body(entry_type, fields, names)
        if not body:
            continue
        entries.append((key, body, {
            'number': '',
            'page': '',
            'title': fields.get('title', ''),
            'author': author_cite,
            'year': year,
            'kind': 'bibitem',
        }))
    return entries


def augment_labels_from_bbl(input_path: Path, labels: Dict[str, LabelInfo]) -> None:
    bbl_path = discover_bbl(input_path)
    if bbl_path is None or not bbl_path.exists():
        return
    bbl_text = bbl_path.read_text(errors='ignore')
    for key, _body, info in parse_biblatex_entries(bbl_text):
        for label in (key, bibliography_anchor_label(key)):
            current = labels.setdefault(label, {'number': '', 'page': '', 'title': '', 'kind': 'bibitem'})
            for field, value in info.items():
                if value and not current.get(field):
                    current[field] = value


def replace_refs(text: str, labels: Dict[str, LabelInfo]) -> str:
    def repl(match: re.Match[str]) -> str:
        cmd = match.group(1)
        raw = match.group(2)
        items = [x.strip() for x in raw.split(',') if x.strip()]
        if not items:
            return match.group(0)
        mode = 'zcref' if cmd in {'zcref', 'cref', 'Cref', 'autoref'} else cmd
        parts = [make_xref_placeholder(label, resolve_label(label, labels, mode)) for label in items]
        return ', '.join(parts)

    return REF_PATTERN.sub(repl, text)


def extract_referenced_equation_labels(text: str, labels: Dict[str, LabelInfo]) -> set[str]:
    referenced: set[str] = set()
    for m in REF_PATTERN.finditer(text):
        cmd = m.group(1)
        for item in m.group(2).split(','):
            label = item.strip()
            if not label:
                continue
            if cmd == 'eqref':
                referenced.add(label)
                continue
            info = labels.get(label, {})
            if classify_label(label, info) == 'Equation' or looks_like_equation_label(label):
                referenced.add(label)
    return referenced



def replace_cites(text: str, labels: Dict[str, LabelInfo]) -> str:
    def repl(match: re.Match[str]) -> str:
        cmd = match.group(1)
        raw = match.group(2)
        items = [x.strip() for x in raw.split(',') if x.strip()]
        if not items:
            return match.group(0)

        def info_for(label: str) -> LabelInfo:
            return labels.get(label, {}) or labels.get(bibliography_anchor_label(label), {})

        def cite_text(label: str) -> str:
            info = info_for(label)
            author = info.get('author', '').strip()
            year = info.get('year', '').strip() or info.get('number', '').strip()
            if cmd in {'cite', 'citep', 'citet'}:
                if author and year:
                    return f'{author} {year}'
                return author or year or label
            if cmd == 'citeauthor':
                return author or label
            if cmd == 'citeyear':
                return year or label
            return label

        parts = [make_xref_placeholder(bibliography_anchor_label(label), cite_text(label)) for label in items]
        if cmd in {'cite', 'citet'}:
            return '; '.join(parts)
        if cmd == 'citep':
            return '（' + '; '.join(parts) + '）'
        if cmd in {'citeauthor', 'citeyear'}:
            return ', '.join(parts)
        return match.group(0)

    return CITE_PATTERN.sub(repl, text)


def clean_bib_body(body: str) -> str:
    body = re.sub(r'\\penalty\d+', '', body)
    body = re.sub(r'\\providecommand\{[^{}]+\}\[[^\]]*\]\{[^{}]*\}', ' ', body)
    body = re.sub(r'\\expandafter.*', ' ', body)
    return clean_bib_value(body)


def parse_bibitem_entries(text: str) -> List[Tuple[str, str]]:
    pattern = re.compile(r'\\bibitem(?:\[[^\]]*\])?\{([^{}]+)\}')
    matches = list(pattern.finditer(text))
    entries: List[Tuple[str, str]] = []
    for i, m in enumerate(matches):
        key = m.group(1).strip()
        start = m.end()
        if i + 1 < len(matches):
            end = matches[i + 1].start()
        else:
            end = text.find('\\end{thebibliography}', start)
            if end == -1:
                end = len(text)
        body = clean_bib_body(text[start:end])
        if body:
            entries.append((key, body))
    return entries


def bibliography_blocks(entries: List[Tuple[str, str]], labels: Dict[str, LabelInfo], heading_cmd: str) -> str:
    blocks = [heading_cmd]
    for key, body in entries:
        anchor_label = bibliography_anchor_label(key)
        number = labels.get(key, {}).get('number', '').strip() or labels.get(anchor_label, {}).get('number', '').strip()
        prefix = f'[{number}] ' if number else ''
        body = finalize_bibliography_text(body)
        blocks.append(f'{LABEL_MARKER_FMT.format(label=anchor_label)} {prefix}{body}\\par')
    return '\n\n'.join(blocks)


def discover_bbl(input_path: Path) -> Optional[Path]:
    candidates: List[Path] = []
    p = input_path.with_suffix('.bbl')
    if p.exists():
        candidates.append(p)

    current = input_path.parent
    seen = {cand.resolve() for cand in candidates if cand.exists()}
    for _ in range(4):
        try:
            for cand in current.glob('*.bbl'):
                if cand.is_file():
                    resolved = cand.resolve()
                    if resolved not in seen:
                        candidates.append(cand)
                        seen.add(resolved)
        except OSError:
            pass
        if current.parent == current:
            break
        current = current.parent

    if not candidates:
        return None

    def score(path: Path) -> Tuple[int, int, int, str]:
        stem_match = int(path.stem == input_path.stem)
        prefix_len = common_prefix_len(path.stem, input_path.stem)
        try:
            bbl_text = path.read_text(errors='ignore')
        except OSError:
            return (stem_match, prefix_len, 0, str(path))
        entries = parse_biblatex_entries(bbl_text)
        if entries:
            entry_count = len(entries)
        else:
            entry_count = len(parse_bibitem_entries(bbl_text))
        return (stem_match, prefix_len, entry_count, str(path))

    return max(candidates, key=score)


def expand_bibliography_commands(text: str, input_path: Path, labels: Dict[str, LabelInfo], cited_keys: List[str]) -> str:
    text = re.sub(r'\\bibliographystyle\{[^{}]+\}', '', text)
    text = re.sub(r'\\nocite\{[^{}]+\}', '', text)
    bbl_path = discover_bbl(input_path)
    if bbl_path is None or not bbl_path.exists():
        return text
    has_bibliography_cmd = bool(re.search(r'\\bibliography\{([^{}]+)\}', text) or re.search(r'\\printbibliography(?:\[[^\]]*\])?', text))
    if not has_bibliography_cmd:
        return text
    bbl_text = bbl_path.read_text(errors='ignore')
    entries3 = parse_biblatex_entries(bbl_text)
    if entries3:
        entries = [(key, body) for key, body, _info in entries3]
    else:
        entries = parse_bibitem_entries(bbl_text)
    existing_keys = {key for key, _body in entries}
    for key in cited_keys:
        if key in existing_keys:
            continue
        info = labels.get(key, {}) or labels.get(bibliography_anchor_label(key), {})
        body = info.get('body', '').strip()
        if body:
            entries.append((key, body))
            existing_keys.add(key)
    if not entries:
        return text
    heading_cmd = '\\chapter*{参考文献}' if '\\chapter{' in text else '\\section*{参考文献}'
    replacement = bibliography_blocks(entries, labels, heading_cmd)
    text = re.sub(r'\\bibliography\{([^{}]+)\}', lambda _m: replacement, text)
    text = re.sub(r'\\printbibliography(?:\[[^\]]*\])?', lambda _m: replacement, text)
    # Remove any explicit bibliography chapter/section from source to avoid duplicates
    text = re.sub(r'\\s*\\\\chapter\\*\\{参考文献\\}\\s*', '\\n\\n', text, flags=re.S)
    text = re.sub(r'\\s*\\\\section\\*\\{参考文献\\}\\s*', '\\n\\n', text, flags=re.S)
    text = re.sub(r'\\s*\\\\chapter\\*\\{References\\}\\s*', '\\n\\n', text, flags=re.S)
    text = re.sub(r'\\s*\\\\section\\*\\{References\\}\\s*', '\\n\\n', text, flags=re.S)
    return text


def expand_inline_thebibliography(text: str, labels: Dict[str, LabelInfo]) -> str:
    pattern = re.compile(r'\\begin\{thebibliography\}\{[^{}]*\}(.*?)\\end\{thebibliography\}', re.S)
    heading_cmd = '\\chapter*{参考文献}' if '\\chapter{' in text else '\\section*{参考文献}'

    def repl(match: re.Match[str]) -> str:
        entries = parse_bibitem_entries(match.group(1))
        if not entries:
            return heading_cmd
        return bibliography_blocks(entries, labels, heading_cmd)

    return pattern.sub(repl, text)


def inject_equation_label_markers(text: str, referenced_labels: Optional[set[str]] = None) -> str:
    for env in EQUATION_ENVS:
        pattern = re.compile(rf'\\begin\{{{re.escape(env)}\}}(?P<body>.*?)\\end\{{{re.escape(env)}\}}', re.S)

        def repl(match: re.Match[str]) -> str:
            block = match.group(0)
            labels = re.findall(r'\\label\{([^{}]+)\}', block)
            if not labels:
                return block
            block_wo_labels = GENERIC_LABEL_RE.sub('', block)
            selected_labels = labels if referenced_labels is None else [label for label in labels if label in referenced_labels]
            if not selected_labels:
                return block_wo_labels
            markers = ''.join(f'\n{EQ_LABEL_MARKER_FMT.format(label=label)}\n' for label in selected_labels)
            trailing_markers = ''.join(f'\n{LABEL_MARKER_FMT.format(label=label)}\n' for label in selected_labels)
            return markers + block_wo_labels + trailing_markers

        text = pattern.sub(repl, text)
    return text


def inject_caption_label_markers(text: str) -> str:
    out: List[str] = []
    i = 0
    needle = '\\caption'
    while True:
        idx = text.find(needle, i)
        if idx == -1:
            out.append(text[i:])
            break
        out.append(text[i:idx])
        j = idx + len(needle)
        while j < len(text) and text[j].isspace():
            j += 1
        optional = ''
        if j < len(text) and text[j] == '[':
            optional, j = extract_balanced(text, j, '[', ']')
            while j < len(text) and text[j].isspace():
                j += 1
        if j >= len(text) or text[j] != '{':
            out.append(needle)
            i = idx + len(needle)
            continue
        group, j = extract_balanced(text, j, '{', '}')
        caption_body = group[1:-1]
        k = j
        while k < len(text) and text[k].isspace():
            k += 1
        m = GENERIC_LABEL_RE.match(text, k)
        if m:
            label = m.group(1)
            rebuilt = f'\\caption{optional}{{{caption_body} {LABEL_MARKER_FMT.format(label=label)}}}'
            out.append(rebuilt)
            i = m.end()
        else:
            out.append(text[idx:j])
            i = j
    return ''.join(out)


def inject_theorem_label_markers(text: str) -> str:
    for env in THEOREM_ENVS:
        pattern = re.compile(rf'(\\begin\{{{re.escape(env)}\}})\s*\\label\{{([^{{}}]+)\}}')
        text = pattern.sub(lambda m: f'\n{LABEL_MARKER_FMT.format(label=m.group(2))}\n{m.group(1)}', text)
    return text


def inject_generic_label_markers(text: str) -> str:
    text = inject_caption_label_markers(text)
    text = inject_theorem_label_markers(text)
    return GENERIC_LABEL_RE.sub(lambda m: f'\n{LABEL_MARKER_FMT.format(label=m.group(1))}\n', text)


def unwrap_theorem_like_envs(text: str, labels: Dict[str, LabelInfo]) -> str:
    display_names = {
        'theorem': '定理',
        'lemma': '引理',
        'proposition': '命题',
        'corollary': '推论',
        'definition': '定义',
        'remark': '注',
        'assumption': '假设',
    }

    marker_re = re.escape(LABEL_MARKER_FMT.split('{label}')[0]) + r'([^\n]+?)__'
    theorem_block_patterns = []
    for env in THEOREM_ENVS:
        theorem_block_patterns.append((
            env,
            re.compile(
                rf'(?P<marker>\s*{marker_re}\s*)?\\begin\{{{re.escape(env)}\}}(?:\[(?P<opt>[^\]]*)\])?(?P<body>.*?)\\end\{{{re.escape(env)}\}}',
                re.S,
            ),
        ))

    chapter_events: List[Tuple[int, str]] = []
    for m in re.finditer(r'\\appendix\b|\\chapter\{', text):
        chapter_events.append((m.start(), 'appendix' if m.group(0).startswith('\\appendix') else 'chapter'))
    chapter_events.sort()

    blocks: List[Tuple[int, int, str, str, re.Match[str]]] = []
    for env, pattern in theorem_block_patterns:
        for match in pattern.finditer(text):
            blocks.append((match.start(), match.end(), 'theorem', env, match))

    proof_pattern = re.compile(r'\\begin\{proof\}(?P<body>.*?)\\end\{proof\}', re.S)
    for match in proof_pattern.finditer(text):
        blocks.append((match.start(), match.end(), 'proof', 'proof', match))

    blocks.sort(key=lambda item: (item[0], item[1]))
    filtered_blocks: List[Tuple[int, int, str, str, re.Match[str]]] = []
    last_end = -1
    for block in blocks:
        if block[0] < last_end:
            continue
        filtered_blocks.append(block)
        last_end = block[1]

    event_idx = 0
    chapter_no = 0
    appendix_mode = False
    appendix_no = 0
    theorem_no = 0

    def advance_context(to_pos: int) -> None:
        nonlocal event_idx, chapter_no, appendix_mode, appendix_no, theorem_no
        while event_idx < len(chapter_events) and chapter_events[event_idx][0] < to_pos:
            event_kind = chapter_events[event_idx][1]
            if event_kind == 'appendix':
                appendix_mode = True
                appendix_no = 0
                theorem_no = 0
            else:
                if appendix_mode:
                    appendix_no += 1
                else:
                    chapter_no += 1
                theorem_no = 0
            event_idx += 1

    out: List[str] = []
    pos = 0
    for start, end, block_kind, env, match in filtered_blocks:
        advance_context(start)
        out.append(text[pos:start])

        if block_kind == 'proof':
            body = match.group('body').strip()
            if body:
                out.append(f'\\par\\noindent\\textit{{证明}}\n\n{body}\n')
            else:
                out.append('\\par\\noindent\\textit{证明}\n')
            pos = end
            continue

        theorem_no += 1
        marker = (match.group('marker') or '').strip()
        label = None
        if marker:
            marker_match = re.search(r'TJUFE_LABEL__(.+?)__', marker, re.S)
            if marker_match and marker_match.group(1).strip():
                label = marker_match.group(1).strip()

        opt = (match.group('opt') or '').strip()
        body = match.group('body').strip()
        display = display_names.get(env, env.title())
        if appendix_mode:
            appendix_label = chr(ord('A') + max(0, appendix_no - 1)) if appendix_no else 'A'
            number = f'{appendix_label}.{theorem_no}'
        else:
            number = f'{chapter_no or 1}.{theorem_no}'

        if label:
            current = labels.setdefault(label, {'number': '', 'page': '', 'title': '', 'kind': ''})
            current['number'] = number
            current['kind'] = env

        heading = f'\\par\\noindent\\textit{{{display}{number}}}'
        if opt:
            heading += f' {opt}'
        prefix = f'{marker}\n' if marker else ''
        if body:
            out.append(f'{prefix}{heading}\n\n{body}\n')
        else:
            out.append(prefix + heading + '\n')
        pos = end

    out.append(text[pos:])
    return ''.join(out)


def main() -> None:
    parser = argparse.ArgumentParser(description='Preprocess TJUFE LaTeX before pandoc DOCX conversion.')
    parser.add_argument('input')
    parser.add_argument('output')
    parser.add_argument('--aux', dest='aux', default=None)
    parser.add_argument('--metadata', dest='metadata', default=None)
    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve()
    aux_path = discover_aux(input_path, Path(args.aux).resolve() if args.aux else None)
    metadata_path = Path(args.metadata).resolve() if args.metadata else None

    text = input_path.read_text(errors='ignore')
    text = unwrap_subfloat(text)
    text = normalize_math_macros(text)
    text = normalize_docx_export_text(text)
    text = strip_decorative_horizontal_rules(text)

    labels = load_aux_labels(aux_path)
    scan_tex_for_missing_labels(input_path, labels)
    referenced_equation_labels = extract_referenced_equation_labels(text, labels)
    text = inject_equation_label_markers(text, referenced_equation_labels)
    text = inject_generic_label_markers(text)

    augment_labels_from_bbl(input_path, labels)
    text = unwrap_theorem_like_envs(text, labels)
    cited_keys = extract_cited_keys(text)
    augment_labels_from_bib(input_path, text, labels, cited_keys, metadata_path)
    text = expand_bibliography_commands(text, input_path, labels, cited_keys)
    text = expand_inline_thebibliography(text, labels)
    text = replace_cites(text, labels)
    text = replace_refs(text, labels)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(text)


if __name__ == '__main__':
    main()
