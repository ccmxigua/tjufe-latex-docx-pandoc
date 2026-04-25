local List = require 'pandoc.List'

local function stringify(x)
  return pandoc.utils.stringify(x or '')
end

local function trim(s)
  s = stringify(s)
  s = s:gsub('^[ \t\r\n\f\v]+', ''):gsub('[ \t\r\n\f\v]+$', '')
  s = s:gsub('[ \t\r\n\f\v]+', ' ')
  return s
end

local function meta_bool(meta, key, default)
  local v = meta[key]
  if v == nil then
    return default
  end
  local s = trim(v):lower()
  if s == '' then
    return default
  end
  return not (s == 'false' or s == '0' or s == 'no')
end

local function meta_str(meta, key)
  return trim(meta[key])
end

local function infer_degree_mode(meta)
  local hints = table.concat({
    meta_str(meta, 'degree_type'),
    meta_str(meta, 'cover_degree_line'),
    meta_str(meta, 'doctoral_apply_degree'),
    meta_str(meta, 'doctoral_subject'),
    meta_str(meta, 'doctoral_major'),
    meta_str(meta, 'doctoral_student_name'),
    meta_str(meta, 'doctoral_committee_chair')
  }, ' ')
  if hints:find('博士', 1, true) then
    return 'doctor'
  end
  if hints:find('专业硕士', 1, true) or hints:find('专业学位', 1, true) then
    return 'professional_master'
  end
  if hints:find('硕士', 1, true) then
    return 'master'
  end
  return ''
end

local function infer_cover_degree_line(meta)
  local mode = infer_degree_mode(meta)
  if mode == 'doctor' then
    return '天津财经大学博士毕业（学位）论文'
  end
  if mode == 'professional_master' then
    return '天津财经大学专业硕士毕业（学位）论文'
  end
  if mode == 'master' then
    return '天津财经大学硕士毕业（学位）论文'
  end
  return ''
end

local function infer_degree_type(meta)
  local mode = infer_degree_mode(meta)
  if mode == 'doctor' then
    return '博士毕业（学位）论文'
  end
  if mode == 'professional_master' then
    return '专业硕士毕业（学位）论文'
  end
  if mode == 'master' then
    return '硕士毕业（学位）论文'
  end
  return ''
end

local YEAR_CN_DIGITS = {
  ['0'] = '○',
  ['1'] = '一',
  ['2'] = '二',
  ['3'] = '三',
  ['4'] = '四',
  ['5'] = '五',
  ['6'] = '六',
  ['7'] = '七',
  ['8'] = '八',
  ['9'] = '九',
}

local CN_DIGITS = { '零', '一', '二', '三', '四', '五', '六', '七', '八', '九' }

local function cn_number_below_100(n)
  n = tonumber(n)
  if not n then
    return ''
  end
  if n < 10 then
    return CN_DIGITS[n + 1]
  end
  if n == 10 then
    return '十'
  end
  if n < 20 then
    return '十' .. CN_DIGITS[n - 10 + 1]
  end
  local tens = math.floor(n / 10)
  local ones = n % 10
  if ones == 0 then
    return CN_DIGITS[tens + 1] .. '十'
  end
  return CN_DIGITS[tens + 1] .. '十' .. CN_DIGITS[ones + 1]
end

local function normalize_cn_date(value)
  local s = trim(value)
  if s == '' or not s:match('%d') then
    return s
  end

  local function format_date(y, m, d)
    local year = tostring(y):gsub('%d', YEAR_CN_DIGITS)
    local out = year .. '年' .. cn_number_below_100(m) .. '月'
    if d and tostring(d) ~= '' then
      out = out .. cn_number_below_100(d) .. '日'
    end
    return out
  end

  local y, m, d = s:match('^(%d%d%d%d)%s*年%s*(%d%d?)%s*月%s*(%d%d?)%s*[日号]?$')
  if y then
    return format_date(y, m, d)
  end
  y, m = s:match('^(%d%d%d%d)%s*年%s*(%d%d?)%s*月?$')
  if y then
    return format_date(y, m)
  end
  y, m, d = s:match('^(%d%d%d%d)%s*[-/%.]%s*(%d%d?)%s*[-/%.]%s*(%d%d?)$')
  if y then
    return format_date(y, m, d)
  end
  y, m = s:match('^(%d%d%d%d)%s*[-/%.]%s*(%d%d?)$')
  if y then
    return format_date(y, m)
  end
  return s
end

local function compact_name(value)
  local s = trim(value)
  if s == '' then
    return ''
  end
  s = s:gsub('%s+', '')
  return s
end

local function normalize_cn_keyword_content(text)
  local s = trim(text)
  -- Lua patterns are byte-oriented for UTF-8. Do not put multibyte
  -- Chinese punctuation inside bracket classes such as [，；、], because
  -- that can split a UTF-8 character and create U+FFFD replacement glyphs
  -- in Pandoc DOCX output. Keep Chinese separators intact and only
  -- normalize ASCII separators that are safe byte characters.
  s = s:gsub('%s*[,;]+%s*', '，')
  s = s:gsub('，,', '，'):gsub(',，', '，')
  s = s:gsub('，;', '，'):gsub(';，', '，')
  s = s:gsub('[,;%.]+%s*$', '')
  return trim(s)
end

local function normalize_en_keyword_content(text)
  local s = trim(text)
  -- Avoid UTF-8 punctuation in Lua bracket classes; replace multibyte
  -- punctuation one literal at a time.
  s = s:gsub('；', ';')
  s = s:gsub('，', ';')
  s = s:gsub('、', ';')
  s = s:gsub('%s*;%s*', '; ')
  s = s:gsub('%s+', ' ')
  s = s:gsub('[;,%.]+%s*$', '')
  return trim(s)
end

local function normalize_en_abstract_text(text)
  local s = trim(text)
  if s == '' then
    return ''
  end
  s = s:gsub('　', ' ')
  s = s:gsub('，', ', ')
  s = s:gsub('；', '; ')
  s = s:gsub('：', ': ')
  s = s:gsub('。', '. ')
  s = s:gsub('？', '? ')
  s = s:gsub('！', '! ')
  s = s:gsub('、', ', ')
  s = s:gsub('（', '('):gsub('）', ')')
  s = s:gsub('【', '['):gsub('】', ']')
  s = s:gsub('“', '"'):gsub('”', '"')
  s = s:gsub('‘', "'"):gsub('’', "'")
  s = s:gsub('%s+', ' ')
  s = s:gsub('%s+([,.;:?!%)%]%}])', '%1')
  s = s:gsub("([,;:?!])([A-Za-z0-9%(%\"'])", '%1 %2')
  s = s:gsub('([%)%]%}])([A-Za-z0-9])', '%1 %2')
  s = s:gsub('%.([A-Z])', '. %1')
  return trim(s)
end

local function para_from_text(text)
  local s = trim(text)
  local inlines = List:new()
  for token in s:gmatch('%S+') do
    if #inlines > 0 then
      inlines:insert(pandoc.Space())
    end
    inlines:insert(pandoc.Str(token))
  end
  return pandoc.Para(inlines)
end

local function para_is_plain_text(block)
  if not block or block.t ~= 'Para' then
    return false
  end
  for _, inline in ipairs(block.content or {}) do
    if inline.t ~= 'Str' and inline.t ~= 'Space' and inline.t ~= 'SoftBreak' and inline.t ~= 'LineBreak' then
      return false
    end
  end
  return true
end

local function make_keyword_para(text, label_text, label_style, normalize_content)
  local raw = trim(text)
  local content = raw
  local lower_raw = raw:lower()
  if label_style == 'Keyword Label EN' then
    if lower_raw:match('^key%s*words?%s*[:：]?') then
      content = raw:gsub('^[Kk][Ee][Yy]%s*[Ww][Oo][Rr][Dd][Ss]?%s*[:：]?', '')
      content = trim(content)
    end
  else
    local lower_label = label_text:lower()
    if lower_raw:sub(1, #lower_label) == lower_label then
      content = trim(raw:sub(#label_text + 1))
      content = content:gsub('^[:：]+', '')
      content = trim(content)
    end
  end
  if normalize_content then
    content = normalize_content(content)
  end
  local inlines = List:new()
  local label_span = pandoc.Span({ pandoc.Str(label_text) }, pandoc.Attr('', {}, { ['custom-style'] = label_style }))
  inlines:insert(label_span)
  inlines:insert(pandoc.Str(content ~= '' and (' ' .. content) or ''))
  return pandoc.Para(inlines)
end

local function meta_rows(meta, key)
  local rows = {}
  local value = meta[key]
  if not value or value.t ~= 'MetaList' then
    return rows
  end
  local items = value.c or value
  for _, item in ipairs(items) do
    if item.t == 'MetaMap' and item.c then
      local label = trim(item.c['label'] or item.c['name'] or item.c['key'])
      local row_value = trim(item.c['value'] or item.c['text'] or item.c['content'])
      if label ~= '' or row_value ~= '' then
        table.insert(rows, { label = label, value = row_value })
      end
    end
  end
  return rows
end

local function meta_str_list(meta, key)
  local out = {}
  local value = meta[key]
  if value == nil then
    return out
  end
  if value.t == 'MetaList' then
    local items = value.c or value
    for _, item in ipairs(items) do
      local s = trim(item)
      if s ~= '' then
        table.insert(out, s)
      end
    end
    return out
  end
  local raw = trim(value)
  if raw == '' then
    return out
  end
  for item in raw:gmatch('[^,，;；]+') do
    local s = trim(item)
    if s ~= '' then
      table.insert(out, s)
    end
  end
  return out
end

local function normalize_label_key(label)
  local s = trim(label)
  s = s:gsub('[：:]', '')
  s = s:gsub('%s+', '')
  return s
end

local function is_probably_name_label(label, extra_labels)
  local s = trim(label)
  if s == '' then
    return false
  end
  local norm = normalize_label_key(s)
  for _, item in ipairs(extra_labels or {}) do
    if normalize_label_key(item) == norm then
      return true
    end
  end
  local keywords = {
    '作者', '导师', '指导教师', '指导老师', '教师', '主席', '评阅人', '姓名', '学生', '研究生', '负责人', '联系人'
  }
  for _, kw in ipairs(keywords) do
    if s:find(kw, 1, true) then
      return true
    end
  end
  return false
end

local function normalize_title_page_rows(rows, extra_name_labels)
  local out = {}
  for _, row in ipairs(rows or {}) do
    local label = trim(row.label)
    local value = trim(row.value)
    if value ~= '' and (label:find('日期', 1, true) or label:find('时间', 1, true)) then
      value = normalize_cn_date(value)
    end
    if value ~= '' and is_probably_name_label(label, extra_name_labels) then
      value = compact_name(value)
    end
    table.insert(out, { label = label, value = value })
  end
  return out
end

local function style_attr(style)
  return pandoc.Attr('', {}, { ['custom-style'] = style })
end

local function styled_para(text_or_inlines, style)
  local para
  if type(text_or_inlines) == 'string' then
    para = pandoc.Para({ pandoc.Str(text_or_inlines) })
  else
    para = pandoc.Para(text_or_inlines)
  end
  return pandoc.Div({ para }, style_attr(style))
end

local function clone_para_with_style(para, style)
  return pandoc.Div({ para }, style_attr(style))
end

local function clone_header_with_style(level, style, new_text)
  local h = pandoc.Header(level, { pandoc.Str(new_text) })
  h.attr = pandoc.Attr('', {}, { ['custom-style'] = style })
  return h
end

local function pagebreak()
  if FORMAT:match('docx') then
    return pandoc.RawBlock('openxml', '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:br w:type="page"/></w:r></w:p>')
  elseif FORMAT:match('latex') then
    return pandoc.RawBlock('tex', '\\newpage')
  else
    return pandoc.HorizontalRule()
  end
end

local function empty_para(style)
  return styled_para('', style or 'Normal')
end

local function starts_with(s, prefix)
  return s:sub(1, #prefix) == prefix
end

local function is_cn_abstract(title)
  local t = trim(title)
  return t == '摘要' or t == '摘 要'
end

local function is_en_abstract(title)
  return trim(title):lower() == 'abstract'
end

local function is_ack(title)
  return trim(title) == '后记' or trim(title) == '后 记'
end

local function is_refs(title)
  return trim(title) == '参考文献'
end

local function is_appendix(title)
  return starts_with(trim(title), '附录')
end

local function is_research_outputs(title)
  local t = trim(title)
  return t == '在学期间发表的学术论文与研究成果'
end

local function make_label_value_blocks(label_style, value_style, label, value)
  local blocks = List:new()
  if trim(label) ~= '' then
    blocks:insert(styled_para(label, label_style))
  end
  if trim(value) ~= '' then
    blocks:insert(styled_para(value, value_style))
  end
  return blocks
end

local function xml_escape(s)
  s = stringify(s or '')
  s = s:gsub('&', '&amp;')
  s = s:gsub('<', '&lt;')
  s = s:gsub('>', '&gt;')
  s = s:gsub('"', '&quot;')
  return s
end

local function with_cn_colon(label)
  label = trim(label)
  if label == '' then
    return ''
  end
  if label:match('[：:]$') then
    return label
  end
  return label .. '：'
end

local function openxml_table_cell(text, width, align, size, bold, east_asia)
  local rpr = string.format(
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="%s" w:cs="Times New Roman"/>%s<w:sz w:val="%s"/><w:szCs w:val="%s"/></w:rPr>',
    xml_escape(east_asia or 'SimSun'),
    bold and '<w:b/><w:bCs/>' or '',
    tostring(size or 28),
    tostring(size or 28)
  )
  return string.format(
    '<w:tc>' ..
      '<w:tcPr><w:tcW w:w="%s" w:type="dxa"/></w:tcPr>' ..
      '<w:p>' ..
        '<w:pPr><w:jc w:val="%s"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>' ..
        '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' ..
      '</w:p>' ..
    '</w:tc>',
    tostring(width),
    align or 'center',
    rpr,
    xml_escape(text ~= '' and text or ' ')
  )
end

local function openxml_table(rows, widths, opts)
  opts = opts or {}
  local parts = {
    '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:tblPr>',
    '<w:tblW w:w="0" w:type="auto"/>',
    string.format('<w:jc w:val="%s"/>', opts.align or 'center'),
    '<w:tblBorders>' ..
      '<w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>' ..
      '<w:insideH w:val="nil"/><w:insideV w:val="nil"/>' ..
    '</w:tblBorders>',
    string.format('<w:tblLayout w:type="%s"/>', opts.layout or 'autofit'),
    '</w:tblPr>',
    '<w:tblGrid>'
  }
  for _, width in ipairs(widths) do
    table.insert(parts, string.format('<w:gridCol w:w="%s"/>', tostring(width)))
  end
  table.insert(parts, '</w:tblGrid>')
  for _, row in ipairs(rows) do
    table.insert(parts, '<w:tr>')
    for idx, cell in ipairs(row) do
      table.insert(parts, openxml_table_cell(cell.text or '', widths[idx], cell.align, cell.size, cell.bold, cell.eastAsia))
    end
    table.insert(parts, '</w:tr>')
  end
  table.insert(parts, '</w:tbl>')
  return pandoc.RawBlock('openxml', table.concat(parts))
end

local function make_cover_info_table(rows)
  local xml_rows = {}
  for _, row in ipairs(rows) do
    table.insert(xml_rows, {
      { text = with_cn_colon(row.label), align = 'distribute', size = 28, bold = false, eastAsia = 'SimSun' },
      { text = trim(row.value), align = 'center', size = 28, bold = false, eastAsia = 'SimSun' },
    })
  end
  return openxml_table(xml_rows, {2292, 3947}, { align = 'center', layout = 'autofit' })
end

local function pair_title_page_rows(rows)
  local paired = {}
  local i = 1
  while i <= #rows do
    local left = rows[i] or { label = '', value = '' }
    local right = rows[i + 1] or { label = '', value = '' }
    table.insert(paired, {
      { text = with_cn_colon(left.label), align = 'distribute', size = 28, bold = true, eastAsia = 'SimSun' },
      { text = trim(left.value), align = 'center', size = 28, bold = true, eastAsia = 'SimSun' },
      { text = with_cn_colon(right.label), align = 'distribute', size = 28, bold = true, eastAsia = 'SimSun' },
      { text = trim(right.value), align = 'center', size = 28, bold = true, eastAsia = 'SimSun' },
    })
    i = i + 2
  end
  return paired
end

local function make_title_page_table(rows)
  return openxml_table(pair_title_page_rows(rows), {2392, 1946, 1842, 2199}, { align = 'center', layout = 'autofit' })
end

local function append_info_rows(blocks, label_style, value_style, rows)
  for _, row in ipairs(rows) do
    for _, b in ipairs(make_label_value_blocks(label_style, value_style, row.label, row.value)) do
      blocks:insert(b)
    end
  end
end

local function split_cover_degree_line(line)
  line = trim(line)
  if line == '' then
    return '', ''
  end
  if starts_with(line, '天津财经大学') then
    local suffix = trim(line:gsub('^天津财经大学', ''))
    if suffix ~= '' then
      return '天津财经大学', suffix
    end
  end
  return line, ''
end

local function inject_blank_lines(blocks, count, style)
  for _ = 1, count do
    blocks:insert(empty_para(style))
  end
end

local function build_frontmatter(meta)
  local blocks = List:new()
  local insert_cover = meta_bool(meta, 'insert_cover', true)
  local insert_title_page = meta_bool(meta, 'insert_title_page', true)

  if insert_cover then
    local cover_degree_line = meta_str(meta, 'cover_degree_line')
    if cover_degree_line == '' then
      cover_degree_line = infer_cover_degree_line(meta)
    end
    local cover_top, cover_degree = split_cover_degree_line(cover_degree_line)
    local cn_title = meta_str(meta, 'cn_title')
    local cn_subtitle = meta_str(meta, 'cn_subtitle')
    local discipline = meta_str(meta, 'discipline')
    local student_id = meta_str(meta, 'student_id')
    local author = compact_name(meta_str(meta, 'author'))
    local advisor = compact_name(meta_str(meta, 'advisor'))
    local submit_date_cn = normalize_cn_date(meta_str(meta, 'submit_date_cn'))

    if cover_top ~= '' or cover_degree ~= '' or cn_title ~= '' then
      if cover_top ~= '' then
        blocks:insert(styled_para(cover_top, 'Cover Top Line'))
      end
      if cover_degree ~= '' then
        blocks:insert(styled_para(cover_degree, 'Cover Top Line'))
      end
      inject_blank_lines(blocks, 1)
      if cn_title ~= '' then
        blocks:insert(styled_para(cn_title, 'Title'))
      end
      if cn_subtitle ~= '' then
        blocks:insert(styled_para('——' .. cn_subtitle, 'Subtitle'))
      end
      inject_blank_lines(blocks, 2)
      blocks:insert(make_cover_info_table({
        { label = '专业名称', value = discipline },
        { label = '作者学号', value = student_id },
        { label = '论文作者', value = author },
        { label = '指导教师', value = advisor },
      }))
      inject_blank_lines(blocks, 4)
      if submit_date_cn ~= '' then
        blocks:insert(styled_para(submit_date_cn, 'Cover Date'))
      end
      blocks:insert(pagebreak())
    end
  end

  if insert_title_page then
    local class_no = meta_str(meta, 'class_no')
    local confidentiality = meta_str(meta, 'confidentiality')
    local degree_type = meta_str(meta, 'degree_type')
    if degree_type == '' then
      degree_type = infer_degree_type(meta)
    end
    local cn_title = meta_str(meta, 'cn_title')
    local cn_subtitle = meta_str(meta, 'cn_subtitle')
    local en_title = meta_str(meta, 'en_title')
    local en_subtitle = meta_str(meta, 'en_subtitle')
    local college = meta_str(meta, 'college')
    local discipline = meta_str(meta, 'discipline')
    local author = compact_name(meta_str(meta, 'author'))
    local advisor = compact_name(meta_str(meta, 'advisor'))
    local doctoral_subject = meta_str(meta, 'doctoral_subject')
    local doctoral_direction = meta_str(meta, 'doctoral_research_direction')
    local doctoral_student_name = compact_name(meta_str(meta, 'doctoral_student_name'))
    local doctoral_defense_date = normalize_cn_date(meta_str(meta, 'doctoral_defense_date'))
    local doctoral_degree_date = normalize_cn_date(meta_str(meta, 'doctoral_degree_date'))
    local title_page_rows = meta_rows(meta, 'title_page_info_rows')
    local extra_name_labels = meta_str_list(meta, 'title_page_compact_name_labels')

    if doctoral_subject ~= '' or doctoral_direction ~= '' or doctoral_student_name ~= '' or doctoral_defense_date ~= '' or doctoral_degree_date ~= '' then
      title_page_rows = {
        { label = '论文作者', value = doctoral_student_name ~= '' and doctoral_student_name or author },
        { label = '指导教师', value = advisor },
        { label = '申请学位', value = meta_str(meta, 'doctoral_apply_degree') ~= '' and meta_str(meta, 'doctoral_apply_degree') or '管理学博士' },
        { label = '培养单位', value = meta_str(meta, 'doctoral_college') ~= '' and meta_str(meta, 'doctoral_college') or college },
        { label = '学科专业', value = meta_str(meta, 'doctoral_major') ~= '' and meta_str(meta, 'doctoral_major') or discipline },
        { label = '研究方向', value = doctoral_direction },
        { label = '答辩日期', value = doctoral_defense_date },
        { label = '学位授予日期', value = doctoral_degree_date },
        { label = '答辩委员会主席', value = compact_name(meta_str(meta, 'doctoral_committee_chair')) },
        { label = '论文评阅人', value = meta_str(meta, 'doctoral_reviewers') ~= '' and compact_name(meta_str(meta, 'doctoral_reviewers')) or '匿名评审' },
      }
    elseif #title_page_rows == 0 then
      title_page_rows = {
        { label = '所属学院', value = college },
        { label = '专业名称', value = discipline },
        { label = '论文作者', value = author },
        { label = '指导教师', value = advisor },
      }
    end

    title_page_rows = normalize_title_page_rows(title_page_rows, extra_name_labels)

    if degree_type ~= '' or cn_title ~= '' then
      if class_no ~= '' then
        blocks:insert(styled_para('分类号：' .. class_no, 'Title Page Meta'))
      end
      if confidentiality ~= '' then
        blocks:insert(styled_para('密  级：' .. confidentiality, 'Title Page Meta'))
      end
      inject_blank_lines(blocks, 1)
      if degree_type ~= '' then
        blocks:insert(styled_para(degree_type, 'Title Page Degree'))
      end
      inject_blank_lines(blocks, 1)
      if cn_title ~= '' then
        blocks:insert(styled_para(cn_title, 'Title'))
      end
      if cn_subtitle ~= '' then
        blocks:insert(styled_para('——' .. cn_subtitle, 'Subtitle'))
      end
      inject_blank_lines(blocks, 3)
      if en_title ~= '' then
        blocks:insert(styled_para(en_title, 'English Title'))
      end
      if en_subtitle ~= '' then
        blocks:insert(styled_para('——' .. en_subtitle, 'English Subtitle'))
      end
      inject_blank_lines(blocks, 5)
      blocks:insert(make_title_page_table(title_page_rows))
      blocks:insert(pagebreak())
    end
  end

  return blocks
end

function Pandoc(doc)
  local out = List:new()
  local last_was_pagebreak = false
  local chapter_no = 0
  local section_no = 0
  local subsection_no = 0
  local appendix_no = 0
  local ack_signature_inserted = false

  local function push(block)
    out:insert(block)
    if block.t == 'RawBlock' and block.format == 'openxml' and block.text:match('w:type="page"') then
      last_was_pagebreak = true
    else
      last_was_pagebreak = false
    end
  end

  local function ensure_pagebreak()
    if #out > 0 and not last_was_pagebreak then
      push(pagebreak())
    end
  end

  local function push_signature_line(text)
    if trim(text) ~= '' then
      push(styled_para(text, 'Signature Line'))
    end
  end

  local function flush_ack_signature()
    if ack_signature_inserted then
      return
    end
    local sign_name = compact_name(meta_str(doc.meta, 'ack_signature_name'))
    if sign_name == '' then
      sign_name = compact_name(meta_str(doc.meta, 'author'))
    end
    local sign_date = normalize_cn_date(meta_str(doc.meta, 'ack_signature_date'))
    if sign_date == '' then
      sign_date = normalize_cn_date(meta_str(doc.meta, 'submit_date_cn'))
    end
    if sign_name == '' and sign_date == '' then
      ack_signature_inserted = true
      return
    end
    inject_blank_lines(out, 2)
    push_signature_line(sign_name)
    push_signature_line(sign_date)
    ack_signature_inserted = true
  end

  for _, b in ipairs(build_frontmatter(doc.meta)) do
    push(b)
  end

  local mode = 'normal'
  local inserted_toc = false
  local want_toc = meta_bool(doc.meta, 'insert_toc', true)
  local appendix_mode = false

  for _, block in ipairs(doc.blocks) do
    if mode == 'ack' and block.t == 'Header' then
      flush_ack_signature()
    end

    if block.t == 'RawBlock' and block.format == 'latex' and trim(block.text) == '\\appendix' then
      appendix_mode = true
    elseif block.t == 'Header' then
      local title = trim(block.content)
      if block.level == 1 then
        if is_cn_abstract(title) then
          mode = 'cn_abstract'
          ensure_pagebreak()
          push(clone_header_with_style(1, 'Abstract Title CN', '摘 要'))
          inject_blank_lines(out, 1)
        elseif is_en_abstract(title) then
          mode = 'en_abstract'
          ensure_pagebreak()
          push(clone_header_with_style(1, 'Abstract Title EN', 'Abstract'))
          inject_blank_lines(out, 1)
        elseif is_ack(title) then
          appendix_mode = false
          mode = 'ack'
          ack_signature_inserted = false
          ensure_pagebreak()
          push(clone_header_with_style(1, 'Acknowledgements Heading', '后 记'))
        elseif is_refs(title) then
          appendix_mode = false
          mode = 'refs'
          ensure_pagebreak()
          push(clone_header_with_style(1, 'References Heading', '参考文献'))
        elseif is_research_outputs(title) then
          appendix_mode = false
          mode = 'research_outputs'
          ensure_pagebreak()
          push(clone_header_with_style(1, 'Research Outputs Heading', '在学期间发表的学术论文与研究成果'))
        elseif appendix_mode or is_appendix(title) then
          mode = 'appendix'
          appendix_no = appendix_no + 1
          section_no = 0
          subsection_no = 0
          ensure_pagebreak()
          local appendix_label = string.char(string.byte('A') + appendix_no - 1)
          local appendix_title = title
          if is_appendix(title) then
            appendix_title = title:gsub('^附录[%sA-ZＡ-Ｚ]*[%s:：-]*', ''):gsub('^%s+', '')
          end
          push(clone_header_with_style(1, 'Appendix Heading', string.format('附录%s  %s', appendix_label, appendix_title)))
        else
          if want_toc and not inserted_toc then
            ensure_pagebreak()
            push(styled_para('目 录', 'TOC Heading'))
            inject_blank_lines(out, 2)
            push(styled_para('__TJUFE_TOC_PLACEHOLDER__', 'Normal'))
            push(pagebreak())
            inserted_toc = true
          end
          mode = 'body'
          chapter_no = chapter_no + 1
          section_no = 0
          subsection_no = 0
          ensure_pagebreak()
          push(clone_header_with_style(1, 'heading 1', string.format('第%d章  %s', chapter_no, title)))
          inject_blank_lines(out, 2)
        end
      elseif block.level == 2 then
        section_no = section_no + 1
        subsection_no = 0
        if mode == 'appendix' then
          push(clone_header_with_style(2, 'heading 2', string.format('%s.%d  %s', string.char(string.byte('A') + appendix_no - 1), section_no, title)))
        else
          push(clone_header_with_style(2, 'heading 2', string.format('%d.%d  %s', chapter_no, section_no, title)))
        end
        inject_blank_lines(out, 1)
      elseif block.level == 3 then
        subsection_no = subsection_no + 1
        if mode == 'appendix' then
          push(clone_header_with_style(3, 'heading 3', string.format('%s.%d.%d  %s', string.char(string.byte('A') + appendix_no - 1), section_no, subsection_no, title)))
        else
          push(clone_header_with_style(3, 'heading 3', string.format('%d.%d.%d  %s', chapter_no, section_no, subsection_no, title)))
        end
      else
        push(block)
      end
    elseif block.t == 'Para' then
      local text = trim(block.content)
      if mode == 'cn_abstract' then
        if starts_with(text, '关键词') then
          inject_blank_lines(out, 1)
          local keyword_para = make_keyword_para(text, '关键词：', 'Keyword Label CN', normalize_cn_keyword_content)
          push(clone_para_with_style(keyword_para, 'Keywords Line CN'))
        else
          push(clone_para_with_style(block, 'Abstract Body CN'))
        end
      elseif mode == 'en_abstract' then
        if starts_with(text:lower(), 'key words') then
          inject_blank_lines(out, 1)
          local keyword_para = make_keyword_para(text, 'Key Words:', 'Keyword Label EN', normalize_en_keyword_content)
          push(clone_para_with_style(keyword_para, 'Keywords Line EN'))
        else
          if para_is_plain_text(block) then
            push(clone_para_with_style(para_from_text(normalize_en_abstract_text(text)), 'Abstract Body EN'))
          else
            push(clone_para_with_style(block, 'Abstract Body EN'))
          end
        end
      elseif mode == 'ack' then
        push(clone_para_with_style(block, 'Acknowledgements Body'))
      elseif mode == 'refs' then
        push(clone_para_with_style(block, 'Bibliography'))
      elseif mode == 'research_outputs' then
        push(clone_para_with_style(block, 'Bibliography'))
      else
        push(block)
      end
    else
      push(block)
    end
  end

  if mode == 'ack' then
    flush_ack_signature()
  end

  return pandoc.Pandoc(out, doc.meta)
end
