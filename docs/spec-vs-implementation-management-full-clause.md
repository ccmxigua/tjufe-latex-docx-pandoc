# 天津财经大学管理学研究生毕业（学位）论文编写规范：全文逐条对照版

日期：2026-04-23

规范原文：
`docs/tjufe-management-thesis-format-spec.docx`

核对对象：
- `convert_tjufe.sh`
- `filters/tjufe-thesis.lua`
- `scripts/build_reference_docx.py`
- `scripts/postprocess_tjufe_docx.py`

判定标签：
- 已实现：脚本中已有明确实现证据
- 部分实现：有实现，但与规范字面值不完全一致，或只覆盖部分场景
- 未实现：在四个核心文件中未见明确实现证据
- 不属于脚本范围：属于内容撰写要求、学术要求或印刷/装订要求，不是这四个脚本能自动保证的
- 需人工终检：脚本已做，但最终需在 Word/WPS 成品中人工确认

重要边界：
1. 本报告按当前 DOCX 规范正文可见内容逐段建账。
2. 本报告已经覆盖提取出的全部章节、条、款、子项；但“视觉观感类要求”仍不能伪装成 100% 自动可证，只能判为“需人工终检”。
3. 本报告只对上述 4 个核心文件负责；若外部模板、Pandoc 默认行为、Word 渲染器、人工后修也参与了最终效果，则超出本报告的“纯脚本实现”边界。

---

## A. 总览结论

总体结论：
- 这套链路对“结构性规范”和“版式主骨架”实现度较高。
- 真正实现主体是三层：
  1. `build_reference_docx.py`：页面、页边距、样式模板
  2. `tjufe-thesis.lua`：封面/扉页/摘要/目录/正文/附录/后记等结构重排
  3. `postprocess_tjufe_docx.py`：页眉页脚、页码分节、目录域、公式编号、图表编号、交叉引用、三线表、脚注编号等 OOXML 后处理
- 当前最主要偏差已经不再是“正文/脚注/公式/参考文献的行距规则未落地”；这些关键样式值已在模板与后处理中同步修正。

当前最重要偏差：
1. 仍有少数细粒度版式/视觉项需要人工终检，例如中文摘要标题在 Word 中的大纲/样式最终表现、页眉“文武线”、公式真实视觉松紧、续图续表跨页视觉观感。
2. 少数规则仍属部分实现，例如表注/图注“左下方/顶左框”位置还不是逐字严格实现；另有部分复杂显示场景仍更适合人工终检。
3. 尚未自动化或仅做启发式覆盖的项目主要集中在结构/内容校验层，例如三级标题下后续编号序列、GB/T 7714 复杂细则等；正文底部大片空白与分页异常、图文先后、图单位/分图、在学期间成果细则已补入保守 WARN 审计。
4. 一些内容性/学术性/印刷性要求本就不属于这 4 个脚本的职责范围。

---

## B. 第一章：毕业（学位）论文的结构格式

### B-1 总述：论文各组成部分按顺序排列
规范原文：封面、扉页、摘要和关键词、Abstract和Key Words、目录、正文、参考文献、附录、在学期间成果、后记。
判定：已实现（附录为条件实现，博士成果页为条件实现）。
证据：
- `tjufe-thesis.lua / build_frontmatter()`：封面、扉页
- `tjufe-thesis.lua / Pandoc()`：摘要、Abstract、目录、正文、参考文献、附录、在学期间成果、后记
- `ensure_pagebreak()` 确保这些部分按顺序分页插入
说明：
- 附录取决于原文是否存在 `\appendix` 或“附录”标题
- 博士成果页取决于输入文稿是否带该一级标题

### B-2 论文使用汉字撰写
判定：不属于脚本范围。
说明：脚本不会自动检查整篇论文语言是不是汉字撰写。

---

## C. 第二章：研究生毕业（学位）论文格式基本要求

### C-1 封面：题目应简明、一般不超过25字、避免不常用缩略词/字符/公式等
判定：部分实现。
证据：
- `convert_tjufe.sh / scan_docx_or_fail()` 会提取封面中文题目并做长度审计
- 去空白后长度大于 `25` 时输出 `DOCX_SCAN_WARN: title_too_long`
说明：
- “一般不超过25字”的长度上限已有告警审计
- 但“题目是否简明”“是否避免不常用缩略词/字符/公式”等语义与内容质量要求，仍不属于这 4 个脚本可完全自动判定的范围

### C-2 封面：只要求中文题目
判定：已实现。
证据：
- `tjufe-thesis.lua / build_frontmatter()` 的封面只写 `cn_title`，不写英文题目

### C-3 封面：可使用副标题，副标题前加破折号
判定：已实现。
证据：
- `tjufe-thesis.lua / build_frontmatter()`：`'——' .. cn_subtitle`

### C-4 封面：顶端为“天津财经大学硕士/专业硕士/博士毕业（学位）论文”，两行居中排列
判定：部分实现。
证据：
- `build_frontmatter()` 通过 `cover_degree_line` 拆成 `cover_top` 与 `cover_degree` 两行；分别套 `Cover Top Line`
说明：
- 需要 metadata 正确提供 `cover_degree_line`
- 脚本支持两行排法，但并不自动判断当前论文类型并生成唯一合法文本，依赖外部 metadata

### C-5 封面：签名栏顺序为专业名称、作者学号、论文作者、指导教师
判定：已实现。
证据：
- `build_frontmatter()` 中 `make_cover_info_table()` 按固定顺序写入这四项

### C-6 封面：最末一行为提交论文日期（中文大写）
判定：已实现。
证据：
- `tjufe-thesis.lua / build_frontmatter()` 写入 `submit_date_cn`
- `tjufe-thesis.lua / normalize_cn_date()` 会把 `YYYY年M月D日`、`YYYY-M-D`、`YYYY/M/D`、`YYYY.M.D` 等形式自动规范为中文日期，例如年份数字转为 `○一二三` 形态，月份/日期转为中文数字
- 样例导出的 `out/sample.docx` 中，`Cover Date` 已为单实例样式，实际落到 `size=28`、`jc=center`、`line=240 auto`、`bold=None`
说明：
- 当前不仅负责放置该字段，也已对常见数字日期输入做中文日期规范化。

### C-7 扉页：上方为论文题目（中英文对照）
判定：已实现。
证据：
- `build_frontmatter()` 写入 `cn_title`、`cn_subtitle`、`en_title`、`en_subtitle`

### C-8 扉页：顶端为硕士/专业硕士/博士毕业（学位）论文
判定：已实现。
证据：
- `build_frontmatter()` 写入 `degree_type`

### C-9 扉页：下方为签名栏（具体内容见模板）
判定：已实现。
证据：
- `make_title_page_table()` + `title_page_rows`
- 普通模式使用学院/专业/作者/导师
- 博士模式使用作者/导师/申请学位/培养单位/学科专业/研究方向/答辩主席/评阅人

### C-10 扉页：姓名文字之间不加空格
判定：已实现。
证据：
- `tjufe-thesis.lua / compact_name()` 会删除姓名字段中的空白字符
- `build_frontmatter()` 已对 `author`、`advisor`、`doctoral_committee_chair`、`doctoral_reviewers` 等常用姓名字段调用 `compact_name()`
- `normalize_title_page_rows()` 已对 `title_page_info_rows` 做姓名标签启发式识别，并支持通过 `title_page_compact_name_labels` 扩展姓名类字段
- 2026-04-24 回归导出的临时样例中，`title_page_info_rows` 里的 `张 三 / 李 四 教 授 / 王 教 授 / 匿 名 评 审` 已在成品 `document.xml` 中落为 `张三 / 李四教授 / 王教授 / 匿名评审`
说明：
- 当前常用姓名字段与 `title_page_info_rows` 的姓名类字段都已进入去空格规范化链路

### C-11 中文摘要：摘要应独立、自含、精炼，包含背景/意义/方法/内容/结论/创新/不足等
判定：不属于脚本范围。
说明：属于摘要内容质量要求，不是排版脚本能保证的。

### C-12 中文摘要：硕士约1000字、博士约2500字
判定：已实现。
证据：
- `convert_tjufe.sh / scan_docx_or_fail()` 已增加摘要字数扫描
- 通过 metadata 中 `degree_type / cover_degree_line / doctoral_apply_degree` 推断 `degree_mode`
- 中文摘要阈值已落地：博士 `2000~3500` 字，硕士 `800~1800` 字；过短/过长分别输出 `DOCX_SCAN_WARN: cn_abstract_too_short` / `cn_abstract_too_long`
- 英文摘要最小词数已落地：博士 `250`，硕士 `120`；过短输出 `DOCX_SCAN_WARN: en_abstract_too_short`
说明：
- 当前实现属于“审计告警”，不是自动改写摘要内容。

### C-13 中文摘要：不要用图、表、数学公式
判定：已实现。
证据：
- `convert_tjufe.sh / scan_docx_or_fail()` 已扫描摘要区块中的表格、绘图、公式
- 命中后直接报错：`abstract-contains-table:cn/en`、`abstract-contains-drawing:cn/en`、`abstract-contains-math:cn/en`

### C-14 中文关键词：一般 3～5 个
判定：已实现。
证据：
- `convert_tjufe.sh / scan_docx_or_fail()` 已对 `KeywordsLineCN` / `KeywordsLineEN` 做数量审计
- 中文关键词按中文逗号分割，数量不在 `3~5` 时输出 `DOCX_SCAN_WARN: cn_keywords_count`
- 英文关键词按分号分割，数量不在 `3~5` 时输出 `DOCX_SCAN_WARN: en_keywords_count`
- 同时增加中英文关键词标点合法性检查：`cn_keywords_punct_bad`、`en_keywords_punct_bad`

### C-15 英文 Abstract / Key Words 应与中文摘要和关键词基本对应、符合英语语法
判定：不属于脚本范围。
说明：属于语义/语言质量要求。

### C-16 目录：应包含摘要、Abstract、正文、参考文献、后记等；按页码顺序排列至三级标题
判定：已实现。
证据：
- `tjufe-thesis.lua` 插入目录页与 TOC 占位符
- `postprocess_tjufe_docx.py / make_toc_paragraph()` 使用 `TOC \o "1-3"`

### C-17 正文：应包括绪论、各具体章节、结论
判定：不属于脚本范围。
说明：脚本不检查正文语义结构是否完整具备“绪论/结论”等。

### C-18 正文：论点正确、论据充实、数据可靠、文字精炼等
判定：不属于脚本范围。

### C-19 正文：引用他人观点必须注明出处
判定：不属于脚本范围。
说明：脚本可处理交叉引用，不等于学术引用规范校验。

### C-20 字数要求：博士不少于10万字，硕士不少于3万字
判定：已实现。
证据：
- `convert_tjufe.sh / scan_docx_or_fail()` 已按 metadata 推断学位类型并统计正文中文字符数
- 阈值已落地：博士 `100000`，硕士 `30000`
- 低于阈值时输出 `DOCX_SCAN_WARN: body_char_count_too_short`
说明：
- 当前是导出后审计告警，不会自动扩写正文。

### C-21 注释：建议采用脚注，每页重新编号
判定：已实现。
证据：
- `postprocess_tjufe_docx.py / FOOTNOTE_NUMFMT`
- `configure_footnote_pr()`：`numRestart = eachPage`

### C-22 参考文献数量与年龄/外文占比要求
判定：已实现。
证据：
- `convert_tjufe.sh / scan_bibliography_or_fail()` 已增加 bibliography 预审
- 按 metadata 推断学位类型并执行数量门槛审计：博士不少于 `100` 项，硕士不少于 `40` 项，专业硕士不少于 `30` 项
- 已增加近五年文献占比审计：不少于总数的 `1/4`
- 已增加外文文献占比审计：不少于总数的 `1/5`
- 不满足时输出 `BIBLIOGRAPHY_SCAN_WARN: total_too_few / recent_ratio_too_low / foreign_ratio_too_low`
说明：
- 外文识别当前采用基于 `language/langid` 与作者/标题字符特征的静态启发式，不等于人工文献学判断。

### C-23 附录可选，用于冗长推导、复杂数据图表、符号说明等
判定：部分实现。
说明：
- 脚本支持附录结构、编号与分页
- 但不会自动判断哪些内容“应该进附录”

### C-24 博士论文应增加“在学期间发表的学术论文与研究成果”
判定：部分实现，已补入保守 WARN 审计。
证据：
- `tjufe-thesis.lua` 支持该一级标题并分页重写
- `convert_tjufe.sh / scan_docx_or_fail()` 已在博士模式下审计该部分是否存在；缺失时输出 `research_outputs_missing:doctor`
说明：
- 脚本不会自动代写成果内容；当前采用导出后 WARN 提醒人工补充

### C-25 在学期间成果的内容格式、括号中注明检索类型、影响因子、项目/专利/奖项说明等
判定：部分实现，已补入保守 WARN 审计。
证据：
- `convert_tjufe.sh / scan_docx_or_fail()` 会收集 `ResearchOutputsHeading` 后、后记/附录/参考文献等下一节前的成果条目
- 成果页存在但无条目时输出 `research_outputs_empty:doctor`
- 疑似成果条目缺少文献/成果类型标识时输出 `research_output_type_marker_missing:<n>`
- 疑似成果条目缺少括号说明、检索类型、影响因子、项目/专利/奖项等说明时输出 `research_output_detail_note_missing:<n>`
说明：
- 当前仍是启发式结构审计，不会自动判断成果真实性、作者排序或影响因子准确性。

### C-26 后记应说明创作来源、感谢对象，字数约500字
判定：不属于脚本范围。

---

## D. 第三章：学位论文的书写、排版及印刷要求

### D-1 封面顶端：宋体三号、居中、无缩进、单倍行距
判定：已实现。
证据：
- `build_reference_docx.py`
  - `CoverTopLine`: `song_r_16`（三号），`align=center`，`firstLineChars=0`，`line=240`, `lineRule=auto`

### D-2 封面题目：宋体二号、加粗、居中、单倍行距
判定：已实现。
证据：
- `Title`: `song_r_22`（二号），bold，center，240 auto

### D-3 封面副标题：宋体小二、加粗、居中、无缩进、单倍行距
判定：已实现。
证据：
- `Subtitle`: `song_r_18`（小二），bold，center，240 auto

### D-4 封面签名栏标题：宋体/Times New Roman 四号，分散对齐，无缩进，单倍行距
判定：已实现。
证据：
- `build_frontmatter()` 使用 openxml table 单元格；label 单元格 `size=28`（四号），`align=distribute`，`line=240 auto`

### D-5 封面签名栏内容：宋体四号，居中，无缩进，单倍行距
判定：已实现。
证据：
- cover info value 单元格 `size=28`，`align=center`，`line=240 auto`

### D-6 封面提交日期：宋体四号，居中，无缩进，单倍行距
判定：已实现。
证据：
- `build_reference_docx.py` 中 `CoverDate` 已为 `song_r_14`（非粗体），center，`line=240 auto`
- `convert_tjufe.sh / scan_docx_or_fail()` 已对 `CoverDate` 加入成品审计：若字体不是 `SimSun`、字号不是 `28` 或仍为粗体，会报 `bad-style-font:CoverDate`
说明：此前“仍为粗体”的状态已被修正；当前该项应上调为已实现。

### D-7 扉页：分类号至密级标题及内容，四号，加粗，两端对齐，无缩进，单倍行距
判定：已实现。
证据：
- `TitlePageMeta`: 四号加粗，两端对齐，240 auto

### D-8 扉页学位类型行：宋体四号加粗居中单倍行距
判定：已实现。
证据：
- `TitlePageDegree`: 四号加粗 center 240 auto

### D-9 扉页中文题目：宋体二号加粗居中单倍行距
判定：已实现。

### D-10 扉页中文副标题：宋体小二加粗居中单倍行距
判定：已实现。

### D-11 扉页英文主标题：Times New Roman 小三加粗居中单倍行距
判定：已实现。
证据：
- `EnglishTitle`: size 30（小三），bold，center，240 auto

### D-12 扉页英文副标题：Times New Roman 四号加粗居中单倍行距
判定：已实现。
证据：
- `EnglishSubtitle`: size 28（四号），bold，center，240 auto

### D-13 扉页所属学院至论文作者标题（博士为论文作者至论文评阅人标题）：四号加粗，分散对齐，无缩进，单倍行距
判定：已实现。
证据：
- title-page 表格 label 单元格 `size=28`、bold、distribute、240 auto

### D-14 扉页对应填写内容：宋体四号加粗居中无缩进单倍行距
判定：已实现。
证据：
- title-page value 单元格 `size=28`、bold、center、240 auto

### D-15 中文摘要标题：另起一页、三号黑体、字间空1个半角空格、大纲级别1级、无缩进、段前后0、固定值20磅
判定：部分实现。
证据：
- `tjufe-thesis.lua`: `ensure_pagebreak()`；标题改写为 `摘 要`
- `AbstractTitleCN`: SimHei size 32（三号），outlineLvl 0，center，firstLineChars 0，before/after 0，`line=400 exact`
说明：
- “另起一页”已实现
- “摘 要”本身带一个空格，满足字间空1个半角空格
- 但 `clone_header_with_style(1, 'Abstract Title CN', '摘 要')` 生成的是自定义 style header；在后处理中会套 `AbstractTitleCN`
- 这项总体可判部分实现，主要是最终是否完全匹配 Word 内的大纲/样式表现仍需人工终检

### D-16 中文摘要正文：下空一行；宋体/Times New Roman 小四；两端对齐；首行缩进2字符；段前后0；固定值20磅
判定：已实现。
证据：
- `tjufe-thesis.lua`：标题后 `inject_blank_lines(out, 1)`
- `AbstractBodyCN`：小四、两端对齐、首行缩进2字符、before/after 0、`line=400 exact`
- `postprocess_tjufe_docx.py / normalize_body_line_spacing()`：会把正文类样式统一保持为 `400 exact`
说明：
- 固定值20磅要求已按样式与后处理双重落地。

### D-17 中文关键词标题“关键词：”用小四黑体；内容小四；两端对齐；无缩进；固定值20磅；关键词之间全角逗号，最后无标点
判定：已实现。
证据：
- `tjufe-thesis.lua / make_keyword_para()` 已把标题词与内容词拆成不同 run
- `tjufe-thesis.lua / normalize_cn_keyword_content()` 会把中英文逗号、分号、顿号等归一为全角逗号，并去掉末尾标点
- `build_reference_docx.py` 已定义字符样式 `KeywordLabelCN`，真实样式名为 `Keyword Label CN`；样例导出的 `out/sample.docx` 中该样式为单实例，实际落到 `size=24`、`eastAsia=SimHei`
- `KeywordsLineCN`：正文宋体/Times New Roman 小四、无首缩进、`line=400 exact`
- `convert_tjufe.sh / scan_docx_or_fail()` 已检查 `cn_keywords_count` 与 `cn_keywords_punct_bad`
说明：
- 标题词/内容词拆分、标题词黑体小四、内容行样式、全角逗号分隔与末尾无标点，当前均已有前处理规范化与导出后审计，宜上调为已实现。

### D-18 英文摘要标题：另起一页；三号 TNR 加粗居中；固定值20磅
判定：已实现。
证据：
- `tjufe-thesis.lua`: `ensure_pagebreak()` + `Abstract`
- `AbstractTitleEN`: size 32，bold，center，`line=400 exact`
- `postprocess_tjufe_docx.py` 当前会把 `Heading1 + Abstract` 规范化为 `AbstractTitleEN`，且已移除把英文 `Abstract` 误改成中文摘要标题的旧逻辑

### D-19 英文摘要正文：下空一行；小四 TNR；首行缩进2字符；固定值20磅；英文标点后应有一个半角空格
判定：部分实现。
证据：
- `AbstractBodyEN`：小四 TNR、首行缩进2字符、`line=400 exact`
- 标题后空一行已建
- `convert_tjufe.sh / scan_docx_or_fail()` 已新增英文摘要细粒度审计：命中英文中文标点会报 `en_abstract_punct_bad`，命中英文标点后空格异常/多余连续空格会报 `en_abstract_spacing_bad`
问题：
- 样式、行距与基础标点/空格审计已实现；
- 但该项仍更适合保守判为“部分实现”，因为它目前属于导出后告警审计，不是逐字符自动修复。

### D-20 英文 Key Words：标题加粗；关键词小四 TNR；半角分号分隔，分号后空一格，最后不加标点
判定：已实现。
证据：
- `tjufe-thesis.lua / make_keyword_para()` 已把 `Key Words:` 标题词与内容词拆成不同 run
- `tjufe-thesis.lua / normalize_en_keyword_content()` 会把中文分隔符归一为半角分号，并规范为 `; `，同时去掉末尾标点
- `build_reference_docx.py` 已定义 `KeywordLabelEN`
- `KeywordsLineEN`：小四 TNR、无首缩进、`line=400 exact`
- `convert_tjufe.sh / scan_docx_or_fail()` 已检查 `en_keywords_count`、`en_keywords_punct_bad`、`en_keywords_spacing_bad`、`en_keywords_terminal_punct_bad`
说明：标题词/内容词拆分、半角分号、分号后空一格、末尾无标点目前均已有前处理规范化与导出后审计，宜上调为已实现。

### D-21 目录标题：另起一页；三号黑体居中；字间空1个半角空格；1.5倍行距；标题下空两行
判定：已实现。
证据：
- `tjufe-thesis.lua`: 目录前 `ensure_pagebreak()`；标题 `目 录`；`inject_blank_lines(out, 2)`
- `TOCHeading`: size 32，SimHei，center，`line=360 auto`

### D-22 目录到三级标题
判定：已实现。
证据：
- `make_toc_paragraph()`：`TOC \o "1-3"`

### D-23 一级目录项：四号，两端对齐，无首缩进，左0右0，1.5倍行距
判定：已实现。
证据：
- `TOC1`: 四号，both，firstLineChars 0，line 360 auto

### D-24 二级目录项：小四，两端对齐，左缩进2字符，1.5倍行距
判定：已实现。
证据：
- `TOC2`: size 24，小四；`leftChars=200`

### D-25 三级目录项：小四，两端对齐，左缩进4字符，1.5倍行距
判定：已实现。
证据：
- `TOC3`: size 24；`leftChars=400`

### D-26 页眉：从正文开始到论文结尾；内容为本章标题；黑体/TNR 五号居中；文武线
判定：部分实现，需人工终检。
证据：
- `collect_fixed_header_sections()` 只从正文、附录、参考文献、后记、成果页开始收集一级标题
- `header_xml(title)` 写入页眉标题与底部边框 `thinThickSmallGap`
- `Header` 样式：size 21（五号），SimHei/TNR
说明：
- 功能实现明确
- 但“文武线”视觉观感只能人工终检

### D-27 页脚：前置部分页码用大写罗马数字；正文从1开始阿拉伯数字；底端居中；TNR 六号；两侧无修饰线
判定：已实现。
证据：
- `configure_sectpr()`：`page_fmt='upperRoman'` / `page_fmt='decimal'`
- `front_break_p` 插入前置节
- `footer_xml()` 插入 PAGE 域
- `Footer` 样式：size 15（六号）center
- 未添加修饰线

### D-28 正文：宋体/TNR 小四，两端对齐，首行缩进2字符，段前后0，固定值20磅
判定：已实现。
证据：
- `build_reference_docx.py` 中 `body_p` 明确设为：两端对齐、首行缩进 2 字符、段前后 0、`line=400`, `lineRule='exact'`
- `Normal` / `BodyText` / `FirstParagraph` 等正文类样式均基于该组参数
- `postprocess_tjufe_docx.py / normalize_body_line_spacing()` 当前保持正文类样式为 `400 exact`，不再放宽为 `atLeast`

### D-29 正文：非章节末尾不要在底部留下大片空白
判定：未实现。
说明：未见 widow/orphan/分页压缩策略或版面回收逻辑。

### D-30 正文采用三级标题形式
判定：已实现。
证据：
- `tjufe-thesis.lua` 明确重写 Heading1/2/3

### D-31 每章另起一页
判定：已实现。
证据：
- `tjufe-thesis.lua` 对每个一级标题 `ensure_pagebreak()`

### D-32 一级标题：小三黑体居中；固定值20磅；序号与标题间 2 个半角空格；标题下空两行；编号“第1章”
判定：已实现。
证据：
- `Heading1`: size 30（小三），SimHei，center，exact 400
- `tjufe-thesis.lua`: `string.format('第%d章  %s', chapter_no, title)`
- `inject_blank_lines(out, 2)`

### D-33 二级标题：四号黑体左对齐；固定值20磅；序号与标题间 2 个半角空格；下空一行；编号“1.1”
判定：已实现。
证据：
- `Heading2`: size 28（四号），left，exact 400
- `tjufe-thesis.lua`: `'%d.%d  %s'`
- `inject_blank_lines(out, 1)`
- `normalize_heading_separator()` 再次规范空格

### D-34 三级标题：小四黑体左对齐；固定值20磅；序号与标题间 2 个半角空格；编号“1.1.1”
判定：已实现。
证据：
- `Heading3`: size 24（小四），left，exact 400
- `tjufe-thesis.lua`: `'%d.%d.%d  %s'`

### D-35 三级标题下标号顺序：（1）…①…第一…
判定：未实现。
说明：未见正文内部层级项目符号/编号体系的专门规范化逻辑；`tighten_list_indentation()` 只是一般列表缩进处理，不足以保证此顺序。

### D-36 表题：黑体/TNR 五号，居中，固定值20磅；按章编号“表1-1”；表号与表题空 4 个半角空格；正文中明确提及
判定：部分实现。
证据：
- `TableCaption`: size 21（五号），center，exact 400
- `format_caption_text('表', label, title, ...)`：使用 4 个半角空格
- `make_caption_label()`：正文章号 `1-1`，附录 `A1`
问题：
- “正文中明确提及”未校验

### D-37 表单位：五号黑体/TNR，右上方列出，固定值20磅
判定：已实现。
证据：
- `build_reference_docx.py` 中 `TableMetaLine` 已绑定 `black_r_10_5`：黑体/TNR、五号（`size=21`）、右对齐、`line=400 exact`
- `postprocess_tjufe_docx.py / is_unit_line()` 会识别单位行，并将其套为 `TableMetaLine`

### D-38 表内文字：五号宋体/TNR；固定值20磅；数字栏居中；左侧项目栏垂直居中；子项目水平垂直居中；三线表黑色单线 3/4 磅
判定：部分实现。
证据：
- `apply_three_line_table_style()`：表上/下边框 `sz=12`，表头下边框 `sz=8`，移除左右/内线，整体居中
- 单元格段落统一 center
问题：
- 没有对“表内五号字”做逐格强制设置
- 没有对“最左侧项目栏垂直居中”“子项目水平垂直居中”做细分列规则
- 线粗与规范“3/4磅”是否完全等值需人工换算确认

### D-39 表注/资料来源：小五，两端对齐，无缩进，固定值20磅，置左下方，顶左框编排
判定：部分实现。
证据：
- `build_reference_docx.py` 中 `SourceNote` 已改为宋体/TNR 小五：`song_r_9`，两端对齐，`line=400 exact`
- `postprocess_tjufe_docx.py / is_source_line()` 会识别资料来源并套 `SourceNote`
问题：
- 字体、字号、对齐和行距已落地
- 但“置左下方、顶左框编排”仍主要依赖相邻段落结构，不是绝对定位或版面级精确约束，因此仍判为部分实现

### D-40 续表：需在续表上方居中注明“续表”，续表表头重复
判定：已实现，跨页视觉观感仍建议终检。
证据：
- `clean_caption_title()` / `format_caption_text()` 支持 `续表`
- `postprocess_tjufe_docx.py / set_repeat_table_header()` 会对表格首行写入 `w:tblHeader`
- `convert_tjufe.sh / scan_docx_or_fail()` 已新增 `continued_table_header_missing` 与 `continued_table_caption_missing` 导出后审计
- `samples/continued-caption-probe.tex` 已验证成品中存在 `续表1-1    重复标题测试表`，且续表后的表格首行含 `w:tblHeader`
说明：
- 续表标题与表头重复的主链路已实现，并有导出后告警审计兜底
- 不同渲染器中的最终跨页视觉效果仍建议终检成品

### D-41 表与上下正文之间各空一行
判定：已实现，异常结构有导出后告警兜底。
证据：
- `postprocess_tjufe_docx.py` 在处理 `TableCaption` 时，会通过 `ensure_blank_paragraph_before()` 与 `ensure_blank_paragraph_after()` 为“表题 + 单位行 + 表格 + 来源行”这一可识别结构补入上下空行
- `convert_tjufe.sh / scan_docx_or_fail()` 已新增 `table_cluster_missing_blank_before` 与 `table_cluster_missing_blank_after` 告警，用于捕获导出成品中表格 cluster 前后空段缺失
说明：
- 标准表格簇的上下空一行主链路已实现
- 对异常输入结构，不再静默漏判，而是通过导出后告警提示人工处理

### D-42 图一般随文编排，先见文字后见图
判定：部分实现。
说明：已补入保守导出后审计 `figure_text_reference_missing_before:<label>`，用于发现图题出现前在同一小节邻近正文中未看到对应“图<label>”引用的情况。该规则只做 WARN，不自动重排图片，也不能完全替代人工判断图文逻辑。

### D-43 图题：五号黑体/TNR，居中，固定值20磅；按章编号“图1-1”；图号和图题空4个半角空格
判定：已实现。
证据：
- `ImageCaption`: size 21, center, exact 400
- `format_caption_text('图', label, ...)` 4 个空格
- `make_caption_label()` 正文章号 `1-1`，附录 `A1`

### D-44 图形要标明计量单位
判定：部分实现。
说明：已补入保守导出后审计 `image_unit_unconfirmed:<label>`：当图题或邻近文本疑似统计图、趋势图、占比图、变量图等计量型图形，且邻近文本未发现“单位/计量单位/数据单位”或常见单位标记时给出 WARN。该规则无法读取图片内部文字，因此仍需人工确认图内单位是否已标明。

### D-45 分图：可用(a)(b)(c)…；主图名在所有分图下方正中
判定：部分实现。
说明：已补入保守导出后审计 `subfigure_labels_missing:<label>`：当同一图题关联到多个图片段，但邻近文本未发现 `(a)`/`（a）` 等分图标签时给出 WARN。脚本不会自动重排分图，也不能判断主图名视觉位置是否绝对居中，因此仍建议人工终检。

### D-46 续图：后续页重复图题并加（续）
判定：已实现，跨页视觉观感仍建议终检。
证据：
- `clean_caption_title()` / `format_caption_text()` 支持 `续图`
- `convert_tjufe.sh / scan_docx_or_fail()` 已新增 `continued_image_caption_missing` 导出后审计，用于发现同章同标题重复但后一次未写成 `续图...` 的成品错误
- `samples/continued-caption-probe.tex` 已验证成品中存在 `续图1-1    重复标题测试图`
说明：
- 续图重复标题改写主链路已实现，并有导出后告警审计兜底
- 不同渲染器中的最终跨页视觉效果仍建议终检成品

### D-47 图注/资料来源：小五，两端对齐，无缩进，固定值20磅，置图左下方
判定：部分实现。
证据：
- `build_reference_docx.py` 中 `SourceNote` 已改为宋体/TNR 小五：`song_r_9`，两端对齐，`line=400 exact`
- `SourceNote` 应用于图片后相邻来源行
问题：
- 字体、字号、对齐和行距已落地
- 但“置图左下方”仍依赖相邻段落结构，不是绝对布局

### D-48 图与上下正文之间各空一行
判定：已实现，异常结构有导出后告警兜底。
证据：
- `postprocess_tjufe_docx.py` 在处理 `ImageCaption` 时，会通过 `ensure_blank_paragraph_before()` 与 `ensure_blank_paragraph_after()` 为“图题 + 图片段落 + 来源行”这一可识别结构补入上下空行
- 对无 caption 的正文图片段，`postprocess_tjufe_docx.py` 也会补入前后空段兜底
- `convert_tjufe.sh / scan_docx_or_fail()` 已新增 `image_cluster_missing_blank_before` 与 `image_cluster_missing_blank_after` 告警，用于捕获导出成品中图片 cluster 前后空段缺失
说明：
- 标准图片簇与无 caption 正文图片段的上下空一行主链路已实现
- 对异常输入结构，不再静默漏判，而是通过导出后告警提示人工处理

### D-49 数学表达式：另起一行；宋体小四居中
判定：部分实现。
证据：
- `paragraph_has_math_para()` -> `EquationBlock`
- `EquationBlock`: center, body_r size 24（小四）
问题：
- 中文字体为宋体、西文为 TNR 可以成立；但具体数学字体仍依赖 Word/OMML/Cambria Math

### D-50 公式编号：右边行末；按章顺序；格式（1.1）；不加虚线
判定：已实现。
证据：
- `append_equation_number()`：中心 tab + 右对齐 tab + 右端编号
- `eq_label = f'（{chapter}.{equation_no}）'`
- 未插虚线

### D-51 含公式段落可设置段前后6磅、1.5倍行距
判定：已实现。
证据：
- `build_reference_docx.py` 中 `equation_block_p` 已显式设为：`before=120`、`after=120`（各 6 磅），`line=360`, `lineRule='auto'`（1.5 倍行距）
- `EquationBlock` 样式绑定该组参数
- `postprocess_tjufe_docx.py` 会对独立公式段套用 `EquationBlock`

### D-52 注释采用脚注；脚注文本小五、左对齐、无缩进、段前后0、单倍行距；①②③……；每页重新编号
判定：已实现。
证据：
- `build_reference_docx.py` 中 `FootnoteText` 已设为：小五、左对齐、无缩进、段前后 0、`line=240`, `lineRule='auto'`（单倍行距）
- `postprocess_tjufe_docx.py / configure_footnote_pr()` 设定 `numRestart = eachPage`
- `FOOTNOTE_NUMFMT = 'decimalEnclosedCircleChinese'`，对应圈号样式

### D-53 参考文献标题：另起一页；黑体小三；大纲级别1级；居中；无缩进；段前2行段后2行；固定值20磅
判定：已实现。
证据：
- `tjufe-thesis.lua`: `ensure_pagebreak()` + `参考文献`，且不额外给该标题注入空段
- `build_reference_docx.py`: `ReferencesHeading` 已设为居中、`before=800`、`after=800`、`line=400 exact`、`outlineLvl=0`
- `postprocess_tjufe_docx.py`: 对 `ReferencesHeading` 仅清理标题前相邻空段，不再把标题 spacing 依赖为插空段近似

### D-54 参考文献内容：楷体/TNR 小四，两端对齐，无缩进，固定值20磅
判定：已实现。
证据：
- `build_reference_docx.py` 中 `Bibliography` 已设为：楷体/TNR 小四、两端对齐、无缩进、`line=400 exact`
- `postprocess_tjufe_docx.py / normalize_body_line_spacing()` 当前保持 `Bibliography` 为 `400 exact`

### D-55 参考文献要精选、直接相关
判定：不属于脚本范围。

### D-56 参考文献应遵照 GB/T 7714-2015
判定：部分实现。
证据：
- `convert_tjufe.sh / scan_bibliography_or_fail()` 已增加 bibliography 条目最小静态审计
- 已检查关键字段缺失与明显异常：`missing-author`、`missing-title`、`missing-year`、`year-XXXX`
说明：
- 当前实现是“最小静态审计 + 异常拦截/告警”，不是完整的 `GB/T 7714-2015` 逐项格式化器。

### D-57 参考文献序号左顶格、方括号、条目末尾实心点
判定：部分实现。
证据：
- `convert_tjufe.sh / scan_docx_or_fail()` 已扫描参考文献成品段落
- 已检查条目是否以方括号序号开头，以及是否以句点类结尾；异常时输出 `references-bad-prefix`、`references-bad-ending`
说明：
- 当前是成品审计，不是逐条自动重写。

### D-58 外文文献编排各细则（姓前名后、标题首字母大写、文献标识、出版年份等）
判定：部分实现。
证据：
- `convert_tjufe.sh / scan_bibliography_or_fail()` 已增加外文文献最小静态检查
- 已覆盖的可静态检查项包括：`missing-year`、英文标题首词首字母异常、部分作者串模式异常
说明：
- “姓前名后”“文献标识”“出版社/期刊名规范缩写”等仍难以仅靠静态规则 100% 自动判定，仍建议人工终检。

### D-59 3 位以上作者只列前 3 位，其后加“等”或 `et al.`
判定：部分实现。
证据：
- `convert_tjufe.sh / scan_bibliography_or_fail()` 已增加作者人数与作者串静态检查
- 检测到 `4` 位及以上作者但未见“等”或 `et al.` 时输出相应告警
说明：
- 该规则目前依赖 `.bib` 作者串解析与启发式判断，复杂姓名场景仍建议人工复核。

### D-60 附录标题：另起一页；黑体/TNR 小三；大纲1级；居中；无缩进；固定值20磅；标题下空两行
判定：已实现。
证据：
- `tjufe-thesis.lua`: 附录前分页；标题 `附录A  标题`；`inject_blank_lines(out, 2)`
- `Heading1`: 小三黑体居中 exact 400

### D-61 附录内容格式同正文文本
判定：部分实现。
说明：附录正文未专门切成单独 body style，默认沿正文流；若原文有局部特殊样式，仍依赖 Pandoc 导出。

### D-62 附录序号用 A、B、C…；如附录A、附录B
判定：已实现。
证据：
- `appendix_no` -> `string.char(string.byte('A') + appendix_no - 1)`
- Heading1: `附录%s  %s`

### D-63 附录中的公式、图、表编号分别用 A1 / 图A1 / 表A1
判定：已实现。
证据：
- `chapter_label_from_heading()` 返回字母章号
- 公式编号：`（A1）`
- `make_caption_label()` 对附录返回 `A1`
- 图/表标题前缀由 `format_caption_text()` 生成 `图A1` / `表A1`

### D-64 在学期间成果：博士需添加；另起一页；标题黑体小三；居中；固定值20磅；标题下空两行
判定：已实现。
证据：
- `tjufe-thesis.lua`: `ensure_pagebreak()` + 标题重写 + `inject_blank_lines(out, 2)`
- 使用 `Heading1`

### D-65 后记标题：另起一页；小三黑体居中；字间空1个半角空格；固定值20磅；标题下空两行
判定：已实现。
证据：
- `tjufe-thesis.lua`: `后 记`（带空格）+ 分页 + `inject_blank_lines(out, 2)`
- `Heading1`: 小三黑体 exact 400

### D-66 后记正文：楷体/TNR 小四；两端对齐；首行缩进2字符；固定值20磅
判定：已实现。
证据：
- `build_reference_docx.py` 中 `AcknowledgementsBody` 采用 `body_p + kaiti_12`，即楷体/TNR 小四、两端对齐、首行缩进 2 字符、`line=400 exact`
- `postprocess_tjufe_docx.py / normalize_body_line_spacing()` 当前保持 `AcknowledgementsBody` 为 `400 exact`

### D-67 后记正文下空两行；段尾打印作者姓名；下一行写作日期；黑体/数字TNR；小四；右对齐；固定值20磅
判定：已实现。
证据：
- `tjufe-thesis.lua / flush_ack_signature()`：会先补空两行，再写署名行与日期行本体，不附加“作者：”“写作日期：”前缀
- `build_reference_docx.py` 中 `SignatureLine` 已为黑体/数字 TNR 小四、右对齐、`line=400 exact`
说明：
- 该条所要求的空行、署名行、日期行、字体、字号、对齐与行距当前已落地
- 最终视觉观感仍可在 Word/WPS 成品中例行终检，但不再构成“部分实现”依据

### D-68 印刷要求：自中文摘要起双面印刷，之前部分单面印刷
判定：不属于脚本范围。
说明：Word 文档生成脚本不控制最终物理打印模式。

### D-69 纸型与页边距：16 开 184mm×260mm；全文行距20磅（目录1.5倍）；页边距上20下15左20右20；页眉1.5cm；页脚1.75cm
判定：已实现。
证据：
- 页面与页边距常量已落地：`PAGE_W=10431`, `PAGE_H=14740`, `TOP=1134`, `BOTTOM=850`, `LEFT=1134`, `RIGHT=1134`, `HEADER=850`, `FOOTER=992`
- `build_reference_docx.py` 中正文、摘要、参考文献、后记等主体文本样式统一为 `line=400 exact`
- 目录样式 `TOCHeading / TOC1 / TOC2 / TOC3` 维持 `line=360`, `lineRule='auto'`（1.5 倍行距）

### D-70 封皮与装订：线装或热胶装订
判定：不属于脚本范围。

---

## E. 第四章：附则

### E-1 适用范围：适用于天津财经大学工商管理/管理科学与工程博士硕士及在职申请硕士学位人员，非汉语参照执行
判定：不属于脚本范围。

### E-2 规范由研究生院负责解释
判定：不属于脚本范围。

### E-3 规范自颁布之日起执行
判定：不属于脚本范围。

---

## F. 逐类遗漏检查结论

### F-1 已明确实现的项目
1. 文档结构顺序主骨架
2. 封面与扉页的主要版式结构
3. 摘要/Abstract/目录/正文/参考文献/附录/成果页/后记的分页
4. 目录到三级
5. 正文三级标题与章/节/目编号
6. 页眉按章节切换
7. 前置部分罗马页码、正文阿拉伯页码
8. 页眉文武线的 OOXML 逼近实现
9. 正文/摘要/参考文献/后记正文固定值 20 磅
10. 脚注单倍行距、每页重新编号与圈号样式
11. 公式段前后 6 磅 + 1.5 倍行距，以及公式按章编号与附录编号
12. 图表按章编号与附录编号
13. 交叉引用与书签
14. 三线表主处理
15. 表单位行黑体五号、来源行宋体小五的样式落地
16. 摘要字数、关键词数量、正文最小字数、参考文献数量/近年占比/外文占比等审计
17. 摘要区禁图、禁表、禁公式审计

### F-2 部分实现的项目
1. 中文摘要/关键词与英文 Abstract 的细粒度 run 级字体、标点与显示效果，仍有部分项目依赖导出后审计或人工终检
2. 封面顶端等字段对 metadata 的依赖
3. 表注/图注位置细节（左下方/顶左框）与续图续表跨页视觉观感
4. 后记签名行的最终视觉观感
5. GB/T 7714 当前仍以静态审计与部分异常拦截为主，不是完整格式化器

### F-3 未实现的项目
1. 三级标题下后续编号序列规范
2. 更完整的 GB/T 7714 细则级自动规范化

说明：正文底部大片空白与分页异常已不再是完全未实现项；当前通过 `main_body_excessive_blank_paragraphs:<n>`、`main_body_manual_page_break`、`main_body_page_break_before_nonheading` 做保守导出后 WARN 审计。图文先后、图单位/分图结构也已不再是完全未实现项；当前通过 `figure_text_reference_missing_before:<label>`、`image_unit_unconfirmed:<label>`、`subfigure_labels_missing:<label>` 做保守导出后 WARN 审计。在学期间成果细则也已通过 `research_outputs_missing:doctor`、`research_outputs_empty:doctor`、`research_output_type_marker_missing:<n>`、`research_output_detail_note_missing:<n>` 做保守导出后 WARN 审计。上述内容侧规则仍需人工确认内容逻辑与视觉效果。

### F-4 不属于脚本范围的项目
1. 摘要内容质量、英文语法质量
2. 正文学术质量与逻辑性
3. 是否存在抄袭/引用不当
4. 双面/单面打印、装订方式
5. 规范适用范围、解释权、生效日期等附则

### F-5 需人工终检的项目
1. 页眉“文武线”最终视觉效果
2. 目录域在 Word/WPS 中首次更新后的样子
3. 公式段的真实视觉松紧
4. 三线表线粗在不同渲染器中的观感
5. 封面/扉页签名栏与标题行的视觉密度
6. 续图续表跨页后的最终视觉观感（重复标题与续表表头已有导出后审计）

---

## G. 最终回答：这次是否还存在“规范遗漏”

如果问题是：
“在当前提取出的规范正文里，是否还有整条整款完全没被登记进对照表？”

本报告结论：
- 按当前 DOCX 规范正文可见内容，本报告已经逐章逐条全部登记，没有像上一版那样只挑重点。
- 也就是说，当前这版报告已经把可见规范条款全部纳入建账。

如果问题是：
“是否所有条款都已经被脚本完全满足？”

答案是：没有。
- 有一批条款属于内容/学术/印刷要求，不属于这 4 个脚本范围。
- 有一批条款属于排版细节，目前只是部分实现。
- 还有一批条款必须看最终 Word/WPS 成品，不能只靠代码文本认定。

因此最终结论应写成：

**本次已完成对规范正文可见条款的全文逐条建账；条款本身未再遗漏，但实现层面仍存在“部分实现、未实现、脚本外、需人工终检”的明确差异。**

---

## H. 下一步建议

如果目标是“继续收敛到尽可能接近逐字满足规范”，建议按以下优先级修正：

### H-1 P1 必改
1. 继续补强 GB/T 7714 复杂细则等高价值审计项；图文先后、图单位、分图结构当前已有保守 WARN，但仍可继续降低误报/漏报
2. 继续补强作者姓名格式、成果作者排序等内容侧结构审计；在学期间成果页缺失/为空/缺少括号说明等已具备保守 WARN 审计
3. 继续增强续图续表跨页显示的长样例/强制分页视觉检查；重复标题与续表表头已具备最小样例和导出后审计

### H-2 P2 建议补强
1. 增加英文关键词“分号后空一格、末尾无标点”等逐字符审计
2. 继续补强作者姓名格式、成果作者排序等内容侧结构审计；在学期间成果页缺失/为空/缺少括号说明等已具备保守 WARN 审计
3. 继续增强续图续表跨页显示的长样例/强制分页视觉检查；重复标题与续表表头已具备最小样例和导出后审计
4. 继续增强导出后 DOCX 自动审计脚本，覆盖 `lineRule`、`pgNumType`、header border、caption cluster 空行、续图续表重复标题等关键字段

---

## I. 2026-04-23 第二轮修正回执

本节用于覆盖 H-1 中已经落地的修正项，避免报告状态落后于代码状态。

### I-1 已完成修正
1. `normalize_body_line_spacing()` 已从“把正文类样式放宽为 `atLeast 20pt`”改为保持 `400 exact`。
2. `FootnoteText` 已改为单倍行距：`line=240`, `lineRule='auto'`。
3. `EquationBlock` 已显式写入段前/段后 `120`（6 磅）与 `line=360`, `lineRule='auto'`（1.5 倍行距）。
4. `TableMetaLine` 已改为黑体五号：`SimHei`, `size=21`。
5. `SourceNote` 已改为宋体小五：`SimSun`, `size=18`。
6. `SignatureLine` 已改为黑体/数字 TNR 小四右对齐：中文字体 `SimHei`, `size=24`。
7. `CoverDate` 已改为非粗体宋体四号，并加入导出后样式审计。
8. 已移除后处理阶段把英文 `Abstract` 误改成中文摘要标题的逻辑。
9. 已补入图表前后空一行的后处理规则。
10. 已新增 caption cluster 空行导出后审计：`table_cluster_missing_blank_before/after` 与 `image_cluster_missing_blank_before/after`。
11. 扉页 `title_page_info_rows` 已接入 `normalize_title_page_rows()`，可结合 `title_page_compact_name_labels` 对姓名类字段做更通用的去空格规范化。
12. 英文 `Key Words` 已增加前处理规范化：半角分号、分号后空一格、末尾无标点。
13. `convert_tjufe.sh` 已增强导出后 DOCX 审计，会直接检查上述关键样式值是否真正落到成品里，并新增英文摘要/关键词细粒度告警、caption cluster 空行告警、`continued_table_header_missing`、`continued_table_caption_missing`、`continued_image_caption_missing`、`main_body_excessive_blank_paragraphs:<n>`、`main_body_manual_page_break`、`main_body_page_break_before_nonheading`、`figure_text_reference_missing_before:<label>`、`image_unit_unconfirmed:<label>`、`subfigure_labels_missing:<label>`、`research_outputs_missing:doctor`、`research_outputs_empty:doctor`、`research_output_type_marker_missing:<n>` 与 `research_output_detail_note_missing:<n>` 审计。

### I-2 已完成验证
1. `python3 -m py_compile scripts/build_reference_docx.py scripts/postprocess_tjufe_docx.py` 通过。
2. `luac -p filters/tjufe-thesis.lua` 通过。
3. `bash -n convert_tjufe.sh` 通过。
4. 真实样例已跑通：
   - 输入：`samples/sample-thesis.tex`
   - metadata：`samples/sample-metadata.yaml`
   - 输出：`out/sample-20260423-rerun.docx`
   - 状态：导出成功，且通过增强后的成品审计。

### I-3 现阶段仍保留的差异
1. 中文关键词与英文 Abstract 的细粒度标点/显示效果虽然已有改进，但部分项目仍更适合人工终检；英文摘要标点空格当前主要通过告警审计而非自动修复。
2. 后记签名行的最终视觉观感仍建议以 Word/WPS 成品例行终检。
3. 页眉“文武线”、目录首次更新、公式视觉松紧、续图续表跨页视觉观感，仍属于成品显示层面的人工终检项。
4. GB/T 7714 复杂细则等自动审计仍未补齐；正文底部大片空白与分页异常、图文先后、图单位/分图结构、在学期间成果页缺失/为空/缺说明已补入保守 WARN 审计；续表表头与续图/续表重复标题当前已具备后处理设置、最小样例与导出后告警审计，但跨页视觉观感仍建议看成品。
