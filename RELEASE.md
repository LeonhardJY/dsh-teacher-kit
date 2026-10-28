# 发布清单（v0.18.0）

面向：中国大陆中学老师，Windows，DeepSeek Harness 桌面版或 `dsh web`。

## 1. 这一版能做什么

面板里 8 张卡（都是一键出成品 + 自动落目录 + 结果卡带预览和「打开所在文件夹」）：

- **出试卷**：粘贴题干 →（可选）套学校空白 `.docx` 模板 → 8 开横排两栏 / A4 等 → 答题卡（A4 竖版单栏）
  → 参考答案页（老师卷口径：题号 + 答案 + 题干摘要 + 解析）→ 印前检查（纸型/方向/页边距/页码域/题号/分值/字体/残留占位符）
  → Word + PDF；可勾 **A/B 两套**（B 卷选项打乱、答案跟着改）与 **教师卷**（答案与解析印在每题下面），两个一起勾一次出四份。卷尾整段 `【答案】1—5 BCBDA` 会被当答案区收走。
- **题库组卷**：直接读 `.xlsx/.xls/.csv/.docx` 题库 → 组卷（表头模糊匹配，合并单元格里的多选项自己拆）；
  可按题型挑题、分套出、套学校模板，A/B 与教师卷也能一起要。
- **听写单**：单词表 → A4 听写纸（英→中 / 中→英 / 混排，1–3 栏，可附答案页）；可**顺手念一段听写音频**
  （Windows 自带语音、离线、每条留 4–10 秒、念 1–3 遍，出 `.wav` 投教室音箱）。
- **教案**：草稿 → 表格式 / 流程版教案（字段名与环节模糊识别）。
- **做课件**：粘贴内容 →（可选）套学校 `.pptx` 模板 → 自动分页 → 「讲：」进备注 → PPTX；
  也可以出**单文件网页课件 HTML**（←/→ 翻页、N 看讲稿、Ctrl+P 一页一张、零外链）。
- **文件加工**：净化格式（超链接/彩色/高亮/字体/段距/空段/批注；下划线默认保留）、**批量净化**、
  合并多份（带图、自动分页、保留自动编号）、提取文字、**检查一份卷子**（只做印前检查）。
- **模板**：把上学期那份卷子反推成空白模板（纸型/栏数/边距/字号/页眉页脚占位符），供「出试卷」复用；
  模板自带原卷的页眉页脚，班级/老师留 `{{className}}`/`{{teacherName}}` 占位符，出卷时填了就印上。
- **整套**：一节课四件套——试卷（题库或粘题干）+ 听写单（词表文件）+ 教案（草稿）+ 课件（内容 / pptx / html），
  填了哪栏出哪件，四件并行、**一件失败不拖累另外三件**，结果卡逐件写清做到哪儿，只给一张试卷预览。

## 1.1 0.1.0 之前修掉的硬伤（评审 + 子代理复核）

- 双栏卷子里表格跨栏、答题卡节继承卷子的横排+两栏（第一份卷子就会中）；
- 大题标题「一、听力（共 20 小题，每小题 1 分，满分 20 分）」整行被吞、满分算错；
- 「（一）文言文阅读」被当大题 → 印前检查一路报跳号，老师只能删自己的小标题；
- 合并带批注/超链接的 Word → Word 报「文件可能已经损坏」；
- 「1. 2. 3.」数字大纲被当小标题 → 课件一半页面只有标题没有正文；
- 课件套模板后旧分节名/自定义放映指向已删页；
- 净化默认删下划线（英语划线题消失）、批注跟着发给学生；
- 默认产物精确到分钟 + 覆盖写 → 一分钟内出两份卷子静默覆盖；已改到秒 + 自动 -2。

## 1.2 0.1.0 真机装上以后修的

真机（本机 DSH 窗口 + CDP）跑出来的问题，假 ctx 的 smoke 一个都照不出来：

- 面板注册用的是 `conversation.input.activity`（单槽、已被语音输入占着）→ 一装就起不来，
  界面卡在「Loading plugins…」。改用 list 槽 `conversation.input.right`，且**必须带 `id`**。
- 宿主回包是信封 `{ok, data}`，面板原来直接读顶层 → 真机全是 undefined。已按信封解析（含失败分支的 code/hint）。
- 导 PDF / 逐页预览在 asar 里找 `libreoffice-kit.exe` 报 ENOENT：改成先试 `app.asar.unpacked` 的绝对路径。
- `inject` 为空导致注册早于 webServer/tools 就绪被静默跳过；`register` 被拆下来传丢 `this`；
  真 `defineTool` 拒绝 `required:false`。三条都在 `lib/index.js` 里修好，并加了 `~/.dsh/teacher-kit/boot.log`。
- 逐页预览改为**从导出的 PDF 出图**：预览即成品，双栏文档预览里那圈边框彻底没了。
- 8 开两栏卷子的**标题块现在跨整页**（标题节单栏 + 正文节两栏的连续分节符）；
  一道题的**题干与选项不再被栏缝切开**（`keep_with_next`）。
- 印前检查的栏数改为看正文节（标题节是单栏），页脚文字按节去重。
- 面板人性化：标签页代替大卡、右上角 ✕、题干框自动聚焦、按钮置灰、结果卡带「打开所在文件夹」。

## 1.3 0.2.0 新增的（功能铺开）

- 面板从 3 张卡扩到 6 张：**题库组卷 / 听写单 / 教案** 三个新引擎模块（`question_bank.py`、`handout.py`、`plan.py`），
  外加**网页课件**（`lesson_html.py`）。
- 宿主新增两个路由：`POST /list-dir`（目录浏览，只列指定类型、跳 `~$` 临时文件）与
  `GET /file?path=`（只回图片，用于结果卡预览）——面板因此有了**文件选择器**，不用再手抄绝对路径。
- `/render` 支持 `pdfPath`：已经导过 PDF 就不再让 LibreOffice 转第二遍。
- 结果卡内嵌**第一页预览图**（从导出的 PDF 出图，所见即所得）。
- 第 7 张卡「模板」：`template:make` 接进面板。
- 模板反推的栏数 bug：`analyse_docx` 原来只看 `sections[0]`（我们自己出的卷子第一节是单栏标题节）→ 反推出 1 栏。
  改成取所有节里最大的栏数。
- `_flatten_sections(document)`：模板里残留的段落级 `w:sectPr` 会在正文里制造空节，套模板出卷时先压平再重新分节。
- 教案自己写页脚（`第 X 页 共 Y 页　授课教师：____`），不再继承卷子版的内置页脚「出卷老师」。

## 1.4 0.3.0 修的（教师视角真人评测 + A/B 卷）

派子 agent 扮七年级英语老师走 HTTP 全流程评测（报告 `.verify/persona-report.md`），修掉：

- **P0：老师把答案整段贴在卷尾 → 答案被当成题目**（56 题变 67 题、答案印进学生卷第 5 页、答案页没生成、
  印前检查仍然 ok）。新增 `_split_answer_block` / `_parse_answer_block`：整段 `【答案】` 或「1—5 BCBDA」
  区间式都认，学生卷不再印答案，答案按题号回填；结果里给 warning `answer-block`（认出 N 条）。
  判据关键：标记行自己带内容时只看这一行（`【答案】A` 是逐题答案，不能把后面的题切走）。
- 勾了答案页但一道答案都没有 → `answers-missing` 中文错误 + 写法提示；
  **默认带的答案页不会因为没答案把整份卷子拦掉**（降级成 warning，不生成答案页）。
- 答案页默认升级成**老师卷口径**（`answerDetail` 默认开）：`19. A　— Can you swim?…　解析：…`。
- 听写单把词表首行标题当第 1 个词条（21/20）→ `skipped-title` warning 并跳过。
- `docx:merge` 参数名同时认 `inputs` 与 `paths`（老说明书写的是 paths），`inputs-required` 带中文 hint。
- 文件加工链路的 `file-not-found` / `unsupported-source` 全部换成中文 + hint（PDF、老 `.doc`、密码 docx 各有话术）。
- **A/B 两套卷**：`exam:build-ab` → `serialize_paper` 把解析结果反序列化后再走一遍 `build_exam`，
  B 卷 `_shuffle_options` 洗选项并同步改答案字母（洗成原样就轮转一位）；无带选项的题时 warning `ab-no-shuffle`。
  A 路径先算出来，B 路径由 `_variant_path` 派生（「（A 卷）」→「（B 卷）」），不会落到「测试卷（B 卷）_时间戳」。
- 面板：出试卷加「出 A/B 两套」勾选框，结果卡同时给 A/B 的 Word 与 PDF 四个「打开所在文件夹」。

## 1.5 0.4.0 修的（老师「印出来不像正经卷子」的那几处）

- **答题卡给作文题留作答格**：以前「六、书面表达」在答题卡上只有一行标题，学生没地方写。
  现在作文小节自动排 **12 行单列作答格**（行高 9mm、卡宽内、提示语与格子拴在一起不跨页），
  `essayLines` 可调；作文节不再出题号格（学生不往格子里写字母），后续小节题号仍连续。
- **学生卷 + 教师卷一次出两份**（`exam:build-teacher`）：教师卷把「【答案】」「【解析】」直接印在每道题下面，
  不印答题卡也不另起答案页；文件名带「（教师卷）」、与学生卷同一分钟成对落盘。面板加勾选框
  「另出一份教师卷（答案和解析印在题后）」。
- **页脚跨纸型编号**：一份文件里同时有 8 开卷子与 A4 答题卡/答案页时，正文页脚**只印「第 N 页」**。
  `NUMPAGES` 会把 A4 那几页算进 8 开（8 开 6 张写成「共 8 页」，老师一眼看出不对）；
  `SECTIONPAGES` 在 LibreOffice 转 PDF 时恒等于 1（实测「第 2 页 共 1 页」）。单纸型（不带卡不带答案）仍保留「共 N 页」。
  答题卡/答案页各自 `w:pgNumType w:start="1"` 从 1 重新编号（`w:pgNumType` 必须插在 `w:cols` 之前，否则 Word 报文件损坏）。
- 印前检查：`NUMPAGES` 或 `SECTIONPAGES` 任一存在就算「有总页数」，多节文档不再误报 `footer-no-total`。
- 选项行判断改看**整行长度**（原来只看单个选项）：四个选项各自短、加起来超一栏宽时会被 Word 折行，
  最后一个选项被拆成两行；现在超长整行改成一行一个选项。
- **教案抬头表改成 3 行**：课题独占一行跨 3 列，第 2 行「课型|课时」，第 3 行「授课教师|班级」。
  原来「课时」落在正文单列表里，渲染成「课时」+下一行「1 课时」，老师以为是漏填；列宽写死并按格宽自动缩字号（8–11pt）防折行。
- **学生活动整列空白时明说**：表头改成「学生活动（老师自填）」，表下补一行 9pt 灰字说明这是留给老师按本班实际写的
  （不编造学生反应）；只空几格时不改表头，避免每份教案都像模板。返回值新增 table 统计，warning 只在整列全空时发。
- **从 Word 卷子里读题干**：出试卷卡的题干行右侧有「从 Word 卷子里读题干」，选一份旧卷子自动把正文填进题干框
  （标题为空时取正文第一行），不用再复制粘贴。
- **每张卡各记各的**：切到「教案」不会再看到「出试卷」填过的卷子名称与题干（真机抓到过：教案直接沿用了上一张卡的标题）。
  切回来仍保留自己填的内容。

## 1.6 0.5.0 新增的（听写单自带音频）

- **听写单可以直接出音频**（`handout:audio`）：勾「同时生成听写音频」，卷子和 `.wav` 一起落盘，
  结果卡上有「打开「听写音频（wav）」所在文件夹」。念法是「第 1 题 … tall … tall …」——中文条目用中文语音报题号，
  英文条目念「Number one」；每条之间停你选的那几秒（4/6/8/10），开头 3 秒、结尾 2 秒，条目内重复处停 0.8 秒。
- **离线零安装**：走 Windows 自带的 `System.Speech`（本机实测 zh-CN「Microsoft Huihui Desktop」、
  en-US「Microsoft Zira Desktop」），不联网、不装 TTS 引擎、不打包任何模型。
- **一个 PowerShell 进程念完整张单子**：先按语音分组写句子表，再一次性合成，Python 侧用 `wave` 拼静音 —— 
  4 条实测 4 秒出 44 秒音频（22.05kHz/16bit/单声道，约 2.6MB/分钟）。
- **失败不拦交付**：没装语音 → `tts-no-voice`（说明缺中文还是英文）、合成超时 → `tts-timeout`、
  进程报错 → `tts-failed`（带 stderr 末行）；这些只影响音频那一步，听写单本身照出。超过 120 条或 30 分钟会先截断并给 `too-many` / `too-long` 提醒。
- 音频设置（每条停几秒、念几遍）在听写表单里直接选，默认 6 秒 / 2 遍。
- **题库组卷可以按题型挑题**（`bank:inspect` + `bank:build` 的 `options.pick`）：
  先点「看看题库里有什么」当场读出「一共 5 题：单项选择 4 题（每题 2 分） · 填空 1 题（每题 2 分）」，
  再在「挑题」里一行写一种题型（`单项选择 10` / `完形填空 1`），不填就全用。
  挑不满只警告（`pick-short`）不拦交付；题型写错直接 `pick-no-type` 并列出能挑的题型；
  题数写错 → `pick-invalid`。挑出来的题按题型顺序排，大题结构与卷子一致。
- **出试卷/题库组卷可以选纸型**：面板「纸型 / 版式」下拉 → `page` 参数（8K/A3 横向两栏或一栏、16K 一栏、A4 一栏），
  没选时仍是 8 开横向两栏。实测 16K→184×260 竖版单栏、A3→420×297 横排（标题单栏节 + 正文两栏节，渲染 3 页正常）。
- **听写单可以直接选词表文件**（`handout.read_word_list`，`handout:build` / `handout:audio` 都认 `source`）：
  Word（所有表格行 + 散装段落）、Excel（`.xlsx/.xlsm/.xltx`，逐行非空单元格拼接）、CSV（依次试 utf-8-sig/utf-8/gbk/gb18030）、TXT。
  表头行（英文/中文/单词/释义/序号/Unit/word/meaning…）与「Unit 4 单词表」这种整行标题会自动跳过，
  不会变成第 1 个词条；`.pdf` 等读不了的格式报 `unsupported-source` 并提示存成 docx/xlsx/csv/txt。
  面板听写卡有「或者直接选一个词表文件（Word / Excel / CSV）」的浏览按钮，选了文件就不用手打。
  文件选择器（`/list-dir`）固定 230px 高、不许被面板的 flex 压缩 —— 真机第一版被挤成一条缝（列表只剩半行），
  已修并复验：点目录进得去、点文件把绝对路径填回输入框、点「取消」收起。
## 1.7 0.6.0 新增的（听力录音稿 → 音频；记住上次设置；结果卡显示用时）

- **听力录音稿念成音频**（`audio:script`）：出试卷卡在题干下面多了一个「听力录音稿」框，一行一句贴进去
  （中英混排也行），勾「把听力录音稿念成音频（wav，上课播放）」，出卷时顺手念出 `.wav`。
  中英混排**按段落切语音**（`_segments` → en/zh/en），中文句子用中文语音、英文句子用英文语音，不是整段硬套一个声音。
  `number=True` 会剥掉行首「1. / 1、」再念「Number one.」，每句之间 `gapSec`（默认 1.2 秒）、每行之后 `pauseSec`（默认 4 秒），
  开头 `leadSec` 2 秒、结尾 `tailSec` 2 秒。参数范围都有校验（`bad-gap`/`bad-pause`/`bad-lead`/`bad-tail`/`bad-rate`），
  超过 400 行或 40 分钟先拦（`too-long`），没装语音 `tts-no-voice`、超时 `tts-timeout`、进程报错 `tts-failed`（带 stderr 末行）。
  默认落 `教师助手/听写/<卷名> 听力音频_年月日-时分秒.wav`（名字里已经有「音频/录音」就不重复叠字）。
- **面板**：结果卡新增一行「听力音频：29 秒，4 句（用 Windows 自带的语音念的，点开就能放）」+ 「打开「听力音频（wav）」所在文件夹」；
  音频这一步失败只写一行提醒（「听力音频没生成：…（卷子不受影响，录音稿还在）」），**不拦卷子也不拦 PDF**。
- **记住上次设置**：纸型、时长、学校模板、听写栏数/模式、教案版式、课件格式、文件加工动作记在本机
  （`dsh-teacher-kit:sticky:v2`），下次打开自动带出来；题干、词表、路径、**勾选框一律不记**（勾选框会被
  记成「下次莫名多出答案页」）。真机复验：选 16 开 → 刷新页面 → 面板仍是「16 开 · 一栏」。
- **结果卡显示用时**：标题行右侧「用了 N 秒」（从点按钮到出结果，含 python 冷启动与 PDF 导出）。
- **发布前全流程自测** `scripts/flow-check.mjs`：打真实 HTTP，一次跑完出卷 / A-B / 教师卷 / 题库看与挑 / 听写单 / 听写音频 /
  听力录音稿音频 / 教案 / 课件 pptx + html / 净化 / 合并 / 提取 / 印前检查 / 模板 / 导 PDF / 逐页预览 + 三条错误路径。
  用法：先起实例，`node scripts/flow-check.mjs <端口>`，末行「全流程自测通过 ✅ N 项」。
- 实测耗时（真机、含 0.4 秒 python 冷启动）：出卷 1.4 秒、A/B 1.6 秒、教师卷 1.5 秒、题库 1.3–1.4 秒、听写 1.1 秒、
  听写音频 1.8 秒、听力音频 ~2 秒、教案 1.1 秒、课件 1.0 秒、净化/合并/提取 ~1.0 秒、印前检查 1.2 秒、导 PDF 1.3 秒、出图 2.1 秒。

## 1.8 0.7.0 修的（新一轮教师视角评测：当场翻车的那条 + 一批"纸面不像正经卷子"）

来自子 agent 扮七年级英语老师走 HTTP 全流程（报告 `.verify/persona-report-050.md`，P0 5 条 / P1 11 条）：

- **P0 卷尾整段答案被当成题**（最要命的一条）：老师把答案整段贴在卷尾（`【答案】1—5 BCBDA 6—10 ABABB`
  ……`56. 略`），原来 56 题被读成 67 题、答案正文印进学生卷第 5 页、答案页没生成、印前检查还判通过。
  现在 `_answer_block_index` 从后往前找整段答案起点，`_parse_answer_block` 认两种写法
  （`1—5 BCBDA` 区间、`1. A` 或 `56. 略（参考范文）` 逐条），`_split_answer_block` 把正文与答案切开；
  **标记行自己带内容时只看这一行**（`【答案】A` 是逐题答案，不能把后面的题切走），标记独占一行才往下看几行；
  至少 2 条答案或一个区间才认。按题号回填，题号下已有答案优先；学生卷不再印那一段，答案页按老师卷口径印。
  实测 41 题 100 分英语卷 → 学生卷 4 页、正文页脚「第 N 页」、答案单独成页。
- **P0 阅读理解的材料（短文）被塞进题干**：粘贴里 `B` / `Passage B` / `材料：` / `【材料】` 开头的段落
  现在认成材料块（`RE_MATERIAL_LABEL` + `RE_MATERIAL_HEAD`），排在对应题目前面、不并进上一题题干；
  题库组卷那边新增「材料 / 短文 / 篇章」列（`MATERIAL_WORDS`），同一篇连着几题只印一遍，
  阅读/完形题没材料会发 `bare-reading` 提醒「学生看不到文章，去题库加一列材料」。
- **P0 页脚 `{{blank}}` 没被替换**：没给 values 时页脚印着 `{{blank}}`。`docx_build._default_values`
  补了 `blank=______`、`schoolYear=本年学年`（8 月起算）与 `className→class / teacher→teacherName / subject→course / title→examTitle` 别名；
  `build_docx` 返回 `values`，`build_exam` 与印前检查用同一份。
- **P0 净化默认原地改文件**：现在默认**另存同目录「原名_净化后.docx」**，原文件不动；真要覆盖原件必须显式
  `overwrite`/`inPlace`，覆盖前自动备份「原名_原件备份.docx」并回传 `backup`；面板结果卡写「原文件没动，副本就在旁边」。
  顺带修掉真崩溃：`_iter_paragraphs` 改成 yield `(paragraph, in_body)` 后旧调用没解包 → `AttributeError: 'tuple' object has no attribute 'find'`。
- **P0 答题卡没有姓名栏 / 作文格被劈成两页**：答题卡抬头补
  「班级：____　姓名：____　学号：____　得分：____」；作文作答格 `page_break_before = True` 整块落一页，默认 10 行（`essayLines` 可调）。
- **P1 一批**：答案页抬头「参考答案 · <卷名>」；题干摘要按词边界截断（不再出现半截单词）；
  顶层 `minutes` 等参数不再被 `options` 吞掉；「勾了答案页却一道答案都没有」降级成提醒（照出卷子、不印空白页）；
  大题写「每题 2 分，共 10 分」不再误报分值不符（`RE_SECTION_TOTAL`）；页眉下载站推广语报 `header-ad`
  （净化会自动删）；Letter 纸型给 `paper-letter` 说明；面板/宿主错误统一 `{ok:false,error:{code,message,hint}}`
  （文件选择器直接显示「请选一个文件夹（要写完整路径，如 D:\教学\第三单元）」这种话，不再是 `[object Object]`）。
- **A/B 卷不再走「序列化→解析」往返**（会丢 12 段材料）：`_parsed` 直接深拷贝，出完再用
  `_missing_material` 逐段核对短文是否真的印进 docx，缺了就发 `ab-materials-lost`「这份卷子有 N 段材料没印出来，请先别印」。

## 1.9 0.8.0 修的（页脚真名 + 三条 persona P2）

- **页脚不再让老师每页手写**：出试卷 / 题库组卷两张卡各加「使用班级」「出卷老师」两格，填了就印在每页页脚
  （`使用班级：七(3)班　出卷老师：王老师`），不填留手写横线。改的是 `docx_build.DEFAULT_FOOTER`（改成
  `{{className}}`/`{{teacherName}}` 占位符）+ `_default_values`（**默认值必须排在别名映射之后**，否则
  `teacher` 写法会被「__________」挡住）+ `exam._own_footer(values=…)` 三处调用；两格进 STICKY_FIELDS，下次自动带出。
- **教师卷不再印「班级/姓名/学号」**：`build_exam_teacher` 里 `options.setdefault("headerLine", False)`，
  学生卷照旧（要印就显式打开）。
- **合并稿默认沿用第一份的页眉页脚**（persona P2-6）：新增 `_inherit_header_footer` 把第一份的 6 个页眉页脚容器
  搬过来（文字/表格/页码域，图片不带以免 rId 悬空，搬完 `_sanitize_rids` 兜底），回传 `headersInherited`；
  warning 从 `headers-not-merged` 改成 `headers-inherited`「沿用第一份文件的页眉页脚（学校抬头、页码都在）」，
  显式 `mergeHeaders:false` 仍是原来的 `headers-not-merged`。
- **课件拆页按内容均分**（persona P2-1）：`lesson._balance_last` 在 `_chunk` 末尾把「6 条 + 1 条孤零零」
  调成 4+3 / 5+2（只动最后一对，前面不动）。
- **模板卡读原卷的满分与时长**（P2）：`template_make` 清正文前先扫前 15 行，`满分 X 分`/`时间 X 分钟`
  命中就把数字当默认值返回，不再一律 100 分 / 90 分钟。

## 1.10 0.9.0 新增的（题库组卷也能出 A/B 与教师卷）

- **题库卡补上「出 A/B 两套」「另出一份教师卷」两个勾**（原来只有出试卷卡有）：老师在题库里已经攥着题，
  想再要一份 B 卷（选项打乱）或教师卷（答案与解析印在题后），不用把题复制回「出试卷」重贴一遍。
- 引擎：`question_bank.bank_build` 认 `options.ab` / `options.teacher`（`ab` 优先），分别转给
  `exam.build_exam_pair` / `exam.build_exam_teacher`，其余参数（纸型、班级、答题卡、答案页）原样透传。
- **A/B 默认文件名修正**：以前 A 卷叫「标题_时间.docx」、B 卷叫「标题_时间-B.docx」，两件东西看着像两份
  不相干的卷子；现在先按「标题（A 卷）」定 A 卷路径，B 卷自然是「标题（B 卷）」。教师卷是学生卷同目录
  「标题_时间（教师卷）.docx」。
- 面板结果卡：A/B 时给「Word（A 卷）」「Word（B 卷）」「PDF」「PDF（B 卷）」四个按钮；教师卷时给
  「Word（学生卷）」「Word（教师卷）」「PDF」「PDF（教师卷）」。
- 测试：`test_bank` 新增 [13] 8 条（两套齐、两文件落盘、标题标 A/B、B 卷题目还在、教师卷落盘且带「（教师卷）」、
  教师卷每题有答案、学生卷无答案）；真机 CDP 复验两张（A/B 5.2 秒、教师卷 6.0 秒，文件名与按钮都对）。

## 1.11 0.10.0 新增的（同一批题一次出好几套）

- **题库卡多了「分成几套」**（就出一份 / 两套 / 三套 / 四套）：把挑中的题**按题型均匀分**，不是从中间一刀切——
  每套里单选几道、阅读几道都差不多，避免「第 1 套全是单选、第 2 套全是阅读」。
- 引擎：`bank_build` 认 `options.sets`（2–6），新增 `split_sets(rows, sets)`（按题型分组轮流发牌，每套内仍按题库原顺序排）；
  返回 `result["sets"] = [{label, title, outputPath, questionCount, totalScore, preflight, warnings}]` 与 `multiSets`，
  第一套的完整结果仍是顶层字段（老调用方不用改）。题不够分 → `sets-too-many`；套数 >6 → `bad-sets`；
  同时勾 A/B/教师卷 → 只按套出并发 `sets-ignored` 提醒。
- 面板：结果卡列出「· 第 1 套：2 题 4 分（检查通过）」并给每套各一个「Word（第 N 套）」「PDF（第 N 套）」按钮，
  PDF 按顺序一串导完（不并发抢 LibreOffice）；分套时不再显示第 1 套的分节明细与单份结论，改成
  「每套的印前检查：第 1 套 通过 · 第 2 套 通过」——否则老师会以为五道题的卷子只剩两道。
- 测试：`test_bank` 新增 [14] 10 条（两套齐、两文件落盘、题量均分 4/4、分值一样、标题标第几套、每套两种题型都有、
  每套自带答案页、总题数累加、题不够报 sets-too-many、套数超限报 bad-sets）；
  真机 CDP：题库 5 题分成两套 → 5.3 秒出结果卡「✓ 分套卷子已生成（共 2 套）｜第 1 套 2 题 4 分 / 第 2 套 3 题 6 分」，
  四件产物（两套 docx + 两套 PDF）都在，`ERRORS: []`。

## 1.12 0.11.0 新增的（听力音频念两遍；题库组卷套学校模板）

- **听力音频「整段念几遍」**：出试卷卡听力那一行多一个下拉（念一遍 / **念两遍（中间停 4 秒，考试常用）** / 念三遍）。
  引擎 `audio:script` 认 `passes`（1–3，越界 `bad-passes`），合成改成「lead + 正文×passes（两遍之间停 pauseSec 秒）+ tail」，
  一个 wav 从头放到尾就是完整听力流程，老师不用手动倒回去再点一次。返回里多 `passes`。
  真机实测同一段 4 句录音稿：念一遍 29 秒、念两遍 55 秒（脚本 `.verify/fills-pass2.json`）。
- **题库组卷也能套学校模板**：题库卡多一格「套学校模板（可不填）」；`bank:build` 把 `baseTemplate` 透传给 `exam:build`。
- **模板纸型真正优先**（以前 README 这么写、代码是反的）：`build_exam` 在「给了模板又没指定 page」时用
  `template_model.analyse_docx` 读模板自己的纸型/方向/栏数；面板在选了模板时**不再下发**下拉里的默认 8 开。
  真机实测：题库卡选 16 开模板但下拉保持默认 8 开 → 产物 184×260mm、单栏（模板赢），没被 8 开盖掉。
- 测试：`test_audio_script` 新增 [6] 3 条（回传 passes、两遍比一遍长 1.6 倍以上、念 5 遍报 bad-passes）；
  `test_bank` 新增 [15] 4 条（16 开原卷 → 模板反推 → 不套模板仍是 8 开 368mm、套了模板产出 184×260mm）。

## 1.13 0.12.0 新增的（答题卡给填空题留作答格；合并时每份保留自己的页眉）

- **答题卡不再只给选择题留格**：原来「用所给词的正确形式填空」「翻译」「简答」那些大题，答题卡上只有
  「题号 3 | 4 | 5」这种字母格，学生写单词没地方下笔。现在 `exam.py` 的 `_is_fill_section()` 会认出来
  （标题命中 填空/简答/完成句子/按首字母/句型转换/翻译/词语运用/语法填空/任务型/补全对话/用所给词/适当形式/根据汉语，
  或整节压根没有选项 → 当填空；作文除外），改印「题号 + 作答区」两列表格：题号一窄列（7% 版心，实测 12.2mm），
  右边整块留白（161.9mm），行高 = 行数 × 8mm。面板上「填空题留几行」选两行（单词短语）/ 三行（写句子）/ 四行（简答翻译），
  对应 `options.fillLines`，默认 2。题号跨节接着选择题数（实测选择题 1–2、填空 3–5）。
- **列宽坑**：只设 `cell.width` 没用，Word 按 `tblGrid` 排——要同时 `table.autofit=False`、`table.allow_autofit=False`
  并写 `table.columns[i].width`，否则作答区被挤成半张纸（第一版就是这么错的，渲图才发现）。
- **合并多份的「页眉页脚怎么带」**：`docx_merge` 新增 `mergeHeaderMode`（也认老名 `mergeHeaders`，first/each/none，
  兼容 逐份/各自/每份/per-source 等别名）。`each` 模式下每份源文件正文各进自己的节（`WD_SECTION.NEW_PAGE`），
  每节 `header/footer.is_linked_to_previous=False`，页眉页脚的文字/表格照搬；页眉里只有图片的整块摘掉
  （源 rId 挂在源 HeaderPart 上，搬进目标包就是悬空引用，Word 会报「文件可能已经损坏」），
  缺抬头的源文件也强制建空白部件（否则会印上一份学校的校名）；复杂 `fldChar` 页码域先压成 `w:fldSimple` + 缓存值，
  不复制源 `pgNumType`，页码按合并稿连续编号。返回新增 `headersMode` / `headerSections` 和 `headers-per-section` warning。
  默认仍是「沿用第一份」，老行为一字未变（测试里 4 条断言钉住）。
- 面板：文件加工 → 合并多份 → 多一个「页眉页脚怎么带」下拉；出试卷与题库卡多一个「填空题留几行」（会记住）。
- 测试：`test_exam` 新增 [20] 4 条（作答格表存在、题号接着选择题数、行高有值、卷子不判 fail），并把「答题卡题号连续」
  改成同时收「题号 + 作答区」表的首列；`test_files` 新增 29 条（[16] 逐份两节各带自己抬头 / [17] 默认口径不变 /
  [18] 别名 / [19] 缺页眉留空 / [20] 图片抬头不悬空 / [21] 复杂页码域），总数 130 → 159。
- 实测：8 题卷（4 选择 + 3 填空 + 1 作文）答题卡第 1 页「题号|作答区」三行、行高 24mm（三行留 3 行 × 8mm）、
  列宽 12.2/161.9mm；两份带不同校名页眉的 docx 用 `each` 合并 → 2 节、`headersMode=each`、`headerSections=2`，
  渲图第 1 页「XX 市第三中学 英语组」、第 2 页「YY 市第一中学 语文组」，页脚「第 1 页 共 2 页 / 第 2 页 共 2 页」连续。

## 1.14 0.13.0 修的（第三轮教师视角评测：模板这条路 + 一批 P1）

- **套模板丢纸型（P0-1）**：多节源卷的 8K 两栏版面藏在**段落 pPr 的 `w:sectPr`** 里，`template_make` 清正文时随段落一起删掉，只剩 body 末尾那张 A4 答题卡节 → 反推出的模板变成 A4 单栏，返回值却还报「8K / 2 栏」（读的是源卷画像）。现在清正文前按纸型+栏数打分挑出该留的 sectPr（`_pick_body_sectpr`）、清完 `_install_sectpr` 装回去，保存后**重新读回实测**（`_measured_paper`）再写进返回，漂移 >3mm 发 `template-paper-drift`。实测反推产物从「210×297 / 1 栏 / 20-18mm」回到「368×260 / 2 栏 / 24-27mm」。
- **模板页脚断链 + 班级老师印不出**：源卷 sec0 带 `w:footerReference`，正文的 8K 节靠「链接到上一节」继承页脚；压成一节后继承链断了 → 套模板出卷 preflight 报「没有页脚」95 分。现在 `_donor_sectpr` 找带 header/footerReference 的那节，`_adopt_header_footer_refs` 把引用深拷贝过来。另一半是占位符打架：`templatize()` 里「使用班级：…」那条 `\S*` 会把已经换好的 `{{className}}` 整个吃掉再塞回 `{{blank}}`；改成 `_label_placeholder()`（已存在 `{{…}}` 就保留，幂等），并在 `commands.py` 里先把页脚换成 `{{className}}` / `{{teacherName}}`。实测模板页脚 `第{{page}}页　　使用班级：{{className}}　　出卷老师：{{teacherName}}`，套用后印「七(3)班 / 王老师」。
- **听写词表把标题和落款收进词条（P0-2）**：`read_word_list()` 的 docx 分支原来先遍历表格再 append 段落，标题行、落款行被排到最后 → 14 个词返回 16 条（第 15 条「Unit 3 … 单词表 ______」、第 16 条「（Unit 3 ______」），`handout:audio` 也照着念 16 条。现在 `_body_blocks()` 按 body 子节点**文档序** yield，段落带位置（first/last/None），`_looks_like_heading(text, position)` 加位置规则（首/末段、≤120 字、命中 单词表|词表|听写|默写|Unit N|姓名|班级|日期|签名 且不是「英文+中文」真词条 → 标题/落款）；`build_dictation_audio` 补 `count` 截断（`items[:wanted]` + `not-enough-items`）。
- **重复导出同名 PDF 报英文 EEXIST（P1）**：`resolvePdfTarget()` 改成「输出按将写文件处理」——父目录 `mkdirSync(recursive)`、非 `.pdf` → `output-not-pdf`、输出=输入 → `output-same-as-input`、目标已存在就清位并返回 `overwritten:true`、目标是文件夹 → `output-is-directory`；`EEXIST/EPERM/EBUSY/EACCES/ENOENT/ENOSPC/EROFS` 逐个翻成中文并给下一步。
- **教师卷【答案】浮到右栏顶（P1）**：`_question_blocks` 的最后一段补 `keepWithNext`，答案/解析跟着题目走。
- **课件多一页只有「操练」的空页（P1）**：`BARE_HEAD_WORDS` 拆成 `SECTION_HEAD_WORDS`（导入/新授/操练/小结/作业/板书…）与 `ACTIVITY_HEAD_WORDS`（pair work / group work / role play / game / discussion / free talk…），活动名不再升成页标题；新增 `_absorb_title_only()` 把只剩标题零正文的块并回上一页（找不到才留 `empty-slide` 告警）。
- **完形填空被当填空题（P2）**：`_is_fill_section()` 先看 `RE_LETTER_SECTION`（完形|cloze|七选五|信息匹配|匹配题|选字母），命中就用字母格而不是书写作答格——完形是选 A/B/C/D，留白格反而让 preflight 报题号不连续。
- **`options.sets` 与顶层 `pick`**：`bank_build` 的 `pick` 也认 payload 顶层（原来只认 options，面板发顶层就被静默忽略）；`sets` 分套输出在给了 `outputPath` 时落在同一目录、沿用老师起的名字。
- 测试：`test_template` [9] 新增 6 条（模板自带 footerReference / 页脚专用占位符 / 套用后有页脚 / 不报 footer-missing / 班级老师印进页脚）、`test_exam` [21] 5 条（完形答题卡 / 教师卷答案 keep_with_next / 模板纸型不一致发 warning）、`test_bank` [16] 4 条（顶层 pick / sets 输出目录）、`test_handout` [13] 15 条、`test_dictation_audio` [8] 6 条、`test_lesson` [16] 16 条；smoke 新增 13 条（PDF 导出：覆盖 / 父目录 / 非 pdf / 同路径）。
- 实测：十二套引擎测试全过；smoke **126 个 ✓**；flow-check 26 项；真机四卡（模板 / 套模板出卷 / 听写单 Word 词表 14 条 / HTML 课件 Pair work 不再空页）`ERRORS: []`。
## 1.15 0.14.0 新增的（整套卡 + A/B 与教师卷可以一起要）

- **一节课四件套（新卡「整套」）**：面板第 8 张卡。试卷（选题库文件或直接粘题干）、听写单（词表文件）、教案（草稿）、
  课件（内容 + 格式）四栏，**填了哪栏就出哪件**；`buildPack()` 用 `Promise.all` 各件独立 catch，
  一件失败只写一行「X 没做成：原因」，另外三件照出；导出 PDF 串行（别几件同时抢 LibreOffice）；四栏全空直接拦住并说清怎么填。
  真机一次出四件 9.4 秒，结果卡逐条：试卷 2 题 4 分 / 听写单 14 条 2 栏 / 教案 3 字段 2 环节 / 课件 4 页 + 试卷第一页预览。
- **A/B 与教师卷能一起勾（四份）**：`build_exam_pair()` 认 `options.teacher`，每套变体再走一遍 `_teacher_copy()`（`inlineAnswers=True`、
  不要答题卡与答案页、不印班级姓名行、标题带「（教师卷）」、路径 `_teacher_path()`），返回体加 `teacherPaths`；
  题库侧 `bank_build` 的 `ab` 分支不再和 `teacher` 互斥（`options["teacher"] = True` 透传）。面板两个 checkbox 都勾 → 四个路径按钮
  （Word（A 卷）/（B 卷）/（A 卷教师）/（B 卷教师））。
- **题库 A/B 文件名**：先按 `title（A 卷）` 定 A 卷路径再算 B 卷，否则 B 卷会变成「标题-B」而不是「（B 卷）」。
- **修掉：整套卡答题卡/答案默认反了**：初版写 `answerCard: draft.card === false ? 'none' : 'card'`，老师没勾也会出答题卡和答案页
  （真机第一次跑出 3 页）；改成和出试卷卡一致的真值判断，未勾就 1 页。
- **修掉：`files.py` 的 `SyntaxWarning: invalid escape sequence`**（docstring 里的 `PAGE \* MERGEFORMAT` 加 `r` 前缀）。
- 测试：`test_exam` [22] 11 条（四件都落盘 / 四个不同文件互不覆盖 / 指定路径时 B 只加 -B、没指定时按（A 卷）（B 卷）（教师卷）命名 /
  两份学生卷无答案 / 两份教师卷每题有答案 / 满分一致 / 无额外 warnings）、`test_bank` [17] 7 条（题库四件）；
  `test/smoke.mjs` **130 个 ✓**（新增整套卡 7 条：一次做齐按钮、四栏都在、纸型下拉、空表单被拦住）。
- 实测：十二套引擎测试全过；smoke 130 ✓ `EXIT=0`；flow-check **26 项全过**（web 53121，`/status` 0.14.0）；
  真机 CDP（实例 9394，`.verify/ud-t10`）整套卡两轮：四件全出 9.4 秒 `ERRORS: []`、只填试卷时 1 页（默认不再夹答题卡）。
- 打包：`dist/dsh-teacher-kit-0.14.0.zip` 234KB / 36 条目 ✅；desktop 与 web 双 profile 已装 0.14.0。

## 1.16 0.15.0 新增的（对话里一步步引导 + 学校画像 + 公众号 / 公文）

这一版改的是**形态**，不是加卡片：从「点按钮 → 填一堆表 → 直接出成品」改成
「点一件事 → 剩下的在对话里被一步步问清」。理由（用户原话）：不希望给一套模板化标准化的答案，
要按每个教师和自己学校去学习适应；不希望点按钮后还要自己输入很多东西，然后给一个流程性的结果。

- **面板从 8 张表单改成「选一件事」**：`src/client/factory.js` 顶部一行画像
  「已记住：临河区第三中学 · 初中英语｜改一下」+ 十行任务（出试卷 / 题库组卷 / 听写单 / 教案 / 做课件 /
  一节课整套 / 文件加工 / **公众号推文** / **公文流转** / 设定我的学校），每行带一句灰色 brief。
  老表单一条不删，折叠在底部「直接出成品（老方式：填表就出，不经过对话）」里，默认收起。
- **点一件事 = 往对话里投一条 user 消息（不弹表单）**：面板 `POST /teacher-kit/api/launch {task, sessionId, note}`
  → 宿主 `deliverGuide()` 取 `ctx.agents.get(sessionId)` → `agent.followup(createUserMessage(...))`
  （退化到 `inject`；都没有则回 `no-channel` 让面板可复制那段话）→ 面板关掉并回执
  「已经把这件事放进下面的话框里了 —— 接着聊就行，缺什么它会问你要。」
  消息体（`lib/guide.js` 的 `buildGuide()`）是**老师口吻的「我想做这件事」**：任务一句话 + 学校画像摘要
  （`profile:get` 的 `data.text`）+ 三栏清单（我要先确认 / 可能会用到的材料 / 我要的成品）+ 六条 `HOW_TO_HELP`：
  能默认就默认别重复问；缺材料直接点名要；关键选择用 `ask_user_question` 一次问 2–3 个且默认值排第一；
  问清就做、做完给路径与预览、别反复确认；**内容按本次材料写、不套一份通用模板、不编没说的事实**；
  最后问一句要不要把版式 / 风格记成学校默认。
- **学校画像（学习适应）**：`engine/teacher_kit/profile.py`（293 行）存本地
  `~/.dsh/teacher-kit/profile.json`（`TEACHER_KIT_PROFILE` 可覆盖）。七节 `SECTIONS`：school / paper / plan /
  deck / wechat / docflow / tasks；`sanitize()` 只收合法键名（`^[A-Za-z][A-Za-z0-9]{0,31}$`）与 str/num/bool/str-list、
  值截 200 字、坏节整节丢；`save()` 临时文件 + `os.replace` 原子写、>64KB 砍 history；`load()` 读坏返回空画像
  **绝不抛错**（记忆不能拖垮生成）；`update(patch, note, task, learn)` 合并并记 history / learned；
  `describe()` 说成人话；`summary()` 给宿主与面板的紧凑版（不带 history）。
- **画像兜底生成（不是死模板）**：`engine/teacher_kit/defaults.py` 的 `apply(payload, command, profile)`
  只补**没说到的**项，八段：纸型（给了 baseTemplate 就不硬塞 page，只留 `_profilePage` 当兜底）/ 页脚两格 /
  考试时长与答题卡 / 听写单栏数与模式 / 教案版式（表格式 → form）/ 课件字号与画布 / 公众号落款 / 公文本单位与科室。
  补了什么写进 `payload["_profileFilled"]`，`cli.py` 转成返回体的 `profileFilled`（面板与模型能说清「这次按你学校的来」）。
- **纸型两种写法归一（真机抓到的 bug）**：画像里记 `{size:'8K', orientation:'landscape', columns:2}`（老师自然写法）时
  `describe` 原来 `PAPER_LABELS.get('8K')` 落空只印「8K」、`defaults` 里 '8K' 不在 `PAPERS` → 画像纸型根本没生效。
  新增 `PAPER_SIZES` / `paper_key_of()` / `paper_page()`：8开 / 8K / A3 / 16开 / 16K / A4 × 朝向 × 栏数 折成
  8K-2 / 8K-1 / A3-2 / A3-1 / 16K-1 / A4-1，认不出给 None（不瞎猜）。
- **公众号推文（新）**：`engine/teacher_kit/wechat.py`（356 行）。`wechat:check` 只体检（贴进来的整段稿按空行分段、
  首行当标题）；`wechat:render` 出 **HTML（全内联样式，公众号编辑器不掉版）+ 审稿 docx**，四套主题
  `school(#1f5fa9) / warm(#c2410c) / green(#0f766e) / mono(#333)`、内容宽 677、块类型 lead/heading/text/list/quote/note/divider；
  体检码：title-missing / title-too-long（>24 warn、>40 error）/ digest-missing / digest-too-long / body-empty /
  paragraph-too-long（>250）/ list-item-too-long（>40）/ absolute-claim（**单个捕获组**正则，点名是哪几个词）/ no-signoff；
  统计 titleChars / digestChars / paragraphs / longestParagraph / bodyChars / readSeconds。不联网、不发布、不配图。
- **公文流转（新）**：`engine/teacher_kit/docflow.py`。一份办理单六段：收文登记表（来文单位 / 发文字号 / 收文日期 /
  密级 / 份数 / 承办科室 / 办理期限 / 登记号）→ 来文要点 → 拟办意见（单独起页，签发要签字）→ 领导批示 → 承办流转单 →
  回复稿（可选；主送机关顶格、首行缩进、落款右对齐、灰字「（盖章处）」）。要素按 GB/T 9704 常用项、页边距
  `GB_MARGINS={top:37,bottom:35,left:28,right:26}`；体检码 field-missing（缺 ≥3 升 error）/ urgent-no-deadline（error）/
  opinion-empty | opinion-short / steps-empty | step-no-due / reply-no-to | reply-no-sign | reply-too-short；红头与印章按需不代盖。
- **修掉：公文 / 推文的 docx 被印上试卷页脚**：`docx_build` 原来默认印「使用班级：__ 出卷老师：__」。新增
  `spec.footer`（整串自定义）与 `spec.defaultFooter`（bool）、`_default_footer(document, values=None, text=None)`；
  docflow 传 `"第 {{page}} 页 共 {{pages}} 页　　临河区第三中学公文办理单"`、wechat 传 `defaultFooter: False`。
- **修掉：拟办意见标题被挤成页脚孤行**：在「二、拟办意见」前插 `pageBreak`，删掉「四、承办流转单」前那条分页。
- **新工具 / 新路由**：`teacher_kit_profile`（action get/set + patch/note/task/learn）；
  `POST /teacher-kit/api/launch`、`GET|POST /teacher-kit/api/profile`（HTTP 路由 7 → **9** 条）；
  `teacher_kit_engine` 描述里补全命令名（exam:build / bank:build / handout:build / plan:build / lesson:deck /
  docx:new / pptx:new / wechat:render / docflow:build / profile:get / profile:set）。
- **测试**：新增两套引擎测试 `test_profile_defaults.py`（72 条：脏数据洗掉、存读 / learn / history、纸型两写法、
  describe 中文、坏文件不拖垮、八段兜底只补缺、不覆盖老师指定、reset）、`test_wechat_docflow.py`（51 条：体检码、
  HTML 无外链、主题切换、docx 不印试卷页脚、六段 / 五段、GB 页边距、模板缺失）；
  `test/smoke.mjs` **183 个 ✓**（新增 6.6 对话式引导 15 条：followup / inject 退化 / no-channel / no-session /
  session-required / unknown-task / GET 405 / GET profile 10 任务、画像记下后再点任务的引导语含校名与「8 开横向两栏」）；
  `scripts/flow-check.mjs` **39 项全过**（10.5 三件新引擎、10.6 launch / profile 路由；
  `GET /profile` 的响应是嵌套的：profile 再包一层 profile）。
- **实测**：十四套引擎测试全过（共 **1239** 条检查）；smoke **183 ✓** `EXIT=0`；flow-check **39 项**；
  真机 CDP（headless Edge 9333 + web 实例 53999，用新建的 `--profile tktest`、画像隔离到 `.verify/flow/profile.json`）：
  点「出试卷」→ 面板回执 + 对话里蓝色引导气泡（带画像：临河区第三中学 · 初中英语 / 8 开横向两栏 / 页脚写法 /
  教案默认版式 / 公众号语气 / 公文单位 + 三栏清单 + 六条引导规则）；面板截图十行任务齐全、画像行在。
  **限制**：测试 profile 没有 API key，模型那一轮报「本轮运行失败 API 密钥无效」——投递由宿主 `followup` 完成、
  不依赖模型，用户自己的 desktop profile 有密钥，所以是环境限制不是插件缺陷。
- **打包**：`dist/dsh-teacher-kit-0.15.0.zip` **271KB / 41 条目** ✅（新增 `lib/guide.js`、`engine/teacher_kit/{profile,defaults}.py`、
  `{wechat,docflow}.py`）；desktop 与 web 双 profile 已装 0.15.0（重启 DSH 生效）。

## 1.17 0.16.0 新增的（面板重做：一屏十件事，两列、深色、窄屏抽屉）

- 面板从「八张卡纵向堆」改成**三段式**：标题栏（图标 + 副题 + ✕）不滚，中间是唯一滚动区，
  底栏「引擎状态 + 收起」永远可见 —— 以前行一多，底栏被挤出视野，老师根本看不到「收起」。
- 任务行按场景分三组：**出卷与练习**（出试卷 / 题库组卷 / 听写单）、**上课用**（教案 / 做课件 / 整套）、
  **学校事务**（文件加工 / 公众号 / 公文），每组两列（面板宽 ≥640px 时一屏放得下十件事）。
- 每行是「图标 + 名字 + 一句人话说明 + 右侧箭头」，不是按钮列表；键盘 Tab 在面板内循环，Esc 收起，
  打开时焦点落在面板本身（以前焦点套在 ✕ 上，看着像选中了关闭）。
- 画像条把记着的习惯翻成中文小标签（学校名 / 纸型 / 教案版式 / 课件配色 / 公文单位，最多 6 个），
  认不出的键直接跳过 —— 真机上曾冒出「orientation landscape」这种半截英文。
  页脚模板里的 `{{className}}` 这类占位符显示成「按学校模板自动填」，不把占位符端给老师。
- 窄屏（≤560px）变底部抽屉；深色模式令牌化（对比度按 WCAG 挑过）；`prefers-reduced-motion` 降级。
- 性能：学校画像按 `路径 + mtime + 大小` 缓存，`/status` 加到 60 秒缓存（`?refresh=1` 旁路），
  Python 探测改异步 —— 面板打开时 GET /status 610ms → 0ms、GET /profile 1111ms → 0.1ms、
  POST /launch 1051ms → 0.1ms（真 HTTP 复测 2.2ms / 10ms / 11ms）。`scripts/perf-check.mjs` 18 项断言守着。
- 排版：公文（docflow）按 GB/T 9704 重排（二号黑体标题、三号仿宋正文、28pt 行距、真 2 字符首行缩进、
  落款右空二字 / 日期右空四字、表格列宽按内容权重、表头浅底跨页重复），4 页 → 3 页；
  公众号行高 1.78、段距 17px、小标题上 22px。

## 1.18 0.17.0 修的（全量审查：四个分区各派一个子代理，逐条复现再修）

四个只读审查（宿主 / 引擎 / 客户端 / 文档打包）共报 **5 个 P0、22 个 P1、60 个 P2**，全部逐条复现后修掉。
每条都留了复现脚本（`.verify/audit-tmp/`）与修复后实测：

**宿主（`lib/index.js`）**

- 引擎参数走 argv：JSON ≥ 32K 必抛 `ENAMETOOLONG`，而 `/engine` 允许 8MB body —— 大题库组卷必失败，
  且只回英文 `spawn ENAMETOOLONG`。改成走 `--args -` 从 stdin 传（200KB 实测通过）。
- `/profile` POST 把引擎 `ok:false` 包成 HTTP 200 `ok:true`（WinError 5 也回「存好了」）→ 画像静默丢失；
  GET 失败回落成「还没记过」→ 引导语按空画像重问。现在失败回 500/502 并带中文提示，GET 失败回 `degraded`。
- 导出 PDF 先 `rmSync` 目标再转换：转换失败旧 PDF 永久消失；输出与输入只差大小写时还会删掉源文件。
  改成先转临时文件（同目录、`.pdf` 结尾，否则转换器按扩展名拒），成功才替换，失败清临时文件。
- `/status?refresh=1` 用 `spawnSync` 冻住整个 Web 壳 0.85s（最坏 20s）→ 改异步 + 并发去重（面板同要
  `/status` 和 `/profile` 时不会起五六个 python）。
- `withService` 先把服务记进「已注册」再跑：服务未就绪时永久静默 404。改成成功才记，失败下次注入再试。
- `/render` 的 `pdfPath`、预览目录、`/launch` 的 `sessionId` 全部收口：预览目录必须带本插件标记文件且内容精确匹配；
  `sessionId` 走白名单正则；超时上限 1s–600s；Windows 超时用 `taskkill /T` 杀整棵树（否则 soffice 占锁）。
- 错误文案中文化：`friendlyText()` 把 `was disposed` / `ENOENT` / `EACCES` / `ENOSPC` / `EBUSY` 翻成人话，
  `bodyFailure()` 接住 body 过大 / JSON 不合法 / 连接断开的三种情况。

**客户端（`src/client/factory.js`）**

- 听力音频回执整体覆盖教师卷路径（A/B + 教师卷四件只剩 5 个按钮）→ 改 concat；单教师卷 Word 按钮不再标成「B 卷」。
- 分套出卷时只看第 1 套的印前检查，却给全部套导 PDF（第 2 套的卷子带着 fail 也导了）→ 逐套判，
  有问题的套写清楚「第 N 套的 PDF 先不导：理由」。
- 导出 PDF 的 `.catch(()=>({ok:false}))` 吞掉宿主给的中文原因 → 失败原因现在逐条写进结果卡（`errorText()`）。
- 听写卡音频失败会把已经生成的 docx 一起吞掉 → 只提醒不拦（「单子不受影响，可以自己念」）。
- 切会话不作废在途请求（上一会话的「✓ 试卷已生成」落到新会话面板上）→ 加 `runSeq` 序号 + `stale()` 守卫。
- `api()` 加 30s / 190s 超时与 `readJson()`（先读 text 再 parse，非 JSON 回「服务回了非预期内容（HTTP N）」）；
  结果卡预览、文件选择器都加了「迟到响应不许赢」的序号守卫。
- 深色模式下「成功/失败」用硬编码绿色红色对比度不足 → 改成令牌 `--tk-ok`/`--tk-bad`。
- 引擎体检失败一次就永久「体检失败」→ 状态条变成可点的「重查」。
- 草稿存储加 400ms 防抖 + 卸载时补存（关面板 / 切会话不丢）；文件选择器记住上次目录；批量净化在没目录时不给「打开文件夹」。

**引擎（`engine/teacher_kit/`，P0 两条都在这里）**

- 听写单：Word 表格题库三列（单词 / 音标 / 词义）把**中文答案直接印在题面上**（`1. apple /ˈæpl/ 苹果 ______`）。
  改成按列语义挑（序号列、音标列、词性列剔除，英文取第一格，释义取其后第一格带中文的）→ `1. apple ______`。
- 教案：环节名「2. 新课呈现」不在词表里 → 并进上一环节、环节名挤进「时间」格。补 14 个常见说法
  （呈现 / 新知 / 新课呈现 / 拓展 / 延伸 / 语言输出 / 迁移 / 应用 …），无标签行不再落进时间格，时间列去重截断。
- 占位符嵌套：模板 `使用班级：{{className}}` 而 `className` 的值又是 `{{blank}}` → 页脚印成
  「使用班级：使用班级：{{blank}}」。`_expand()` 递归填 4 层（值里套占位符也能填到底），`template:make` 不再自己套标签。
- 8 开两栏 + 长题干：正文中途改栏数让 LibreOffice 把整节推到下一页（第 1 页只剩三行卷头、白印一张），
  加空段后变成左右栏乱序。改成**卷头写进第一页页眉**（页眉天然跨整页宽），正文一栏到底 —— 实测 2 页 → 1 页。
- 完形填空「21. A. small B. beautiful …」整行进题干、`options=[]` → B 卷与 A 卷逐字相同且无警告。行首即选项时改成收进选项。
- 不给卷名时卷首大字印成第一个大题名（`examTitle` 被 `setdefault` 让位）→ 无卷名按「学科 + 测试卷」兜底。
- 模板页眉「XX中学七年三班 英语」原样带到新卷每页；`values.schoolYear` 是数组（画像里存的列表）。
  页眉也跑一遍班级 / 老师正则（中文数字班号也认），`schoolYear` 是列表就拼成字符串，没有年份就留占位符。
- Word 题库没表头时 5 列死序切分（答案列变成选项 A、解析变选项 B）→ 新增列体检 `_repair_roles()`：
  整列像答案就往左找题干、往右找分值与解析；判断题两列也对。xlsx 多 sheet 不再只读第一张。
- 认不出题号也照样出卷（`questionCount=0` 的空卷 `ok:true`、印前检查 pass）→ 加空卷闸，回 `no-questions` 并教老师怎么写题号。
- `paper.py` 一串静默行为：`page` 传字符串报 `internal`、`marginsMm` 非数字报 `internal`、`columns: 0` 静默变 1 栏、
  页边距写「2.7cm」不认 → 全部改成中文错误码（`bad-margin` / `bad-columns` / `bad-score`）。
- 画像：`paper.footer` 记了从不注入（回执却说已记上）→ 现在出卷时生效；`profile:reset` 从不真清；
  中文键静默丢弃；`note` 给数字报 `internal`；`profile:set` 结构不对静默 `saved: []` → 各自补上校验与提示。
- 公众号：配图与 caption 静默丢 → HTML 留「【配图】说明」占位（公众号编辑器不认本地图，假图更糟），
  docx 真贴图（宽度默认 140mm，手机大图不会撑爆版面），缺图时发 `image-missing` 警告。
- 输出名带半角冒号写成 NTFS 数据流（文件夹里只剩 0 字节同名文件）→ 32 个命令入口统一过 `_guard_raw_paths()`，
  回中文「文件名里有不能用的符号「:」」。
- 印前检查：文本框里的 `{{schoolName}}`、脚注里的占位符以前完全看不见（漏报）；字体统计漏表格 run 与样式继承字体；
  题出在表格里时分值统计误报 → `_every_text_node()` 全量扫，字体补 eastAsia 与样式回落。
- 一批输入校验与编号：`--debug` 参数从来使不上、`docx:merge` 覆盖默认值与同族相反、第 11 条承办步骤编号回到「（一）」、
  `_cn_number` 超过 99 `IndexError`、题库题号写中文不认、非法颜色静默变黑（现在回 `color-fallback` 说明）、
  只传 `title` 时教案「课题」空、空白版式的课件标题凭空消失、同一秒两次生成撞文件名（精确到毫秒）。

**文档与打包**

- README 让人装 0.12.0.zip、RELEASE 自报 0.14.0 → 三处版本齐步 0.17.0。
- `npm pack` 会带 `engine/tests` 与 26 个 `.pyc`（zip 只要 41 个文件）→ `package.json files` 收窄 + 新建 `.npmignore`。
- 打包脚本加 `assertConsistent()`：三处版本不一致、或 `client/client.js` 比 `src/client/factory.js` 旧，直接拒绝出包。
- `install-local.mjs` 按标记去重漏了「市场装过再本地装」→ 再按插件 id 扫一遍顶层列表项。

## 1.19 0.18.0 修的（端到端：先出成品逐页看，再拿坏输入撞）

- **模板卷的「共 N 页」印错**：8 开模板 + 答题卡 + 答案页时，页脚的 `{{pages}}` 被换成真 `NUMPAGES` 域，LibreOffice 在多节 + 页码重起的文档里算成 5 页而实际 4 页（`exam.py` 新增 `_strip_page_total()`：多节时摘掉「共 N 页」只留「第 N 页」，单节模板照旧；`test_exam` 的 `[21b]` 用带 PAGE/NUMPAGES 的模板锁住这条）。

这一轮不再只看代码：每条都先用真样本出成品、把 PDF 逐页渲成图看，再补一轮对抗输入（241 例）。
共修 **3 个 P0、20 个 P1、30 个 P2**，复现与修复证据全在 `.verify/audit18/` 与 `.verify/lead18/`。

**考试线（出卷 / 组卷 / 模板）**

- 卷中出现「【答案】1—5 ABCDA」这种小节答案行，整行被当成卷尾答案区：**其后所有大题被吞进答案文字**
  （11 题只剩 5 题，答题卡与答案页跟着缩水，印前检查还报 pass）。现在先看答案行之后还有没有大题/选项行，
  判成小节答案就只回填本节题号；「带头 + 区间」的写法也认了（答案页不再印成「5. 1—5 ABCDA」）。
- 勾了答题卡/答案页，**这三节的第一页都没有页脚**（开了「首页不同」，页脚只写进了 `section.footer`）
  → `_own_footer` 把页脚同时镜像进 `first_page_footer`；新分节不再继承正文卷头（以前答题卡顶部把卷名、
  满分、班级又印一遍）。
- 套学校模板后**模板纸型与页边距被默认 8 开横排盖掉**（270×390 竖版 → 368×260 横版、2 栏）；
  未显式指定纸型时按模板自己的 sectPr 走，模板页眉的校名也并入第一页页眉（以前首页看不到校名）。
- 选项后面的说明行被拼进最后一个选项（「D. Chinese books 阅读 B 篇，选出最佳选项。」）→ 说明行先挂起，
  遇到材料/分篇标签再前置；实在无处可放就放到本节开头并回 `stray-line` 警告。
- 题库 xlsx 用合并单元格跨多行存材料时，分到第 2 套的完形题**完全没有短文** → 只补「跨行」合并的锚点值
  （单行横向合并仍不补，避免把「一、单项选择题」灌进每一列）。
- 卷首写「满分 100 分」而小题只加到 75 时纸面被静默改成 75 → 回 `total-score-mismatch`，写明声明值与实算值；
  完形「26.」后直接跟选项的空题干改成印「26. ______」并回 `blank-stem-filled`；
  模板学年占位符不再把「学年」二字一起吞掉（「2025—2026学年第一学期」）。

**上课线（教案 / 课件）**

- 教学过程表跨页被撕开、第二页起没有表头：根因是**两张表在 XML 里紧挨着**，Word/LibreOffice 当成一张表，
  后面那张的 `w:tblHeader` 就不算数 → 相邻表之间垫一个 1 磅空段落（肉眼不可见，全局生效）；
  超长环节按 240 字切成「（续）」行，表头逐页重现。
- 一份教案第 2 页只剩板书框 + 作业布置（约 90% 空白）：板书块加了「不许跨页撕开」，但整页留白没有强行
  拉伸版面（会牺牲正常单页教案）——已知限制，写进 §6。
- 教案「课题」写进标题行但信息表课题格空、学科名与课题重复前缀、合并单元格时「学生活动」列表头已填内容却标成
  空白、课件把「小结 / 作业」另起一页重复、封面标题与第一页标题近似时多一页同标题空页 → 各自修掉，
  近似判据只放宽第一块（避免把正文行误判成标题）。
- 课件备注传布尔值/数组往备注里塞 Python 字面量、自定义小节名不升页、空白版式标题凭空消失 → 校验 + 兜底。

**教务线（公文 / 公众号 / 文件加工 / 画像）**

- 公文：没有「来文要点」时四节编号从「二、」起（跳号）；回复稿里混对象时报英文 internal；
  承办流转单字段别名（title/department/person/deadline/body）以前不认；第 11 条编号回到「（一）」。
- 公众号：配图是坏图也照出稿 → 出稿前用 Pillow 验一遍（打不开就回 `picture-unreadable`，两份成品都不落盘）；
  HTML 与 docx 改成先写临时文件、两份都成功才替换（写一半失败不再留下半份）。
- 文件加工：广告段落识别补「代写 / 加群 / 扫码加 / 加微信 / 联系QQ」（故意不收单字「微信」——学校通知里会写
  「微信扫描班级群」）；带页码域的段落只摘广告 run，页码域保留。
- 画像：`profile:reset` 真清空；`describe()` 的课件段翻成中文（主题 / 标题字号 / 正文字号）。

**对抗输入（241 例，0 例英文 internal）**

- 路径里带 NUL（JSON 的 `\u0000`）一路穿到 zipfile 报 `embedded null character` → 入口拦下并说人话；
  `runs` 写成字符串报 `'str' object has no attribute 'get'` → 按纯文本处理或回 `invalid-blocks`。
- `payload` 不是对象（null/数组/字符串/数字）以前 32 条命令全吐英文 internal → 统一 `bad-payload`；
  JSON 截断、`--args-file` 不存在、输出盘符不在、`outputPath` 是数字/文件夹、`page` 传字符串、
  `options` 传字符串、行距/段距/缩进传字母、图片宽度为负、CMYK 图 —— 全部中文错误码 + hint。
- 宿主 9 条路由里 8 条回 `{ok:false, error:"英文代码"}` → 统一 `{ok:false, error:{code, message, hint}}`
  （面板「复制这段话」照旧能用）；引擎超时用 `taskkill /T` 杀整棵树，不再留 soffice 占锁。

## 2. 发布前必须全绿的四项

1. `node scripts/build-client.mjs`（面板源码改过就必须重打，产物 `client/client.js`）
2. `ELECTRON_RUN_AS_NODE=1 "D:/Deepseek Harnees/DeepSeek Harness.exe" test/smoke.mjs`
   → 末行 `全部通过 ✅`、`EXIT=0`（用 `> smoke.log 2>&1` 收日志，别用 `tail` 看失败项）
3. 引擎十四套（`engine/tests` 下，用 DSH 自带 Python）：
   `test_engine.py`、`test_template.py`、`test_preflight.py`、`test_exam.py`、`test_lesson.py`、
   `test_lesson_html.py`、`test_files.py`、`test_handout.py`、`test_plan.py`、`test_bank.py`、
   `test_dictation_audio.py`、`test_audio_script.py`、`test_profile_defaults.py`、`test_wechat_docflow.py`
   各末行 `全部通过 ✅`（0.15.0 起共 1239 条检查）
4. 人工看一遍成品：出 1 份 8 开双栏 100 分卷（带答题卡+答案页）导 PDF，逐页翻；再看 1 份课件在 PowerPoint/WPS 里放映。

最近一次全绿记录（2026-10-26，0.15.0）：十四套引擎测试全过（新两套 `test_profile_defaults.py` 72 条、`test_wechat_docflow.py` 51 条，合计 1239 条）；
`test/smoke.mjs` **183 个 ✓、`EXIT=0`**（新增 6.6 对话式引导 15 条）；`scripts/flow-check.mjs` **39 项全过**（新增 10.5 三件新引擎 / 10.6 launch 与 profile 路由）；
真机（headless Edge CDP 9333 + web 实例 53999，新 profile `tktest`，画像隔离 `.verify/flow/profile.json`）：面板「选一件事」十行任务 + 画像行「已记住：临河区第三中学 · 初中英语」，
点「出试卷」→ 面板回执「已经把这件事放进下面的话框里了」+ 对话里出现带画像与三栏清单的引导消息（截图 `.verify/tk15-panel.png` / `.verify/tk15-cdp.png`）。
`GET /profile` 返回 10 个任务、`POST /launch` 缺 sessionId → 400 `session-required`、未知任务 → 400 `unknown-task`。
已知环境限制：`tktest` 这个新 profile 没配 API key，模型那一轮报「本轮运行失败 API 密钥无效」（投递由宿主 `followup` 完成，不依赖模型）。

上一次全绿记录（2026-10-26，0.14.0）：十二套引擎测试全过（`test_exam` [22] A/B + 两份教师卷四件、`test_bank` [17] 题库四件）；
`test/smoke.mjs` **130 个 ✓、`EXIT=0`**（新增整套卡 7 条断言）；`scripts/flow-check.mjs` **26 项全过**（web 实例 53121，`/status` version 0.14.0）；
真机（本机 DSH 窗口 + CDP，实例 9394，user-data-dir `.verify/ud-t10`）整套卡两轮：四栏都填 → 9.4 秒出结果卡
「✓ 整节课的东西都做好了｜试卷：共 2 题，合计 4 分｜听写单：共排进 14 条，2 栏｜教案：认出了 3 个字段、2 个教学环节｜课件：共 4 页」，
产物 `整套联调 Unit 3_…docx/.pdf`、`… 听写_…docx/.pdf`、教案、`…pptx/.pdf` 全在，`ERRORS: []`；只填试卷时 1 页（默认不夹答题卡）。

上一次全绿记录（2026-10-26，0.13.0）：十二套引擎测试全过（`test_template` [9] 模板页脚断链与班级占位符、`test_exam` [21] 完形答题卡 / 教师卷答案跟题 / 模板纸型 warning、`test_bank` [16] 顶层 pick 与 sets 输出目录、`test_handout` [13] Word 词表按文档序、`test_lesson` [16] 活动名不升标题）；
`test/smoke.mjs` **126 个 ✓、`EXIT=0`**（新增 13 条 PDF 导出断言）；`scripts/flow-check.mjs` **26 项全过**（web 实例 53111，`/status` version 0.13.0）；
真机（本机 DSH 窗口 + CDP，实例 9390，user-data-dir `.verify/ud-t9`）四卡：模板「8K · 2 栏 · 横向 / 页眉页脚 3 行换成占位符」、套该模板出卷「印前检查通过（100 分）」且页脚实测「使用班级：七(3)班　出卷老师：王老师」、听写单 Word 词表「共排进 14 条，2 栏」无标题落款混入、HTML 课件「共 4 页」Pair work 不再单独占页，`ERRORS: []`。

上一次全绿记录（2026-10-26，0.12.0）：十二套引擎测试全过（`test_exam` [20] 答题卡作答格、`test_files` 159 条含逐份页眉）；
`test/smoke.mjs` **113 个 ✓、`EXIT=0`**（新增 5 条：填空题留几行下拉 2 条、合并页眉页脚下拉 3 条）；`scripts/flow-check.mjs` **26 项全过**（web 实例 53081，`/status` version 0.12.0）；
真机（本机 DSH 窗口 + CDP，实例 9384，user-data-dir `.verify/ud-t8`）：出卷勾答题卡 + 填空题留三行 → 3.6 秒出结果卡，
产物答题卡实测「题号 + 作答区」三行、行高 24mm、列宽 12.2/161.9mm；文件加工合并多份的「页眉页脚怎么带」下拉在真机可见，
`each` 模式合两份不同校名页眉的卷子 → 2 节各带自己抬头、页码连续（渲图核对）。

上一次全绿记录（2026-10-26，0.11.0）：十二套引擎测试全过（`test_audio_script` [6] 念两遍、`test_bank` [15] 模板纸型优先）；
`test/smoke.mjs` **108 个 ✓、`EXIT=0`**；`scripts/flow-check.mjs` **26 项全过**（web 实例 53061，`/status` version 0.11.0，各命令 1.0–2.1 秒）；
真机（本机 DSH 窗口 + CDP，实例 9382，user-data-dir `.verify/ud-t7`）：出试卷勾「念两遍」→ 5.4 秒结果卡「听力音频：55 秒，4 句」，
三个路径按钮齐、`ERRORS: []`；题库卡套 16 开模板（下拉保持默认 8 开）→ 3.6 秒出结果卡，产物实测 184×260mm 单栏，模板纸型生效。

上一次全绿记录（2026-10-26，0.10.0）：十二套引擎测试全过（`test_bank` 新增 [14] 10 条分套：题量均分 4/4、
每套两种题型都有、每套自带答案页、sets-too-many / bad-sets）；`test/smoke.mjs` **108 个 ✓、`EXIT=0`**；
真机（实例 9378，user-data-dir `.verify/ud-t6`）：题库 5 题分成两套 → 5.3 秒出结果卡
「✓ 分套卷子已生成（共 2 套）｜第 1 套 2 题 4 分 / 第 2 套 3 题 6 分｜每套的印前检查：第 1 套 通过 · 第 2 套 通过」，
产物 `分套真机卷 第 1 套_…docx`、`第 2 套_…docx` 与两份 PDF 四件齐，`ERRORS: []`。

上一次全绿记录（2026-10-26，0.9.0）：十二套引擎测试全过（`test_bank` 新增 [13] 8 条题库 A/B 与教师卷）；
`test/smoke.mjs` **108 个 ✓、`EXIT=0`**；`scripts/flow-check.mjs` **26 项全过**（`/status` version 0.9.0）；真机（本机 DSH 窗口 + CDP，实例 9376，user-data-dir `.verify/ud-t5`）：
题库卡勾「出 A/B 两套」→ 5.2 秒出结果卡「A/B 两卷已生成 / 从题库读进 5 题（xlsx），合计 10 分 / 印前检查通过（100 分）」，
产物 `题库AB真机卷（A 卷）_20261026-042124.docx` 与 `（B 卷）` 成对；勾「另出一份教师卷」→ 6.0 秒出结果卡
「学生卷 + 教师卷已生成」，产物 `题库教师卷真机_20261026-042051.docx` 与 `…（教师卷）.docx`，两件 PDF 也都在，`ERRORS: []`。

上一次全绿记录（2026-10-26，0.8.0）：十二套引擎测试全过（`test_exam` [17] 加 2 条教师卷不印班级行/学生卷仍印、
[19] 5 条页脚真名；`test_files` 新增 [15] 7 条合并沿用第一份页眉（首份抬头在、正文两份都在、headers-inherited、
mergeHeaders:false 不带、无悬空 rId）；`test_lesson` 新增 [15] 3 条 7 条要点拆成 5+2 且一条不少）。
`test/smoke.mjs` **108 个 ✓、`EXIT=0`**；`scripts/flow-check.mjs` **26 项全过**（`/status` version 0.8.0）。
真机（本机 DSH 窗口 + CDP，实例 9374，全新 user-data-dir `.verify/ud-t4`）：出试卷填「七(3)班 / 王老师」→
3.6 秒出结果卡「共 2 题，合计 4 分 / 印前检查通过（100 分）/ 共 1 页」，`ERRORS: []`；
渲染 PDF 页脚实测「第 1 页 共 1 页　使用班级：七(3)班　出卷老师：王老师」。

上一次全绿记录（2026-10-26，0.7.0）：十二套引擎测试全过（`test_exam` 新增 [18] 材料段不被当题：
B 篇 / 「材料：」两种写法都单独成段、短文排在题前、第 2 题题干里没有整篇短文；
`test_bank` 新增 [12] 材料列：两行同一篇只印一遍、空材料是 `None`、缺材料发 `bare-reading`、
出卷后 docx 里短文真的排在题目前面；`test_preflight` 新增 [16] 页眉广告 / [17] 学校抬头不误报 + Letter 说明；
`test_files` 123 条（净化默认副本 + 原件备份 + 页眉广告 + 批注/下划线/编号/分页 五组）；
`test_plan` [13] 环节带时长、`test_lesson` [14] 裸小标题）。
`test/smoke.mjs` **108 个 ✓、`EXIT=0`**；`scripts/flow-check.mjs` **26 项全过**（`/status` version 0.7.0）。
真机（本机 DSH 窗口 + CDP，实例 9372，全新 user-data-dir）：出试卷 41 题 100 分 → 3.8 秒出结果卡
「共 41 题，合计 100 分 / 印前检查通过（100 分）/ 提醒：卷尾那段【答案】我当成答案区了（认出 40 条），不会印进学生卷 /
共 4 页 · 这是第一页」，`ERRORS: []`。

上一次全绿记录（2026-10-26，0.6.0）：十二套引擎测试全过（新增 `test_audio_script.py` 6 组：中英混排段切分、
6 个参数错误、真合成 3 行（单声道/16bit/22050/时长与 wav 头一致/3 行/bytes>10000/报语音名）、
每句停 9 秒比停 1 秒长 10 秒以上、念题号更长、同名不覆盖 + `overwrite=True` 才覆盖；
`test_handout.py` 新增 [12]：`Unit 3 单词听写` 这种「英文+标题词」混合首行不再被当词条，
`dictation 听写`/`practice 练习`/`test 测试` 仍按词条收 —— flow-check 首跑 23/24 就是这条）。
`test/smoke.mjs` **108 个 ✓、`EXIT=0`**。
真机（本机 DSH 窗口 + CDP，实例 9362）出试卷勾「把听力录音稿念成音频」端到端：
结果卡「共 2 题，合计 4 分 / 听力音频：29 秒，4 句（用 Windows 自带的语音念的，点开就能放）/
打开「Word」/「听力音频（wav）」/「PDF」所在文件夹」，用了 5 秒，`ERRORS: []`；
产物 `听力音频联调卷 听力 听力音频_20261026-032204.wav`（1.27MB / 29 秒，22.05kHz 单声道）。
`scripts/flow-check.mjs` 新增 `audio:script` 两项断言（落盘 + 3 句有声音）。

上一次全绿记录（2026-10-26，0.5.0）：十一套引擎测试全过（新增 `test_dictation_audio.py` 8 组：语音列表、
真合 2 条（单声道/16bit/22050/时长与 wav 头一致）、念 2 遍更长、不报题号更短、cn2en、六个错误码、120 条上限）；
`test/smoke.mjs` **108 个 ✓、`EXIT=0`**（新增题库卡「看看题库里有什么」+「挑题」三条断言）；
真机（本机 DSH 窗口 + CDP，实例 9352）题库挑题端到端：点「看看题库里有什么」读到 5 题分布 →
挑题框写「单项选择 3」→ 结果卡「按题型挑出 3 题（题库里一共 5 题），合计 6 分」，`ERRORS: []`。
真机听写单勾音频端到端：
结果卡「共排进 4 条，2 栏 / 听写音频：44 秒，4 条 / 打开「听写音频（wav）」所在文件夹」，`ERRORS: []`，
产物 `Unit 3 单词听写 音频_20261026-021047.wav`（1.9MB / 44.3 秒，与 wav 头一致）。
`test_handout.py` 新增 [11] 十条断言（Word 两列表格 + 散装行读得出、标题行不当词条、CSV 表头跳过、Excel 表头跳过、
只给 `source` 也能出单、`file-not-found`、`.pdf` 报 `unsupported-source` 且 hint 给了能用的格式）；
真机（实例 9358）听写单**只填词表文件**（不粘正文）端到端：「共排进 5 条，2 栏」，`ERRORS: []`；
文件选择器真机复验：点「浏览…」→ 列出文档目录（目录/文件图标、大小、快捷位置、上一级）→ 点 `上次的卷子.docx`
→ 绝对路径填回输入框、选择器收起；修掉「选择器被面板 flex 挤成一条缝」的版式 bug（固定 230px 高 + flexShrink: 0）。

上一次全绿记录（2026-10-26，0.4.0）：十套引擎测试全过（test_exam 新增 [16] 作文作答格、[17] 学生卷+教师卷；
test_plan 新增 [11] 学生活动整列空白 / [12] 长课题缩字号）；
`test/smoke.mjs` **105 个 ✓、`EXIT=0`**（新增「各卡各记各的」两条断言）；真机（本机 DSH 窗口 + CDP，实例 9344/9348/9350）逐卡端到端：
出试卷、A/B 两套、学生卷+教师卷、从 Word 读题干、题库组卷、听写单、教案、模板、文件加工（净化/批量/检查）、HTML 课件
全部出成品且 `ERRORS: []`。
渲染复核：41 题 100 分 8 开两栏卷 9 页（正文页脚只印「第 N 页」、答题卡第 1/2 页各自从 1 编、作文格跨卡第 1–2 页），
教师卷 5 页（每道题下面 `【答案】` 跟解析、标题「（教师卷）」、单纸型页脚保留「共 5 页」），
教案 1 页（抬头表 3 行、课时与课型同行、学生活动列标「老师自填」+ 灰字说明、板书框、页脚「第 1 页 共 1 页　授课教师：____」）。

## 3. 打包

```sh
cd teacher-kit
node scripts/build-client.mjs
node scripts/package.mjs        # 产出 dist/dsh-teacher-kit-0.18.0.zip
```

zip 里只放运行所需：`lib/ client/ engine/cli.py engine/teacher_kit/ cordis.patch.yml package.json README.md RELEASE.md scripts/{build-client,install-local,package}.mjs`。
**不打进去**：`engine/tests/ test/ src/ reviews/ .verify/ .perf/ dist/ __pycache__ *.pyc smoke.log node_modules`。
`package.json` 的 `files` 与 `.npmignore` 是同一口径，所以市场（npm）装法与本地 zip 装法内容一致。

**打包前自检**（0.17 起，`package.mjs` 自己会拦）：`package.json` / `lib/index.js` 的 VERSION /
`engine/teacher_kit/__init__.py` 的 `__version__` 三处必须相同；`client/client.js` 不能比
`src/client/factory.js` 旧。任一项不满足就打印 ✗ 并退出，不生成新包。

## 4. 安装（两条路）

- 市场：设置 → 插件市场 → 安装 → 刷新页面。
- 本地：解压后 `node scripts/install-local.mjs --profile desktop`（或 `web`），然后**重启 DSH**。
  脚本做两件事：把插件拷进 `<profile>/node_modules/dsh-teacher-kit`、在 `<profile>/cordis.patch.yml`
  的 insert 里加 `{id: teacher-kit, name: dsh-teacher-kit}`（空列表 `[]` 也能正确追加）。
  去重按**插件 id**，不只按注释标记：市场装过再本地装，会把那条没有标记的同 id 记录换掉，不会写出两条。

## 5. 市场提交要准备

- `package.json`：`name` `dsh-teacher-kit`、`version`、`description`（中文一句话）、`keywords`
  （教师/试卷/课件/Word/PPT/教育）、`dsh.client.platform=web`、`dsh.bundle.patch`。
- `engines.dsh` 声明宿主版本（`>=0.2.1-rc.2`），否则市场无法判断兼容。
- README 里要有：这是什么、装完看到什么、每张卡怎么用、常见问题、隐私说明（纯本地）。
- 截图建议：输入框下方「教师助手」按钮、出试卷表单+结果卡（带预览图）、生成的 8 开两栏卷子 PDF 首页、
  HTML 课件投屏页。
- 许可：待定（建议 MIT，等作者确认）。

## 6. 已知限制（写进 README，也写进市场说明）

- 公式不做转换（OMML 原样保留，跨软件复制会变形）；数学卷出完需人工补公式。
- 扫描件/图片文字不做 OCR。
- 合并是「拼正文」：批注不跟过来，脚注/尾注不带（会在结果里说明）。
- 没有学校模板时，页脚是内置的「第 X 页 共 Y 页 · 使用班级／出卷老师（手写线）」，字号 9pt；
  要学校自己的页脚，请套模板。多纸型混排（正文 8 开 + 答题卡 A4）时正文页脚只印「第 N 页」，
  答题卡与答案页各自从 1 起（LibreOffice 转 PDF 时 SECTIONPAGES 恒为 1、NUMPAGES 会把握把答题卡算进卷子，
  所以这两种混排文档不印总页数）。
- 答题卡：选择题印「题号 + 答案」两行格；作文/书面表达印作答格（0.4.0 起 `essayLines` 可调，默认 10 行）；填空/简答/翻译/用所给词填空这类大题 0.12.0 起印「题号 + 作答区」，留几行由 `fillLines` 定（面板下拉，默认 2 行）；其它认不出来的大题仍是题号格。
- **音频是 `.wav` 不是 `.mp3`**：本机没有 mp3 编码器（ffmpeg/lame 都没装），wav 约 2.6MB/分钟。
  要小体积就在微信/剪辑软件里转一下，或以后装 ffmpeg 再加 mp3 输出。
- 语音依赖 Windows 自带的语音包（中英各一个）；非 Windows 或精简系统会报 `tts-unavailable`/`tts-no-voice`，
  只影响音频那一步，Word/PDF 照出。
- A/B 卷只打乱带选项的题（选择题、判断等），填空/作文两套一致。
- 网页课件是单文件、无动画/无音频；要动画请在 PowerPoint 里做。
- 教案的板书块较大时，第 2 页可能只有板书框与作业布置（约一页留白）：板书块保证不被跨页撕开，
  但没有为「凑满一页」去拉伸版面（那样正常单页教案会变形）。

## 7. 后续（v0.19 候选）

0.15.0 把形态改成「对话里一步步引导 + 学校画像 + 公众号 / 公文」，0.16.0 重做了面板，0.17.0 全量审查修掉 5 个 P0 / 22 个 P1 / 60 个 P2，0.18.0 又端到端修掉 3 个 P0 / 20 个 P1 / 30 个 P2（详见 §1.17–§1.19）。剩下：
画像面板（现在只能对话里改或改 JSON，缺一个能看见全部条目的编辑页）→ 一校多班 / 一人多校的画像切换 →
公众号配图位与二维码（要联网或本地图库，得先定边界）→ 公文红头模板库（各校红头文件版式不一样）→
教案 / 课件的「风格预设」做成可选清单（现在靠画像 + 对话里说）→ 从题库一次出整套单元卷 →
模板批量套用 → 页眉里带图片的合并 → 教师卷/学生卷分页对齐 → 音频转 mp3（本机没有 ffmpeg/lame）→ 市场发布。
