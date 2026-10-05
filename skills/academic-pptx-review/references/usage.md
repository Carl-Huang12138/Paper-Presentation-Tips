# 学术汇报 PPTX 审阅使用说明

入口先准备同一输入、同一渲染快照的资料，由审阅者实际看图，再导入有来源的逐页判断。输入 PPTX 保持原样。通用判断见 [SKILL.md](../SKILL.md)。安装和框架调用见 [框架说明](frameworks.md)。

## 准备、审阅、导入

将 `<SKILL_DIR>` 替换为已加载 Skill 的绝对路径，输入与输出路径同样使用绝对路径；命令可从任意目录运行。安装包内 `requirements.txt` 中的依赖，并准备 LibreOffice。可执行文件依次查 --soffice、PPTX_SOFFICE 环境变量、PATH 和常见安装路径；DPI 默认 144，可指定 72–600。可先运行 `python "<SKILL_DIR>/scripts/analyze_pptx.py" --help` 查看当前选项。

```powershell
python "<SKILL_DIR>/scripts/analyze_pptx.py" input.pptx --out review-run-1
# 需要指定环境时：
python "<SKILL_DIR>/scripts/analyze_pptx.py" input.pptx --out review-run-1 --soffice "C:/Program Files/LibreOffice/program/soffice.exe" --dpi 160
```

上面两条是不同环境的二选一示例；每次准备使用新目录。查看 `review_packet/README.md`，逐页打开 PNG，再结合对象视图、原媒体和必要局部填写 `review_template.json` 的副本 `review.json`。资料包、自动线索和空模板不声称已经完成视觉审阅。

先对代表页实际看PNG，覆盖主要语言（尤其中文）、公式/符号、主图与动态字段；随后仍逐页核对。PDF文字层 matched、mapping verified 与文件产出不证明字形实际显示。主体缺字/缺图时保留受影响页的 partial/failed，不用不完整画面判断作者版面。字体替代、字形缺失与深目录导出失败分别调查；本流程不自动安装字体，也不声称代表页通过能保证其他页。

默认 --render-scope all 包含隐藏页；仅审放映页可指定 slides，报告保留原页号和排除范围。原页号、PDF 页号与 PNG 必须由 manifest 对应，不能用数组位置猜页。默认保留 PDF，--no-keep-pdf 可省略副本。

```powershell
python "<SKILL_DIR>/scripts/analyze_pptx.py" input.pptx --out review-run-1 --review review-run-1/review.json
```

导入默认复用已看过的同一快照，验证输入、snapshot_id、页码与图像绑定。动态日期、字体变化或重新导出可能改变 PNG，不能一边更换画面一边沿用旧审阅结论。显式重渲染使用新目录，生成新 snapshot_id，并按新图重新核对：

```powershell
python "<SKILL_DIR>/scripts/analyze_pptx.py" input.pptx --out review-run-2 --rerender
```

旧 v1–v3 只有加 `--legacy-review` 才可导入，仍须满足适用的绑定检查；报告标明历史模式，不当作 schema v4 新鲜验收。旧文件不补一个版本字段就冒充新审阅。

## 产物与读取顺序

| 产物 | 用途 |
|---|---|
| `review_packet/README.md` | 本快照的逐页看图入口、来源和读取边界 |
| `snapshot_manifest.json` / `render_manifest.json` | 输入与原页身份、渲染范围、页图映射、哈希及快照来源 |
| `renders/slide-XX.png` | 按原页号绑定的整页画面 |
| `inventory.json` | 完整对象/来源清单、有效可见性、解析与渲染状态 |
| `review_template.json` | 与当前快照绑定的完整 v4 记录框架 |
| `report.md` | 作者先看的精简报告：整套判断、逐页状态和必要方案 |
| `report-details.md` | 四项判断、11 项提示、读取证据、可选建议与所有未知 |
| `diagnosis.json` | 机器读取的完整记录，不截去详细项 |

无 `--review` 时是 preliminary_only，线索不能称为确定缺陷。正式导入仍可能有 partial/failed 页面，报告须保留具体完成范围，不能因命令退出、priorities 为空或某项字段通过就宣称整套质量通过。

退出0表示本次流程没有阻塞，不表示页面无需修改。退出2可能是已生成 `partial_visual_review`（具体页面仍有核心未知）、`failed_visual_review`（全部页面审阅失败），也可能是准备/校验失败；结合报告中的 mode/completion 和终端原因判断。校验失败时旧结果保留，原因写入同级 `OUTPUT.last-failure.json`。

## schema v4：完成度与质量分别记录

新审阅顶层必须为：

```json
{
  "review_schema_version": 4,
  "media_review_version": 1,
  "evidence_review_version": 1,
  "render_fidelity_review_version": 1,
  "input_sha256": "当前输入 SHA-256",
  "snapshot_id": "当前资料包的 snapshot_id",
  "deck_assessment": {
    "purpose": "材料用途和汇报目标",
    "storyline": "结合实际原页号说明问题、方法、证据和结论的衔接",
    "visual_system": "整套主次与同角色组件是否协调",
    "priorities": []
  },
  "slides": []
}
```

上述及下文代码块是字段说明片段；完整逐页记录从当前 `review_template.json` 填写，不将空 slides 或格式示例作为真实审阅提交。

正式导入前将模板的 `preparation_only` 改为 false，清除已解决的顶层阻塞项，并同步顶层完成度和质量状态。整套完成度只有全部页面 complete 才为 complete；全部 failed 才为 failed，其余为 partial。未完整审阅时顶层 quality_status 为“未完成”；全部完成时按“有确定问题 → 需要人工确认 → 无需修改”汇总。各页若保留 snapshot_id，须与顶层一致；模板占位标记不能冒充正式结果。

| 每页字段 | 含义 |
|---|---|
| `slide`、`render_sha256` | 原页号及该页有效 PNG 的哈希 |
| `role`、`basis` | 本页作用、判断依据及已检查范围 |
| `review_completion` | complete / partial / failed |
| `quality_status`、`status` | 两字段同值：无需修改 / 有确定问题 / 需要人工确认 / 未完成 |
| `assessment` | core_message、content_organization、visual_structure、reading_path |
| `tip_checks` | 11 项齐全；adequate / problem / not_applicable / uncertain，每项有页面依据 |
| `issues` | necessary 问题及来源、障碍、收益和方案 |
| `optional_suggestions` | optional 改善，不改变质量状态 |
| `context_questions` | needs_context；具体 evidence 与 action 说明未知如何影响判断或修法 |
| `blocking_unknowns` | 影响完整判断的核心证据缺口 |
| `media_reviews`、`visual_observations` | 实际看图结果、来源与置信度 |
| `structural_limitations` | 结构提取、来源或变换的能力边界 |
| `render_fidelity` | 实际PNG字形、符号、主图与动态字段核对；未确认核心内容阻塞完整结论 |

v7新准备仍使用schema v4，同时将 `render_fidelity_review_version: 1` 绑定进 inventory 和模板，不能删掉字段绕过画面核对。既有v4固定快照未启用此字段时仍可读，不将旧报告冒充本轮新规则执行。每页按实际工具查看填写：

```json
{
  "render_fidelity": {
    "status": "confirmed",
    "evidence": "实际整页PNG中标题、正文与关键图完整，核对所需字形",
    "checks": {
      "text_glyphs": "填写实际可见标题/正文与语言，不能仅引用PDF文字层",
      "math_symbols": "填写关键数学符号，或说明本页不适用",
      "main_media": "填写主要图/表显示范围，或说明本页不适用",
      "dynamic_fields": "填写日期/页码的显示与来源差异，或说明不适用"
    },
    "blocking_unknowns": []
  }
}
```

status为confirmed / partial / failed / not_checked。后三者须具体列阻塞项，不能与complete同时出现；confirmed也不能残留核心阻塞。动态日期不同但正常显示可confirmed。校验只验证声明的状态与结构，不自动识别缺字，实际图片工具检查仍必需。

核心画面、映射或证据未完成时，review_completion 为 partial/failed，quality_status/status 只能为“需要人工确认”或“未完成”。basis 写清已检查范围；已看到的局部问题仍可记录，不将其扩展为整页质量结论。

完成度表示读取与审阅覆盖，不要求已经实施改稿或复证论文。已经读清的直接冲突可以 complete，并在修改方向中保留作者核定正确版本的步骤；未提供论文或尚未修改本身不构成读取未完成。

核心检查完成、无必要问题时可以“无需修改”，可选建议不改变它。有可定位必要问题时记录“有确定问题”。用途、前提或比较身份未知而会改变判断时写具体待确认项；没有全面复证论文、无法恢复位图原数据本身都不构成问题。priorities 只表示修改顺序，不能作为整套状态开关。

11 项是提示，不评分。纯文字、留白、同权和多部分本身不等于障碍；problem 须与对应的 necessary issue 一致，可接受/不适用不扣分。只有局部遮挡时采用局部修复，不为它强制设计整页。

默认现场学术讲解。用户确认讲义、附录、技术复习或教学模板用途时记来源再调整判断，不凭空设想用途降级必要问题。每个适用Tip的短依据应说明听众如何识别主旨、理解机制、找比较、连接证据或记住结论；“文字可读、同一研究、无遮挡”只是观察。已定位障碍、影响讲解任务、有明确理解收益时给最小必要方向。贡献主张埋在细节、完整表无主比较导读、说明与图往返竞争等是典型阅读任务障碍；有效小表、专业公式页、概览与Q&A仍可保留。前页背景或专家用途改变结论时引用具体已读页/用户说明。

用途证据也可来自材料说明和实际页面中面向制作者的填写/组件指导，不限于用户显式指定。示例集合里的研究页面仍按其讲解任务检查；用指导语解释版式的页面按教学角色判断，不把它直接当成研究内容漏填。正反标签本身不证明用途或质量。依据不足时保留会改变结论的针对性确认，不任选正式草稿或教学页作为答案。

围绕主比较记录极值、并列、明显例外和原句范围；局部限定不否定有效主趋势，不扩成整篇论文核验。作者简报在必要方案附近显示对应Tip，11项完整依据仍留详情。

数值主比较须留下四项记录：原主张与总体/多数/全部等范围；同组、同指标、同单位及方向的显示值；工具对齐比较后的例外与并列；结论对原句范围的影响。已有文案的比较也须核对，读对数值不等于验证了“每行均更高”。只检查支撑主比较的可读值，未知不补造。将输入、来源和实际工具输出存入本次计算记录，在 `visual_observations.source` 或相关问题中引用；Python逐组比较记录不冒称通过了未支持的 numeric_checks 运算。主趋势有效但有例外时保留主体，按原句范围给局部限定、可选说明或作者确认，不强制改判全称事实错误。

可读取与适合投影是两项检查。整页先判断听众能否同步定位主信息和主要证据；局部放大用于读清内容，不替代对整页字号、密度与比例的判断。关键表格/图内标签显著小于正文并妨碍比较时，提出保留材料的放大或重组；不用固定字号自动扣分。比较总句还须结合已播放/可引用的前页，不能仅凭当前图例只展示子集判结论错误。

## 图片、来源与实际可见范围

先查看整页，再按需查看局部或原图，最后回到整页。核对图片裁切、旋转、透明度、组合坐标、前景填充遮挡及图文关系。同素材的多个实例分别核对；原图中未显示的内容不作放映画面证据。

每个可见图片记录：

```json
{
  "object": "shape-7",
  "importance": "core",
  "status": "reviewed",
  "source": "rendered_png+rendered_png_detail",
  "confidence": "高",
  "recognized": [
    {"kind": "visual_content", "content": "实际读到的图示关系、标签或趋势"}
  ],
  "unreadable_items": []
}
```

reviewed 表示所需可见信息已经核对；partial 同时列已读内容与具体未读项；not_reviewed 须给 reason，不能伪装成未知或已看。recognized.kind 为 visual_content / exact_value / approximate_value；估值须有 basis，说明实际可读尺度。没有数值的照片可仅写可见内容，三类无需凑齐。清晰印刷数字是页面显示值，原生缓存是文件中的值，都不等于外部实验真实性已核验。

core 的 partial 默认阻塞。只有核心 recognized 内容已读，该 media_review 中 `unreadable_scope: "secondary"`、`blocking_unknowns: []` 及具体 `nonblocking_reason` 同时成立，才可将剩余次要未读项视为非阻塞。secondary 的 partial 也要 nonblocking_reason；secondary 的 not_reviewed 如实报告其范围，不能声称已看。读取边界与作者内容缺陷分别记录。

解析未支持 AlternateContent、OMML、母版继承或组变换时，按实际读取范围写 structural_limitations。PNG 可读的主体可补区域/raw XML 证据；不能因清单漏读判页面空白，也不能因目视读清宣称完整恢复了原生结构。局部图应使用页面坐标，来源变换未知时须回看整页。

## 引用绑定到段落、单元格或实际子对象

source_ref 的共同结构为 `{source,slide,object,selector?,quote?/description?}`。quote 指向真正承载待操作内容的单一读取单位；只有空白可以归一化，不能拼接其他格/段落/子对象或消除有意义的公式符号。同页多处证据通过 refs 列表逐条绑定，外层 source 为 multiple_sources；跨页来源使用 issue.source_refs。

```json
{
  "target_evidence": {
    "source": "multiple_sources",
    "refs": [
      {
        "source": "pptx_object",
        "slide": 3,
        "object": "shape-5",
        "selector": {"kind": "cell", "row": 2, "column": 3},
        "quote": "该单元格确实存在的原文"
      }
    ]
  }
}
```

段落 selector 为 `{kind:"paragraph",index:1}`；组合子对象为 `{kind:"object",path:"实际 ledger 子对象路径"}`。行、列、段落编号均从 1 开始。chart_point 使用 plot/series/point（均从 1 开始）和 component 定位实际缓存分量，不将缺失数据伪装为空值。selector 和引用对象必须与真实清单对应。

单目标也可写 `target_evidence={source,selector?,quote?/description?}`，由 issue.object 定位。图片和没有可读文本的目标描述具体子区。清单漏读的公式/正文使用 `object: "region:具体位置"` 与来源 description，必要时加 raw XML id/name；不能借存在的页眉对象。

字段检查能验证部分引用存在和读取单位，不能证明图片描述真实或操作语义正确，仍须对照整页。

## 必要问题与修改方案

issues 中 recommendation_kind 必须为 necessary，并写 obstacle（现有结构的具体障碍）、benefit（改动收益）、evidence、impact、direction、severity、confidence 与实际目标。scope 可为 visual / presentation / critical_logic / factual_error。tips 引用对应提示；没有合适编号时可用空列表并填 tip_mapping_reason，不硬塞编号。

optional_suggestions 每项必须有 recommendation_kind=optional、object/evidence/direction/benefit；context_questions 每项必须有 recommendation_kind=needs_context、具体 evidence/action。两者都不混入 necessary 问题。

| action | 方案要求 |
|---|---|
| local | revision_plan.steps 与 check 说明针对目标的操作及复核；不要求重做完整布局 |
| revise | revision_plan 给 goal/layout/steps/typography/check，说明分组、顺序、主信息与原内容对应 |
| split | 完整 revision_plan 和 split_plan 至少两页；每页 title/message/content/layout/components 及非空 source_refs，说明顺序和去向 |

条件式备用方案在 issue.alternatives 里写 action=split、activation_condition、pages；每个新页的字段与非空 source_refs 与 split_plan 同样完整。首选 revise 不必为了备用改 action，也不混用顶层 split_plan。仅结构字段决定是否要求拆页方案，自然语言、否定句或建议中出现“拆页”不作强制触发。

结构方案通过 issue.source_refs 与 issue.content_disposition 登记原内容去向：

```json
{
  "content_disposition": [
    {
      "source_ref": {
        "source": "pptx_object",
        "slide": 3,
        "object": "shape-5",
        "selector": {"kind": "paragraph", "index": 1},
        "quote": "该段中实际存在的待分配文字"
      },
      "disposition": "move",
      "destination": "new-page-1",
      "reason": "该原文承担第一页的任务定义"
    }
  ]
}
```

disposition 为 retain / move / backup / deduplicate；destination 写 current-slide、新页或明确备份位置。去重注明保留在哪里；图子区、关键数值、单位、限定条件与核心公式逐项有去向。列表齐全仍不能证明语义保真，审阅者需检查新旧内容与图归属。

每个 issue.source_refs 中的原材料都须在 content_disposition 登记对应去向；复用同一来源引用，不能用另一个无关对象代替。这个检查仍不能判断未登记材料是否被遗漏。

先腾空间、分组或去重复，再选择大且协调的字体；主要比较数据和符号解释不能当脚注缩小。普通方案不要求英寸/像素坐标，必要参数写 revision_plan.parameters。位图字高不能换算为精确字体 pt。放大图表保持比例并安排标题、正文、公式及边距；未实施改稿渲染时明确“待验证”，不声称已能放下。

## numeric_checks：算术与输入来源分别检查

涉及新增数值例子、修改后复算或“提升最大”等比较时，调用计算工具。numeric_checks 支持有限登记运算，不执行 JSON 任意表达式，也不证明论文公式或未登记运算。

v4 每个 inputs 的数值叶节点都要有 source_refs 的对应 `input` 路径和 `value`；标签身份与单位等另在原材料和 basis 核对。真实输入注明 source、slide、object、selector 和独立表示该输入值的 quote；数值转换另保留工具步骤，不能只引用含多值的长句。假设代入用 test_assumption 与 basis，不伪装成原页实验数据。以下只展示合法假设输入的复算记录：

```json
{
  "operation": "geometric_sum",
  "inputs": {"coefficient": 3, "ratio": 0.5, "start": 0, "end": 1},
  "reported": 4.5,
  "basis": "假设代入有限求和；不是实验数据，也不声称满足原论文全部假设",
  "source_refs": [
    {"input": "coefficient", "source": "test_assumption", "value": 3, "basis": "本次复算假设系数"},
    {"input": "ratio", "source": "test_assumption", "value": 0.5, "basis": "本次复算假设比值"},
    {"input": "start", "source": "test_assumption", "value": 0, "basis": "本次复算起始索引"},
    {"input": "end", "source": "test_assumption", "value": 1, "basis": "本次复算终止索引"}
  ]
}
```

geometric_sum 的 start/end 是满足 0≤start≤end≤10000 的整数，复算容差为 1e-9。compare_gains 的 inputs.rows 含 label/before/after，来源路径按 rows.0.before 等逐输入绑定，reported 保留输入范围内全部并列最大绝对提升的标签；它不评价相对提升、统计显著性或其他模型。复杂运算保留实际工具记录，前提不明写待确认。

算术正确不证明来源正确。图上/表中数值、单位、基线、范围与例外须再次核对；数值示例不能扩大原研究结论。

## 交付与回归边界

作者先看 report.md；详情和机器记录保留四项整体判断、全部 11 项提示、所有读取证据、可选建议和未知。报告摘要由逐页质量、完成度及阻塞未知形成，不以问题数或 priorities 为空决定。

哈希/字段/有限运算校验能说明绑定和契约满足，不能证明宿主实际调用看图工具、视觉语义正确或方案保真。LibreOffice 可能替换字体，静态审阅不能确认动画与讲述节奏；疑点写具体影响，不凭这些未知编造缺陷。

原仓库的 `validation/v7_acceptance.md` 记录验证范围；这些验收材料不是独立运行依赖。真实材料的独立审阅应先冻结判断，再读旧审阅记录进行对照；旧 v2/v3 兼容结果与固定问题数不是新版行为验收。
