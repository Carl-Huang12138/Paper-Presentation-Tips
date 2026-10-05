# 学术 PPTX 审阅：v7 通用包验证

2026-10-05 将已校准的 v7 Skill 整理为可独立安装的 Agent Skills 包。唯一维护源为 `skills/academic-pptx-review/`；根目录 Python 文件提供兼容转发，仓库内 Codex/Claude 入口由工具生成并读取同一份规则。

## 本轮实际检查

| 检查 | 结果与范围 |
|---|---|
| 原仓库完整回归 | 123 项通过、0 跳过；含本地真实 PPTX 验收与 7 项安装/可移植性测试 |
| 实际拟上传副本 | 115 项通过、8 项真实材料验收因缺少私人输入明确跳过、0 失败 |
| 安装目标 | Codex、Claude Code、OpenCode 的项目目录及自定义完整包目录；逐文件字节核对 |
| 独立运行 | 在临时项目测试中，无原仓库工作目录也能读取文本、表格；输入哈希不变 |
| 原始运行引擎 | 5 个包内 Python 模块与此前冻结 v7 的字节完全一致；路径重组未改诊断实现 |
| 真实 22 页输入 | 独立安装目录执行完整提取与渲染，22 页对象清单和 22 张 PNG；6 个表格、3 个图表、19 个图片对象 |
| 原始 PPTX | SHA-256 `97072d949785756fc582421e5b56afdb9f8a57d536f6862794a355567944cff1`，保持不变 |
| 渲染目视检查 | 实际查看新渲染的第 2、5、16 页，确认中文、表格和图示实际显示；本次未新作 22 页语义诊断 |
| 格式与资源 | 标准 frontmatter、包内资源与文档链接核对通过；仓库入口明确将资源解析到通用包 |

新增测试先在缺少完整包/安装器时观察到安装与独立运行失败，实施后通过；默认拒绝覆盖现有安装，显式更新保留旧目录备份。包内已含运行代码、Python 依赖说明、11 条建议和 schema 说明，不依赖仓库外文档。

本地旧版 quick_validate 不认识标准的 compatibility 字段；本轮按 Agent Skills 规范检查该字段类型与长度，以及 name/description/允许字段，未以删除标准字段来冒充该旧工具通过。

## 运行方法

```bash
python -m pip install -r requirements.txt
python -m unittest discover -s tests -p 'test_*.py' -v
python tools/install_skill.py --agent claude --project "PROJECT"
python "SKILL_DIR/scripts/analyze_pptx.py" "INPUT.pptx" --out "OUTPUT_DIR"
```

真实材料验收需要本地 22 页示例、27 页实际汇报及旧的绑定清单；公开仓库不附这些私人输入和生成结果。缺少材料的跳过不算真实材料验收通过。实际渲染回归依赖 LibreOffice。

## 原有视觉验收与能力边界

此前 v7 开发已核对 22 页示例全部 11 组与 27 页实际汇报全部页面，另有新的受控材料检查保持良好页面与限定数值比较范围；最终校准结果不宣称是新盲测。本轮验证针对包的安装、资源解析、入口兼容与运行行为，不宣称重新执行整套模型语义审阅，也未启动各框架客户端的真实审阅会话。

脚本准备资料不等于 Agent 已看图。完整流程要求能读取文件、执行命令、渲染并查看图片的宿主；文本/OCR、哈希绑定或 PDF 文字层不能替代实际视觉核对。核心字形/主图无法确认时保留 partial/failed。

SmartArt、OLE、动画阶段、链接外部数据、复杂继承与字体替代仍可能需要人工确认。图片内容须实际辨认，并区分精确读数、近似值与无法辨认项。未提供论文原文不推造引用或研究结论。旧 v2/v3 文件仅为历史兼容资料，不能给旧判断绑定新渲染并冒充新审阅。

见 [使用说明](../skills/academic-pptx-review/references/usage.md)与[框架说明](../skills/academic-pptx-review/references/frameworks.md)。
