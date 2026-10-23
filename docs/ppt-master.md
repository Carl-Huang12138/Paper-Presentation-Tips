# 学术 PPT 生成与润色

面向学术汇报的 PPT 生成与润色 Skill，支持将论文和研究材料组织为汇报演示文稿，也可优化已有 PPTX 的内容结构、页面布局、字体和图表表达。

先规划每页的核心任务与视觉呈现，在聊天中确认逐页参考图，再制作原生可编辑 PPTX 并检查实际显示效果。在保留事实、数据和来源的基础上，提升信息层级、阅读顺序与整体观感。

## 成品展示

### 已有 PPT 润色：思维链鲁棒性

**Prompt**

> Ppt Master 请修改润色组会1.pptx的页面

[![思维链鲁棒性汇报：左侧润色前，右侧润色后](images/meeting1-before-after.jpg)](images/meeting1-before-after.jpg)

左侧为原稿，右侧为润色版。

沿用原稿的学术模板，调整内容组织、公式呈现和表格布局。文字、表格和组合图形可以继续编辑。

### 从论文生成：中国菜品消费与温室气体减排

**Prompt**

> Ppt Master 请根据GHG mitigation strategies on China’s diverse dish consumption are key to meet the Paris Agreement targets.pptx的内容生成一份PPT

[![中国菜品消费与温室气体减排汇报预览](images/ghg-preview.jpg)](images/ghg-preview.jpg)

用图解串起研究问题、核算方法和减排情景，包含可编辑图表、表格及独立的论文图片。

## 开始使用

需要 Git、Python 3.10+，以及支持 Agent Skills、具备文件读写和命令执行能力的 Agent。制作时还需要图片查看工具，以及 PowerPoint、WPS 或 LibreOffice 等 PPTX 渲染工具。

### 1. 从 GitHub 获取

```powershell
git clone https://github.com/MLNLP-World/Paper-Presentation-Tips.git
cd Paper-Presentation-Tips
python -m pip install -r skills/ppt-master/requirements.txt
```

### 2. 安装到 Agent

在克隆目录中启动 Codex 或 Claude Code，可直接使用仓库入口。安装到其他项目时，在仓库根目录选择对应命令，将 PROJECT 换成目标项目的绝对路径：

```powershell
# Codex
python tools/install_skill.py --skill ppt-master --agent codex --project "PROJECT"

# Claude Code
python tools/install_skill.py --skill ppt-master --agent claude --project "PROJECT"
```

Codex 使用 `$ppt-master`，Claude Code 使用 `/ppt-master`。其他支持 Agent Skills 的工具也可以使用整个 `skills/ppt-master/` 目录。

### 3. 开始制作

在对话中附上材料或提供准确路径，例如：

> 用这篇论文制作一份中文组会 PPT，面向本领域研究生，汇报约 15 分钟。沿用提供的背景，先展示逐页参考图，确认后再制作可编辑 PPTX。

已有 PPT 润色示例：

> 润色这份 PPT，重点检查内容组织、布局与整体观感，保留已经讲清楚的页面；需要调整的页先给参考图，确认后再修改。

## 设计重点

- 一页有一个清楚的核心任务，标题能说明本页讲什么。
- 先用案例说明问题，再介绍概念、方法和公式。
- 文字尽量少而大，用图示表达关系，让听众一眼看到重点。
- 实验突出主比较；数据可靠、表达更清楚时，将大表格改成合适的图。
- 总结回扣问题、做法和已有证据。足够好的页面可以保留。

详细的 12 条原则见 [学术设计规则](../skills/ppt-master/references/academic-chat-design.md#4-设计规则逐页应用)。原仓库的建议与资源整理在 [Presentation Tips](../README.md#-presentation-tips)。

## 制作方式

**理解材料 → 规划汇报 → 设计参考页 → 聊天确认 → 可编辑制作 → 实际渲染与调整**

规划阶段查看合适的风格、版式、图表和图标。生成参考图前，将每页的文字、主图、数据、阅读顺序和背景整理为视觉简报。

默认使用 Agent 自带的绘图工具。Claude 等宿主没有内置生图工具时，可选用已配置的外部 API，或提供已有参考图。配置和具体用法见 [使用说明](usage.md)。

确认后按参考页制作原生文字、形状、表格、数据图表和公式，再查看实际 PPTX 的显示效果。照片和复杂论文原图保留为独立图片，其内部内容不会自动变成可编辑组件。

## 更多用法

- [安装、材料准备、生图配置与按页反馈](usage.md)
- [单页设计简报与参考图提示词](../skills/ppt-master/references/page-reference-generation.md)

## 来源

基于 [Hugo He / PPT Master](https://github.com/hugohe3/ppt-master)，保留其模板资源和原生 PPTX 转换后端，在此基础上加入学术汇报的内容组织、参考页设计和聊天确认流程。

原代码遵循 [MIT 许可证](../skills/ppt-master/LICENSE)；图标等素材见 [第三方许可说明](../skills/ppt-master/templates/icons/THIRD_PARTY_NOTICES.md)。
