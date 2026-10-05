# 安装与不同 Agent 宿主的使用

## 通用完整包

唯一维护源是 `skills/academic-pptx-review/`，包括 SKILL.md、requirements.txt、scripts/ 和 references/。可以整体复制、作为 ZIP 上传到支持资源文件的宿主，或由安装工具复制到框架的技能目录。只复制 SKILL.md 不包含运行依赖。

Agent Skills 格式可移植不等于任意框架自动支持。宿主必须加载完整指令并能访问包内资源；需要 Python 3.10+、包内 Python 依赖、LibreOffice、文件读写、执行命令和真正的图片查看能力。这里的路径加载兼容、脚本运行与模型语义审阅分别验收；本次安装测试不声称已运行各客户端的真实会话。

## 在其他项目安装

在原仓库根目录执行下列命令，PROJECT 替换为实际项目绝对路径：

```bash
python tools/install_skill.py --agent codex --project "PROJECT"
python tools/install_skill.py --agent claude --project "PROJECT"
python tools/install_skill.py --agent opencode --project "PROJECT"
```

分别安装到 PROJECT 下的 `.agents/skills/academic-pptx-review/`、`.claude/skills/academic-pptx-review/`、`.opencode/skills/academic-pptx-review/`。每次只安装到实际使用的目标；有些宿主扫描多个目录，应避免重复安装同名 Skill。

其他宿主或用户级安装可指定确切技能目录（末尾目录名须为 academic-pptx-review）：

```bash
python tools/install_skill.py --destination "ABSOLUTE_SKILLS_DIR/academic-pptx-review"
```

目标已有内容时默认拒绝覆盖。明确需要更新时加 `--overwrite`，旧目录会移动到同级带时间戳的备份，便于恢复。安装工具不安装 Python 依赖、不配置宿主权限、不自动注册自定义 Agent。

## 执行与调用

先安装依赖；SKILL_DIR 替换为实际安装目录绝对路径：

```bash
python -m pip install -r "SKILL_DIR/requirements.txt"
python "SKILL_DIR/scripts/analyze_pptx.py" "INPUT.pptx" --out "OUTPUT_DIR"
```

Claude Code 在可发现的项目入口中可使用 `/academic-pptx-review`；Codex CLI/IDE 可使用 `$academic-pptx-review` 或技能选择器。其他宿主按其激活方式加载 SKILL.md。通用提示词例如：

> 使用 academic-pptx-review 审阅 INPUT.pptx，输出到 OUTPUT_DIR。先准备对象和渲染资料，实际查看每一页及必要局部，按包内规则写本次快照的 review.json，再用 --review 与 --require-evidence 导入。只分析报告，不修改输入；足够好的页面明确无需修改；核心未知保留未完成范围。

无 `/`、`$` 或原生技能加载器的框架，可将 SKILL.md 注入指令上下文，并告知其包所在绝对路径；按需提供 references/，保留全部审阅和能力边界。把完整包挂载到运行环境，并提供文件读取、执行命令、查看图片的工具。接口工具名可以不同，能力须对应。

## 能力不足时

只有文本/OCR或没有本地执行能力时，可提供有限内容审阅或请求用户提供渲染材料，但不能声称已完成本 Skill 的完整流程。已生成 PNG 不代表审阅者看过图。核心字形、图表或证据无法确认时保留 partial/failed 和具体阻塞项，不能编造数值或以文件哈希替代视觉检查。

## 仓库维护入口

仓库内 `.agents` 与 `.claude` 的 SKILL.md 是小型发现入口，由下列命令生成。它们让宿主读取同一份通用 SKILL.md，完整安装使用前面的安装命令。

```bash
python tools/install_skill.py --sync-entries --project .
```

修改规则与代码只修改通用包，随后同步入口；根目录的 Python 文件只提供兼容转发。不在客户端目录另写一套规则。根目录仍可运行 `python analyze_pptx.py input.pptx --out output_dir`。

## 官方格式与加载资料

- [Agent Skills 规范](https://agentskills.io/specification)：包内格式与资源约定。
- [集成说明](https://agentskills.io/client-implementation/adding-skills-support)：自定义框架的发现、加载和资源访问。
- [Codex](https://learn.chatgpt.com/docs/build-skills)、[Claude Code](https://code.claude.com/docs/en/skills)、[OpenCode](https://opencode.ai/docs/skills/)：各自的发现目录和调用方式。支持范围会随客户端版本变化，安装后应确认宿主加载的实际路径。
