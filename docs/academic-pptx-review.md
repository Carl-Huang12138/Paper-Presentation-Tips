# 学术汇报 PPTX 诊断 Skill 使用说明

本阶段只做两件事：先读取 PPTX 并核对放映画面，再判断哪些问题会妨碍听众理解。入口不会保存或改写输入文件。

## 环境和运行

需要 Python 3.10+、`python-pptx`、PyMuPDF (`fitz`) 与 LibreOffice。先执行 `python -m pip install -r requirements.txt`。Windows 上脚本默认使用 `C:/Program Files/LibreOffice/program/soffice.com`；若安装在别处，请修改 `render()` 中的程序路径。运行：

```powershell
python analyze_pptx.py input.pptx --out output_dir
```

第一轮得到 `inventory.json`、`diagnosis.json`、`report.md` 和 `renders/slide-XX.png`。`diagnosis.json` 的 `review_mode=preliminary_only` 意味着只给出需要人工核对的线索，尚不能称为最终质量结论。

## 人工核对并生成最终诊断

逐页看 PNG，再从 `inventory.json` 找对应对象的 `path`。原生表格各单元格在 `table.rows`，图表类别和缓存数据在 `chart`，文字段落和可读取的字体、颜色在 `text_frame`。组合对象在 `children`。`notes_not_visible` 和隐藏状态不属于放映内容。图片内文字和图形要从 PNG 查看；无法确认时写“未知”。

审阅 JSON 须覆盖全部页：

```json
{
  "input_sha256": "输入 PPTX 的 SHA-256",
  "slides": [
    {
      "slide": 1,
      "render_sha256": "第 1 页 PNG 的 SHA-256",
      "role": "背景/动机",
      "status": "有确定问题",
      "basis": "从对象与渲染图共同观察到的判断依据",
      "visual_observations": [
        {"observation": "实际看到的元素", "source": "rendered_png", "confidence": "高"}
      ],
      "issues": [
        {
          "object": "shape-3",
          "category": "内容叙事",
          "evidence": "对象和画面上的事实",
          "impact": "对听众理解的具体影响",
          "severity": "中",
          "confidence": "高",
          "tips": [2, 9],
          "direction": "修改方向"
        }
      ],
      "uncertainties": []
    }
  ]
}
```

`status` 只能为 `无需修改`、`有确定问题`、`需要人工确认`。前者必须没有 `issues`。问题对象要么是清单里的 `path`，要么是 `region:` 开头的可定位区域；严重程度和置信度为高、中、低。每页 PNG 哈希和输入哈希不匹配时脚本拒绝导入审阅记录，避免拿旧审阅覆盖新文件。

生成最终报告：

```powershell
python analyze_pptx.py input.pptx --out output_dir --review review.json
```

报告先列整套主要发现，随后逐页给出角色、结论、证据、影响和修改方向，最后列解析盲区。`diagnosis.json` 可供后续处理，但不要把它视为论文事实核查结论。

## 本仓库测试材料

在 `Paper-Presentation-Tips-editable-22-v2.pptx` 上逐页独立审阅后，验收记录保存在 `validation/v2_review.json`。它与输入和渲染图的哈希绑定。`README-editable-testdeck.md` 和源码 ZIP 仅在独立审阅后用于交叉检查对象数和正反例覆盖；测试脚本或标签不是诊断答案。

## 已知能力边界

LibreOffice 的版式和字体替换可能与 PowerPoint 不完全相同。图表缓存可能缺失；自动轴刻度、SmartArt 结构、OLE、图片里的原始数据无法由普通 PPTX 对象 API 完整读取。脚本不对位图图表猜测数值，且不会凭颜色数、字数或表格存在直接定罪。渲染失败会在对应页和报告中显式记录。
