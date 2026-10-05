"""Author-facing report: conclusions and actions, with evidence in a separate file."""
from collections import Counter
from pptx_review_contract import aggregate_completion


def compact_report(inventory, diagnoses, render_issues, deck_assessment=None, *, mode='preliminary_only'):
    counts = Counter(d['status'] for d in diagnoses)
    completion = Counter(d.get('review_completion', 'legacy_unknown') for d in diagnoses)
    total_completion = (aggregate_completion(diagnoses) if diagnoses and
                        all(d.get('review_completion') in ('complete', 'partial', 'failed') for d in diagnoses)
                        else 'legacy_unknown')
    lines = ['# 学术汇报 PPTX 审阅', '',
             f"输入：`{inventory['input']}`", f"输入 SHA-256：`{inventory['sha256']}`",
             f"快照：`{inventory.get('snapshot_id', '历史记录未绑定快照')}`",
             f"模式：{mode}；页面结论：{dict(counts)}；检查完成度：{dict(completion)}", '']
    lines += [f'整套审阅完成度：{total_completion}', '']
    if mode == 'failed_visual_review':
        lines += ['本次全部页面未能完成视觉审阅；以下只报告已检查范围与阻塞项，不能认证页面质量。', '']
    if mode.startswith('legacy'):
        lines += ['这是历史记录导入，不代表最新版 Skill 新鲜执行或页图映射重新验证。', '']
    if mode in ('preliminary_only', 'failed_preparation'):
        lines += ['已准备读取资料，尚未完成逐页视觉质量审阅。请打开 review_packet.md，实际看图后填写新记录。', '']
    if deck_assessment:
        lines += [f"汇报目标：{deck_assessment['purpose']}", '',
                  f"论述结构：{deck_assessment['storyline']}", '',
                  f"整体观感：{deck_assessment['visual_system']}", '']
        priorities = deck_assessment.get('priorities', [])
        if priorities:
            lines += ['优先修改：', ''] + [f'{i}. {p}' for i, p in enumerate(priorities, 1)] + ['']
        elif any(d.get('issues') for d in diagnoses):
            lines += ['未列修改顺序；必要问题见逐页结论。', '']
    lines += ['每页结论和方案如下。完整对象、11项检查、图片证据及数值核对见 [详细记录](report-details.md)。', '']
    for slide, d in zip(inventory['slides'], diagnoses):
        title = d.get('confirmed_title') or slide.get('title') or '标题未确认'
        if not d.get('confirmed_title') and slide.get('title_status') == 'candidate':
            title += '（候选）'
        lines += [f"## 第 {d['slide']} 页 · {title}", '',
                  f"**{d['status']}**；检查：{d.get('review_completion', '历史完成度未知')}；作用：{d['role']}", '', d['basis'], '']
        if slide.get('render', {}).get('status') == 'ok':
            lines += [f"[整页渲染](renders/slide-{d['slide']:02d}.png)", '']
        if message := d.get('assessment', {}).get('core_message'):
            lines += [f'核心信息：{message}', '']
        assessment = d.get('assessment', {})
        if assessment.get('content_organization') and assessment.get('visual_structure'):
            lines += [f"整体判断：{assessment['content_organization']} {assessment['visual_structure']}", '']
        for issue in d.get('issues', []):
            lines += [f"- **必要修改 · {issue['severity']} · {issue['object']}**：{issue['evidence']}",
                      f"  - 影响：{issue['impact']}", f"  - 修改：{issue['direction']}"]
            if issue.get('tips'):
                labels = {1:'信息提炼', 2:'核心任务', 3:'标题', 4:'色彩', 5:'布局', 6:'动机案例',
                          7:'例子与公式', 8:'创新提炼', 9:'重点定位', 10:'比较表达', 11:'总结'}
                lines.append('  - 建议依据：' + ' / '.join(f'Tip {tip} {labels[tip]}' for tip in issue['tips']))
            plan = issue.get('revision_plan', {})
            if plan.get('layout'):
                lines.append(f"  - 结构：{plan['layout']}")
            lines += [f'  - 操作：{step}' for step in plan.get('steps', [])]
            if plan.get('typography'):
                lines.append(f"  - 字体：{plan['typography']}")
            lines.append(f"  - 复核：{plan.get('check', '改稿后渲染核对')}")
            for label, pages in [('拆页', issue.get('split_plan', []))] + [
                    (f"备选拆页（启用条件：{a['activation_condition']}）", a['pages'])
                    for a in issue.get('alternatives', [])]:
                if pages:
                    lines.append(f'  - {label}：')
                    for i, page in enumerate(pages, 1):
                        lines += [f"    - 新页{i} **{page['title']}**：{page['message']}；内容：{'；'.join(page['content'])}；结构：{page['layout']}；组件：{'；'.join(page['components'])}"]
        for suggestion in d.get('optional_suggestions', []):
            lines.append(f"- 可选改善（不影响当前质量结论）：{suggestion['direction']}；收益：{suggestion['benefit']}")
        for question in d.get('context_questions', []):
            lines.append(f"- 上下文待确认：{question['evidence']}；{question['action']}")
        for item in d.get('blocking_unknowns', []):
            lines.append(f'- 阻塞检查：{item}')
        if record := d.get('render_fidelity'):
            lines.append(f"- 画面核对：{record['status']}；{record['evidence']}")
        lines.append('')
    limits = []
    for slide in inventory['slides']:
        limits.extend(f"第{slide['slide']}页：{x}" for x in slide.get('parse_limitations', []))
    lines += ['## 读取边界', '', *[f'- {error}' for error in render_issues],
              f'- 解析边界已逐页登记（共 {len(limits)} 条），详见读取资料包和详细记录；它们不自动是页面缺陷。',
              '- 哈希和字段校验不证明实际看过图，也不证明语义判断正确。',
              '- 所有未实施的布局和拆页提案都需改稿后渲染验证。', '']
    return '\n'.join(lines)
