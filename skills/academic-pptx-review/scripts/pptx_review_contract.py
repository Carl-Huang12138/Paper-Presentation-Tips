"""Validate evidence units and completion independently of presentation quality."""
from __future__ import annotations
import hashlib
import re
from pathlib import Path


def aggregate_completion(entries):
    """One rule for record validation, CLI and author summary."""
    states = [entry.get('review_completion') for entry in entries]
    if not states or any(state not in ('complete', 'partial', 'failed') for state in states):
        raise ValueError('逐页 review_completion 须为 complete/partial/failed')
    if all(state == 'complete' for state in states):
        return 'complete'
    if all(state == 'failed' for state in states):
        return 'failed'
    return 'partial'


def validate_render_fidelity(entry, *, required=False):
    """Check the declared visual inspection; never infer glyphs from PDF text."""
    record = entry.get('render_fidelity')
    if record is None and not required:
        return []
    if not isinstance(record, dict):
        raise ValueError('新快照须记录 render_fidelity 实际画面核对')
    if record.get('status') not in ('confirmed', 'partial', 'failed', 'not_checked'):
        raise ValueError('render_fidelity.status 无效')
    require_text(record, ['evidence'], 'render_fidelity')
    checks = record.get('checks')
    if not isinstance(checks, dict):
        raise ValueError('render_fidelity.checks 须记录字形、符号、主图与动态字段')
    require_text(checks, ['text_glyphs', 'math_symbols', 'main_media', 'dynamic_fields'], 'render_fidelity.checks')
    pending = record.get('blocking_unknowns')
    strings(pending, 'render_fidelity.blocking_unknowns')
    if record['status'] == 'confirmed':
        if pending:
            raise ValueError('画面 confirmed 不能仍有主体缺失/阻塞未知')
        return []
    if not pending:
        raise ValueError('画面未确认须具体列出 render_fidelity.blocking_unknowns')
    return pending


def nonempty(value):
    return isinstance(value, str) and bool(value.strip())


def require_text(obj, names, prefix):
    if not isinstance(obj, dict):
        raise ValueError(f'{prefix} 须为对象')
    for key in names:
        if not nonempty(obj.get(key)):
            raise ValueError(f'{prefix}.{key} 须为非空文字')


def strings(value, prefix, allow_empty=True):
    if not isinstance(value, list) or (not value and not allow_empty) or any(not nonempty(x) for x in value):
        raise ValueError(f'{prefix} 须为非空文字组成的列表')


def descendants(obj):
    yield obj
    for child in obj.get('children', []):
        yield from descendants(child)


def text_units(obj):
    """Never create a contiguous quote across paragraphs, cells or children."""
    for item in descendants(obj):
        frame = item.get('text_frame', {})
        paragraphs = frame.get('paragraphs')
        if isinstance(paragraphs, list) and paragraphs:
            yield from (p.get('text', '') for p in paragraphs)
        else:
            yield from frame.get('text', '').splitlines()
        for row in item.get('table', {}).get('rows', []):
            for cell in row:
                yield cell.get('text', '')


def selected_units(obj, selector, objects):
    if selector is None:
        return list(text_units(obj))
    if not isinstance(selector, dict):
        raise ValueError('selector 须为对象')
    kind = selector.get('kind')
    if kind == 'object':
        path = selector.get('path')
        if not nonempty(path) or path not in objects or not (path == obj['path'] or path.startswith(obj['path'] + '/')):
            raise ValueError('子对象 selector 不属于实际目标')
        return list(text_units(objects[path]))
    if kind == 'cell':
        row, col = selector.get('row'), selector.get('column')
        rows = obj.get('table', {}).get('rows', [])
        if type(row) is not int or type(col) is not int or row < 1 or col < 1 or row > len(rows) or col > len(rows[row-1]):
            raise ValueError('cell selector 须指向实际单元格（行列从1开始）')
        cell = rows[row-1][col-1]
        origin = cell.get('merge_origin')
        if cell.get('is_spanned') and origin:
            raise ValueError(f'合并续格不可充当独立值来源，请引用 merge origin {origin}')
        return [cell.get('text', '')]
    if kind == 'paragraph':
        index = selector.get('index')
        paragraphs = obj.get('text_frame', {}).get('paragraphs', [])
        if type(index) is not int or not 1 <= index <= len(paragraphs):
            raise ValueError('paragraph selector 须指向实际段落（从1开始）')
        return [paragraphs[index-1].get('text', '')]
    if kind == 'chart_point':
        p, s, n = (selector.get(k) for k in ('plot', 'series', 'point'))
        if any(type(x) is not int or x < 1 for x in (p, s, n)):
            raise ValueError('chart_point 的 plot/series/point 从1开始')
        try:
            plots = obj['chart']['plots']
            point = plots[p-1]['series'][s-1]['points'][n-1]
            component = selector.get('component', 'y')
            value = point.get(component)
            if component not in ('x', 'y', 'size', 'value') or value is None or value == 'unknown':
                raise KeyError(component)
            return [str(value)]
        except (KeyError, IndexError, TypeError):
            raise ValueError('chart_point 来源未读取到指定数据') from None
    raise ValueError('selector.kind 仅支持 cell/paragraph/object/chart_point')


def validate_anchor(anchor, target, objects, page, *, strict=False):
    prefix = f'第 {page} 页 target_evidence'
    require_text(anchor, ['source'], prefix)
    if not nonempty(target):
        raise ValueError(f'{prefix}缺少 object')
    if target.startswith('region:'):
        if not target[7:].strip():
            raise ValueError(f'{prefix} region 必须具体定位')
        require_text(anchor, ['description'], prefix)
        return
    if target not in objects:
        raise ValueError(f'{prefix}对象不在清单中：{target}')
    obj = objects[target]
    selector = anchor.get('selector')
    if strict and anchor.get('quote') and ('table' in obj or obj.get('children')) and selector is None:
        raise ValueError(f'{prefix}表格/组合引用须有具体 cell/子对象 selector')
    units = selected_units(obj, selector, objects)
    # Whitespace within one real unit may differ between OOXML and the render.
    norm = lambda x: re.sub(r'\s+', '', x)
    readable = [norm(x) for x in units if isinstance(x, str) and norm(x)]
    if readable:
        require_text(anchor, ['quote'], prefix)
        if not any(norm(anchor['quote']) in unit for unit in readable):
            raise ValueError(f'{prefix}引用不在目标对象的指定段落/单元格/子对象中；不能跨来源拼接')
    else:
        require_text(anchor, ['description'], prefix)


def validate_target(issue, objects, page, required, *, strict=False):
    anchor = issue.get('target_evidence')
    if anchor is None and not required:
        return
    if isinstance(anchor, dict) and 'refs' in anchor:
        refs = anchor['refs']
        if not isinstance(refs, list) or not refs or any(not isinstance(ref, dict) for ref in refs):
            raise ValueError(f'第 {page} 页 target_evidence.refs 须为非空引用列表')
        require_text(anchor, ['source'], 'target_evidence')
        for ref in refs:
            if ref.get('slide', page) != page:
                raise ValueError('target_evidence.refs 跨页请使用 source_refs')
            validate_anchor(ref, ref.get('object', issue['object']), objects, page, strict=strict)
        return
    validate_anchor(anchor, issue['object'], objects, page, strict=strict)


def validate_observations(entry):
    observations = entry.get('visual_observations', [])
    if not isinstance(observations, list):
        raise ValueError('visual_observations 须为列表')
    for item in observations:
        require_text(item, ['observation', 'source'], 'visual_observations')
        if item.get('confidence') not in ('高', '中', '低'):
            raise ValueError('visual_observations.confidence 无效')
    uncertainties = entry.get('uncertainties', [])
    if not isinstance(uncertainties, list):
        raise ValueError('uncertainties 须为列表')
    for doubt in uncertainties:
        if isinstance(doubt, dict):
            require_text(doubt, ['evidence', 'action'], 'uncertainties')
        elif not nonempty(doubt):
            raise ValueError('uncertainties 须有具体内容')


def media_blockers(entry, *, strict=False):
    result = []
    for record in entry.get('media_reviews', []):
        importance = record.get('importance', 'core')
        if strict and importance not in ('core', 'secondary'):
            raise ValueError('media_reviews.importance 须为 core/secondary')
        if strict and 'importance' not in record:
            raise ValueError('v4 media_reviews 须声明核心/次要作用')
        pending = record.get('blocking_unknowns', [])
        strings(pending, 'media_reviews.blocking_unknowns')
        result.extend(pending)
        if record['status'] == 'reviewed':
            continue
        nonblocking = (importance == 'secondary' or
                       (record.get('unreadable_scope') == 'secondary' and bool(record.get('recognized'))))
        if nonblocking and nonempty(record.get('nonblocking_reason')) and not pending:
            continue
        result.append(f"{record['object']}：核心可见内容未完成检查")
    return result


def validate_completion(entry, slide, inventory, version, *, require_fidelity=False):
    render = slide.get('render', {})
    blockers = media_blockers(entry, strict=version == 4)
    unavailable = (render.get('status') != 'ok' or not render.get('image') or
                   (render.get('mapping_status') != 'verified' if version == 4 else
                    render.get('mapping_status') in ('mapping_unverified', 'unverified')))
    if version == 4:
        blockers.extend(validate_render_fidelity(entry, required=require_fidelity))
        completion = entry.get('review_completion')
        if completion not in ('complete', 'partial', 'failed'):
            raise ValueError('v4 review_completion 须为 complete/partial/failed')
        if entry.get('quality_status') != entry.get('status'):
            raise ValueError('quality_status 和 status 不一致')
        pending = entry.get('blocking_unknowns')
        strings(pending, 'blocking_unknowns')
        if (unavailable or blockers or pending) and completion == 'complete':
            raise ValueError('核心证据或页图映射未完成，不能标 complete')
        if completion != 'complete' and entry['status'] not in ('未完成', '需要人工确认'):
            raise ValueError('部分/失败审阅不能给出完整页面质量结论')
        if completion == 'complete' and entry['status'] == '未完成':
            raise ValueError('complete 与未完成状态矛盾')
        if unavailable and not pending:
            raise ValueError('未完成渲染/映射须列 blocking_unknowns')
    elif version >= 3 and (unavailable or (blockers and entry.get('status') == '无需修改')):
        raise ValueError('渲染/映射/核心图片未完成，旧记录不能充当完整质量结论')
    image = render.get('image')
    if image:
        path = Path(image)
        if not path.is_absolute():
            path = Path(inventory.get('output_root', '.')) / path
        if not path.is_file() or entry.get('render_sha256') != hashlib.sha256(path.read_bytes()).hexdigest():
            raise ValueError(f"第 {entry['slide']} 页审阅图像哈希不匹配或不存在")
    elif version >= 3 and entry.get('review_completion') == 'complete':
        raise ValueError('完整审阅必须绑定真实渲染图')


def validate_suggestions(entry):
    for key, kind, fields in [('optional_suggestions', 'optional', ['object', 'evidence', 'direction', 'benefit']),
                              ('context_questions', 'needs_context', ['evidence', 'action'])]:
        items = entry.get(key, [])
        if not isinstance(items, list):
            raise ValueError(f'{key} 须为列表')
        for item in items:
            require_text(item, fields, key)
            if item.get('recommendation_kind') != kind:
                raise ValueError(f'{key}.recommendation_kind 须为 {kind}')


def validate_source_refs(refs, inventory, *, prefix='source_refs'):
    if not isinstance(refs, list) or not refs:
        raise ValueError(f'{prefix} 须为非空来源列表')
    from analyze_pptx import flatten
    for ref in refs:
        if not isinstance(ref, dict) or type(ref.get('slide')) is not int or not 1 <= ref['slide'] <= inventory['slide_count']:
            raise ValueError(f'{prefix} 页码无效')
        slide = inventory['slides'][ref['slide']-1]
        objects = {o['path']: o for o in flatten(slide['objects'])}
        validate_anchor(ref, ref.get('object'), objects, ref['slide'], strict=True)


def numeric_leaves(data, path=''):
    if isinstance(data, dict):
        for key, value in data.items():
            yield from numeric_leaves(value, f'{path}.{key}' if path else key)
    elif isinstance(data, list):
        for index, value in enumerate(data):
            yield from numeric_leaves(value, f'{path}.{index}')
    elif type(data) in (int, float):
        yield path, data


def validate_numeric_sources(issue, inventory):
    for check in issue.get('numeric_checks', []):
        refs = check.get('source_refs')
        if not isinstance(refs, list):
            raise ValueError('v4 numeric_checks 须登记每项 input 的 source_refs')
        expected = dict(numeric_leaves(check['inputs']))
        seen = set()
        for ref in refs:
            if not isinstance(ref, dict) or ref.get('input') not in expected or ref['input'] in seen:
                raise ValueError('numeric_checks source_refs input 无效或重复')
            key = ref['input']
            seen.add(key)
            if type(ref.get('value')) not in (int, float) or ref['value'] != expected[key]:
                raise ValueError('numeric_checks source_refs value 与 input 不一致')
            if ref.get('source') == 'test_assumption':
                require_text(ref, ['basis'], 'test_assumption')
            else:
                validate_source_refs([ref], inventory, prefix='numeric source')
                quote = ref.get('quote')
                try:
                    matched = float(quote.strip().rstrip('%')) == float(ref['value'])
                except (ValueError, AttributeError):
                    matched = False
                if not matched:
                    raise ValueError('numeric source quote 须独立表示该输入值；未知转换请另记录工具步骤')
                if not ref['object'].startswith('region:'):
                    from analyze_pptx import flatten
                    source_slide = inventory['slides'][ref['slide']-1]
                    objects = {o['path']: o for o in flatten(source_slide['objects'])}
                    units = selected_units(objects[ref['object']], ref.get('selector'), objects)
                    tokens = [token for unit in units for token in re.findall(
                        r'(?<![\w.])[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?(?![\w.])', unit)]
                    if not any(float(token) == float(ref['value']) for token in tokens):
                        raise ValueError('数值来源单位没有该独立数值，不能截取另一数字的子串')
        if seen != set(expected):
            raise ValueError(f'numeric_checks 来源未覆盖 inputs：{sorted(set(expected)-seen)}')


def validate_disposition(issue, inventory):
    refs = issue.get('source_refs')
    if refs is not None or issue.get('action') != 'local':
        validate_source_refs(refs, inventory)
    disposition = issue.get('content_disposition')
    if issue.get('action') != 'local' and (not isinstance(disposition, list) or not disposition):
        raise ValueError('v4 结构修改须登记 content_disposition，记录原内容去向')
    if disposition is not None:
        if not isinstance(disposition, list):
            raise ValueError('content_disposition 须为列表')
        for item in disposition:
            require_text(item, ['destination', 'reason'], 'content_disposition')
            if item.get('disposition') not in ('retain', 'move', 'backup', 'deduplicate'):
                raise ValueError('content_disposition.disposition 无效')
            validate_source_refs([item.get('source_ref')], inventory, prefix='content_disposition.source_ref')
        if refs is not None:
            def identity(ref):
                import json
                return (ref.get('source'), ref.get('slide'), ref.get('object'),
                        json.dumps(ref.get('selector'), sort_keys=True),
                        re.sub(r'\s+', '', ref.get('quote', '')),
                        re.sub(r'\s+', '', ref.get('description', '')))
            allocated = {identity(item['source_ref']) for item in disposition}
            if any(identity(ref) not in allocated for ref in refs):
                raise ValueError('content_disposition 须为每个 source_ref 登记对应去向，不能借无关对象代替')
    pages = issue.get('split_plan', [])
    if pages:
        for proposed in pages:
            validate_source_refs(proposed.get('source_refs'), inventory, prefix='split page.source_refs')
    for alternative in issue.get('alternatives', []):
        for proposed in alternative.get('pages', []):
            validate_source_refs(proposed.get('source_refs'), inventory, prefix='alternative page.source_refs')


def validate_plan_v4(issue, page):
    prefix = f'第 {page} 页修改方案'
    action = issue.get('action')
    if action not in ('local', 'revise', 'split'):
        raise ValueError(f'{prefix} action 须为 local/revise/split')
    plan = issue.get('revision_plan')
    require_text(plan, ['check'] if action == 'local' else ['goal', 'layout', 'typography', 'check'], prefix)
    strings(plan.get('steps'), prefix + '.steps', allow_empty=False)
    if 'parameters' in issue:
        raise ValueError('parameters 须放 revision_plan 内')
    if 'parameters' in plan:
        strings(plan['parameters'], prefix + '.parameters')

    def pages_valid(pages):
        if not isinstance(pages, list) or len(pages) < 2:
            raise ValueError('split 必须给出至少两页完整方案')
        for proposed in pages:
            require_text(proposed, ['title', 'message', 'layout'], 'split page')
            strings(proposed.get('content'), 'split.content', False)
            strings(proposed.get('components'), 'split.components', False)
    if action == 'split':
        pages_valid(issue.get('split_plan'))
    elif 'split_plan' in issue:
        raise ValueError('v4 备用拆页请放 alternatives，不混用顶层 split_plan')
    alternatives = issue.get('alternatives', [])
    if not isinstance(alternatives, list):
        raise ValueError('alternatives 须为列表')
    for alternative in alternatives:
        require_text(alternative, ['activation_condition'], 'alternative')
        if alternative.get('action') == 'split':
            pages_valid(alternative.get('pages'))
        else:
            raise ValueError('当前 alternatives 仅接受明确的 split 方案')
