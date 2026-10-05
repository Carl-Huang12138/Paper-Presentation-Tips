"""Evidence and quality-state regressions from the independent 2026-10-02 audit."""
import copy
import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from analyze_pptx import validate_review, validate_target_evidence, validate_polish_plan, markdown
from pptx_report import compact_report


def review_fixture(folder, version=3):
    image = Path(folder) / 'slide.png'
    image.write_bytes(b'render-fixture-for-contract-testing')
    sha = hashlib.sha256(image.read_bytes()).hexdigest()
    inventory = {'input': 'input.pptx', 'sha256': 'input', 'slide_count': 1,
                 'snapshot_id': 'snapshot', 'slides': [{'slide': 1, 'title': '方法',
                 'parse_limitations': [], 'objects': [{'path': 'shape-1', 'visible_in_slideshow': True,
                 'text_frame': {'text': 'CORE CONTENT', 'paragraphs': [{'text': 'CORE CONTENT'}]}}],
                 'render': {'status': 'ok', 'image': str(image), 'mapping_status': 'verified'}}]}
    entry = {'slide': 1, 'render_sha256': sha, 'role': '方法', 'status': '无需修改',
             'basis': '核心概念有清楚顺序和视觉区分。', 'issues': [], 'uncertainties': [],
             'media_reviews': [], 'visual_observations': [],
             'assessment': {'core_message': '解释一个概念', 'content_organization': '定义接解释',
                            'visual_structure': '大字号分层清楚', 'reading_path': '标题接核心概念'},
             'tip_checks': [{'tip': i, 'result': 'adequate', 'basis': '本页这项组织有效'} for i in range(1, 12)]}
    if version == 4:
        entry.update(review_completion='complete', quality_status='无需修改', blocking_unknowns=[],
                     optional_suggestions=[], context_questions=[])
    review = {'review_schema_version': version, 'input_sha256': 'input', 'media_review_version': 1,
              'evidence_review_version': 1, 'snapshot_id': 'snapshot', 'slides': [entry],
              'deck_assessment': {'purpose': '解释机制', 'storyline': '问题接方法和证据',
                                  'visual_system': '统一的文字层级', 'priorities': []}}
    return inventory, review


def issue_fixture():
    return {'object': 'shape-1', 'target_evidence': {'source': 'pptx_object', 'quote': 'CORE'},
            'category': '视觉层级', 'scope': 'presentation', 'action': 'revise', 'severity': '中',
            'confidence': '高', 'evidence': '核心句和细节难以区分', 'impact': '重点难定位',
            'direction': '让核心句成为主区域', 'tips': [9],
            'revision_plan': {'goal': '形成清楚主次', 'layout': '单一主区域',
                              'steps': ['保留核心句并减少重复'], 'typography': '优先放大主体',
                              'check': '重新渲染验证主次及内容保留'}}


class ReviewIntegrity(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.inventory, self.review = review_fixture(self.temp.name)

    def test_v4_rejects_conflicting_per_page_snapshot(self):
        inventory, review = review_fixture(self.temp.name, 4)
        review['slides'][0]['snapshot_id'] = 'DIFFERENT_BATCH'
        with self.assertRaises(ValueError):
            validate_review(review, inventory)

    def fidelity_record(self, status='confirmed'):
        inventory, review = review_fixture(self.temp.name, 4)
        review['render_fidelity_review_version'] = 1
        review['slides'][0]['render_fidelity'] = {
            'status': status, 'evidence': '实际查看整页PNG对照原文字，标题与正文均显示',
            'checks': {'text_glyphs': '标题与正文可见', 'math_symbols': '本页无公式',
                       'main_media': '本页无主体图片', 'dynamic_fields': '本页无动态字段'},
            'blocking_unknowns': []}
        return inventory, review

    def test_matched_pdf_text_does_not_override_missing_visible_glyphs(self):
        inventory, review = self.fidelity_record('failed')
        inventory['slides'][0]['render'].update(comparison_status='matched', pdf_text='CORE CONTENT')
        review['slides'][0]['render_fidelity'].update(
            evidence='文字层匹配，但PNG主体字形没有绘出', blocking_unknowns=['主体字形缺失'])
        with self.assertRaises(ValueError):
            validate_review(review, inventory)

    def test_new_snapshot_cannot_drop_fidelity_review_contract(self):
        inventory, review = review_fixture(self.temp.name, 4)
        inventory['render_fidelity_review_version'] = 1
        with self.assertRaises(ValueError):
            validate_review(review, inventory)

    def test_unknown_fidelity_version_is_rejected(self):
        inventory, review = self.fidelity_record()
        review['render_fidelity_review_version'] = 9
        with self.assertRaises(ValueError):
            validate_review(review, inventory)

    def test_verified_visible_glyphs_and_date_difference_can_complete(self):
        inventory, review = self.fidelity_record()
        review['slides'][0]['render_fidelity']['checks']['dynamic_fields'] = '日期取当前渲染日，与缓存不同；无主体缺失'
        validate_review(review, inventory)

    def test_partial_glyph_review_keeps_core_blocker(self):
        inventory, review = self.fidelity_record('partial')
        entry = review['slides'][0]
        entry.update(review_completion='partial', status='需要人工确认', quality_status='需要人工确认',
                     blocking_unknowns=['数学符号显示待确认'])
        entry['render_fidelity']['blocking_unknowns'] = ['数学符号显示待确认']
        validate_review(review, inventory)

    def test_v4_rejects_stale_preparation_flags_on_complete_review(self):
        inventory, review = review_fixture(self.temp.name, 4)
        review.update(preparation_only=True, review_completion='partial', quality_status='未完成',
                      blocking_unknowns=['尚未检查'])
        with self.assertRaises(ValueError):
            validate_review(review, inventory)

    def test_null_numeric_checks_is_a_validation_error(self):
        inventory, review = review_fixture(self.temp.name, 4)
        entry = review['slides'][0]
        entry.update(status='有确定问题', quality_status='有确定问题')
        item = issue_fixture()
        item.update(action='local', recommendation_kind='necessary', obstacle='核心句被挡住',
                    benefit='恢复读取', numeric_checks=None)
        entry['issues'] = [item]
        entry['tip_checks'][8]['result'] = 'problem'
        with self.assertRaises(ValueError):
            validate_review(review, inventory)

    def with_issue(self):
        item = self.review['slides'][0]
        item['status'] = '有确定问题'
        item['issues'] = [issue_fixture()]
        item['tip_checks'][8]['result'] = 'problem'
        return item

    def test_failed_render_cannot_certify_no_change(self):
        self.inventory['slides'][0]['render'] = {'status': 'failed', 'reason': 'renderer unavailable'}
        with self.assertRaises(ValueError):
            validate_review(self.review, self.inventory)

    def test_unverified_page_mapping_cannot_certify_quality(self):
        self.inventory['slides'][0]['render']['mapping_status'] = 'mapping_unverified'
        with self.assertRaises(ValueError):
            validate_review(self.review, self.inventory)

    def test_unreviewed_core_picture_blocks_no_change(self):
        self.inventory['slides'][0]['objects'].append({'path': 'shape-2', 'visible_in_slideshow': True, 'picture': {}})
        self.review['slides'][0]['media_reviews'] = [{'object': 'shape-2', 'status': 'not_reviewed',
            'source': 'tool_unavailable', 'confidence': '低', 'recognized': [], 'unreadable_items': [],
            'importance': 'core', 'reason': '尚未查看主要结果图'}]
        with self.assertRaises(ValueError):
            validate_review(self.review, self.inventory)

    def test_empty_critical_issue_fields_are_rejected(self):
        entry = self.with_issue()
        for key in ('evidence', 'impact', 'direction', 'category'):
            changed = copy.deepcopy(self.review)
            changed['slides'][0]['issues'][0][key] = ' '
            with self.subTest(key=key), self.assertRaises(ValueError):
                validate_review(changed, self.inventory)

    def test_unmapped_tips_need_an_explicit_reason(self):
        entry = self.with_issue()
        entry['issues'][0]['tips'] = []
        entry['tip_checks'][8]['result'] = 'adequate'
        with self.assertRaises(ValueError):
            validate_review(self.review, self.inventory)
        entry['issues'][0]['tip_mapping_reason'] = '该错误是来源冲突，11项设计提示没有直接覆盖。'
        validate_review(self.review, self.inventory)

    def test_tip_issue_reverse_mapping_is_consistent(self):
        entry = self.with_issue()
        entry['tip_checks'][8]['result'] = 'adequate'
        with self.assertRaises(ValueError):
            validate_review(self.review, self.inventory)

    def test_malformed_observation_is_rejected_before_formatting(self):
        self.review['slides'][0]['visual_observations'] = [{'observation': '已看核心句'}]
        with self.assertRaises(ValueError):
            validate_review(self.review, self.inventory)

    def test_cross_cell_pseudo_quote_is_rejected(self):
        objects = {'shape-1': {'path': 'shape-1', 'table': {'rows': [[{'text': '4'}, {'text': '2'}]]}}}
        issue = issue_fixture()
        issue['target_evidence']['quote'] = '42'
        with self.assertRaises(ValueError):
            validate_target_evidence(issue, objects, 1, True)

    def test_cross_paragraph_pseudo_quote_is_rejected(self):
        objects = {'shape-1': {'path': 'shape-1', 'text_frame': {'text': '4\n2',
                    'paragraphs': [{'text': '4'}, {'text': '2'}]}}}
        issue = issue_fixture()
        issue['target_evidence']['quote'] = '42'
        with self.assertRaises(ValueError):
            validate_target_evidence(issue, objects, 1, True)

    def test_cell_selector_cannot_reference_another_cell(self):
        objects = {'shape-1': {'path': 'shape-1', 'table': {'rows': [[{'text': '4'}, {'text': '2'}]]}}}
        issue = issue_fixture()
        issue['target_evidence'].update(quote='4', selector={'kind': 'cell', 'row': 1, 'column': 2})
        with self.assertRaises(ValueError):
            validate_target_evidence(issue, objects, 1, True)
        issue['target_evidence']['selector']['column'] = 1
        validate_target_evidence(issue, objects, 1, True)

    def test_empty_priority_list_does_not_claim_deck_passed(self):
        self.with_issue()
        text = markdown(self.inventory, self.review['slides'], [], self.review['deck_assessment'])
        self.assertNotIn('整套表达已足够有效，无需修改。', text)

    def test_negative_split_sentence_is_not_an_instruction(self):
        issue = issue_fixture()
        issue['direction'] = '不需要将内容拆分为两页；保留当前概览。'
        validate_polish_plan(issue, 1, strict=True)

    def test_v4_optional_does_not_invalidate_good_page(self):
        inventory, review = review_fixture(self.temp.name, 4)
        review['slides'][0]['optional_suggestions'] = [{'recommendation_kind': 'optional',
            'object': 'shape-1', 'evidence': '四条观点已有清楚顺序',
            'direction': '若想更突出顺序，可只为四条加编号。', 'benefit': '略微增加阅读锚点'}]
        validate_review(review, inventory)

    def test_v4_partial_core_review_cannot_be_complete_no_change(self):
        inventory, review = review_fixture(self.temp.name, 4)
        review['slides'][0]['review_completion'] = 'partial'
        with self.assertRaises(ValueError):
            validate_review(review, inventory)

    def test_v4_secondary_unreadable_is_nonblocking(self):
        inventory, review = review_fixture(self.temp.name, 4)
        inventory['slides'][0]['objects'].append({'path': 'shape-2', 'visible_in_slideshow': True, 'picture': {}})
        review['slides'][0]['media_reviews'] = [{'object': 'shape-2', 'status': 'partial',
            'source': 'render_and_original_image', 'confidence': '高', 'recognized': [
                {'kind': 'visual_content', 'content': '核心图形和标签清楚'}],
            'importance': 'secondary', 'unreadable_items': ['来源脚注年份'],
            'blocking_unknowns': [], 'nonblocking_reason': '年份仅是附注，不影响主要结论和布局判断。'}]
        validate_review(review, inventory)

    def v4_issue(self):
        inv, review = review_fixture(self.temp.name, 4)
        issue = issue_fixture()
        issue.update(action='local', recommendation_kind='necessary',
                     obstacle='核心句被一个前景块遮住', benefit='恢复已有关键内容的可见性',
                     revision_plan={'steps': ['将前景块改成无填充'], 'check': '重看整页确认核心句完整可见'})
        review['slides'][0].update(status='有确定问题', quality_status='有确定问题', issues=[issue])
        review['slides'][0]['tip_checks'][8]['result'] = 'problem'
        return inv, review, issue

    def test_v4_local_fix_needs_no_full_page_redesign(self):
        inv, review, issue = self.v4_issue()
        validate_review(review, inv)
        detailed = markdown(inv, review['slides'], [], review['deck_assessment'])
        compact = compact_report(inv, review['slides'], [], review['deck_assessment'], mode='visual_review')
        self.assertIn('将前景块改成无填充', compact)
        self.assertIn('将前景块改成无填充', detailed)
        self.assertNotIn('11 项逐项判断', compact)

    def test_v4_structured_alternative_not_negation_regex_drives_split(self):
        inv, review, issue = self.v4_issue()
        issue['direction'] = '不需要将内容拆分为两页；保留当前概览。'
        validate_review(review, inv)
        issue['alternatives'] = [{'action': 'split', 'activation_condition': '若主体仍难读', 'pages': []}]
        with self.assertRaises(ValueError):
            validate_review(review, inv)

    def test_v4_disposition_cannot_substitute_an_unrelated_source(self):
        inv, review, issue = self.v4_issue()
        inv['slides'][0]['objects'].append({'path': 'shape-2', 'text_frame': {'text': 'OTHER CONTENT'}})
        issue.update(action='revise', source_refs=[{'source': 'pptx_object', 'slide': 1,
            'object': 'shape-1', 'quote': 'CORE CONTENT'}], content_disposition=[{
            'source_ref': {'source': 'pptx_object', 'slide': 1, 'object': 'shape-2', 'quote': 'OTHER CONTENT'},
            'disposition': 'retain', 'destination': 'current-slide', 'reason': '保留材料'}])
        issue['revision_plan'].update(goal='重组主次', layout='主区域接说明', typography='主体大且协调')
        with self.assertRaises(ValueError):
            validate_review(review, inv)
        issue['content_disposition'][0]['source_ref'] = issue['source_refs'][0]
        validate_review(review, inv)

    def test_v4_numeric_inputs_need_actual_or_declared_assumption_sources(self):
        inv, review, issue = self.v4_issue()
        issue['numeric_checks'] = [{'operation': 'geometric_sum', 'inputs': {
            'coefficient': 1, 'ratio': 0.5, 'start': 0, 'end': 1}, 'reported': 1.5,
            'basis': '仅代数反例，不假称满足作者全部研究前提。'}]
        with self.assertRaises(ValueError):
            validate_review(review, inv)
        check = issue['numeric_checks'][0]
        check['source_refs'] = [{'input': key, 'source': 'test_assumption', 'value': value,
                                'basis': '用于局部有限求和特例，不是实验数据。'}
                               for key, value in check['inputs'].items()]
        validate_review(review, inv)
        check['source_refs'][1]['value'] = 0.6
        with self.assertRaises(ValueError):
            validate_review(review, inv)

    def test_v4_partial_state_can_report_local_evidence_without_certifying_whole(self):
        inv, review, issue = self.v4_issue()
        review['slides'][0].update(review_completion='partial', status='未完成', quality_status='未完成',
                                  blocking_unknowns=['主要图像尚未确认'])
        validate_review(review, inv)

    def test_v4_snapshot_cannot_be_borrowed(self):
        inv, review = review_fixture(self.temp.name, 4)
        review['snapshot_id'] = 'another-snapshot'
        with self.assertRaises(ValueError):
            validate_review(review, inv)

    def test_numeric_source_cannot_borrow_a_digit_from_another_value(self):
        inv, review, issue = self.v4_issue()
        inv['slides'][0]['objects'][0]['table'] = {'rows': [[{'text': '67'}]]}
        issue['target_evidence'] = {'source': 'pptx_object', 'selector': {'kind': 'cell', 'row': 1, 'column': 1}, 'quote': '67'}
        check = {'operation': 'geometric_sum', 'inputs': {'coefficient': 6, 'ratio': 0.5, 'start': 0, 'end': 1},
                 'reported': 9, 'basis': '错误来源要被拒绝，即使算术成立。', 'source_refs': [
                     {'input': 'coefficient', 'source': 'pptx_object', 'slide': 1, 'object': 'shape-1',
                      'selector': {'kind': 'cell', 'row': 1, 'column': 1}, 'quote': '6', 'value': 6},
                     *[{'input': k, 'source': 'test_assumption', 'basis': '有限求和特例', 'value': v}
                       for k, v in [('ratio', 0.5), ('start', 0), ('end', 1)]]]}
        issue['numeric_checks'] = [check]
        with self.assertRaises(ValueError):
            validate_review(review, inv)


if __name__ == '__main__':
    unittest.main()
