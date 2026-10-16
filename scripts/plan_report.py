#!/usr/bin/env python3
"""Validate declared report options. Does not establish permission or interpret scores."""
from __future__ import annotations
import argparse
import json
from pathlib import Path
import re
import sys

class PlanError(ValueError):
    pass


def _text(value: object) -> bool:
    return isinstance(value, str) and bool(value.strip())


def plan(request: dict) -> dict:
    if not isinstance(request, dict):
        raise PlanError('任务单必须是JSON对象')
    mode = request.get('task_mode', 'group')
    if mode not in {'personal', 'group', 'both', 'comparison'}:
        raise PlanError('未知任务模式')
    profile = request.get('group_profile', 'management12')
    if profile not in {'management12', 'compact10', 'custom'}:
        raise PlanError('未知群体结构')
    if profile == 'custom':
        outline = request.get('outline')
        if not isinstance(outline, list) or not outline or not all(_text(x) for x in outline):
            raise PlanError('自定义结构须给出非空页面提纲')
    auth = request.get('authorization', {})
    if not isinstance(auth, dict):
        raise PlanError('authorization须是对象')
    names = request.get('display_names', 'random_codes')
    if names not in {'random_codes', 'aliases', 'full_names'}:
        raise PlanError('未知姓名显示模式')
    if type(request.get('show_identity_mapping', False)) is not bool:
        raise PlanError('show_identity_mapping须是布尔值')
    if names == 'full_names' or request.get('show_identity_mapping') is True:
        required = ('authorized_to_provide', 'approved_environment_confirmed', 'full_name_display_allowed')
        if not all(auth.get(k) is True for k in required):
            raise PlanError('实名或身份对照缺少本次授权记录')
        if not _text(request.get('audience')) or request.get('identity_mapping_confirmed') is not True:
            raise PlanError('实名或身份对照需要明确受众和已确认映射')
    charts = request.get('charts', {})
    if not isinstance(charts, dict):
        raise PlanError('charts须是对象')
    median = charts.get('include_median', mode in {'group', 'both'})
    if type(median) is not bool:
        raise PlanError('include_median须是布尔值')
    if request.get('data_availability') == 'means_only' and median:
        raise PlanError('只有均值，不能生成中位值')
    heat = request.get('heatmap', {})
    if not isinstance(heat, dict):
        raise PlanError('heatmap须是对象')
    scheme = heat.get('label_scheme', 'source_labels')
    if scheme not in {'source_labels', 'custom_display'}:
        raise PlanError('未知热力图标签方案')
    heat_ready = heat.get('source_intervals_confirmed') is True and _text(heat.get('source_reference'))
    if scheme == 'custom_display':
        if not heat_ready or heat.get('display_labels_confirmed') is not True or not _text(heat.get('display_reference')):
            raise PlanError('标签映射必须同时确认源区间、显示要求及各自来源')
        bands = heat.get('bands', [])
        if not isinstance(bands, list) or not bands:
            raise PlanError('自定义标签缺少区间映射')
        covered, source_labels, display_labels = [], [], []
        for b in bands:
            if not isinstance(b, dict):
                raise PlanError('每个区间须是对象')
            lo, hi = b.get('min'), b.get('max')
            if type(lo) is not int or type(hi) is not int or not (0 <= lo <= hi <= 100):
                raise PlanError('区间应为0–100内整数')
            if not _text(b.get('source_label')) or not _text(b.get('display_label')):
                raise PlanError('源标签和显示标签须分别填写')
            if not re.fullmatch(r'[0-9A-Fa-f]{6}', str(b.get('color', ''))):
                raise PlanError('颜色须为六位十六进制')
            covered.extend(range(lo, hi+1));source_labels.append(b['source_label']);display_labels.append(b['display_label'])
        if sorted(covered) != list(range(101)) or len(set(source_labels)) != len(bands) or len(set(display_labels)) != len(bands):
            raise PlanError('区间不得遗漏重叠，标签不得混用')
    advice = request.get('include_development_advice', True)
    if type(advice) is not bool:
        raise PlanError('include_development_advice须是布尔值')
    templates = {}
    if mode in {'personal', 'both'}:
        templates['personal'] = 'assets/templates/individual-report-template.pptx'
    if mode in {'group', 'both'}:
        templates['group'] = {'management12': 'assets/templates/group-management-template.pptx',
                              'compact10': 'assets/templates/group-report-template.pptx',
                              'custom': 'user_selected_template'}[profile]
    if mode == 'comparison':
        templates['comparison'] = 'two_pages_means_chart_and_management_interpretation'
    template_choice = request.get('template_choice', 'bundled_blank_templates')
    if template_choice not in {'bundled_blank_templates', 'user_selected'}:
        raise PlanError('未知模板来源；请明确内置或用户指定')
    if template_choice == 'user_selected':
        overrides = request.get('template_overrides')
        if not isinstance(overrides, dict) or not overrides:
            raise PlanError('用户指定模板需要实际可访问的模板引用')
        for key in templates:
            if not _text(overrides.get(key)):
                raise PlanError('用户指定模板缺少当前报告类型的引用')
            templates[key] = overrides[key]
    return {'schema_version': '2.1', 'task_mode': mode, 'group_profile': profile,
            'templates': templates, 'display_names': names, 'show_identity_mapping': request.get('show_identity_mapping', False),
            'include_median': median, 'label_on_point': True,
            'color_encoding': 'cohort' if mode == 'comparison' else 'statistic',
            'heatmap_label_scheme': scheme, 'heatmap_ready': bool(heat_ready),
            'summary_columns': ['可放大的优势', '需管理的风险', '对应的发展建议' if advice else '重点观察'],
            'release_real_data_allowed': False,
            'warnings': ([] if heat_ready or mode == 'personal' else ['热力图区间来源未确认，不应据此定级作图']),
            'limitation': '只检查已填写选项；不能证明真实权限、来源或解释正确；不生成报告或推送文件。'}


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('request', type=Path);p.add_argument('--out', type=Path, required=True)
    a = p.parse_args()
    try:
        if a.out.resolve() == a.request.resolve():
            raise PlanError('不能覆盖输入')
        result = plan(json.loads(a.request.read_text(encoding='utf-8')))
        a.out.parent.mkdir(parents=True, exist_ok=True)
        a.out.write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
        print('报告选项检查完成；真实授权与参考来源仍须核对。')
        return 0
    except (OSError, ValueError, TypeError, KeyError):
        print('报告计划未通过，请在授权目录核对输入、名称授权及标签映射。', file=sys.stderr)
        return 2
if __name__ == '__main__':
    raise SystemExit(main())
