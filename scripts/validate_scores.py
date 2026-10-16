#!/usr/bin/env python3
"""Validate supplied Hogan score records; calculate exact current-cohort differences.
This validates entered records, NOT the truth of the underlying PDF extraction.
Python 3.10+, standard library only.
"""
from __future__ import annotations
import argparse
import json
import re
import sys
from decimal import Decimal, ROUND_HALF_UP
from fractions import Fraction
from pathlib import Path
from typing import Any

DIMENSIONS = {
    'HPI': ['调适', '抱负', '社交', '人际敏感度', '审慎', '好奇', '学习方式'],
    'HDS': ['激动', '多疑', '谨慎', '内敛', '消极', '自大', '狡猾', '戏剧化', '幻想', '苛求', '恭顺'],
    'MVPI': ['认可', '权力', '享乐', '利他', '归属', '传统', '保障', '商业', '美感', '科学'],
}

class ValidationError(ValueError):
    """The supplied dataset cannot be used for current-cohort analysis."""

def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValidationError(message)

def _nonempty(value: Any) -> bool:
    return isinstance(value, str) and bool(value.strip())

def _page(value: Any) -> bool:
    return type(value) is int and value >= 1

def _unique_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValidationError(f'JSON存在重复键：{key}')
        result[key] = value
    return result

def load_json(path: Path) -> dict[str, Any]:
    with path.open(encoding='utf-8-sig') as f:
        value = json.load(f, object_pairs_hook=_unique_object)
    _require(isinstance(value, dict), 'JSON顶层必须是对象')
    return value

def display_one(value: Fraction) -> str:
    number = Decimal(value.numerator) / Decimal(value.denominator)
    return str(number.quantize(Decimal('0.1'), rounding=ROUND_HALF_UP))

def validate_and_analyze(data: dict[str, Any]) -> dict[str, Any]:
    _require(data.get('schema_version') == '1.0', 'schema_version必须为1.0')
    batch = data.get('batch')
    _require(isinstance(batch, dict), '缺少batch对象')
    for field in ('label', 'year', 'norm'):
        _require(_nonempty(batch.get(field)), f'batch.{field}必须是非空字符串')
    _require(batch.get('score_type') == 'percentile', '只能处理原报告百分位，不能混用原始分')
    _require(batch.get('mean_includes_self') is True, '本校验器要求均值包含本人')
    _require(batch.get('comparison_scope') == 'current_batch', '本校验器仅支持当前完整批次')
    people = data.get('participants')
    _require(isinstance(people, list) and len(people) > 0, 'participants必须为非空数组')
    expected = batch.get('expected_count')
    if expected is not None:
        _require(type(expected) is int and expected >= 1, 'expected_count须为正整数')
        _require(expected == len(people), f'实际人数{len(people)}与期望人数{expected}不一致')
    ids: set[str] = set()
    aliases: set[str] = set()
    warnings: list[str] = []
    for index, p in enumerate(people):
        _require(isinstance(p, dict), f'participants[{index}]必须为对象')
        pid, alias = p.get('id'), p.get('alias')
        _require(_nonempty(pid), f'第{index+1}人缺少id')
        _require(pid not in ids, f'重复id：{pid}')
        _require(_nonempty(alias) and re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]{0,63}', alias) is not None,
                 f'人员{pid}缩写无效，需安全且唯一的文件名缩写')
        _require(alias.casefold() not in aliases, f'重复展示缩写：{alias}')
        ids.add(pid)
        aliases.add(alias.casefold())
        _require(p.get('norm', batch['norm']) == batch['norm'], f'{alias}常模与本批不一致')
        scores, sources, checks = p.get('scores'), p.get('sources'), p.get('crosscheck')
        _require(isinstance(scores, dict) and set(scores) == set(DIMENSIONS), f'{alias}量表不完整或多出未知量表')
        _require(isinstance(sources, dict), f'{alias}缺少sources')
        _require(isinstance(checks, dict), f'{alias}缺少crosscheck')
        for scale, dimensions in DIMENSIONS.items():
            values = scores[scale]
            _require(isinstance(values, dict), f'{alias}/{scale}分值须为字典')
            _require(set(values) == set(dimensions), f'{alias}/{scale}维度名或维度数不匹配')
            for dim in dimensions:
                value = values[dim]
                _require(type(value) is int and 0 <= value <= 100,
                         f'{alias}/{scale}/{dim}须为0–100整数，不能为布尔值/空值/字符串')
            source = sources.get(scale)
            _require(isinstance(source, dict), f'{alias}/{scale}缺少来源记录')
            kind = source.get('kind', 'pdf')
            _require(kind in ('pdf', 'table'), f'{alias}/{scale}来源类型须为pdf或table')
            if kind == 'pdf':
                _require(_nonempty(source.get('file')) and source['file'].lower().endswith('.pdf'),
                         f'{alias}/{scale}缺少PDF来源名')
                _require(_page(source.get('overview_page')), f'{alias}/{scale}总览页码无效')
            else:
                _require(_nonempty(source.get('file')) and Path(source['file']).suffix.lower() in ('.xlsx','.xls','.csv','.tsv','.json'),
                         f'{alias}/{scale}表来源文件名无效')
                _require(_nonempty(source.get('locator')), f'{alias}/{scale}缺少实际表格定位')
            detail = source.get('detail_pages', {})
            _require(isinstance(detail, dict), f'{alias}/{scale}detail_pages须为对象')
            for dim, page in detail.items():
                _require(dim in dimensions and _page(page), f'{alias}/{scale}维度页码映射无效：{dim}')
            check = checks.get(scale)
            _require(isinstance(check, dict), f'{alias}/{scale}缺少核验状态')
            status = check.get('status')
            _require(status in ('matched', 'single_source_reviewed'), f'{alias}/{scale}未完成源分数核验')
            _require(_nonempty(check.get('note')), f'{alias}/{scale}核验说明不能为空')
            if status == 'single_source_reviewed':
                warnings.append(f'{alias}/{scale}仅核验单一来源：{check["note"]}')
    count = len(people)
    means = {scale: {dim: Fraction(sum(p['scores'][scale][dim] for p in people), count)
                     for dim in dims} for scale, dims in DIMENSIONS.items()}
    analyses = []
    for p in people:
        record: dict[str, Any] = {'id': p['id'], 'alias': p['alias'], 'scales': {}}
        for scale, dims in DIMENSIONS.items():
            deltas = {d: Fraction(p['scores'][scale][d]) - means[scale][d] for d in dims}
            maximum = max(abs(v) for v in deltas.values())
            ties = [d for d in dims if abs(deltas[d]) == maximum]
            record['scales'][scale] = {
                'deltas': {d: float(v) for d, v in deltas.items()},
                'maximum_absolute_difference': {
                    'dimension': ties[0], 'ties': ties,
                    'personal_score': p['scores'][scale][ties[0]],
                    'cohort_mean': float(means[scale][ties[0]]),
                    'cohort_mean_display': display_one(means[scale][ties[0]]),
                    'delta': float(deltas[ties[0]]),
                    'exact_delta': str(deltas[ties[0]]),
                },
            }
        analyses.append(record)
    return {
        'schema_version': '1.0', 'batch': batch, 'count': count, 'score_count': 28 * count,
        'means': {s: {d: float(v) for d, v in vals.items()} for s, vals in means.items()},
        'means_display': {s: {d: display_one(v) for d, v in vals.items()} for s, vals in means.items()},
        'participants': analyses, 'warnings': warnings,
        'limitations': '仅校验已输入数据及算术；不验证PDF真实性、语义解释或核验声明是否真实。',
    }

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('data', type=Path)
    parser.add_argument('--out', type=Path)
    args = parser.parse_args()
    try:
        if args.out and args.data.resolve() == args.out.resolve():
            raise ValidationError('输出文件不得覆盖输入数据')
        result = validate_and_analyze(load_json(args.data))
        text = json.dumps(result, ensure_ascii=False, indent=2)
        if args.out:
            args.out.parent.mkdir(parents=True, exist_ok=True)
            args.out.write_text(text + '\n', encoding='utf-8')
        else:
            print(json.dumps({'count': result['count'], 'score_count': result['score_count'], 'warning_count': len(result['warnings']), 'note': '详细结果请用--out写入获授权目录。'}, ensure_ascii=False))
        return 0
    except (ValidationError, OSError, json.JSONDecodeError) as exc:
        print(f'校验未通过：{exc}', file=sys.stderr)
        return 2

if __name__ == '__main__':
    raise SystemExit(main())
