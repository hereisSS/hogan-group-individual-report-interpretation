#!/usr/bin/env python3
"""Find missing fields in interpretation cards. No validity or semantic certification."""
from __future__ import annotations
import argparse
import json
from pathlib import Path
import sys
from style_check import inspect

KINDS = {'core_profile','advantage','risk','group_common','group_difference','motivation','comparison'}
BASIS = {'report_fact','calculation','work_hypothesis','observed_work_record'}


def inspect_cards(document: dict) -> list[dict]:
    findings = []
    if not isinstance(document, dict) or not isinstance(document.get('cards'), list) or not document['cards']:
        return [{'card': None, 'field': 'cards', 'message': '需要非空解读卡列表'}]
    for i, card in enumerate(document['cards']):
        def add(field, message):
            findings.append({'card': i, 'field': field, 'message': message})
        if not isinstance(card, dict):
            add('card','每张卡须为对象');continue
        kind = card.get('kind')
        if kind not in KINDS:
            add('kind','未知解读卡类型')
        # Core descriptions do not need to be turned into risk narratives.
        required = ['title','behavior','basis_type','evidence']
        if kind != 'core_profile':
            required += ['scene','work_impact','observation']
        if kind == 'group_difference':
            required += ['selection_reason','comparison_basis']
        if kind == 'comparison':
            required += ['common_ground','group_a_focus','group_b_focus']
        for field in required:
            value=card.get(field)
            if field=='evidence':
                if not isinstance(value,list) or not value:
                    add(field,'需要至少一个真实来源定位；测试仅用虚构示例')
                else:
                    for entry in value:
                        if not isinstance(entry,dict) or not all(isinstance(entry.get(x),str) and entry[x].strip() for x in ('source','locator')):
                            add(field,'每项依据需要source和locator')
            elif not isinstance(value,str) or not value.strip():
                add(field,'缺少非空字段')
        if card.get('basis_type') not in BASIS:
            add('basis_type','须区分报告事实、计算、工作假设与实际观察')
        if kind == 'core_profile' and card.get('primary_scale') != 'HPI':
            add('primary_scale','本流程核心画像以HPI为主；用户另有明确要求需人工确认')
        if card.get('basis_type') == 'work_hypothesis' and not card.get('uncertainty_disclosed') is True:
            add('uncertainty_disclosed','工作假设须说明不代表已发生事实')
        entries = card.get('evidence') if isinstance(card.get('evidence'), list) else []
        if card.get('basis_type') == 'observed_work_record' and not any(isinstance(e,dict) and e.get('kind')=='work_record' for e in entries):
            add('evidence','实际工作陈述需要工作记录，而不是只有测评分数')
        for field in ('title','behavior','work_impact','support_action'):
            value=card.get(field)
            if isinstance(value,str):
                for hit in inspect(value):
                    add(field,'措辞待审：'+hit['phrase'])
        if card.get('requires_support') is True and not isinstance(card.get('support_action'),str):
            add('support_action','已要求发展建议，需要对应支持动作')
        elif card.get('requires_support') is True and not card['support_action'].strip():
            add('support_action','支持动作不能为空')
    return findings


def main() -> int:
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('cards',type=Path);p.add_argument('--out',type=Path,required=True);a=p.parse_args()
    try:
        if a.out.resolve()==a.cards.resolve():
            raise ValueError('不能覆盖输入')
        findings=inspect_cards(json.loads(a.cards.read_text(encoding='utf-8')))
        a.out.parent.mkdir(parents=True,exist_ok=True)
        a.out.write_text(json.dumps({'findings':findings,'limitation':'仅结构和少量关键词；不能验证来源、因果、专业解释或自然语言质量。'},ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        print(f'解读卡发现{len(findings)}项待审；未打印具体人员内容。')
        return 1 if findings else 0
    except (OSError,ValueError,TypeError,KeyError):
        print('解读卡检查失败，请在授权目录核对文件。',file=sys.stderr);return 2
if __name__=='__main__':
    raise SystemExit(main())
