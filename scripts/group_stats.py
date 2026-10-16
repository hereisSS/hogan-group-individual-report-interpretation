#!/usr/bin/env python3
"""Describe one verified cohort; no PDF extraction or automatic psychological interpretation.
Python 3.10+, standard library only. Use an authorized run directory OUTSIDE the repo.
"""
from __future__ import annotations
import argparse
from fractions import Fraction
import json
from pathlib import Path
import sys
from validate_scores import DIMENSIONS, ValidationError, display_one, load_json, validate_and_analyze


def _median(values: list[int]) -> Fraction:
    vals=sorted(values); n=len(vals)
    return Fraction(vals[n//2]) if n%2 else Fraction(vals[n//2-1]+vals[n//2],2)


def _check_rules(rules: dict | None) -> None:
    if rules is None: return
    if rules.get('confirmed') is not True or not str(rules.get('source','')).strip():
        raise ValidationError('分档须确认并写明来源；不能用示例文件代替确认')
    for scale in DIMENSIONS:
        bands=rules.get('bands',{}).get(scale)
        if not isinstance(bands,list) or not bands:
            raise ValidationError(f'{scale}缺少分档')
        covered=[]; labels=[]
        for b in bands:
            lo,hi,label=b.get('min'),b.get('max'),b.get('label')
            if type(lo) is not int or type(hi) is not int or not (0<=lo<=hi<=100) or not isinstance(label,str) or not label.strip():
                raise ValidationError('分档须有整数上下限及非空标签')
            covered.extend(range(lo,hi+1));labels.append(label)
        if sorted(covered)!=list(range(101)) or len(set(labels))!=len(labels):
            raise ValidationError(f'{scale}分档须无遗漏、无重叠且标签唯一')


def summarize(data: dict, rules: dict | None=None, include_roles: bool=False, role_source: str='') -> dict:
    audited=validate_and_analyze(data)
    _check_rules(rules)
    if include_roles and not role_source.strip():
        raise ValidationError('使用项目角色规则需明确其经认可的来源')
    n=audited['count']; people=data['participants']
    result={'schema_version':'2.0','batch':data['batch'],'count':n,'dimensions':{},
            'warnings':audited['warnings'],
            'limitations':'仅输入记录的描述统计；不是PDF真实性、心理解释、能力或绩效的验证。'}
    for scale,dims in DIMENSIONS.items():
        result['dimensions'][scale]={}
        for dim in dims:
            vals=[p['scores'][scale][dim] for p in people]
            mean=Fraction(sum(vals),n);median=_median(vals)
            row={'mean':float(mean),'mean_display':display_one(mean),'median':float(median),
                 'median_display':display_one(median),'min':min(vals),'max':max(vals),'valid_n':n}
            if rules:
                row['bands']=[{'label':b['label'],'min':b['min'],'max':b['max'],
                               'count':sum(b['min']<=v<=b['max'] for v in vals),
                               'percent':float(Fraction(100*sum(b['min']<=v<=b['max'] for v in vals),n))}
                              for b in rules['bands'][scale]]
            result['dimensions'][scale][dim]=row
    if rules:result['display_rule_source']=rules['source']
    if include_roles:
        conditions={
            '结果':lambda s:s['抱负']>=65,
            '过程':lambda s:s['审慎']>=65,
            '创新':lambda s:s['好奇']>=50 and s['抱负']>=35,
            '务实':lambda s:s['人际敏感度']<65 and s['好奇']<50,
            '关系':lambda s:s['人际敏感度']>=50 and s['社交']>=35,
        }
        result['project_roles']={name:{'count':sum(fn(p['scores']['HPI']) for p in people),
                  'percent':float(Fraction(100*sum(fn(p['scores']['HPI']) for p in people),n))}
                  for name,fn in conditions.items()}
        result['project_role_source']=role_source
        result['project_role_note']='角色可重叠；项目规则，不等同实际能力、绩效或Hogan官方角色判定。'
    return result


def main()->int:
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('data',type=Path);p.add_argument('--out',type=Path,required=True)
    p.add_argument('--display-rules',type=Path)
    p.add_argument('--include-project-roles',action='store_true')
    p.add_argument('--role-rule-source',default='')
    a=p.parse_args()
    try:
        if a.out.resolve()==a.data.resolve():raise ValidationError('不能覆盖输入')
        r=summarize(load_json(a.data),load_json(a.display_rules) if a.display_rules else None,
                    a.include_project_roles,a.role_rule_source)
        a.out.parent.mkdir(parents=True,exist_ok=True)
        a.out.write_text(json.dumps(r,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        print(f'描述统计已写入指定文件；样本{r["count"]}人。未在日志打印个人分数。')
        return 0
    except (OSError,ValueError,KeyError,TypeError) as exc:
        print('群体统计未通过：请在获授权环境检查输入、来源和分档参数。',file=sys.stderr)
        return 2
if __name__=='__main__':raise SystemExit(main())
