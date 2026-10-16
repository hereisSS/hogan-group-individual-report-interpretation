#!/usr/bin/env python3
"""Flag some prohibited vague phrases in a draft, not a complete semantic reviewer."""
from __future__ import annotations
import argparse
import json
from pathlib import Path

PHRASES={
 '把想法接上闭环':'写清谁负责、何时交、是否跟进到完成。',
 '沉着不等于反馈充分':'写明困难是否及时说明、合作方能否调整安排。',
 '把想法接上验证':'写明要验证什么、怎样判断、谁持续跟进。',
 '前端发散':'写明是否不断增加研究方向。',
 '后端收敛':'写明如何取舍任务、确定交付范围。',
 '赋能结果转化':'写明具体动作和工作后果。',
 '拉通协同':'写明各方分工、信息或排期哪里需要协作。',
 '传统低所以不守规则':'价值偏好不是违规证据。',
 '嘴快':'改成具体的发言、需求确认或承诺行为，不作随意评价。',
 '抢戏':'说明是否占据表达时间及是否影响倾听，不给人贴标签。',
 '死磕':'说明是否持续修改细节，以及对排期有什么影响。',
 '光说不干':'用可核对的跟进行为和来源，不作笼统定性。',
 '乖巧听话':'说明实际的规则/权威偏好，不使用幼态化描述。',
 '创新能力强':'核对是否有能力证据；仅HPI不能得出此结论。',
}

def inspect(text: str)->list[dict]:
    findings=[]
    for num,line in enumerate(text.splitlines(),1):
        for phrase,tip in PHRASES.items():
            if phrase in line:findings.append({'line':num,'phrase':phrase,'rewrite_tip':tip})
    return findings

def main()->int:
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('draft',type=Path);p.add_argument('--out',type=Path)
    a=p.parse_args()
    findings=inspect(a.draft.read_text(encoding='utf-8'))
    if a.out:
        if a.out.resolve()==a.draft.resolve():raise SystemExit('不能覆盖文稿')
        a.out.parent.mkdir(parents=True,exist_ok=True)
        a.out.write_text(json.dumps({'findings':findings,'limitation':'仅关键词检查；工作影响、证据与语义仍须人工审稿。'},ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(f'发现{len(findings)}处需检查的表达。无命中不等于文稿合格。')
    return 1 if findings else 0
if __name__=='__main__':raise SystemExit(main())
