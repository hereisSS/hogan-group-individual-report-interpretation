#!/usr/bin/env python3
"""Record human radar approval and detect asset changes.
This is a workflow guard, not an authorization or tamper-proof security system.
The caller must obtain a real user confirmation before invoking approve.
"""
from __future__ import annotations
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import sys
from validate_scores import DIMENSIONS, ValidationError, load_json, validate_and_analyze

PNG_SIGNATURE=b'\x89PNG\r\n\x1a\n'

def timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()

def digest(path: Path) -> str:
    h=hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda:f.read(1024*1024),b''):h.update(block)
    return h.hexdigest()

def asset(path: Path, kind: str, **extra) -> dict:
    if path.is_symlink() or not path.is_file():
        raise ValidationError(f'素材不存在或是链接：{path}')
    return {'kind':kind,'path':str(path.resolve()),'sha256':digest(path),**extra}

def create_manifest(data_path: Path, template: Path, charts: Path) -> dict:
    data=load_json(data_path)
    result=validate_and_analyze(data)
    if template.suffix.lower()!='.pptx':raise ValidationError('模板必须为PPTX')
    if charts.is_symlink() or not charts.is_dir():raise ValidationError('图表路径必须是实际目录')
    records=[asset(data_path,'scores'),asset(template,'template')]
    root=charts.resolve()
    for person in data['participants']:
        for scale in DIMENSIONS:
            filename=f'{person["alias"]}_{scale}.png'
            hits=list(charts.rglob(filename))
            if len(hits)!=1:raise ValidationError(f'{filename}需要唯一文件，实际找到{len(hits)}个')
            path=hits[0]
            if path.is_symlink() or not path.resolve().is_relative_to(root):
                raise ValidationError(f'图表路径超出授权目录：{path}')
            with path.open('rb') as f:
                if f.read(8)!=PNG_SIGNATURE:raise ValidationError(f'不是有效PNG文件头：{path}')
            records.append(asset(path,'radar',alias=person['alias'],scale=scale))
    return {'schema_version':'1.0','status':'awaiting_radar_confirmation',
            'created_at':timestamp(),'batch_label':data['batch']['label'],
            'participant_count':result['count'],'chart_count':3*result['count'],
            'warnings':result['warnings'],'assets':records,'approval':None}

def check_manifest(manifest: dict, require_approved: bool = True) -> dict:
    if manifest.get('schema_version')!='1.0':raise ValidationError('清单版本无效')
    if require_approved:
        if manifest.get('status')!='radar_approved':raise ValidationError('尚未记录用户对当前雷达图的确认')
        a=manifest.get('approval')
        if not isinstance(a,dict) or not str(a.get('confirmation','')).strip():
            raise ValidationError('缺少真实确认原话记录')
    records=manifest.get('assets')
    if not isinstance(records,list) or not records:raise ValidationError('素材清单为空')
    count=manifest.get('participant_count')
    if type(count) is not int or count<1:raise ValidationError('清单人数无效')
    radar=[r for r in records if r.get('kind')=='radar']
    if len(radar)!=3*count or manifest.get('chart_count')!=3*count:
        raise ValidationError('图像数量与人数不一致')
    if sum(r.get('kind')=='scores' for r in records)!=1 or sum(r.get('kind')=='template' for r in records)!=1:
        raise ValidationError('数据及模板记录必须各1份')
    pairs=[(r.get('alias'),r.get('scale')) for r in radar]
    if len(set(pairs))!=len(pairs):raise ValidationError('存在重复人/量表图')
    for record in records:
        path=Path(record['path'])
        if path.is_symlink() or not path.is_file():raise ValidationError(f'素材缺失或变为链接：{path}')
        if digest(path)!=record.get('sha256'):raise ValidationError(f'素材已经改变，旧确认失效：{path.name}')
    score_path=Path(next(r['path'] for r in records if r['kind']=='scores'))
    dataset=load_json(score_path)
    analysis=validate_and_analyze(dataset)
    expected={(p['alias'],s) for p in dataset['participants'] for s in DIMENSIONS}
    if set(pairs)!=expected or analysis['count']!=count:
        raise ValidationError('清单与当前人员/量表不一致')
    return {'ok':True,'status':manifest['status'],'participant_count':count,
            'verified_assets':len(records),'warnings':analysis['warnings']}

def approve(manifest: dict, confirmation: str) -> dict:
    if not confirmation.strip():raise ValidationError('确认原话不能为空')
    if manifest.get('status')!='awaiting_radar_confirmation':
        raise ValidationError('当前清单不在等待确认状态；不要覆盖既有确认')
    check_manifest(manifest,require_approved=False)
    manifest['status']='radar_approved'
    manifest['approval']={'recorded_at':timestamp(),'confirmation':confirmation,
                          'source':'user_confirmation_recorded_by_executing_agent'}
    return manifest

def write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True,exist_ok=True)
    tmp=path.with_name(path.name+'.writing')
    if tmp.exists():raise ValidationError(f'临时文件已存在，需人工检查：{tmp}')
    tmp.write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    tmp.replace(path)

def main() -> int:
    p=argparse.ArgumentParser(description=__doc__)
    sub=p.add_subparsers(dest='command',required=True)
    init=sub.add_parser('init')
    init.add_argument('--data',type=Path,required=True);init.add_argument('--template',type=Path,required=True)
    init.add_argument('--charts',type=Path,required=True);init.add_argument('--out',type=Path,required=True)
    ap=sub.add_parser('approve');ap.add_argument('--manifest',type=Path,required=True)
    ap.add_argument('--confirmation',required=True)
    ck=sub.add_parser('check');ck.add_argument('--manifest',type=Path,required=True)
    args=p.parse_args()
    try:
        if args.command=='init':
            if args.out.exists():raise ValidationError('清单已存在，不能覆盖；请使用新文件名')
            result=create_manifest(args.data,args.template,args.charts)
            write_json(args.out,result)
            print(json.dumps({'status':result['status'],'chart_count':result['chart_count']},ensure_ascii=False))
        elif args.command=='approve':
            result=approve(load_json(args.manifest),args.confirmation)
            write_json(args.manifest,result)
            print('已记录用户确认；PPT制作前和交付前仍需运行check。')
        else:print(json.dumps(check_manifest(load_json(args.manifest)),ensure_ascii=False,indent=2))
        return 0
    except (OSError,ValidationError,json.JSONDecodeError,KeyError,TypeError) as exc:
        print(f'审核检查未通过：{exc}',file=sys.stderr);return 2
if __name__=='__main__':raise SystemExit(main())
