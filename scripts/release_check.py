#!/usr/bin/env python3
"""Limited release audit: structure, tracked run-data indicators and reviewed binary hashes.
Not a DLP product; cannot prove that all identifying information is absent.
"""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET

PRIVATE_DIRS={'private','data','input','inputs','output','outputs','reports','runs','cache','audits','tmp'}
SKIP_DIRS={'.git','__pycache__','.pytest_cache','.venv'}
FORBIDDEN_SUFFIX={'.pdf','.csv','.tsv','.xls','.xlsx','.zip','.pem','.key','.p12','.log'}
RUN_FILES={'scores.json','score_audit.json','group_stats.json','review_manifest.json'}
TEXT_SUFFIX={'.md','.py','.json','.txt','.yaml','.yml'}

def sha(path:Path)->str:
    h=hashlib.sha256()
    with path.open('rb') as f:
        for chunk in iter(lambda:f.read(1024*1024),b''):h.update(chunk)
    return h.hexdigest()

def audit(root:Path)->list[str]:
    root=root.resolve();errors=[]
    try:
        manifest=json.loads((root/'assets/release-allowlist.json').read_text(encoding='utf-8'))
        approved=manifest['reviewed_binaries']
    except (OSError,ValueError,KeyError):return ['缺少或损坏的二进制审查清单。']
    for rel,expected in approved.items():
        path=root/rel
        if not path.is_file() or path.is_symlink() or sha(path)!=expected:
            errors.append(f'二进制文件未通过版本检查：{rel}')
    for path in root.rglob('*'):
        rel=path.relative_to(root);parts=rel.parts
        if any(x in SKIP_DIRS for x in parts):continue
        if path.is_symlink():errors.append(f'禁止链接：{rel}');continue
        if not path.is_file():continue
        if parts[0] in PRIVATE_DIRS or path.suffix.lower() in FORBIDDEN_SUFFIX or path.name in RUN_FILES or path.name.startswith('.env'):
            errors.append(f'疑似运行数据或秘密文件：{rel}');continue
        if path.suffix.lower() in {'.pptx','.png','.jpg','.jpeg','.webp','.gif'} and str(rel) not in approved:
            errors.append(f'未审查二进制：{rel}');continue
        if path.suffix.lower()=='.pptx':
            try:
                with zipfile.ZipFile(path) as z:
                    for name in z.namelist():
                        if name.startswith(('ppt/media/','ppt/embeddings/','ppt/comments','ppt/commentAuthors')):
                            errors.append(f'空白模板含媒体/嵌入/批注：{rel}:{name}')
                        if name.endswith('.rels'):
                            rt=ET.fromstring(z.read(name))
                            if any(el.get('TargetMode')=='External' for el in rt):errors.append(f'模板含外链：{rel}')
                    cp=z.read('docProps/core.xml').decode('utf-8') if 'docProps/core.xml' in z.namelist() else ''
                    if re.search(r'<dc:creator[^>]*>(?!Hogan Workflow Template|</)[^<]+',cp):
                        errors.append(f'模板作者属性需人工确认：{rel}')
            except (OSError,zipfile.BadZipFile,ET.ParseError):errors.append(f'无法读取模板：{rel}')
        if path.suffix.lower() in TEXT_SUFFIX or path.name in {'.gitignore','.gitattributes'}:
            text=path.read_text(encoding='utf-8',errors='replace')
            # Local absolute paths, private keys and access tokens do not belong in distributable docs.
            for label,pat in [('运行时绝对路径',r'/(?:mnt/data|Users/|home/oai)/[^\s\'\"]+'),
                              ('私钥',r'-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----'),
                              ('访问令牌',r'\b(?:ghp|github_pat)_[A-Za-z0-9_]{25,}\b')]:
                if re.search(pat,text):errors.append(f'{label}需人工排查：{rel}')
            if path.suffix.lower()=='.json' and path.name not in {'release-allowlist.json','template-map.json'}:
                try:
                    obj=json.loads(text)
                    if isinstance(obj,dict) and 'participants' in obj and obj.get('data_kind')!='synthetic':
                        errors.append(f'样例含人员记录但未明确标为虚构：{rel}')
                except ValueError:errors.append(f'JSON无效：{rel}')
    # Read-only check of the index. Does not inspect or rewrite Git history.
    if (root/'.git').exists():
        try:
            r=subprocess.run(['git','ls-files','-z'],cwd=root,capture_output=True,check=True)
            for raw in r.stdout.decode().split('\0'):
                if not raw:continue
                p=Path(raw)
                if p.parts[0] in PRIVATE_DIRS or p.suffix.lower() in FORBIDDEN_SUFFIX or p.name in RUN_FILES:
                    errors.append(f'Git仍跟踪运行数据：{raw}')
        except (OSError,subprocess.SubprocessError):errors.append('未能读取Git索引；需人工检查。')
    return sorted(set(errors))

def main()->int:
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('root',type=Path,nargs='?',default=Path('.'));a=p.parse_args()
    findings=audit(a.root)
    for item in findings:print(item)
    print(f'有限发布检查：{len(findings)}项待处理。仍须人工检查内容、图片、受众权限和Git历史。')
    return 2 if findings else 0
if __name__=='__main__':raise SystemExit(main())
