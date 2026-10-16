#!/usr/bin/env python3
"""Read PowerPoint OOXML shape inventory without changing the source file.
Coordinates use inches. Nested group coordinates are local, not global.
"""
from __future__ import annotations
import argparse
from collections import Counter
import json
from pathlib import Path
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
EMU = 914400

def inventory(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        pres = ET.fromstring(z.read('ppt/presentation.xml'))
        size = pres.find('p:sldSz', NS)
        if size is None:
            raise ValueError('PPT缺少页面尺寸')
        width, height = int(size.attrib['cx']), int(size.attrib['cy'])
        rels = ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        relation = {r.attrib['Id']: posixpath.normpath(posixpath.join('ppt', r.attrib['Target']))
                    for r in rels.findall(f'{REL}Relationship')
                    if r.attrib.get('TargetMode') != 'External'}
        slides = []
        for page, sid in enumerate(pres.findall('p:sldIdLst/p:sldId', NS), 1):
            rid = sid.attrib.get(f'{{{NS["r"]}}}id')
            part = relation[rid]
            if part.startswith('/'):
                part = part.lstrip('/')
            root = ET.fromstring(z.read(part))
            tree = root.find('p:cSld/p:spTree', NS)
            if tree is None:
                raise ValueError(f'第{page}页缺少形状树')
            records = []
            def walk(parent: ET.Element, depth: int = 0) -> None:
                for shape in parent:
                    kind = shape.tag.split('}')[-1]
                    if kind not in ('sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'):
                        continue
                    props = shape.find('.//p:cNvPr', NS)
                    xform = shape.find('p:spPr/a:xfrm', NS)
                    if xform is None:
                        xform = shape.find('p:xfrm', NS)
                    if xform is None:
                        xform = shape.find('p:grpSpPr/a:xfrm', NS)
                    bounds = None
                    outside = None
                    if xform is not None:
                        off, ext = xform.find('a:off', NS), xform.find('a:ext', NS)
                        if off is not None and ext is not None:
                            x,y = int(off.attrib['x']),int(off.attrib['y'])
                            w,h = int(ext.attrib['cx']),int(ext.attrib['cy'])
                            bounds = dict(x=round(x/EMU,5),y=round(y/EMU,5),
                                          width=round(w/EMU,5),height=round(h/EMU,5))
                            outside = bool(x < 0 or y < 0 or x+w > width+EMU/1000 or y+h > height+EMU/1000) if depth == 0 else None
                    text = '' if kind == 'grpSp' else '\n'.join(
                        ''.join(node.itertext()) for node in shape.findall('.//a:t', NS))
                    blips = [b.attrib.get(f'{{{NS["r"]}}}embed') for b in shape.findall('.//a:blip',NS)] if kind != 'grpSp' else []
                    records.append({'id': props.attrib.get('id') if props is not None else None,
                                    'name': props.attrib.get('name') if props is not None else None,
                                    'kind': kind, 'depth': depth, 'bounds_inches': bounds,
                                    'out_of_bounds': outside, 'text': text,
                                    'image_relationship_ids': blips})
                    if kind == 'grpSp':
                        walk(shape, depth+1)
            walk(tree)
            names = Counter(r['name'] for r in records if r['name'])
            slides.append({'page':page,'part':part,'shapes':records,
                           'duplicate_names':[n for n,c in names.items() if c>1],
                           'out_of_bounds_ids':[r['id'] for r in records if r['out_of_bounds']]})
        return {'file':str(path.resolve()), 'slide_width_inches':width/EMU,
                'slide_height_inches':height/EMU,'slide_count':len(slides),'slides':slides,
                'limitations':'几何和结构清点不替代实际渲染；组合图形子坐标为局部坐标。'}

def main() -> int:
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('pptx',type=Path);p.add_argument('--out',type=Path)
    args=p.parse_args()
    try:
        if args.out and args.pptx.resolve()==args.out.resolve():
            raise ValueError('输出不得覆盖原PPT')
        text=json.dumps(inventory(args.pptx),ensure_ascii=False,indent=2)
        if args.out:
            args.out.parent.mkdir(parents=True,exist_ok=True)
            args.out.write_text(text+'\n',encoding='utf-8')
        else:print(text)
        return 0
    except (OSError,ValueError,KeyError,zipfile.BadZipFile,ET.ParseError) as exc:
        print(f'模板清点失败：{exc}',file=sys.stderr);return 2
if __name__=='__main__':raise SystemExit(main())
