"""Synthetic tests only: no data from actual assessment participants."""
from __future__ import annotations
import base64
import copy
import json
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from validate_scores import DIMENSIONS, ValidationError, load_json, validate_and_analyze
from review_gate import approve, check_manifest, create_manifest
from inspect_pptx import inventory

PNG=base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=')

def dataset() -> dict:
    people=[]
    for i,value in enumerate((10,90),1):
        alias=f'TEST{i}'
        people.append({'id':alias,'alias':alias,
                       'scores':{s:{d:value for d in dims} for s,dims in DIMENSIONS.items()},
                       'sources':{s:{'file':f'{s}-{alias}.pdf','overview_page':4} for s in DIMENSIONS},
                       'crosscheck':{s:{'status':'matched','note':'Synthetic fixture, not a real PDF verification.'} for s in DIMENSIONS}})
    return {'schema_version':'1.0',
            'batch':{'label':'Synthetic fixture','year':'2099','norm':'TEST-NORM',
                     'score_type':'percentile','comparison_scope':'current_batch',
                     'mean_includes_self':True,'expected_count':2},'participants':people}

class ScoreTests(unittest.TestCase):
    def setUp(self): self.data=dataset()
    def test_means_include_both_people_and_preserve_ties(self):
        result=validate_and_analyze(self.data)
        self.assertEqual(result['count'],2)
        self.assertEqual(result['score_count'],56)
        self.assertEqual(result['means_display']['HPI']['调适'],'50.0')
        self.assertEqual(result['participants'][0]['scales']['HPI']['maximum_absolute_difference']['ties'],DIMENSIONS['HPI'])
    def test_decimal_mean_and_exact_difference(self):
        self.data['participants'][0]['scores']['HPI']['调适']=11
        result=validate_and_analyze(self.data)
        self.assertEqual(result['means_display']['HPI']['调适'],'50.5')
        self.assertEqual(result['participants'][0]['scales']['HPI']['maximum_absolute_difference']['dimension'],'抱负')
    def test_invalid_numbers(self):
        for value in (-1,101,None,True,50.0,'50',float('nan')):
            with self.subTest(value=value):
                d=copy.deepcopy(self.data);d['participants'][0]['scores']['HPI']['调适']=value
                with self.assertRaises(ValidationError):validate_and_analyze(d)
    def test_missing_dimension(self):
        del self.data['participants'][0]['scores']['HPI']['调适']
        with self.assertRaises(ValidationError):validate_and_analyze(self.data)
    def test_wrong_population_size(self):
        self.data['batch']['expected_count']=10
        with self.assertRaises(ValidationError):validate_and_analyze(self.data)
    def test_duplicate_alias_case_insensitive(self):
        self.data['participants'][1]['alias']='test1'
        with self.assertRaises(ValidationError):validate_and_analyze(self.data)
    def test_duplicate_id(self):
        self.data['participants'][1]['id']='TEST1'
        with self.assertRaises(ValidationError):validate_and_analyze(self.data)
    def test_norm_mismatch(self):
        self.data['participants'][1]['norm']='OTHER'
        with self.assertRaises(ValidationError):validate_and_analyze(self.data)
    def test_wrong_comparison_scope(self):
        self.data['batch']['comparison_scope']='historical'
        with self.assertRaises(ValidationError):validate_and_analyze(self.data)
    def test_unresolved_source_blocks(self):
        self.data['participants'][0]['crosscheck']['HPI']['status']='unresolved'
        with self.assertRaises(ValidationError):validate_and_analyze(self.data)
    def test_single_source_emits_warning(self):
        self.data['participants'][0]['crosscheck']['HPI']['status']='single_source_reviewed'
        self.assertEqual(len(validate_and_analyze(self.data)['warnings']),1)
    def test_invalid_source_page(self):
        self.data['participants'][0]['sources']['HDS']['overview_page']=0
        with self.assertRaises(ValidationError):validate_and_analyze(self.data)
    def test_duplicate_json_keys_are_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)/'test.json';p.write_text('{"a":1,"a":2}')
            with self.assertRaises(ValidationError):load_json(p)
    def test_unsafe_alias_is_rejected(self):
        self.data['participants'][0]['alias']='../TEST'
        with self.assertRaises(ValidationError):validate_and_analyze(self.data)

class ReviewTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.root=Path(self.tmp.name)
        self.data=self.root/'scores.json';self.data.write_text(json.dumps(dataset()),encoding='utf-8')
        self.template=self.root/'template.pptx';self.template.write_bytes(b'synthetic hash fixture, not a renderable pptx')
        self.charts=self.root/'charts';self.charts.mkdir()
        for person in dataset()['participants']:
            for scale in DIMENSIONS:(self.charts/f'{person["alias"]}_{scale}.png').write_bytes(PNG)
        self.manifest=create_manifest(self.data,self.template,self.charts)
    def test_unapproved_blocks(self):
        with self.assertRaises(ValidationError):check_manifest(self.manifest)
    def test_approval_and_check(self):
        m=approve(self.manifest,'Synthetic user confirmation for unit test')
        self.assertEqual(check_manifest(m)['verified_assets'],8)
    def test_empty_confirmation_blocks(self):
        with self.assertRaises(ValidationError):approve(self.manifest,'  ')
    def test_changed_png_invalidates_approval(self):
        m=approve(self.manifest,'Synthetic confirmation')
        p=self.charts/'TEST1_HPI.png';p.write_bytes(PNG+b'changed')
        with self.assertRaises(ValidationError):check_manifest(m)
    def test_changed_data_invalidates_approval(self):
        m=approve(self.manifest,'Synthetic confirmation')
        self.data.write_text(self.data.read_text()+'\n')
        with self.assertRaises(ValidationError):check_manifest(m)
    def test_changed_template_invalidates_approval(self):
        m=approve(self.manifest,'Synthetic confirmation');self.template.write_bytes(b'new')
        with self.assertRaises(ValidationError):check_manifest(m)
    def test_missing_chart_blocks(self):
        (self.charts/'TEST2_MVPI.png').unlink()
        with self.assertRaises(ValidationError):create_manifest(self.data,self.template,self.charts)
    def test_duplicate_chart_blocks(self):
        sub=self.charts/'duplicate';sub.mkdir();(sub/'TEST1_HPI.png').write_bytes(PNG)
        with self.assertRaises(ValidationError):create_manifest(self.data,self.template,self.charts)
    def test_fake_png_blocks(self):
        (self.charts/'TEST1_HPI.png').write_bytes(b'not png')
        with self.assertRaises(ValidationError):create_manifest(self.data,self.template,self.charts)

class PptxTests(unittest.TestCase):
    def test_duplicate_shape_and_overflow_detection(self):
        with tempfile.TemporaryDirectory() as temp:
            file=Path(temp)/'template.pptx'
            pres='''<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>'''
            rel='''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="slides/slide1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/></Relationships>'''
            shapes=''
            for i,x in ((1,0),(2,9000000)):
                shapes+=f'''<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="duplicate"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>test</a:t></a:r></a:p></p:txBody></p:sp>'''
            slide=f'''<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>{shapes}</p:spTree></p:cSld></p:sld>'''
            with zipfile.ZipFile(file,'w') as z:
                z.writestr('ppt/presentation.xml',pres)
                z.writestr('ppt/_rels/presentation.xml.rels',rel)
                z.writestr('ppt/slides/slide1.xml',slide)
            result=inventory(file)
            self.assertEqual(result['slide_count'],1)
            self.assertEqual(result['slides'][0]['duplicate_names'],['duplicate'])
            self.assertEqual(result['slides'][0]['out_of_bounds_ids'],['2'])
if __name__=='__main__':unittest.main()
