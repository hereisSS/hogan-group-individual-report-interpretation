"""Synthetic v2 tests; no actual assessment data or external network access."""
from __future__ import annotations
import copy
import importlib.util
import json
from pathlib import Path
import re
import sys
import tempfile
import unittest
import zipfile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
sys.path.insert(0,str(Path(__file__).resolve().parent))
from test_workflow import dataset
from validate_scores import DIMENSIONS, ValidationError, validate_and_analyze
from group_stats import summarize
from style_check import inspect
from inspect_pptx import inventory
from release_check import audit
ROOT=Path(__file__).resolve().parents[1]

class GroupTests(unittest.TestCase):
    def test_even_median(self):
        r=summarize(dataset())
        self.assertEqual(r['dimensions']['HPI']['调适']['median'],50)
        self.assertEqual(r['dimensions']['HDS']['自大']['valid_n'],2)
        self.assertNotIn('bands',r['dimensions']['HDS']['自大'])
        self.assertNotIn('project_roles',r)
    def test_odd_median(self):
        d=dataset();p=copy.deepcopy(d['participants'][0]);p['id']=p['alias']='SYNTHETIC3'
        p['scores']={s:{dim:51 for dim in ds} for s,ds in DIMENSIONS.items()}
        d['participants'].append(p);d['batch']['expected_count']=3
        self.assertEqual(summarize(d)['dimensions']['HPI']['调适']['median'],51)
    def test_unconfirmed_rules_block(self):
        rules=json.loads((ROOT/'assets/input/display-rules.example.json').read_text())
        with self.assertRaises(ValidationError):summarize(dataset(),rules)
    def test_confirmed_rules_counts(self):
        rules=json.loads((ROOT/'assets/input/display-rules.example.json').read_text());rules['confirmed']=True;rules['source']='synthetic test only'
        r=summarize(dataset(),rules)
        self.assertEqual([x['count'] for x in r['dimensions']['HDS']['自大']['bands']],[1,0,0,1])
    def test_overlapping_rules_block(self):
        rules=json.loads((ROOT/'assets/input/display-rules.example.json').read_text());rules['confirmed']=True
        rules['bands']['HPI'][1]['min']=25
        with self.assertRaises(ValidationError):summarize(dataset(),rules)
    def test_roles_require_source(self):
        with self.assertRaises(ValidationError):summarize(dataset(),include_roles=True)
    def test_roles_overlap(self):
        r=summarize(dataset(),include_roles=True,role_source='synthetic fixture rule')
        self.assertEqual(r['project_roles']['结果']['count'],1)
        self.assertEqual(r['project_roles']['过程']['count'],1)
        self.assertEqual(r['project_roles']['创新']['percent'],50)
    def test_table_source_supported_without_fake_pdf(self):
        d=dataset()
        for p in d['participants']:
            p['sources']={s:{'kind':'table','file':'synthetic.xlsx','locator':'Sheet1:synthetic test row'} for s in DIMENSIONS}
        self.assertEqual(validate_and_analyze(d)['count'],2)
    def test_table_requires_actual_locator(self):
        d=dataset();d['participants'][0]['sources']['HPI']={'kind':'table','file':'synthetic.csv'}
        with self.assertRaises(ValidationError):validate_and_analyze(d)

class AssetTests(unittest.TestCase):
    def test_skill_name_and_description(self):
        t=(ROOT/'SKILL.md').read_text();self.assertTrue(t.startswith('---\n'))
        name=re.search(r'^name:\s*(.+)$',t,re.M).group(1)
        self.assertEqual(name,ROOT.name);self.assertRegex(name,r'^[a-z0-9]+(?:-[a-z0-9]+)*$')
        desc=re.search(r'^description:\s*(.+)$',t,re.M).group(1);self.assertLessEqual(len(desc),1024)
    def test_template_counts_and_geometry(self):
        for file,n in [('individual-report-template.pptx',1),('group-report-template.pptx',10)]:
            r=inventory(ROOT/'assets/templates'/file);self.assertEqual(r['slide_count'],n)
            for page in r['slides']:
                self.assertEqual(page['duplicate_names'],[]);self.assertEqual(page['out_of_bounds_ids'],[])
    def test_templates_contain_no_original_media_or_embeddings(self):
        for file in (ROOT/'assets/templates').glob('*.pptx'):
            with zipfile.ZipFile(file) as z:
                self.assertFalse(any(n.startswith(('ppt/media/','ppt/embeddings/','ppt/comments')) for n in z.namelist()))
    def test_all_input_people_are_marked_synthetic(self):
        d=json.loads((ROOT/'assets/input/scores.synthetic.json').read_text());self.assertEqual(d['data_kind'],'synthetic')
        self.assertEqual(validate_and_analyze(d)['count'],3)
    def test_style_blocks_known_vague_phrases(self):
        self.assertEqual(len(inspect('把想法接上闭环\n沉着不等于反馈充分')),2)
    def test_concrete_statement_no_keyword_flag(self):
        self.assertEqual(inspect('提出方案后，明确谁负责、何时交，并持续跟进到完成。'),[])
    def test_release_bundle_checks(self):
        self.assertEqual(audit(ROOT),[])
    def test_run_data_detected_when_copied(self):
        # Isolated minimal audit fixture, not real cohort data.
        with tempfile.TemporaryDirectory() as tmp:
            r=Path(tmp);(r/'assets').mkdir();(r/'assets/release-allowlist.json').write_text('{"reviewed_binaries":{}}')
            (r/'scores.json').write_text('{}')
            self.assertTrue(any('scores.json' in x for x in audit(r)))

if __name__=='__main__':unittest.main()
