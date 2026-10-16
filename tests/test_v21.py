"""v2.1 regression tests. Synthetic config/content only; no employee data."""
from __future__ import annotations
import copy
import json
from pathlib import Path
import sys
import unittest
import zipfile
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from plan_report import plan, PlanError
from narrative_check import inspect_cards
from style_check import inspect
from inspect_pptx import inventory


def request():
    return {'task_mode':'group','group_profile':'management12','display_names':'random_codes'}


def mapping():
    return {'label_scheme':'custom_display','source_intervals_confirmed':True,'source_reference':'synthetic source',
            'display_labels_confirmed':True,'display_reference':'synthetic explicit style request',
            'bands':[{'min':lo,'max':hi,'source_label':src,'display_label':dst,'color':col}
                     for lo,hi,src,dst,col in [(0,39,'none','low','56AA86'),(40,69,'low','midlow','D5BE47'),
                         (70,89,'moderate','midhigh','F0A04B'),(90,100,'high','high','EF5B64')]]}


class PlanTests(unittest.TestCase):
    def test_default_management12(self):
        p=plan(request());self.assertIn('group-management-template.pptx',p['templates']['group']);self.assertTrue(p['include_median'])
    def test_personal_does_not_add_median(self):
        d=request();d['task_mode']='personal';self.assertFalse(plan(d)['include_median'])
    def test_user_template_override_kept(self):
        d=request();d.update(template_choice='user_selected',template_overrides={'group':'synthetic-template-reference'})
        self.assertEqual(plan(d)['templates']['group'],'synthetic-template-reference')
    def test_mapping_flag_must_be_bool(self):
        d=request();d['show_identity_mapping']='true'
        with self.assertRaises(PlanError):plan(d)
    def test_compact_10_is_retained(self):
        d=request();d['group_profile']='compact10';self.assertIn('group-report-template',plan(d)['templates']['group'])
    def test_comparison_defaults_means_only(self):
        d=request();d['task_mode']='comparison';p=plan(d);self.assertFalse(p['include_median']);self.assertEqual(p['color_encoding'],'cohort')
    def test_explicit_median_kept(self):
        d=request();d['task_mode']='comparison';d['charts']={'include_median':True};self.assertTrue(plan(d)['include_median'])
    def test_means_only_cannot_invent_median(self):
        d=request();d['data_availability']='means_only'
        with self.assertRaises(PlanError):plan(d)
    def test_unknown_profile_rejected(self):
        d=request();d['group_profile']='invented'
        with self.assertRaises(PlanError):plan(d)
    def test_custom_outline_required(self):
        d=request();d['group_profile']='custom'
        with self.assertRaises(PlanError):plan(d)
        d['outline']=['Synthetic page'];self.assertEqual(plan(d)['group_profile'],'custom')
    def test_full_names_require_authorization(self):
        d=request();d['display_names']='full_names'
        with self.assertRaises(PlanError):plan(d)
    def test_authorized_name_mode_not_overridden(self):
        d=request();d.update(display_names='full_names',audience='synthetic audience',identity_mapping_confirmed=True,
           authorization={'authorized_to_provide':True,'approved_environment_confirmed':True,'full_name_display_allowed':True})
        self.assertEqual(plan(d)['display_names'],'full_names');self.assertFalse(plan(d)['release_real_data_allowed'])
    def test_mapping_appendix_requires_same_guard(self):
        d=request();d['show_identity_mapping']=True
        with self.assertRaises(PlanError):plan(d)
    def test_conclusion_only_no_advice(self):
        d=request();d['include_development_advice']=False;self.assertEqual(plan(d)['summary_columns'][-1],'重点观察')
    def test_unconfirmed_heatmap_warns_not_fake_confirmation(self):
        p=plan(request());self.assertFalse(p['heatmap_ready']);self.assertTrue(p['warnings'])
    def test_custom_labels_require_mapping(self):
        d=request();d['heatmap']={'label_scheme':'custom_display'}
        with self.assertRaises(PlanError):plan(d)
    def test_custom_labels_keep_original_ranges(self):
        d=request();d['heatmap']=mapping();self.assertTrue(plan(d)['heatmap_ready'])
    def test_overlapping_heat_bands_rejected(self):
        d=request();d['heatmap']=mapping();d['heatmap']['bands'][1]['min']=39
        with self.assertRaises(PlanError):plan(d)
    def test_duplicate_display_labels_rejected(self):
        d=request();d['heatmap']=mapping();d['heatmap']['bands'][1]['display_label']='low'
        with self.assertRaises(PlanError):plan(d)

class NarrativeTests(unittest.TestCase):
    def setUp(self):
        self.cards=json.loads((ROOT/'assets/input/analysis-card.example.json').read_text())
    def test_synthetic_cards_pass_structure(self):self.assertEqual(inspect_cards(self.cards),[])
    def test_core_not_forced_into_risk_template(self):
        c=self.cards['cards'][0];self.assertNotIn('work_impact',c);self.assertEqual(inspect_cards({'cards':[c]}),[])
    def test_core_hpi_primary(self):
        self.cards['cards'][0]['primary_scale']='MVPI';self.assertTrue(inspect_cards(self.cards))
    def test_risk_impact_required(self):
        del self.cards['cards'][1]['work_impact'];self.assertTrue(any(x['field']=='work_impact' for x in inspect_cards(self.cards)))
    def test_advantage_needs_work_value_too(self):
        self.cards['cards'][2]['work_impact']='';self.assertTrue(inspect_cards(self.cards))
    def test_inferred_is_not_actual_record(self):
        self.cards['cards'][1]['basis_type']='observed_work_record';self.assertTrue(inspect_cards(self.cards))
    def test_null_evidence_is_safe(self):
        self.cards['cards'][1].update(basis_type='observed_work_record',evidence=None);self.assertTrue(inspect_cards(self.cards))
    def test_hypothesis_needs_disclosure(self):
        self.cards['cards'][1]['uncertainty_disclosed']=False;self.assertTrue(inspect_cards(self.cards))
    def test_difference_requires_selection_reason(self):
        self.cards['cards'][1]['kind']='group_difference';self.assertTrue(any(x['field']=='selection_reason' for x in inspect_cards(self.cards)))
    def test_support_must_be_concrete_field(self):
        self.cards['cards'][1]['support_action']='';self.assertTrue(inspect_cards(self.cards))
    def test_colloquial_label_flag(self):self.assertTrue(inspect('嘴快、抢戏、死磕，光说不干。'))
    def test_professional_terms_not_blanket_banned(self):self.assertEqual(inspect('明确判断依据与优先级，观察关键节点、协作方式和交付完整性。'),[])

class TemplateTests(unittest.TestCase):
    def test_12_pages_geometry_and_names(self):
        inv=inventory(ROOT/'assets/templates/group-management-template.pptx');self.assertEqual(inv['slide_count'],12)
        for s in inv['slides']:
            self.assertEqual(s['out_of_bounds_ids'],[]);self.assertEqual(s['duplicate_names'],[])
    def test_required_hpi_cards_and_summary(self):
        d=json.loads((ROOT/'assets/templates/management-template-map.json').read_text())
        fields=d['pages'][3]['fields'];self.assertTrue(any(x['name']=='m4.diff0.title' for x in fields))
        texts=''.join(x['placeholder'] for x in d['pages'][10]['fields'])
        for label in ['可放大的优势','需管理的风险','对应的发展建议','正向价值','需要防范']:self.assertIn(label,texts)
    def test_no_private_media_or_data_in_template(self):
        with zipfile.ZipFile(ROOT/'assets/templates/group-management-template.pptx') as z:
            self.assertFalse(any(n.startswith(('ppt/media/','ppt/embeddings/','ppt/comments')) for n in z.namelist()))
    def test_all_local_markdown_links_resolve(self):
        import re
        for md in ROOT.rglob('*.md'):
            for target in re.findall(r'\]\(([^)]+)\)',md.read_text()):
                if target.startswith(('http:','https:','#','mailto:')):continue
                self.assertTrue((md.parent/target.split('#')[0]).exists(),f'{md.name}:{target}')
if __name__=='__main__':unittest.main()
