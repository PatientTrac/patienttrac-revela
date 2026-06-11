import React, { useState, useEffect } from 'react';
import { supabase } from '../../lib/supabase';
import { Save, CheckCircle, Search, Plus, X } from 'lucide-react';
import { getOpNoteDraft } from '../../lib/revelai';
import type { OpNoteDraftResult } from '../../lib/revelai';

interface PatientContext {
  patient_id: number;
  encounter_id: number;
  patient_name: string;
  patient_sex: 'male' | 'female' | 'non_binary';
  org_id: string;
  provider_id: number;
}

interface OperativeNoteData {
  pre_op_diagnosis_icd10: string[];
  procedures_performed: string[];
  procedures_performed_cpt: string[];
  laterality: 'left' | 'right' | 'bilateral' | 'midline' | 'na';
  anesthesia_type: 'general_endotracheal' | 'general_lma' | 'mac' | 'local';
  anesthesia_provider: string;
  surgeon_primary: string;
  positioning: 'supine' | 'prone' | 'lateral_right' | 'lateral_left' | 'sitting';
  prep_solution: 'betadine' | 'chlorhexidine' | 'alcohol';
  implant_details?: {
    manufacturer: string;
    product_line: string;
    type: 'saline' | 'silicone_cohesive_gel' | 'structured';
    surface: 'smooth' | 'textured';
    shape: 'round' | 'anatomic';
    profile: 'low' | 'moderate' | 'moderate_plus' | 'high';
    size_right_cc: number;
    size_left_cc: number;
    serial_right: string;
    serial_left: string;
  };
  tissue_removed?: {
    specimen_type: string;
    weight_right_grams: number;
    weight_left_grams: number;
    sent_to_pathology: boolean;
  };
  fat_harvested?: {
    donor_sites: string[];
    total_volume_ml: number;
    processing_method: 'centrifuge' | 'filtration' | 'decantation';
    volume_injected_ml: number;
    injection_sites: string[];
  };
  estimated_blood_loss_ml: number;
  crystalloid_given_ml: number;
  drains_placed: boolean;
  drain_details?: Array<{
    location: string;
    type: 'jackson_pratt' | 'blake' | 'penrose';
    size_french: number;
  }>;
  complications: string;
  condition: 'stable' | 'critical';
  disposition: 'pacu' | 'icu' | 'floor' | 'home';
}

interface Props {
  patientContext: PatientContext;
}

export default function BasicOperativeNote({ patientContext }: Props) {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  const [searchingCpt, setSearchingCpt] = useState(false);
  const [searchingIcd, setSearchingIcd] = useState(false);
  
  const [formData, setFormData] = useState<OperativeNoteData>({
    pre_op_diagnosis_icd10: [],
    procedures_performed: [],
    procedures_performed_cpt: [],
    laterality: 'na',
    anesthesia_type: 'general_endotracheal',
    anesthesia_provider: '',
    surgeon_primary: '',
    positioning: 'supine',
    prep_solution: 'betadine',
    estimated_blood_loss_ml: 0,
    crystalloid_given_ml: 0,
    drains_placed: false,
    complications: 'none',
    condition: 'stable',
    disposition: 'pacu'
  });

  const [cptSearch, setCptSearch] = useState('');
  const [cptResults, setCptResults] = useState<any[]>([]);
  const [icdSearch, setIcdSearch] = useState('');
  const [icdResults, setIcdResults] = useState<any[]>([]);
  const [signed, setSigned] = useState(false);
  const [signedAt, setSignedAt] = useState<string | null>(null);
  const [consentObtained, setConsentObtained] = useState(false);

  // Caprini VTE risk factors (each checked = points shown)
  const [caprini, setCaprini] = useState({
    age_41_60: false, age_61_74: false, age_75_plus: false,
    minor_surgery: false, major_surgery_lt45: false, major_surgery_gt45: false,
    prior_dvt_pe: false, family_history_clot: false,
    obesity_bmi_gt25: false, varicose_veins: false,
    immobility_gt72h: false, central_line: false,
    malignancy: false, congestive_heart_failure: false,
    sepsis: false, inflammatory_bowel: false,
    oral_contraceptives: false, hrt: false,
    thrombophilia: false, stroke: false,
  });

  // STOP-BANG OSA screen (each = 1 point)
  const [stopBang, setStopBang] = useState({
    snoring: false,
    tired: false,
    observed: false,
    pressure: false,
    bmi_gt35: false,
    age_gt50: false,
    neck_gt40: false,
    gender_male: false,
  });

  // RCRI — Revised Cardiac Risk Index (each = 1 point; Lee 1999)
  const [rcri, setRcri] = useState({
    high_risk_surgery: false,       // supra-inguinal vascular, intraperitoneal, intrathoracic
    ischemic_heart_disease: false,  // h/o MI, positive stress test, use of nitrates, ECG Q-waves
    congestive_heart_failure: false,// pulmonary edema, PND, bilateral rales, S3, CXR vascular redistribution
    cerebrovascular_disease: false, // h/o stroke or TIA
    insulin_dependent_dm: false,    // pre-op insulin use
    creatinine_gt2: false,          // pre-op creatinine > 2.0 mg/dL (>177 μmol/L)
  });

  // Universal Protocol — Joint Commission (NPSG.01.01.01, UP.01.01.01–UP.01.03.01)
  const [universalProtocol, setUniversalProtocol] = useState({
    patient_identity_verified: false,   // two identifiers confirmed
    procedure_verified: false,          // consent matches planned procedure
    site_marked: false,                 // site/side marked by surgeon
    timeout_performed: false,           // team time-out performed
    sponge_count_correct: false,        // sponge count verified correct
    instrument_count_correct: false,    // instrument count verified correct
    needle_count_correct: false,        // needle/sharps count verified correct
    timeout_time: '',                   // time time-out was performed
  });

  // SSI Prophylaxis — SCIP Inf-1/Inf-2 (antibiotic within 60 min of incision, d/c within 24h)
  const [ssiProphylaxis, setSsiProphylaxis] = useState({
    antibiotic_given: false,
    antibiotic_agent: 'cefazolin' as 'cefazolin' | 'clindamycin' | 'vancomycin' | 'other',
    antibiotic_other: '',
    minutes_before_incision: '' as string,
    discontinued_within_24h: false,
  });

  const [asaClass, setAsaClass] = useState<'I'|'II'|'III'|'IV'|'V'|null>(null);
  const [woundClass, setWoundClass] = useState<'I_clean'|'II_clean_contaminated'|'III_contaminated'|'IV_dirty'|''>('');
  const [fireRisk, setFireRisk] = useState<'low'|'high'|''>('');

  const [narrativeDraft, setNarrativeDraft] = useState<OpNoteDraftResult | null>(null);
  const [generatingDraft, setGeneratingDraft] = useState(false);
  const [narrativeText, setNarrativeText] = useState('');
  const [showNarrative, setShowNarrative] = useState(false);

  useEffect(() => {
    loadExistingNote();
  }, [patientContext.encounter_id]);

  const loadExistingNote = async () => {
    try {
      const { data, error } = await supabase
        .schema('cr').from('operative_note_detailed')
        .select('*')
        .eq('encounter_id', patientContext.encounter_id)
        .single();

      if (data) {
        setFormData({
          pre_op_diagnosis_icd10:   data.pre_op_diagnosis_icd10 || [],
          procedures_performed:     data.procedures_performed_cpt?.map((_: string, i: number) => data.procedures_performed_cpt[i]) || [],
          procedures_performed_cpt: data.procedures_performed_cpt || [],
          laterality:               data.laterality || 'na',
          anesthesia_type:          data.anesthesia_type || 'general_endotracheal',
          anesthesia_provider:      data.anesthesia_provider || '',
          surgeon_primary:          data.surgeon_primary || '',
          positioning:              data.positioning || 'supine',
          prep_solution:            data.prep_solution || 'betadine',
          implant_details:          data.implant_details,
          tissue_removed:           data.tissue_removed_details,
          fat_harvested:            data.fat_harvested_details,
          estimated_blood_loss_ml:  data.estimated_blood_loss_ml || 0,
          crystalloid_given_ml:     data.crystalloid_given_ml || 0,
          drains_placed:            data.drains_placed || false,
          drain_details:            data.drain_details,
          complications:            data.complications || 'none',
          condition:                data.condition || 'stable',
          disposition:              data.disposition || 'pacu',
        });
        if (data.asa_class) setAsaClass(data.asa_class);
        if (data.wound_class) setWoundClass(data.wound_class);
        if (data.fire_risk) setFireRisk(data.fire_risk);
      }
    } catch (err) {
      console.error('Error loading operative note:', err);
    } finally {
      setLoading(false);
    }
  };

  const searchCptCodes = async (query: string) => {
    if (!query || query.length < 3) return;
    setSearchingCpt(true);
    try {
      const { data, error } = await supabase
        .rpc('search_cpt', { search_query: query });
      if (!error && data) {
        setCptResults(data.slice(0, 10));
      }
    } catch (err) {
      console.error('Error searching CPT:', err);
    } finally {
      setSearchingCpt(false);
    }
  };

  const searchIcdCodes = async (query: string) => {
    if (!query || query.length < 3) return;
    setSearchingIcd(true);
    try {
      const { data, error } = await supabase
        .rpc('search_icd10', { search_query: query });
      if (!error && data) {
        setIcdResults(data.slice(0, 10));
      }
    } catch (err) {
      console.error('Error searching ICD-10:', err);
    } finally {
      setSearchingIcd(false);
    }
  };

  const addCptCode = (code: string, description: string) => {
    if (!formData.procedures_performed_cpt.includes(code)) {
      setFormData(prev => ({
        ...prev,
        procedures_performed_cpt: [...prev.procedures_performed_cpt, code],
        procedures_performed: [...prev.procedures_performed, description]
      }));
    }
    setCptSearch('');
    setCptResults([]);
  };

  const addIcdCode = (code: string) => {
    if (!formData.pre_op_diagnosis_icd10.includes(code)) {
      setFormData(prev => ({
        ...prev,
        pre_op_diagnosis_icd10: [...prev.pre_op_diagnosis_icd10, code]
      }));
    }
    setIcdSearch('');
    setIcdResults([]);
  };

  // Caprini score: age points + factor points
  const capriniScore = (() => {
    let s = 0
    if (caprini.age_41_60) s += 1; if (caprini.age_61_74) s += 2; if (caprini.age_75_plus) s += 3
    if (caprini.minor_surgery) s += 1; if (caprini.major_surgery_lt45) s += 2; if (caprini.major_surgery_gt45) s += 3
    if (caprini.prior_dvt_pe) s += 3; if (caprini.family_history_clot) s += 3
    if (caprini.obesity_bmi_gt25) s += 1; if (caprini.varicose_veins) s += 1
    if (caprini.immobility_gt72h) s += 1; if (caprini.central_line) s += 2
    if (caprini.malignancy) s += 2; if (caprini.congestive_heart_failure) s += 1
    if (caprini.sepsis) s += 1; if (caprini.inflammatory_bowel) s += 1
    if (caprini.oral_contraceptives) s += 1; if (caprini.hrt) s += 1
    if (caprini.thrombophilia) s += 3; if (caprini.stroke) s += 5
    return s
  })()
  const capriniRisk = capriniScore <= 1 ? 'Low' : capriniScore <= 2 ? 'Moderate' : capriniScore <= 4 ? 'High' : 'Highest'
  const capriniProphylaxis = capriniScore <= 1 ? 'Early ambulation' : capriniScore <= 2 ? 'SCDs + ambulation' : 'SCDs + pharmacologic (LMWH/UFH)'

  // STOP-BANG score
  const stopBangScore = Object.values(stopBang).filter(Boolean).length
  const stopBangRisk = stopBangScore <= 2 ? 'Low' : stopBangScore <= 4 ? 'Intermediate' : 'High'

  // RCRI score (Lee 1999) — predicts major adverse cardiac events (MACE)
  const rcriScore = Object.values(rcri).filter(Boolean).length
  const rcriRisk = rcriScore === 0 ? 'Very Low (0.4%)' : rcriScore === 1 ? 'Low (1.0%)' : rcriScore === 2 ? 'Moderate (2.4%)' : 'High (≥5.4%)'
  const rcriClass = rcriScore === 0 ? 'green' : rcriScore === 1 ? 'green' : rcriScore === 2 ? 'amber' : 'red'

  // Universal Protocol — all 7 checkboxes required before completing
  const upComplete = universalProtocol.patient_identity_verified && universalProtocol.procedure_verified &&
    universalProtocol.site_marked && universalProtocol.timeout_performed &&
    universalProtocol.sponge_count_correct && universalProtocol.instrument_count_correct &&
    universalProtocol.needle_count_correct

  const handleSign = () => {
    const now = new Date().toISOString()
    setSigned(true)
    setSignedAt(now)
  }

  const handleSave = async () => {
    setSaving(true);
    try {
      const operativeNoteData = {
        org_id: patientContext.org_id,
        patient_id: patientContext.patient_id,
        encounter_id: patientContext.encounter_id,
        provider_id: patientContext.provider_id,
        // explicit column mapping — avoids tissue_removed boolean conflict and unknown columns
        pre_op_diagnosis_icd10:   formData.pre_op_diagnosis_icd10,
        procedures_performed_cpt: formData.procedures_performed_cpt,
        laterality:               formData.laterality,
        anesthesia_type:          formData.anesthesia_type,
        anesthesia_provider:      formData.anesthesia_provider,
        surgeon_primary:          formData.surgeon_primary,
        positioning:              formData.positioning,
        prep_solution:            formData.prep_solution,
        estimated_blood_loss_ml:  formData.estimated_blood_loss_ml,
        crystalloid_given_ml:     formData.crystalloid_given_ml,
        drains_placed:            formData.drains_placed,
        implant_details:          formData.implant_details ?? null,
        tissue_removed_details:   formData.tissue_removed ?? null,
        fat_harvested_details:    formData.fat_harvested ?? null,
        drain_details:            formData.drain_details ?? null,
        complications:            formData.complications,
        condition:                formData.condition,
        disposition:              formData.disposition,
        operative_date: new Date().toISOString(),
        surgical_consent_obtained: consentObtained,
        caprini_score: capriniScore,
        caprini_risk: capriniRisk,
        stop_bang_score: stopBangScore,
        stop_bang_risk: stopBangRisk,
        rcri_score: rcriScore,
        rcri_risk: rcriRisk,
        universal_protocol: universalProtocol,
        ssi_prophylaxis: ssiProphylaxis,
        asa_class: asaClass,
        wound_class: woundClass || null,
        fire_risk: fireRisk || null,
        electronically_signed: signed,
        signed_at: signedAt ?? null,
      };

      const { error } = await supabase
        .schema('cr').from('operative_note_detailed')
        .upsert(operativeNoteData, {
          onConflict: 'encounter_id'
        });
      
      if (error) throw error;

      // Auto-populate superbill
      await autoPopulateSuperbill();

      setSaved(true);
      setTimeout(() => setSaved(false), 3000);
    } catch (err) {
      console.error('Error saving operative note:', err);
      alert('Error saving. Please try again.');
    } finally {
      setSaving(false);
    }
  };

  const autoPopulateSuperbill = async () => {
    try {
      const superbillData = {
        org_id: patientContext.org_id,
        patient_id: patientContext.patient_id,
        encounter_id: patientContext.encounter_id,
        provider_id: patientContext.provider_id,
        cpt_codes: formData.procedures_performed_cpt,
        icd10_codes: formData.pre_op_diagnosis_icd10,
        date_of_service: new Date().toISOString().split('T')[0]
      };

      await supabase
        .schema('cr').from('superbill')
        .upsert(superbillData, { onConflict: 'encounter_id' });
    } catch (err) {
      console.error('Error auto-populating superbill:', err);
    }
  };

  const handleGenerateDraft = async () => {
    setGeneratingDraft(true);
    setShowNarrative(true);
    try {
      const draft = await getOpNoteDraft({
        procedureName: formData.procedures_performed_cpt?.map((p: any) => p.description || p).join(', ') || 'Procedure',
        procedureType: formData.pre_op_diagnosis_icd10 || 'plastic_surgery',
        laterality: formData.laterality,
        anesthesiaType: formData.anesthesia_type,
        anesthesiaProvider: formData.anesthesia_provider,
        surgeonPrimary: formData.surgeon_primary,
        positioning: formData.positioning,
        prepSolution: formData.prep_solution,
        estimatedBloodLossML: formData.estimated_blood_loss_ml ? Number(formData.estimated_blood_loss_ml) : undefined,
        crystalloidGivenML: formData.crystalloid_given_ml ? Number(formData.crystalloid_given_ml) : undefined,
        drainsPlaced: formData.drains_placed,
        drainDetails: formData.drain_details,
        implantDetails: formData.implant_details,
        tissueRemovedDetails: formData.tissue_removed,
        fatHarvestedDetails: formData.fat_harvested,
        complications: formData.complications,
        condition: formData.condition,
        disposition: formData.disposition,
        asaClass: asaClass ?? undefined,
        capriniScore: Object.values(caprini).filter(Boolean).length,
        stopBangScore: Object.values(stopBang).filter(Boolean).length,
        rcriScore: Object.values(rcri).filter(Boolean).length,
        woundClass: woundClass || undefined,
        universalProtocolComplete: universalProtocol.patient_identity_verified &&
          universalProtocol.procedure_verified && universalProtocol.site_marked &&
          universalProtocol.timeout_performed,
        ssiAntibiotic: ssiProphylaxis.antibiotic_given ? ssiProphylaxis.antibiotic_agent : undefined,
        encounterId: patientContext.encounter_id?.toString(),
        orgId: patientContext.org_id,
        providerId: patientContext.provider_id?.toString(),
      });
      setNarrativeDraft(draft);
      // Pre-populate narrative text area with the operative narrative paragraph
      setNarrativeText(draft.operativeNarrative);
    } catch (err) {
      console.error('Draft generation error:', err);
      setNarrativeText('Error generating draft. Please try again.');
    } finally {
      setGeneratingDraft(false);
    }
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center py-12">
        <div className="animate-spin rounded-full h-12 w-12 border-t-2 border-b-2 border-[#c9a96e]"></div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-rajdhani font-bold text-white">Operative Note</h2>
          <p className="text-gray-400 text-sm mt-1">Comprehensive OR documentation</p>
        </div>
        <button
          onClick={handleSave}
          disabled={saving}
          className="flex items-center gap-2 px-6 py-2 bg-[#c9a96e] hover:bg-[#b39960] text-[#060e1c] font-rajdhani font-semibold rounded transition-colors disabled:opacity-50"
        >
          {saving ? (
            <>
              <div className="animate-spin rounded-full h-4 w-4 border-t-2 border-b-2 border-[#060e1c]"></div>
              Saving...
            </>
          ) : saved ? (
            <>
              <CheckCircle className="w-4 h-4" />
              Saved!
            </>
          ) : (
            <>
              <Save className="w-4 h-4" />
              Save & Bill
            </>
          )}
        </button>
      </div>

      {/* Pre-Op Diagnosis */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Pre-Operative Diagnosis (ICD-10)</h3>
        
        <div className="relative mb-4">
          <input
            type="text"
            value={icdSearch}
            onChange={(e) => {
              setIcdSearch(e.target.value);
              searchIcdCodes(e.target.value);
            }}
            placeholder="Search ICD-10 codes (e.g., breast ptosis, gynecomastia)..."
            className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white pr-10"
          />
          <Search className="absolute right-3 top-2.5 w-5 h-5 text-gray-500" />
          
          {icdResults.length > 0 && (
            <div className="absolute z-10 w-full mt-1 bg-[#0a1628] border border-[#c9a96e] rounded-lg shadow-lg max-h-60 overflow-y-auto">
              {icdResults.map((result) => (
                <button
                  key={result.code}
                  onClick={() => addIcdCode(result.code)}
                  className="w-full text-left px-4 py-3 hover:bg-[#c9a96e]/20 border-b border-gray-700 last:border-0"
                >
                  <div className="font-mono text-[#c9a96e]">{result.code}</div>
                  <div className="text-sm text-white">{result.description}</div>
                </button>
              ))}
            </div>
          )}
        </div>

        <div className="flex flex-wrap gap-2">
          {formData.pre_op_diagnosis_icd10.map((code) => (
            <div key={code} className="flex items-center gap-2 bg-[#060e1c] border border-[#c9a96e] rounded px-3 py-1">
              <span className="text-white font-mono">{code}</span>
              <button
                onClick={() => setFormData(prev => ({
                  ...prev,
                  pre_op_diagnosis_icd10: prev.pre_op_diagnosis_icd10.filter(c => c !== code)
                }))}
                className="text-red-500 hover:text-red-400"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          ))}
        </div>
      </div>

      {/* Procedures Performed */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Procedures Performed (CPT)</h3>
        
        <div className="relative mb-4">
          <input
            type="text"
            value={cptSearch}
            onChange={(e) => {
              setCptSearch(e.target.value);
              searchCptCodes(e.target.value);
            }}
            placeholder="Search CPT codes (e.g., breast augmentation, liposuction)..."
            className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white pr-10"
          />
          <Search className="absolute right-3 top-2.5 w-5 h-5 text-gray-500" />
          
          {cptResults.length > 0 && (
            <div className="absolute z-10 w-full mt-1 bg-[#0a1628] border border-[#c9a96e] rounded-lg shadow-lg max-h-60 overflow-y-auto">
              {cptResults.map((result) => (
                <button
                  key={result.code}
                  onClick={() => addCptCode(result.code, result.description)}
                  className="w-full text-left px-4 py-3 hover:bg-[#c9a96e]/20 border-b border-gray-700 last:border-0"
                >
                  <div className="font-mono text-[#c9a96e]">{result.code}</div>
                  <div className="text-sm text-white">{result.description}</div>
                  <div className="text-xs text-gray-400 mt-1">Fee: ${result.medicare_fee_2025}</div>
                </button>
              ))}
            </div>
          )}
        </div>

        <div className="space-y-2">
          {formData.procedures_performed.map((proc, idx) => (
            <div key={idx} className="flex items-center justify-between bg-[#060e1c] border border-gray-700 rounded px-4 py-3">
              <div>
                <span className="text-[#c9a96e] font-mono font-semibold">{formData.procedures_performed_cpt[idx]}</span>
                <span className="text-white ml-3">{proc}</span>
              </div>
              <button
                onClick={() => setFormData(prev => ({
                  ...prev,
                  procedures_performed: prev.procedures_performed.filter((_, i) => i !== idx),
                  procedures_performed_cpt: prev.procedures_performed_cpt.filter((_, i) => i !== idx)
                }))}
                className="text-red-500 hover:text-red-400"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          ))}
        </div>
      </div>

      {/* Anesthesia & Team */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Anesthesia & Surgical Team</h3>
        
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Anesthesia Type
            </label>
            <select
              value={formData.anesthesia_type}
              onChange={(e) => setFormData(prev => ({ ...prev, anesthesia_type: e.target.value as any }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="general_endotracheal">General Endotracheal</option>
              <option value="general_lma">General LMA</option>
              <option value="mac">MAC (Monitored Anesthesia Care)</option>
              <option value="local">Local</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Anesthesia Provider
            </label>
            <input
              type="text"
              value={formData.anesthesia_provider}
              onChange={(e) => setFormData(prev => ({ ...prev, anesthesia_provider: e.target.value }))}
              placeholder="Anesthesiologist name"
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            />
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Primary Surgeon
            </label>
            <input
              type="text"
              value={formData.surgeon_primary}
              onChange={(e) => setFormData(prev => ({ ...prev, surgeon_primary: e.target.value }))}
              placeholder="Surgeon name"
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            />
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Laterality
            </label>
            <select
              value={formData.laterality}
              onChange={(e) => setFormData(prev => ({ ...prev, laterality: e.target.value as any }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="na">Not Applicable</option>
              <option value="left">Left</option>
              <option value="right">Right</option>
              <option value="bilateral">Bilateral</option>
              <option value="midline">Midline</option>
            </select>
          </div>
        </div>
      </div>

      {/* Intraoperative Details */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Intraoperative Details</h3>
        
        <div className="grid grid-cols-3 gap-4 mb-4">
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Positioning
            </label>
            <select
              value={formData.positioning}
              onChange={(e) => setFormData(prev => ({ ...prev, positioning: e.target.value as any }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="supine">Supine</option>
              <option value="prone">Prone</option>
              <option value="lateral_right">Lateral Right</option>
              <option value="lateral_left">Lateral Left</option>
              <option value="sitting">Sitting</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Prep Solution
            </label>
            <select
              value={formData.prep_solution}
              onChange={(e) => setFormData(prev => ({ ...prev, prep_solution: e.target.value as any }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="betadine">Betadine</option>
              <option value="chlorhexidine">Chlorhexidine</option>
              <option value="alcohol">Alcohol</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              EBL (ml)
            </label>
            <input
              type="number"
              value={formData.estimated_blood_loss_ml || ''}
              onChange={(e) => setFormData(prev => ({ ...prev, estimated_blood_loss_ml: parseInt(e.target.value) || 0 }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            />
          </div>
        </div>

        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Crystalloid Given (ml)
            </label>
            <input
              type="number"
              value={formData.crystalloid_given_ml || ''}
              onChange={(e) => setFormData(prev => ({ ...prev, crystalloid_given_ml: parseInt(e.target.value) || 0 }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            />
          </div>

          <div>
            <label className="flex items-center gap-3 h-full items-end pb-2">
              <input
                type="checkbox"
                checked={formData.drains_placed}
                onChange={(e) => setFormData(prev => ({ ...prev, drains_placed: e.target.checked }))}
                className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
              />
              <span className="text-white">Drains placed</span>
            </label>
          </div>
        </div>
      </div>

      {/* Complications & Disposition */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Post-Operative</h3>
        
        <div className="grid grid-cols-3 gap-4 mb-4">
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Complications
            </label>
            <input
              type="text"
              value={formData.complications}
              onChange={(e) => setFormData(prev => ({ ...prev, complications: e.target.value }))}
              placeholder="None"
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            />
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Condition
            </label>
            <select
              value={formData.condition}
              onChange={(e) => setFormData(prev => ({ ...prev, condition: e.target.value as any }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="stable">Stable</option>
              <option value="critical">Critical</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Disposition
            </label>
            <select
              value={formData.disposition}
              onChange={(e) => setFormData(prev => ({ ...prev, disposition: e.target.value as any }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="pacu">PACU</option>
              <option value="icu">ICU</option>
              <option value="floor">Floor</option>
              <option value="home">Home</option>
            </select>
          </div>
        </div>
      </div>

      {/* ── RCRI — Revised Cardiac Risk Index ── */}
      <div className="border border-gray-700 rounded-lg p-6">
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-lg font-rajdhani font-semibold text-white">RCRI — Revised Cardiac Risk Index</h3>
          <div className={`px-3 py-1 rounded text-xs font-bold ${
            rcriClass === 'red' ? 'bg-red-900/40 text-red-400' :
            rcriClass === 'amber' ? 'bg-amber-900/40 text-amber-400' :
            'bg-green-900/40 text-green-400'
          }`}>
            {rcriScore}/6 — {rcriRisk} MACE Risk
          </div>
        </div>
        <p className="text-xs text-gray-500 mb-3">Lee 1999. Predicts major adverse cardiac events (MI, pulmonary edema, VF/VT, heart block, cardiac arrest) within 30 days of non-cardiac surgery.</p>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
          {([
            ['high_risk_surgery',        'High-risk surgery (suprainguinal vascular, intraperitoneal, intrathoracic)'],
            ['ischemic_heart_disease',   'History of ischemic heart disease (MI, nitrates, Q-waves, positive stress test)'],
            ['congestive_heart_failure', 'History of congestive heart failure (pulmonary edema, PND, S3, bilateral rales)'],
            ['cerebrovascular_disease',  'History of cerebrovascular disease (stroke or TIA)'],
            ['insulin_dependent_dm',     'Insulin-dependent diabetes mellitus'],
            ['creatinine_gt2',           'Pre-op creatinine > 2.0 mg/dL (> 177 μmol/L)'],
          ] as [keyof typeof rcri, string][]).map(([key, label]) => (
            <label key={key} className="flex items-start gap-2 cursor-pointer text-sm text-gray-300">
              <input type="checkbox" checked={rcri[key]}
                onChange={e => setRcri(p => ({ ...p, [key]: e.target.checked }))}
                className="mt-0.5 w-4 h-4 accent-[#c9a96e]" />
              {label}
            </label>
          ))}
        </div>
      </div>

      {/* ── Universal Protocol — Joint Commission ── */}
      <div className={`border rounded-lg p-6 ${upComplete ? 'border-green-700/50 bg-green-900/10' : 'border-red-700/40 bg-red-900/10'}`}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-lg font-rajdhani font-semibold text-white">Universal Protocol <span className="text-xs text-gray-500 font-normal ml-2">Joint Commission NPSG · UP.01.01–UP.01.03</span></h3>
          {upComplete
            ? <span className="text-green-400 text-xs font-bold flex items-center gap-1"><CheckCircle className="w-4 h-4" /> Complete</span>
            : <span className="text-red-400 text-xs font-bold">Incomplete — required before proceeding</span>
          }
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4">
          {([
            ['patient_identity_verified',  'Patient identity verified (2 identifiers)'],
            ['procedure_verified',         'Consent matches planned procedure'],
            ['site_marked',                'Surgical site/side marked by surgeon'],
            ['timeout_performed',          'Team time-out performed'],
            ['sponge_count_correct',       'Sponge count verified correct'],
            ['instrument_count_correct',   'Instrument count verified correct'],
            ['needle_count_correct',       'Needle/sharps count verified correct'],
          ] as [keyof typeof universalProtocol, string][]).map(([key, label]) => (
            <label key={key} className="flex items-center gap-2 cursor-pointer text-sm text-gray-300">
              <input type="checkbox" checked={universalProtocol[key] as boolean}
                onChange={e => setUniversalProtocol(p => ({ ...p, [key]: e.target.checked }))}
                className="w-4 h-4 accent-green-500" />
              {label}
            </label>
          ))}
        </div>
        <div className="flex items-center gap-3">
          <label className="text-xs text-gray-400 whitespace-nowrap">Time-out time:</label>
          <input type="time" value={universalProtocol.timeout_time}
            onChange={e => setUniversalProtocol(p => ({ ...p, timeout_time: e.target.value }))}
            className="bg-[#060e1c] border border-gray-700 rounded px-3 py-1 text-white text-sm w-36" />
        </div>
      </div>

      {/* ── SSI Prophylaxis — SCIP Inf-1/Inf-2 ── */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-3">
          SSI Prophylaxis <span className="text-xs text-gray-500 font-normal ml-2">SCIP Inf-1 / Inf-2</span>
        </h3>
        <label className="flex items-center gap-3 mb-4 cursor-pointer">
          <input type="checkbox" checked={ssiProphylaxis.antibiotic_given}
            onChange={e => setSsiProphylaxis(p => ({ ...p, antibiotic_given: e.target.checked }))}
            className="w-4 h-4 accent-[#c9a96e]" />
          <span className="text-white text-sm">Prophylactic antibiotic administered</span>
        </label>
        {ssiProphylaxis.antibiotic_given && (
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mt-2">
            <div>
              <label className="block text-xs text-[#c9a96e] mb-1">Agent</label>
              <select value={ssiProphylaxis.antibiotic_agent}
                onChange={e => setSsiProphylaxis(p => ({ ...p, antibiotic_agent: e.target.value as any }))}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm">
                <option value="cefazolin">Cefazolin (Ancef)</option>
                <option value="clindamycin">Clindamycin (PCN allergy)</option>
                <option value="vancomycin">Vancomycin (MRSA risk)</option>
                <option value="other">Other</option>
              </select>
            </div>
            {ssiProphylaxis.antibiotic_agent === 'other' && (
              <div>
                <label className="block text-xs text-[#c9a96e] mb-1">Specify</label>
                <input type="text" value={ssiProphylaxis.antibiotic_other}
                  onChange={e => setSsiProphylaxis(p => ({ ...p, antibiotic_other: e.target.value }))}
                  className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm"
                  placeholder="Drug name" />
              </div>
            )}
            <div>
              <label className="block text-xs text-[#c9a96e] mb-1">Min before incision</label>
              <input type="number" value={ssiProphylaxis.minutes_before_incision}
                onChange={e => setSsiProphylaxis(p => ({ ...p, minutes_before_incision: e.target.value }))}
                className={`w-full bg-[#060e1c] border rounded px-3 py-2 text-white text-sm ${
                  Number(ssiProphylaxis.minutes_before_incision) > 60 ? 'border-red-500' : 'border-gray-700'
                }`}
                placeholder="e.g. 30" />
              {Number(ssiProphylaxis.minutes_before_incision) > 60 && (
                <p className="text-red-400 text-xs mt-1">SCIP: must be ≤ 60 min (≤ 120 for vanco)</p>
              )}
            </div>
            <div className="flex items-end pb-2">
              <label className="flex items-center gap-2 cursor-pointer text-sm text-gray-300">
                <input type="checkbox" checked={ssiProphylaxis.discontinued_within_24h}
                  onChange={e => setSsiProphylaxis(p => ({ ...p, discontinued_within_24h: e.target.checked }))}
                  className="w-4 h-4 accent-[#c9a96e]" />
                Discontinued within 24h
              </label>
            </div>
          </div>
        )}
      </div>

      {/* OR Classification: ASA, Wound Class, Fire Risk */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">OR Classification</h3>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">ASA Physical Status</label>
            <select value={asaClass ?? ''}
              onChange={e => setAsaClass((e.target.value as 'I'|'II'|'III'|'IV'|'V') || null)}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm">
              <option value="">Select...</option>
              <option value="I">ASA I — Healthy</option>
              <option value="II">ASA II — Mild disease</option>
              <option value="III">ASA III — Severe disease</option>
              <option value="IV">ASA IV — Threat to life</option>
              <option value="V">ASA V — Moribund</option>
            </select>
          </div>
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">Wound Classification (Altemeier)</label>
            <select value={woundClass}
              onChange={e => setWoundClass(e.target.value as typeof woundClass)}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm">
              <option value="">Select...</option>
              <option value="I_clean">Class I — Clean</option>
              <option value="II_clean_contaminated">Class II — Clean-Contaminated</option>
              <option value="III_contaminated">Class III — Contaminated</option>
              <option value="IV_dirty">Class IV — Dirty/Infected</option>
            </select>
          </div>
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">Fire Risk Assessment <span className="text-xs text-gray-500">(AORN / JC NPSG.03.04.01)</span></label>
            <div className="flex gap-3">
              {([['low','Low'],['high','High']] as const).map(([val,label]) => (
                <button key={val} onClick={() => setFireRisk(val)}
                  className={`flex-1 px-3 py-2 rounded border text-sm transition-colors ${
                    fireRisk === val
                      ? val === 'high' ? 'bg-red-800 border-red-500 text-white' : 'bg-green-800 border-green-500 text-white'
                      : 'border-gray-600 text-gray-300 hover:border-gray-400'
                  }`}>{label}</button>
              ))}
            </div>
            <p className="text-xs text-gray-500 mt-1">High = O&#x2082;/N&#x2082;O enrichment near ignition source</p>
          </div>
        </div>
      </div>

      {/* ── Pre-Op Safety Screening ── */}
      <div className="border border-gray-700 rounded-lg p-6 space-y-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white">Pre-Op Safety Screening</h3>

        {/* Caprini VTE */}
        <div>
          <div className="flex items-center justify-between mb-3">
            <h4 className="text-sm font-rajdhani font-semibold text-[#c9a96e] uppercase tracking-wider">Caprini VTE Risk Score</h4>
            <div className={`px-3 py-1 rounded text-xs font-bold ${
              capriniRisk === 'Low' ? 'bg-green-900/40 text-green-400' :
              capriniRisk === 'Moderate' ? 'bg-yellow-900/40 text-yellow-400' :
              capriniRisk === 'High' ? 'bg-orange-900/40 text-orange-400' :
              'bg-red-900/40 text-red-400'
            }`}>
              {capriniScore} pts — {capriniRisk} · {capriniProphylaxis}
            </div>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
            {([
              ['age_41_60','Age 41–60 (+1)'], ['age_61_74','Age 61–74 (+2)'], ['age_75_plus','Age ≥75 (+3)'],
              ['minor_surgery','Minor surgery (+1)'], ['major_surgery_lt45','Major surgery <45 min (+2)'], ['major_surgery_gt45','Major surgery >45 min (+3)'],
              ['prior_dvt_pe','Prior DVT/PE (+3)'], ['family_history_clot','Family Hx clot (+3)'], ['thrombophilia','Thrombophilia (+3)'],
              ['obesity_bmi_gt25','BMI >25 (+1)'], ['varicose_veins','Varicose veins (+1)'], ['immobility_gt72h','Immobility >72h (+1)'],
              ['central_line','Central line (+2)'], ['malignancy','Active malignancy (+2)'], ['stroke','Prior stroke (+5)'],
              ['congestive_heart_failure','CHF (+1)'], ['sepsis','Sepsis (+1)'], ['inflammatory_bowel','IBD (+1)'],
              ['oral_contraceptives','OCP use (+1)'], ['hrt','HRT (+1)'],
            ] as [keyof typeof caprini, string][]).map(([key, label]) => (
              <label key={key} className="flex items-center gap-2 cursor-pointer text-xs text-gray-300">
                <input type="checkbox" checked={caprini[key]}
                  onChange={e => setCaprini(p => ({ ...p, [key]: e.target.checked }))}
                  className="w-3.5 h-3.5 accent-[#c9a96e]" />
                {label}
              </label>
            ))}
          </div>
        </div>

        {/* STOP-BANG */}
        <div>
          <div className="flex items-center justify-between mb-3">
            <h4 className="text-sm font-rajdhani font-semibold text-[#c9a96e] uppercase tracking-wider">STOP-BANG OSA Screen</h4>
            <div className={`px-3 py-1 rounded text-xs font-bold ${
              stopBangRisk === 'Low' ? 'bg-green-900/40 text-green-400' :
              stopBangRisk === 'Intermediate' ? 'bg-yellow-900/40 text-yellow-400' :
              'bg-red-900/40 text-red-400'
            }`}>
              {stopBangScore}/8 — {stopBangRisk} OSA Risk
            </div>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
            {([
              ['snoring','S — Loud snoring'],
              ['tired','T — Tired/fatigued daytime'],
              ['observed','O — Observed apnea'],
              ['pressure','P — Hypertension/treatment'],
              ['bmi_gt35','B — BMI > 35'],
              ['age_gt50','A — Age > 50'],
              ['neck_gt40','N — Neck circ. > 40 cm'],
              ['gender_male','G — Male gender'],
            ] as [keyof typeof stopBang, string][]).map(([key, label]) => (
              <label key={key} className="flex items-center gap-2 cursor-pointer text-xs text-gray-300">
                <input type="checkbox" checked={stopBang[key]}
                  onChange={e => setStopBang(p => ({ ...p, [key]: e.target.checked }))}
                  className="w-3.5 h-3.5 accent-[#c9a96e]" />
                {label}
              </label>
            ))}
          </div>
        </div>
      </div>

      {/* Surgical Informed Consent */}
      <div className="border border-amber-700/50 rounded-lg p-5 bg-amber-900/10">
        <h3 className="text-base font-rajdhani font-semibold text-amber-400 mb-3">Surgical Informed Consent</h3>
        <label className="flex items-start gap-3 cursor-pointer">
          <input type="checkbox" checked={consentObtained}
            onChange={e => setConsentObtained(e.target.checked)}
            className="mt-1 w-4 h-4 accent-amber-400" />
          <span className="text-sm text-gray-300 leading-relaxed">
            Patient has been counseled on procedure risks, benefits, alternatives, and expected outcomes.
            Written informed consent has been obtained, signed, and placed in the chart.
          </span>
        </label>
      </div>

      {/* AI Op Note Draft */}
      <div className="border border-[#c9a96e]/30 rounded-lg p-6">
        <div className="flex items-center justify-between mb-4">
          <div>
            <h3 className="text-lg font-rajdhani font-semibold text-white">AI Narrative Draft</h3>
            <p className="text-xs text-gray-500 mt-0.5">Generate a dictation-quality operative note from the data above. Review and edit before saving.</p>
          </div>
          <button
            onClick={handleGenerateDraft}
            disabled={generatingDraft}
            className="flex items-center gap-2 px-4 py-2 bg-[#c9a96e] hover:bg-[#b39960] disabled:opacity-50 text-[#060e1c] font-rajdhani font-bold text-sm rounded transition-colors"
          >
            {generatingDraft ? (
              <>
                <div className="animate-spin rounded-full h-4 w-4 border-t-2 border-b-2 border-[#060e1c]" />
                Generating…
              </>
            ) : (
              <>✦ Generate Draft</>
            )}
          </button>
        </div>

        {showNarrative && (
          <div className="space-y-4">
            {narrativeDraft && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-sm">
                <div className="bg-[#060e1c] rounded p-3 border border-gray-800">
                  <div className="text-xs text-[#c9a96e] mb-1">Pre-op Diagnosis</div>
                  <div className="text-gray-200">{narrativeDraft.preoperativeDiagnosis}</div>
                </div>
                <div className="bg-[#060e1c] rounded p-3 border border-gray-800">
                  <div className="text-xs text-[#c9a96e] mb-1">Procedure Performed</div>
                  <div className="text-gray-200">{narrativeDraft.procedurePerformed}</div>
                </div>
                <div className="bg-[#060e1c] rounded p-3 border border-gray-800">
                  <div className="text-xs text-[#c9a96e] mb-1">Intraoperative Findings</div>
                  <div className="text-gray-200">{narrativeDraft.findings}</div>
                </div>
                <div className="bg-[#060e1c] rounded p-3 border border-gray-800">
                  <div className="text-xs text-[#c9a96e] mb-1">Complications / EBL / Disposition</div>
                  <div className="text-gray-200">{narrativeDraft.complications} · EBL {narrativeDraft.ebl} · {narrativeDraft.disposition}</div>
                </div>
              </div>
            )}
            <div>
              <label className="block text-xs text-[#c9a96e] mb-2">Operative Narrative <span className="text-gray-500">(editable — review before signing)</span></label>
              <textarea
                value={narrativeText}
                onChange={e => setNarrativeText(e.target.value)}
                rows={10}
                placeholder={generatingDraft ? 'Generating narrative…' : 'Click Generate Draft to create AI-assisted operative note…'}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-3 text-white text-sm leading-relaxed font-mono focus:border-[#c9a96e] focus:outline-none"
              />
            </div>
          </div>
        )}
      </div>

      {/* Electronic Signature */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-3">Provider Attestation</h3>
        {signed ? (
          <div className="flex items-center gap-3 text-green-400">
            <CheckCircle className="w-5 h-5" />
            <span className="text-sm">Electronically signed by provider — {signedAt ? new Date(signedAt).toLocaleString() : ''}</span>
          </div>
        ) : (
          <div className="flex items-center gap-4">
            <span className="text-sm text-gray-400">
              By clicking "Sign Note" you attest that this operative note is accurate and complete to the best of your knowledge.
            </span>
            <button onClick={handleSign}
              className="flex-shrink-0 px-5 py-2 bg-green-700 hover:bg-green-600 text-white font-rajdhani font-semibold rounded transition-colors text-sm">
              Sign Note
            </button>
          </div>
        )}
      </div>

      {/* Billing Preview */}
      {(formData.procedures_performed_cpt.length > 0 || formData.pre_op_diagnosis_icd10.length > 0) && (
        <div className="border border-[#c9a96e] bg-[#c9a96e]/10 rounded-lg p-6">
          <h3 className="text-lg font-rajdhani font-semibold text-[#c9a96e] mb-3">
            📋 Superbill Preview
          </h3>
          <div className="grid grid-cols-2 gap-6">
            <div>
              <div className="text-sm text-gray-400 mb-2">CPT Codes</div>
              {formData.procedures_performed_cpt.map((code) => (
                <div key={code} className="text-white font-mono">{code}</div>
              ))}
            </div>
            <div>
              <div className="text-sm text-gray-400 mb-2">ICD-10 Codes</div>
              {formData.pre_op_diagnosis_icd10.map((code) => (
                <div key={code} className="text-white font-mono">{code}</div>
              ))}
            </div>
          </div>
          <p className="text-xs text-gray-400 mt-3">
            ✓ Will auto-populate superbill on save
          </p>
        </div>
      )}
    </div>
  );
}
