import React, { useState, useEffect } from 'react';
import { supabase } from '../../lib/supabase';
import { Save, CheckCircle, Search, Plus, X } from 'lucide-react';

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
    snoring: false,        // S — Do you snore loudly?
    tired: false,          // T — Do you often feel tired/fatigued?
    observed: false,       // O — Observed stop breathing during sleep?
    pressure: false,       // P — High blood pressure or treatment?
    bmi_gt35: false,       // B — BMI > 35?
    age_gt50: false,       // A — Age > 50?
    neck_gt40: false,      // N — Neck circumference > 40 cm?
    gender_male: false,    // G — Male gender?
  });

  useEffect(() => {
    loadExistingNote();
  }, [patientContext.encounter_id]);

  const loadExistingNote = async () => {
    try {
      const { data, error } = await supabase
        .from('cr.operative_note_detailed')
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
        electronically_signed: signed,
        signed_at: signedAt ?? null,
      };

      const { error } = await supabase
        .from('cr.operative_note_detailed')
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
        .from('cr.superbill')
        .upsert(superbillData, { onConflict: 'encounter_id' });
    } catch (err) {
      console.error('Error auto-populating superbill:', err);
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
