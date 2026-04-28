import React, { useState, useEffect } from 'react';
import { supabase } from '@/lib/supabase';
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
          pre_op_diagnosis_icd10: data.pre_op_diagnosis_icd10 || [],
          procedures_performed: data.procedures_performed || [],
          procedures_performed_cpt: data.procedures_performed_cpt || [],
          laterality: data.laterality || 'na',
          anesthesia_type: data.anesthesia_type || 'general_endotracheal',
          anesthesia_provider: data.anesthesia_provider || '',
          surgeon_primary: data.surgeon_primary || '',
          positioning: data.positioning || 'supine',
          prep_solution: data.prep_solution || 'betadine',
          implant_details: data.implant_details,
          tissue_removed: data.tissue_removed,
          fat_harvested: data.fat_harvested,
          estimated_blood_loss_ml: data.estimated_blood_loss_ml || 0,
          crystalloid_given_ml: data.crystalloid_given_ml || 0,
          drains_placed: data.drains_placed || false,
          drain_details: data.drain_details,
          complications: data.complications || 'none',
          condition: data.condition || 'stable',
          disposition: data.disposition || 'pacu'
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

  const handleSave = async () => {
    setSaving(true);
    try {
      const operativeNoteData = {
        org_id: patientContext.org_id,
        patient_id: patientContext.patient_id,
        encounter_id: patientContext.encounter_id,
        provider_id: patientContext.provider_id,
        ...formData,
        operative_date: new Date().toISOString(),
        electronically_signed: true,
        signed_at: new Date().toISOString()
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
