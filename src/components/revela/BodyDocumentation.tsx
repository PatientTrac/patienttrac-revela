import React, { useState, useEffect } from 'react';
import { supabase } from '../../lib/supabase';
import { Save, CheckCircle, AlertTriangle } from 'lucide-react';

interface PatientContext {
  patient_id: number;
  encounter_id: number;
  patient_name: string;
  patient_sex: 'male' | 'female' | 'non_binary';
  org_id: string;
  provider_id: number;
}

interface BodyContouringData {
  abdomen: {
    skin_laxity: 'none' | 'mild' | 'moderate' | 'severe';
    striae: boolean;
    pannus_present: boolean;
    pannus_grade: 'I' | 'II' | 'III' | 'IV';
    diastasis_recti: boolean;
    diastasis_width_cm: number;
    measurements: {
      waist_circumference_cm: number;
      hip_circumference_cm: number;
      waist_to_hip_ratio: number;
      abdominal_projection_cm: number;
    };
  };
  buttocks: {
    volume: 'deficient' | 'adequate' | 'full';
    shape: 'round' | 'square' | 'flat' | 'v_shape';
    ptosis: 'none' | 'mild' | 'moderate' | 'severe';
    asymmetry: boolean;
    cellulite: 'none' | 'mild' | 'moderate' | 'severe';
    hip_dips_present: boolean;
    banana_roll: boolean;
  };
  flanks: {
    fullness: 'none' | 'mild' | 'moderate' | 'severe';
    love_handles: boolean;
  };
  thighs: {
    inner_thigh_laxity: 'none' | 'mild' | 'moderate' | 'severe';
    outer_thigh_fullness: 'none' | 'mild' | 'moderate';
  };
  post_bariatric: {
    is_post_bariatric: boolean;
    weight_loss_kg: number;
    time_since_surgery_months: number;
    weight_stable: boolean;
    stable_duration_months: number;
  };
}

interface Props {
  patientContext: PatientContext;
}

export default function BodyDocumentation({ patientContext }: Props) {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  
  const [formData, setFormData] = useState<BodyContouringData>({
    abdomen: {
      skin_laxity: 'none',
      striae: false,
      pannus_present: false,
      pannus_grade: 'I',
      diastasis_recti: false,
      diastasis_width_cm: 0,
      measurements: {
        waist_circumference_cm: 0,
        hip_circumference_cm: 0,
        waist_to_hip_ratio: 0,
        abdominal_projection_cm: 0
      }
    },
    buttocks: {
      volume: 'adequate',
      shape: 'round',
      ptosis: 'none',
      asymmetry: false,
      cellulite: 'none',
      hip_dips_present: false,
      banana_roll: false
    },
    flanks: {
      fullness: 'none',
      love_handles: false
    },
    thighs: {
      inner_thigh_laxity: 'none',
      outer_thigh_fullness: 'none'
    },
    post_bariatric: {
      is_post_bariatric: false,
      weight_loss_kg: 0,
      time_since_surgery_months: 0,
      weight_stable: false,
      stable_duration_months: 0
    }
  });

  const [suggestedProcedures, setSuggestedProcedures] = useState<string[]>([]);
  const [consentObtained, setConsentObtained] = useState(false);
  const [asaClass, setAsaClass] = useState<'I'|'II'|'III'|'IV'|'V'|null>(null);
  const [herniaPresent, setHerniaPresent] = useState<boolean|null>(null);
  const [herniaDescription, setHerniaDescription] = useState('');
  const [keloidHistory, setKeloidHistory] = useState<boolean|null>(null);
  const [bodyQ, setBodyQ] = useState({
    body_image: 0,
    weight_appearance: 0,
    clothes_fit: 0,
    physical_function: 0,
    psychological_wellbeing: 0,
    satisfaction_with_outcome: 0,
  });

  useEffect(() => {
    loadExistingDocumentation();
  }, [patientContext.patient_id]);

  // Auto-calculate waist-to-hip ratio
  useEffect(() => {
    const { waist_circumference_cm, hip_circumference_cm } = formData.abdomen.measurements;
    if (waist_circumference_cm > 0 && hip_circumference_cm > 0) {
      const ratio = waist_circumference_cm / hip_circumference_cm;
      setFormData(prev => ({
        ...prev,
        abdomen: {
          ...prev.abdomen,
          measurements: {
            ...prev.abdomen.measurements,
            waist_to_hip_ratio: parseFloat(ratio.toFixed(2))
          }
        }
      }));
    }
  }, [formData.abdomen.measurements.waist_circumference_cm, formData.abdomen.measurements.hip_circumference_cm]);

  // AI Procedure Suggestions
  useEffect(() => {
    const suggestions = [];
    
    if (formData.abdomen.diastasis_recti && formData.abdomen.skin_laxity !== 'none') {
      suggestions.push('Abdominoplasty with Muscle Repair (15847)');
    } else if (formData.abdomen.skin_laxity === 'moderate' || formData.abdomen.skin_laxity === 'severe') {
      suggestions.push('Abdominoplasty (15847)');
    }
    
    if (formData.abdomen.pannus_present && formData.abdomen.pannus_grade !== 'I') {
      suggestions.push('Panniculectomy (15830)');
    }
    
    if (formData.buttocks.volume === 'deficient' && formData.flanks.fullness !== 'none') {
      suggestions.push('Brazilian Butt Lift (15876 + 15877)');
    }
    
    if (formData.thighs.inner_thigh_laxity === 'moderate' || formData.thighs.inner_thigh_laxity === 'severe') {
      suggestions.push('Thigh Lift (15838)');
    }
    
    if (formData.post_bariatric.is_post_bariatric && formData.post_bariatric.weight_stable) {
      suggestions.push('Body Lift (15839) - Circumferential');
      suggestions.push('Arm Lift (15842)');
    }
    
    setSuggestedProcedures(suggestions);
  }, [formData]);

  const loadExistingDocumentation = async () => {
    try {
      const { data, error } = await supabase
        .from('cr.body_part_documentation')
        .select('*')
        .eq('patient_id', patientContext.patient_id)
        .eq('encounter_id', patientContext.encounter_id)
        .eq('body_part', 'body_contour')
        .single();

      if (data && data.physical_findings) {
        setFormData(data.physical_findings);
      }
      if (data) {
        if (data.asa_class) setAsaClass(data.asa_class);
        if (data.hernia_present !== null && data.hernia_present !== undefined) setHerniaPresent(data.hernia_present);
        if (data.hernia_description) setHerniaDescription(data.hernia_description);
        if (data.keloid_history !== null && data.keloid_history !== undefined) setKeloidHistory(data.keloid_history);
        if (data.prom_score?.domains) setBodyQ(data.prom_score.domains);
      }
    } catch (err) {
      console.error('Error loading body documentation:', err);
    } finally {
      setLoading(false);
    }
  };

  const handleSave = async () => {
    setSaving(true);
    try {
      const documentationData = {
        org_id: patientContext.org_id,
        patient_id: patientContext.patient_id,
        encounter_id: patientContext.encounter_id,
        provider_id: patientContext.provider_id,
        body_part: 'body_contour',
        documentation_type: 'consultation',
        patient_sex: patientContext.patient_sex,
        physical_findings: formData,
        ai_suggested_procedures: suggestedProcedures,
        ai_suggested_cpt_codes: extractCptCodes(suggestedProcedures),
        surgical_consent_obtained: consentObtained,
        asa_class: asaClass,
        hernia_present: herniaPresent,
        hernia_description: herniaDescription || null,
        keloid_history: keloidHistory,
        prom_score: { instrument: 'BODY-Q', domains: bodyQ, date: new Date().toISOString().split('T')[0] },
        // created_at — let DB default now() set server-side timestamp
      };

      const { error } = await supabase
        .from('cr.body_part_documentation')
        .upsert(documentationData, {
          onConflict: 'patient_id,encounter_id,body_part'
        });
      
      if (error) throw error;

      setSaved(true);
      setTimeout(() => setSaved(false), 3000);
    } catch (err) {
      console.error('Error saving body documentation:', err);
      alert('Error saving. Please try again.');
    } finally {
      setSaving(false);
    }
  };

  const extractCptCodes = (procedures: string[]): string[] => {
    const codes: string[] = [];
    procedures.forEach(proc => {
      const matches = proc.match(/\d{5}/g);
      if (matches) codes.push(...matches);
    });
    return codes;
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
          <h2 className="text-2xl font-rajdhani font-bold text-white">Body Contouring Assessment</h2>
          <p className="text-gray-400 text-sm mt-1">Abdomen, buttocks, flanks, and thighs</p>
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
              Save
            </>
          )}
        </button>
      </div>

      {/* Post-Bariatric Assessment */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Post-Bariatric History</h3>
        
        <label className="flex items-center gap-3 mb-4">
          <input
            type="checkbox"
            checked={formData.post_bariatric.is_post_bariatric}
            onChange={(e) => setFormData(prev => ({
              ...prev,
              post_bariatric: { ...prev.post_bariatric, is_post_bariatric: e.target.checked }
            }))}
            className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
          />
          <span className="text-white">Post-bariatric surgery patient</span>
        </label>

        {formData.post_bariatric.is_post_bariatric && (
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                Weight Loss (kg)
              </label>
              <input
                type="number"
                value={formData.post_bariatric.weight_loss_kg || ''}
                onChange={(e) => setFormData(prev => ({
                  ...prev,
                  post_bariatric: { ...prev.post_bariatric, weight_loss_kg: parseFloat(e.target.value) || 0 }
                }))}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
              />
            </div>

            <div>
              <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                Months Since Surgery
              </label>
              <input
                type="number"
                value={formData.post_bariatric.time_since_surgery_months || ''}
                onChange={(e) => setFormData(prev => ({
                  ...prev,
                  post_bariatric: { ...prev.post_bariatric, time_since_surgery_months: parseInt(e.target.value) || 0 }
                }))}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
              />
            </div>

            <div>
              <label className="flex items-center gap-3">
                <input
                  type="checkbox"
                  checked={formData.post_bariatric.weight_stable}
                  onChange={(e) => setFormData(prev => ({
                    ...prev,
                    post_bariatric: { ...prev.post_bariatric, weight_stable: e.target.checked }
                  }))}
                  className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
                />
                <span className="text-white text-sm">Weight stable</span>
              </label>
            </div>

            {formData.post_bariatric.weight_stable && (
              <div>
                <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                  Stable Duration (months)
                </label>
                <input
                  type="number"
                  value={formData.post_bariatric.stable_duration_months || ''}
                  onChange={(e) => setFormData(prev => ({
                    ...prev,
                    post_bariatric: { ...prev.post_bariatric, stable_duration_months: parseInt(e.target.value) || 0 }
                  }))}
                  className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
                />
                {formData.post_bariatric.stable_duration_months < 6 && (
                  <p className="text-xs text-yellow-500 mt-1">
                    ⚠️ Recommend 6+ months weight stability before body contouring
                  </p>
                )}
              </div>
            )}
          </div>
        )}
      </div>

      {/* Abdomen Assessment */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Abdominal Assessment</h3>
        
        <div className="grid grid-cols-2 gap-4 mb-4">
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Skin Laxity
            </label>
            <select
              value={formData.abdomen.skin_laxity}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                abdomen: { ...prev.abdomen, skin_laxity: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="none">None</option>
              <option value="mild">Mild</option>
              <option value="moderate">Moderate</option>
              <option value="severe">Severe</option>
            </select>
          </div>

          <div>
            <label className="flex items-center gap-3 h-full items-end pb-2">
              <input
                type="checkbox"
                checked={formData.abdomen.striae}
                onChange={(e) => setFormData(prev => ({
                  ...prev,
                  abdomen: { ...prev.abdomen, striae: e.target.checked }
                }))}
                className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
              />
              <span className="text-white">Striae (stretch marks) present</span>
            </label>
          </div>
        </div>

        {/* Pannus */}
        <div className="mb-4">
          <label className="flex items-center gap-3 mb-2">
            <input
              type="checkbox"
              checked={formData.abdomen.pannus_present}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                abdomen: { ...prev.abdomen, pannus_present: e.target.checked }
              }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
            />
            <span className="text-white">Pannus present</span>
          </label>

          {formData.abdomen.pannus_present && (
            <div className="ml-7">
              <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                Pannus Grade
              </label>
              <select
                value={formData.abdomen.pannus_grade}
                onChange={(e) => setFormData(prev => ({
                  ...prev,
                  abdomen: { ...prev.abdomen, pannus_grade: e.target.value as any }
                }))}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
              >
                <option value="I">Grade I - To pubis</option>
                <option value="II">Grade II - To upper thigh</option>
                <option value="III">Grade III - To mid-thigh</option>
                <option value="IV">Grade IV - To knee or below</option>
              </select>
            </div>
          )}
        </div>

        {/* Diastasis Recti */}
        <div className="mb-4">
          <label className="flex items-center gap-3 mb-2">
            <input
              type="checkbox"
              checked={formData.abdomen.diastasis_recti}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                abdomen: { ...prev.abdomen, diastasis_recti: e.target.checked }
              }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
            />
            <span className="text-white">Diastasis recti present</span>
          </label>

          {formData.abdomen.diastasis_recti && (
            <div className="ml-7">
              <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                Diastasis Width (cm)
              </label>
              <input
                type="number"
                step="0.1"
                value={formData.abdomen.diastasis_width_cm || ''}
                onChange={(e) => setFormData(prev => ({
                  ...prev,
                  abdomen: { ...prev.abdomen, diastasis_width_cm: parseFloat(e.target.value) || 0 }
                }))}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
              />
            </div>
          )}
        </div>

        {/* Measurements */}
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Waist Circumference (cm)
            </label>
            <input
              type="number"
              step="0.1"
              value={formData.abdomen.measurements.waist_circumference_cm || ''}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                abdomen: {
                  ...prev.abdomen,
                  measurements: { ...prev.abdomen.measurements, waist_circumference_cm: parseFloat(e.target.value) || 0 }
                }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            />
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Hip Circumference (cm)
            </label>
            <input
              type="number"
              step="0.1"
              value={formData.abdomen.measurements.hip_circumference_cm || ''}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                abdomen: {
                  ...prev.abdomen,
                  measurements: { ...prev.abdomen.measurements, hip_circumference_cm: parseFloat(e.target.value) || 0 }
                }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            />
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Waist-to-Hip Ratio (auto)
            </label>
            <input
              type="text"
              value={formData.abdomen.measurements.waist_to_hip_ratio || ''}
              readOnly
              className="w-full bg-gray-800 border border-gray-700 rounded px-4 py-2 text-white cursor-not-allowed"
            />
            <p className="text-xs text-gray-500 mt-1">
              Ideal female: 0.7-0.8 · Ideal male: 0.9-1.0
            </p>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Abdominal Projection (cm)
            </label>
            <input
              type="number"
              step="0.1"
              value={formData.abdomen.measurements.abdominal_projection_cm || ''}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                abdomen: {
                  ...prev.abdomen,
                  measurements: { ...prev.abdomen.measurements, abdominal_projection_cm: parseFloat(e.target.value) || 0 }
                }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            />
          </div>
        </div>
      </div>

      {/* Buttocks Assessment */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Buttocks & Gluteal Region</h3>
        
        <div className="grid grid-cols-2 gap-4 mb-4">
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Volume
            </label>
            <select
              value={formData.buttocks.volume}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                buttocks: { ...prev.buttocks, volume: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="deficient">Deficient</option>
              <option value="adequate">Adequate</option>
              <option value="full">Full</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Shape
            </label>
            <select
              value={formData.buttocks.shape}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                buttocks: { ...prev.buttocks, shape: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="round">Round</option>
              <option value="square">Square</option>
              <option value="flat">Flat</option>
              <option value="v_shape">V-Shape</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Ptosis
            </label>
            <select
              value={formData.buttocks.ptosis}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                buttocks: { ...prev.buttocks, ptosis: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="none">None</option>
              <option value="mild">Mild</option>
              <option value="moderate">Moderate</option>
              <option value="severe">Severe</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Cellulite
            </label>
            <select
              value={formData.buttocks.cellulite}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                buttocks: { ...prev.buttocks, cellulite: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="none">None</option>
              <option value="mild">Mild</option>
              <option value="moderate">Moderate</option>
              <option value="severe">Severe</option>
            </select>
          </div>
        </div>

        <div className="flex gap-4">
          <label className="flex items-center gap-3">
            <input
              type="checkbox"
              checked={formData.buttocks.hip_dips_present}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                buttocks: { ...prev.buttocks, hip_dips_present: e.target.checked }
              }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
            />
            <span className="text-white text-sm">Hip dips present</span>
          </label>

          <label className="flex items-center gap-3">
            <input
              type="checkbox"
              checked={formData.buttocks.banana_roll}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                buttocks: { ...prev.buttocks, banana_roll: e.target.checked }
              }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
            />
            <span className="text-white text-sm">Banana roll</span>
          </label>

          <label className="flex items-center gap-3">
            <input
              type="checkbox"
              checked={formData.buttocks.asymmetry}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                buttocks: { ...prev.buttocks, asymmetry: e.target.checked }
              }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
            />
            <span className="text-white text-sm">Asymmetry</span>
          </label>
        </div>
      </div>

      {/* Flanks & Thighs */}
      <div className="grid grid-cols-2 gap-6">
        <div className="border border-gray-700 rounded-lg p-6">
          <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Flanks</h3>
          
          <div className="mb-4">
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Fullness
            </label>
            <select
              value={formData.flanks.fullness}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                flanks: { ...prev.flanks, fullness: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="none">None</option>
              <option value="mild">Mild</option>
              <option value="moderate">Moderate</option>
              <option value="severe">Severe</option>
            </select>
          </div>

          <label className="flex items-center gap-3">
            <input
              type="checkbox"
              checked={formData.flanks.love_handles}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                flanks: { ...prev.flanks, love_handles: e.target.checked }
              }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
            />
            <span className="text-white text-sm">Love handles present</span>
          </label>
        </div>

        <div className="border border-gray-700 rounded-lg p-6">
          <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Thighs</h3>
          
          <div className="mb-4">
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Inner Thigh Laxity
            </label>
            <select
              value={formData.thighs.inner_thigh_laxity}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                thighs: { ...prev.thighs, inner_thigh_laxity: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="none">None</option>
              <option value="mild">Mild</option>
              <option value="moderate">Moderate</option>
              <option value="severe">Severe</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Outer Thigh Fullness
            </label>
            <select
              value={formData.thighs.outer_thigh_fullness}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                thighs: { ...prev.thighs, outer_thigh_fullness: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="none">None</option>
              <option value="mild">Mild</option>
              <option value="moderate">Moderate</option>
            </select>
          </div>
        </div>
      </div>

      {/* AI Suggested Procedures */}
      {suggestedProcedures.length > 0 && (
        <div className="border border-[#c9a96e] bg-[#c9a96e]/10 rounded-lg p-6">
          <h3 className="text-lg font-rajdhani font-semibold text-[#c9a96e] mb-3">
            🤖 AI Suggested Procedures
          </h3>
          <div className="space-y-2">
            {suggestedProcedures.map((proc, idx) => (
              <div key={idx} className="flex items-center gap-2 text-white">
                <CheckCircle className="w-4 h-4 text-[#c9a96e]" />
                <span>{proc}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ASA Physical Status */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-3">ASA Physical Status</h3>
        <div className="flex flex-wrap gap-2">
          {(['I','II','III','IV','V'] as const).map(cls => (
            <button key={cls} onClick={() => setAsaClass(cls)}
              className={`px-4 py-2 rounded border font-rajdhani font-semibold text-sm transition-colors ${
                asaClass === cls ? 'bg-[#c9a96e] border-[#c9a96e] text-[#060e1c]' : 'border-gray-600 text-gray-300 hover:border-[#c9a96e]'
              }`}>ASA {cls}</button>
          ))}
        </div>
        <p className="text-xs text-gray-500 mt-2">
          {asaClass === 'I' && 'Normal healthy patient'}{asaClass === 'II' && 'Mild systemic disease'}
          {asaClass === 'III' && 'Severe systemic disease'}{asaClass === 'IV' && 'Severe, constant threat to life'}
          {asaClass === 'V' && 'Moribund'}
        </p>
      </div>

      {/* Hernia Assessment */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-3">Hernia Assessment</h3>
        <div className="flex gap-4 mb-3">
          {([['true','Present'],['false','Absent']] as const).map(([val,label]) => (
            <button key={val} onClick={() => setHerniaPresent(val === 'true')}
              className={`px-4 py-2 rounded border text-sm transition-colors ${
                herniaPresent === (val === 'true')
                  ? val === 'true' ? 'bg-red-800 border-red-500 text-white' : 'bg-green-800 border-green-500 text-white'
                  : 'border-gray-600 text-gray-300 hover:border-gray-400'
              }`}>{label}</button>
          ))}
        </div>
        {herniaPresent && (
          <textarea value={herniaDescription}
            onChange={e => setHerniaDescription(e.target.value)}
            placeholder="Describe type, location, size (e.g., umbilical hernia 2cm, reducible, no bowel sounds)..."
            rows={2}
            className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm" />
        )}
        {herniaPresent && <p className="text-xs text-amber-400 mt-1">Note: abdominoplasty patients with hernias require surgical consent for concurrent repair and may need general surgery consult.</p>}
      </div>

      {/* Keloid / Hypertrophic Scar History */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-3">Keloid / Hypertrophic Scar History</h3>
        <div className="flex gap-4">
          {([['true','Yes — keloid/hypertrophic history'],['false','No — normal scar formation']] as const).map(([val,label]) => (
            <button key={val} onClick={() => setKeloidHistory(val === 'true')}
              className={`px-4 py-2 rounded border text-sm transition-colors ${
                keloidHistory === (val === 'true')
                  ? val === 'true' ? 'bg-amber-800 border-amber-500 text-white' : 'bg-green-800 border-green-500 text-white'
                  : 'border-gray-600 text-gray-300 hover:border-gray-400'
              }`}>{label}</button>
          ))}
        </div>
        {keloidHistory && <p className="text-xs text-amber-400 mt-2">Counsel patient on keloid risk at incision sites. Consider silicone sheeting and steroid injection protocol post-op.</p>}
      </div>

      {/* BODY-Q Patient-Reported Outcomes */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-1">BODY-Q — Patient-Reported Outcomes</h3>
        <p className="text-xs text-gray-500 mb-4">Rate each domain 1 (Very dissatisfied) → 5 (Very satisfied).</p>
        <div className="space-y-3">
          {([
            ['body_image',               'Body image (how you feel about your body)'],
            ['weight_appearance',        'Weight / appearance satisfaction'],
            ['clothes_fit',              'Clothes fit and dressing comfort'],
            ['physical_function',        'Physical function (mobility, activity)'],
            ['psychological_wellbeing',  'Psychological wellbeing (mood, confidence)'],
            ['satisfaction_with_outcome','Overall satisfaction with surgical outcome'],
          ] as [keyof typeof bodyQ, string][]).map(([key, label]) => (
            <div key={key} className="flex items-center justify-between gap-4">
              <span className="text-sm text-gray-300 flex-1">{label}</span>
              <div className="flex gap-1">
                {[1,2,3,4,5].map(n => (
                  <button key={n} onClick={() => setBodyQ(p => ({ ...p, [key]: n }))}
                    className={`w-8 h-8 rounded text-sm font-bold transition-colors ${
                      bodyQ[key] === n ? 'bg-[#c9a96e] text-[#060e1c]' : 'bg-gray-800 text-gray-400 hover:bg-gray-700'
                    }`}>{n}</button>
                ))}
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Surgical Informed Consent */}
      <div className="border border-amber-700/50 rounded-lg p-5 bg-amber-900/10">
        <h3 className="text-base font-rajdhani font-semibold text-amber-400 mb-3">Surgical Informed Consent</h3>
        <label className="flex items-start gap-3 cursor-pointer">
          <input
            type="checkbox"
            checked={consentObtained}
            onChange={e => setConsentObtained(e.target.checked)}
            className="mt-1 w-4 h-4 accent-amber-400"
          />
          <span className="text-sm text-gray-300 leading-relaxed">
            Patient has been counseled on procedure risks, benefits, alternatives, and expected outcomes.
            Written informed consent has been obtained, signed, and placed in the chart.
          </span>
        </label>
      </div>
    </div>
  );
}
