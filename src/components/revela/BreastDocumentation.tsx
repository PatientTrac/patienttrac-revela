import React, { useState, useEffect } from 'react';
import { supabase } from '../../lib/supabase';
import { Save, Camera, AlertTriangle, CheckCircle, Ruler } from 'lucide-react';
import SurgicalDrawingTool from '../SurgicalDrawingTool';

interface PatientContext {
  patient_id: number;
  encounter_id: number;
  patient_name: string;
  patient_sex: 'male' | 'female' | 'non_binary';
  org_id: string;
  provider_id: number;
}

interface BreastMeasurements {
  sternal_notch_to_nipple_right_cm: number;
  sternal_notch_to_nipple_left_cm: number;
  nipple_to_imf_right_cm: number;
  nipple_to_imf_left_cm: number;
  breast_width_at_base_right_cm: number;
  breast_width_at_base_left_cm: number;
  anterior_pull_test_right_cm: number;
  anterior_pull_test_left_cm: number;
  imf_to_imf_cm: number;
  chest_wall_diameter_cm: number;
}

interface BreastFindings {
  symmetry: {
    symmetric: boolean;
    asymmetry_description: string;
    asymmetry_type: 'volume' | 'shape' | 'position' | 'nipple';
  };
  ptosis: {
    right_grade: 'none' | 'pseudoptosis' | 'I' | 'II' | 'III';
    left_grade: 'none' | 'pseudoptosis' | 'I' | 'II' | 'III';
    nipple_position_vs_imf: string;
  };
  skin_quality: {
    elasticity: 'excellent' | 'good' | 'moderate' | 'poor';
    striae_present: boolean;
    striae_location: string;
    skin_laxity: 'none' | 'mild' | 'moderate' | 'severe';
  };
  tissue_quality: {
    density: 'dense_fibrous' | 'mixed' | 'soft_fatty';
    upper_pole_fullness: 'excellent' | 'good' | 'adequate' | 'deficient';
    lower_pole_fullness: 'excellent' | 'good' | 'adequate' | 'deficient';
    glandular_tissue_distribution: string;
  };
  nipple_areola: {
    nipple_position: 'normal' | 'inverted' | 'high' | 'low';
    areola_diameter_right_cm: number;
    areola_diameter_left_cm: number;
    nipple_sensation: 'intact' | 'diminished' | 'absent';
    discharge: boolean;
  };
  implants_if_present?: {
    implants_present: boolean;
    type: string;
    size_cc_right: number;
    size_cc_left: number;
    position: string;
    capsular_contracture: boolean;
    baker_grade: '1' | '2' | '3' | '4';
    rippling: boolean;
    malposition: boolean;
  };
}

interface Props {
  patientContext: PatientContext;
}

export default function BreastDocumentation({ patientContext }: Props) {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  
  const [measurements, setMeasurements] = useState<BreastMeasurements>({
    sternal_notch_to_nipple_right_cm: 0,
    sternal_notch_to_nipple_left_cm: 0,
    nipple_to_imf_right_cm: 0,
    nipple_to_imf_left_cm: 0,
    breast_width_at_base_right_cm: 0,
    breast_width_at_base_left_cm: 0,
    anterior_pull_test_right_cm: 0,
    anterior_pull_test_left_cm: 0,
    imf_to_imf_cm: 0,
    chest_wall_diameter_cm: 0
  });

  const [findings, setFindings] = useState<BreastFindings>({
    symmetry: {
      symmetric: true,
      asymmetry_description: '',
      asymmetry_type: 'volume'
    },
    ptosis: {
      right_grade: 'none',
      left_grade: 'none',
      nipple_position_vs_imf: 'at or above'
    },
    skin_quality: {
      elasticity: 'good',
      striae_present: false,
      striae_location: '',
      skin_laxity: 'none'
    },
    tissue_quality: {
      density: 'mixed',
      upper_pole_fullness: 'adequate',
      lower_pole_fullness: 'adequate',
      glandular_tissue_distribution: 'normal'
    },
    nipple_areola: {
      nipple_position: 'normal',
      areola_diameter_right_cm: 0,
      areola_diameter_left_cm: 0,
      nipple_sensation: 'intact',
      discharge: false
    }
  });

  const [photoUrls, setPhotoUrls] = useState<string[]>([]);
  const [suggestedProcedures, setSuggestedProcedures] = useState<string[]>([]);
  const [gynecomastia, setGynecomastia] = useState<{ simon_grade: string; pectoral_definition: string }>({
    simon_grade: 'none',
    pectoral_definition: 'excellent',
  });
  const [consentObtained, setConsentObtained] = useState(false);
  const [asaClass, setAsaClass] = useState<'I'|'II'|'III'|'IV'|'V'|null>(null);
  const [mammogramStatus, setMammogramStatus] = useState<'current'|'overdue'|'never'|'not_indicated'|''>('');
  const [mammogramDate, setMammogramDate] = useState('');
  const [birads, setBirads] = useState<0|1|2|3|4|5|6|null>(null);
  const [priorBreastProcedures, setPriorBreastProcedures] = useState('');
  const [breastQ, setBreastQ] = useState({
    satisfaction_with_breasts: 0,
    psychosocial_wellbeing: 0,
    sexual_wellbeing: 0,
    physical_wellbeing_chest: 0,
    satisfaction_with_outcome: 0,
  });
  const [breastQInsights, setBreastQInsights] = useState<import('../../lib/revelai').PromsInsightsResult | null>(null);
  const [breastQInsightsLoading, setBreastQInsightsLoading] = useState(false);

  useEffect(() => {
    loadExistingDocumentation();
  }, [patientContext.patient_id]);

  // AI suggestion based on findings
  useEffect(() => {
    const suggestions = [];
    
    if (findings.ptosis.right_grade !== 'none' || findings.ptosis.left_grade !== 'none') {
      if (findings.tissue_quality.upper_pole_fullness === 'deficient') {
        suggestions.push('Mastopexy with Augmentation (19316 + 19325)');
      } else {
        suggestions.push('Mastopexy (19316)');
      }
    }
    
    if (findings.tissue_quality.upper_pole_fullness === 'deficient' && 
        findings.ptosis.right_grade === 'none') {
      suggestions.push('Breast Augmentation (19325)');
    }
    
    if (!findings.symmetry.symmetric) {
      suggestions.push('Breast Asymmetry Correction');
    }
    
    if (patientContext.patient_sex === 'male') {
      suggestions.push('Gynecomastia Excision (19300)');
      suggestions.push('Gynecomastia with Liposuction (19300 + 15877)');
    }
    
    setSuggestedProcedures(suggestions);
  }, [findings, patientContext.patient_sex]);

  const loadExistingDocumentation = async () => {
    try {
      const { data, error } = await supabase
        .schema('cr').from('body_part_documentation')
        .select('*')
        .eq('patient_id', patientContext.patient_id)
        .eq('encounter_id', patientContext.encounter_id)
        .eq('body_part', 'breast')
        .single();

      if (data) {
        setMeasurements(data.measurements || measurements);
        setFindings(data.physical_findings || findings);
        setPhotoUrls(data.photo_urls || []);
        if (data.asa_class) setAsaClass(data.asa_class);
        if (data.mammogram_status) setMammogramStatus(data.mammogram_status);
        if (data.mammogram_date) setMammogramDate(data.mammogram_date);
        if (data.birads_category !== null && data.birads_category !== undefined) setBirads(data.birads_category);
        if (data.prior_procedures?.breast) setPriorBreastProcedures(data.prior_procedures.breast);
        if (data.prom_score?.domains) setBreastQ(data.prom_score.domains);
      }
    } catch (err) {
      console.error('Error loading breast documentation:', err);
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
        body_part: 'breast',
        documentation_type: 'consultation',
        patient_sex: patientContext.patient_sex,
        measurements: measurements,
        physical_findings: findings,
        gynecomastia_assessment: patientContext.patient_sex === 'male' ? gynecomastia : null,
        photo_urls: photoUrls,
        photo_consent_obtained: photoUrls.length > 0,
        surgical_consent_obtained: consentObtained,
        ai_suggested_procedures: suggestedProcedures,
        ai_suggested_cpt_codes: extractCptCodes(suggestedProcedures),
        asa_class: asaClass,
        mammogram_status: mammogramStatus || null,
        mammogram_date: mammogramDate || null,
        birads_category: birads,
        prior_procedures: priorBreastProcedures ? { breast: priorBreastProcedures } : null,
        prom_score: { instrument: 'BREAST-Q', domains: breastQ, date: new Date().toISOString().split('T')[0] },
        // created_at — let DB default now() set server-side timestamp
      };

      const { error } = await supabase
        .schema('cr').from('body_part_documentation')
        .upsert(documentationData, {
          onConflict: 'patient_id,encounter_id,body_part'
        });
      
      if (error) throw error;

      setSaved(true);
      setTimeout(() => setSaved(false), 3000);
    } catch (err) {
      console.error('Error saving breast documentation:', err);
      alert('Error saving. Please try again.');
    } finally {
      setSaving(false);
    }
  };

  const handleBreastQInsights = async () => {
    setBreastQInsightsLoading(true);
    try {
      const { getPromsInsights } = await import('../../lib/revelai');
      const insights = await getPromsInsights({
        instrument: 'BREAST-Q',
        domains: breastQ,
        procedureName: 'Breast procedure',
        procedureType: 'breast',
        encounterId: patientContext.encounter_id?.toString(),
        orgId: patientContext.org_id,
      });
      setBreastQInsights(insights);
    } catch (err) {
      console.error('BREAST-Q insights error:', err);
    } finally {
      setBreastQInsightsLoading(false);
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

  const isFemale = patientContext.patient_sex === 'female';
  const isMale = patientContext.patient_sex === 'male';
  const [showDrawing, setShowDrawing] = useState(false);

  return (
    <div className="space-y-6">
      {/* Surgical Drawing Panel */}
      <div className="border border-[rgba(201,169,110,0.2)] rounded-lg overflow-hidden">
        <button
          onClick={() => setShowDrawing(v => !v)}
          className="w-full flex items-center justify-between px-5 py-3 bg-[#0a1628] hover:bg-[#0d1e36] transition-colors"
        >
          <span className="text-[#c9a96e] font-rajdhani font-semibold text-sm tracking-wide flex items-center gap-2">
            ✏ Surgical Drawings
          </span>
          <span className="text-gray-400 text-xs">{showDrawing ? '▲ Collapse' : '▼ Expand'}</span>
        </button>
        {showDrawing && (
          <SurgicalDrawingTool
            encounterId={patientContext.encounter_id.toString()}
            patientId={patientContext.patient_id.toString()}
            orgId={patientContext.org_id}
            procedureType="breast"
          />
        )}
      </div>

      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-rajdhani font-bold text-white">
            {isFemale ? 'Breast' : isMale ? 'Male Chest / Gynecomastia' : 'Chest'} Documentation
          </h2>
          <p className="text-gray-400 text-sm mt-1">
            {isFemale ? 'Female breast examination and measurements' : 
             isMale ? 'Male chest contour and gynecomastia assessment' :
             'Chest examination and measurements'}
          </p>
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

      {/* Measurements Grid */}
      <div className="border border-gray-700 rounded-lg p-6">
        <div className="flex items-center gap-2 mb-4">
          <Ruler className="w-5 h-5 text-[#c9a96e]" />
          <h3 className="text-lg font-rajdhani font-semibold text-white">
            {isFemale ? 'Breast Measurements (cm)' : 'Chest Measurements (cm)'}
          </h3>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Sternal Notch to Nipple - Right
            </label>
            <input
              type="number"
              step="0.1"
              value={measurements.sternal_notch_to_nipple_right_cm || ''}
              onChange={(e) => setMeasurements(prev => ({
                ...prev,
                sternal_notch_to_nipple_right_cm: parseFloat(e.target.value) || 0
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e]"
            />
            <p className="text-xs text-gray-500 mt-1">Typical: 19-23 cm</p>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Sternal Notch to Nipple - Left
            </label>
            <input
              type="number"
              step="0.1"
              value={measurements.sternal_notch_to_nipple_left_cm || ''}
              onChange={(e) => setMeasurements(prev => ({
                ...prev,
                sternal_notch_to_nipple_left_cm: parseFloat(e.target.value) || 0
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e]"
            />
            <p className="text-xs text-gray-500 mt-1">Typical: 19-23 cm</p>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Nipple to IMF - Right
            </label>
            <input
              type="number"
              step="0.1"
              value={measurements.nipple_to_imf_right_cm || ''}
              onChange={(e) => setMeasurements(prev => ({
                ...prev,
                nipple_to_imf_right_cm: parseFloat(e.target.value) || 0
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e]"
            />
            <p className="text-xs text-gray-500 mt-1">Typical: 5-7 cm</p>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Nipple to IMF - Left
            </label>
            <input
              type="number"
              step="0.1"
              value={measurements.nipple_to_imf_left_cm || ''}
              onChange={(e) => setMeasurements(prev => ({
                ...prev,
                nipple_to_imf_left_cm: parseFloat(e.target.value) || 0
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e]"
            />
            <p className="text-xs text-gray-500 mt-1">Typical: 5-7 cm</p>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Breast Width at Base - Right
            </label>
            <input
              type="number"
              step="0.1"
              value={measurements.breast_width_at_base_right_cm || ''}
              onChange={(e) => setMeasurements(prev => ({
                ...prev,
                breast_width_at_base_right_cm: parseFloat(e.target.value) || 0
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e]"
            />
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Breast Width at Base - Left
            </label>
            <input
              type="number"
              step="0.1"
              value={measurements.breast_width_at_base_left_cm || ''}
              onChange={(e) => setMeasurements(prev => ({
                ...prev,
                breast_width_at_base_left_cm: parseFloat(e.target.value) || 0
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e]"
            />
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              IMF to IMF Distance
            </label>
            <input
              type="number"
              step="0.1"
              value={measurements.imf_to_imf_cm || ''}
              onChange={(e) => setMeasurements(prev => ({
                ...prev,
                imf_to_imf_cm: parseFloat(e.target.value) || 0
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e]"
            />
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Chest Wall Diameter
            </label>
            <input
              type="number"
              step="0.1"
              value={measurements.chest_wall_diameter_cm || ''}
              onChange={(e) => setMeasurements(prev => ({
                ...prev,
                chest_wall_diameter_cm: parseFloat(e.target.value) || 0
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e]"
            />
          </div>
        </div>
      </div>

      {/* Physical Findings */}
      {isFemale && (
        <div className="border border-gray-700 rounded-lg p-6">
          <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Physical Examination</h3>

          {/* Symmetry */}
          <div className="mb-4">
            <label className="flex items-center gap-3 mb-2">
              <input
                type="checkbox"
                checked={findings.symmetry.symmetric}
                onChange={(e) => setFindings(prev => ({
                  ...prev,
                  symmetry: { ...prev.symmetry, symmetric: e.target.checked }
                }))}
                className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
              />
              <span className="text-white">Breasts are symmetric</span>
            </label>
            
            {!findings.symmetry.symmetric && (
              <textarea
                value={findings.symmetry.asymmetry_description}
                onChange={(e) => setFindings(prev => ({
                  ...prev,
                  symmetry: { ...prev.symmetry, asymmetry_description: e.target.value }
                }))}
                placeholder="Describe asymmetry (e.g., Left breast 1 cup size larger)"
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
                rows={2}
              />
            )}
          </div>

          {/* Ptosis Grade */}
          <div className="grid grid-cols-2 gap-4 mb-4">
            <div>
              <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                Ptosis Grade - Right
              </label>
              <select
                value={findings.ptosis.right_grade}
                onChange={(e) => setFindings(prev => ({
                  ...prev,
                  ptosis: { ...prev.ptosis, right_grade: e.target.value as any }
                }))}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
              >
                <option value="none">None</option>
                <option value="pseudoptosis">Pseudoptosis</option>
                <option value="I">Grade I</option>
                <option value="II">Grade II</option>
                <option value="III">Grade III</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                Ptosis Grade - Left
              </label>
              <select
                value={findings.ptosis.left_grade}
                onChange={(e) => setFindings(prev => ({
                  ...prev,
                  ptosis: { ...prev.ptosis, left_grade: e.target.value as any }
                }))}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
              >
                <option value="none">None</option>
                <option value="pseudoptosis">Pseudoptosis</option>
                <option value="I">Grade I</option>
                <option value="II">Grade II</option>
                <option value="III">Grade III</option>
              </select>
            </div>
          </div>

          {/* Skin Quality */}
          <div className="grid grid-cols-2 gap-4 mb-4">
            <div>
              <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                Skin Elasticity
              </label>
              <select
                value={findings.skin_quality.elasticity}
                onChange={(e) => setFindings(prev => ({
                  ...prev,
                  skin_quality: { ...prev.skin_quality, elasticity: e.target.value as any }
                }))}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
              >
                <option value="excellent">Excellent</option>
                <option value="good">Good</option>
                <option value="moderate">Moderate</option>
                <option value="poor">Poor</option>
              </select>
            </div>

            <div>
              <label className="flex items-center gap-3 h-full items-end pb-2">
                <input
                  type="checkbox"
                  checked={findings.skin_quality.striae_present}
                  onChange={(e) => setFindings(prev => ({
                    ...prev,
                    skin_quality: { ...prev.skin_quality, striae_present: e.target.checked }
                  }))}
                  className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
                />
                <span className="text-white">Striae present</span>
              </label>
            </div>
          </div>
        </div>
      )}

      {/* Male Gynecomastia Assessment */}
      {isMale && (
        <div className="border border-gray-700 rounded-lg p-6">
          <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">
            Gynecomastia Assessment
          </h3>
          
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                Simon Grade
              </label>
              <select
                value={gynecomastia.simon_grade}
                onChange={e => setGynecomastia(prev => ({ ...prev, simon_grade: e.target.value }))}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
              >
                <option value="none">No Gynecomastia</option>
                <option value="I">Grade I - Minimal</option>
                <option value="IIA">Grade IIA - Moderate, no skin excess</option>
                <option value="IIB">Grade IIB - Moderate with skin excess</option>
                <option value="III">Grade III - Marked with redundant skin</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                Pectoral Definition
              </label>
              <select
                value={gynecomastia.pectoral_definition}
                onChange={e => setGynecomastia(prev => ({ ...prev, pectoral_definition: e.target.value }))}
                className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
              >
                <option value="excellent">Excellent</option>
                <option value="good">Good</option>
                <option value="fair">Fair</option>
                <option value="poor">Poor</option>
              </select>
            </div>
          </div>
        </div>
      )}

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

      {/* ASA Physical Status */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-3">ASA Physical Status</h3>
        <div className="flex flex-wrap gap-2">
          {(['I','II','III','IV','V'] as const).map(cls => (
            <button key={cls} onClick={() => setAsaClass(cls)}
              className={`px-4 py-2 rounded border font-rajdhani font-semibold text-sm transition-colors ${
                asaClass === cls ? 'bg-[#c9a96e] border-[#c9a96e] text-[#060e1c]' : 'border-gray-600 text-gray-300 hover:border-[#c9a96e]'
              }`}>
              ASA {cls}
            </button>
          ))}
        </div>
        <p className="text-xs text-gray-500 mt-2">
          {asaClass === 'I' && 'Normal healthy patient'}
          {asaClass === 'II' && 'Mild systemic disease'}
          {asaClass === 'III' && 'Severe systemic disease'}
          {asaClass === 'IV' && 'Severe, constant threat to life'}
          {asaClass === 'V' && 'Moribund'}
        </p>
      </div>

      {/* Mammogram Status */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-3">
          Mammogram Status <span className="text-xs text-gray-500 font-normal">ACR / ASPS pre-op guidelines</span>
        </h3>
        <div className="grid grid-cols-2 gap-4 mb-3">
          <div>
            <label className="block text-xs text-[#c9a96e] mb-1">Status</label>
            <select value={mammogramStatus}
              onChange={e => setMammogramStatus(e.target.value as typeof mammogramStatus)}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm">
              <option value="">Select...</option>
              <option value="current">Current (within 12 months)</option>
              <option value="overdue">Overdue — ordered pre-op</option>
              <option value="never">Never — ordered pre-op</option>
              <option value="not_indicated">Not indicated (age &lt; 40, male)</option>
            </select>
          </div>
          <div>
            <label className="block text-xs text-[#c9a96e] mb-1">Most recent date</label>
            <input type="date" value={mammogramDate}
              onChange={e => setMammogramDate(e.target.value)}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm" />
          </div>
        </div>
        <div>
          <label className="block text-xs text-[#c9a96e] mb-2">BI-RADS Category</label>
          <div className="flex flex-wrap gap-2">
            {([0,1,2,3,4,5,6] as const).map(cat => (
              <button key={cat} onClick={() => setBirads(cat)}
                className={`px-3 py-1 rounded border text-sm font-mono transition-colors ${
                  birads === cat
                    ? cat <= 2 ? 'bg-green-700 border-green-500 text-white'
                      : cat === 3 ? 'bg-yellow-700 border-yellow-500 text-white'
                      : 'bg-red-800 border-red-500 text-white'
                    : 'border-gray-600 text-gray-400 hover:border-gray-400'
                }`}>
                {cat}
              </button>
            ))}
          </div>
          <p className="text-xs text-gray-500 mt-1">
            0=Incomplete · 1=Negative · 2=Benign · 3=Probably benign · 4=Suspicious · 5=Highly suspicious · 6=Known malignancy
          </p>
        </div>
      </div>

      {/* Prior Breast Procedures */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-3">Prior Breast Procedures</h3>
        <textarea value={priorBreastProcedures}
          onChange={e => setPriorBreastProcedures(e.target.value)}
          placeholder="e.g., Breast augmentation 2018 (400cc smooth round silicone, dual plane), biopsy right breast 2020 (benign fibroadenoma)..."
          rows={3}
          className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm" />
      </div>

      {/* BREAST-Q Patient-Reported Outcomes */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-1">BREAST-Q — Patient-Reported Outcomes</h3>
        <p className="text-xs text-gray-500 mb-4">Rate each domain 1 (Very dissatisfied) → 5 (Very satisfied). ASPS registry requirement.</p>
        <div className="space-y-3">
          {([
            ['satisfaction_with_breasts',  'Satisfaction with breasts (appearance, shape, size, feel)'],
            ['psychosocial_wellbeing',     'Psychosocial wellbeing (confidence, self-consciousness)'],
            ['sexual_wellbeing',           'Sexual wellbeing (comfortable, attractive, sexuality)'],
            ['physical_wellbeing_chest',   'Physical wellbeing — chest (tightness, discomfort, sensitivity)'],
            ['satisfaction_with_outcome',  'Overall satisfaction with outcome / would recommend'],
          ] as [keyof typeof breastQ, string][]).map(([key, label]) => (
            <div key={key} className="flex items-center justify-between gap-4">
              <span className="text-sm text-gray-300 flex-1">{label}</span>
              <div className="flex gap-1">
                {[1,2,3,4,5].map(n => (
                  <button key={n} onClick={() => setBreastQ(p => ({ ...p, [key]: n }))}
                    className={`w-8 h-8 rounded text-sm font-bold transition-colors ${
                      breastQ[key] === n ? 'bg-[#c9a96e] text-[#060e1c]' : 'bg-gray-800 text-gray-400 hover:bg-gray-700'
                    }`}>{n}</button>
                ))}
              </div>
            </div>
          ))}
        </div>
        {Object.values(breastQ).some(v => v > 0) && (
          <p className="text-xs text-gray-500 mt-3">
            Mean score: {(Object.values(breastQ).reduce((a, b) => a + b, 0) / Object.values(breastQ).filter(v => v > 0).length).toFixed(1)} / 5
          </p>
        )}
      </div>

      {/* BREAST-Q AI Insights */}
      {Object.values(breastQ).some(v => v > 0) && (
        <div className="border border-gray-700 rounded-lg p-4">
          <div className="flex items-center justify-between mb-3">
            <span className="text-sm font-rajdhani font-semibold text-white">BREAST-Q Clinical Insights</span>
            <button onClick={handleBreastQInsights} disabled={breastQInsightsLoading}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-[#c9a96e]/20 hover:bg-[#c9a96e]/30 border border-[#c9a96e]/40 text-[#c9a96e] text-xs font-rajdhani rounded transition-colors disabled:opacity-50">
              {breastQInsightsLoading ? <><div className="animate-spin h-3 w-3 border border-[#c9a96e] border-t-transparent rounded-full"/>Analyzing…</> : <>✦ Get AI Insights</>}
            </button>
          </div>
          {breastQInsights && (
            <div className="space-y-3 text-sm">
              <div className={`px-3 py-2 rounded text-xs font-medium ${
                breastQInsights.proceedWithSurgery === 'yes' ? 'bg-green-900/30 text-green-400 border border-green-700/40' :
                breastQInsights.proceedWithSurgery === 'with_caution' ? 'bg-amber-900/30 text-amber-400 border border-amber-700/40' :
                'bg-red-900/30 text-red-400 border border-red-700/40'
              }`}>
                {breastQInsights.proceedWithSurgery === 'yes' ? '✓' : breastQInsights.proceedWithSurgery === 'with_caution' ? '⚠' : '⛔'} {breastQInsights.proceedRationale}
              </div>
              <p className="text-gray-400 text-xs">{breastQInsights.overallInterpretation}</p>
              {breastQInsights.clinicalConcerns.length > 0 && (
                <div>
                  <div className="text-xs text-amber-400 font-medium mb-1">Clinical Concerns</div>
                  <ul className="space-y-1">{breastQInsights.clinicalConcerns.map((c, i) => <li key={i} className="text-xs text-gray-400">• {c}</li>)}</ul>
                </div>
              )}
              {breastQInsights.providerRecommendations.length > 0 && (
                <div>
                  <div className="text-xs text-[#c9a96e] font-medium mb-1">Recommendations</div>
                  <ul className="space-y-1">{breastQInsights.providerRecommendations.map((r, i) => <li key={i} className="text-xs text-gray-400">→ {r}</li>)}</ul>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* Photo Upload */}
      <div className="border border-gray-700 rounded-lg p-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-lg font-rajdhani font-semibold text-white">Photo Documentation</h3>
          <button className="flex items-center gap-2 px-4 py-2 bg-[#c9a96e]/20 hover:bg-[#c9a96e]/30 text-[#c9a96e] rounded transition-colors">
            <Camera className="w-4 h-4" />
            Upload Photos
          </button>
        </div>
        
        <div className="grid grid-cols-3 gap-4">
          {photoUrls.map((url, idx) => (
            <div key={idx} className="aspect-square bg-gray-800 rounded border border-gray-700">
              <img src={url} alt={`Photo ${idx + 1}`} className="w-full h-full object-cover rounded" />
            </div>
          ))}
          <div className="aspect-square bg-gray-800/50 rounded border-2 border-dashed border-gray-600 flex items-center justify-center">
            <Camera className="w-8 h-8 text-gray-600" />
          </div>
        </div>
        
        <p className="text-xs text-gray-500 mt-3">
          Recommended: Front, Lateral Right, Lateral Left, Base views
        </p>
      </div>
    </div>
  );
}
