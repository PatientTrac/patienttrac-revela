import React, { useState, useEffect } from 'react';
import { supabase } from '@/lib/supabase';
import { Save, Camera, AlertTriangle, CheckCircle, Ruler } from 'lucide-react';

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
        .from('cr.body_part_documentation')
        .select('*')
        .eq('patient_id', patientContext.patient_id)
        .eq('encounter_id', patientContext.encounter_id)
        .eq('body_part', 'breast')
        .single();

      if (data) {
        setMeasurements(data.measurements || measurements);
        setFindings(data.physical_findings || findings);
        setPhotoUrls(data.photo_urls || []);
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
        photo_urls: photoUrls,
        photo_consent_obtained: photoUrls.length > 0,
        ai_suggested_procedures: suggestedProcedures,
        ai_suggested_cpt_codes: extractCptCodes(suggestedProcedures),
        ai_confidence_score: 0.85,
        created_at: new Date().toISOString()
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
      console.error('Error saving breast documentation:', err);
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

  const isFemale = patientContext.patient_sex === 'female';
  const isMale = patientContext.patient_sex === 'male';

  return (
    <div className="space-y-6">
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
              <select className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white">
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
              <select className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white">
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
          <p className="text-xs text-gray-400 mt-3">
            Confidence: 85% • Based on physical examination findings
          </p>
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
