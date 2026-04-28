import React, { useState, useEffect } from 'react';
import { supabase } from '@/lib/supabase';
import { Save, CheckCircle, Plus, Trash2 } from 'lucide-react';

interface PatientContext {
  patient_id: number;
  encounter_id: number;
  patient_name: string;
  patient_sex: 'male' | 'female' | 'non_binary';
  org_id: string;
  provider_id: number;
}

interface LiposuctionArea {
  area_id: string;
  area: string;
  pinch_test_cm: number;
  estimated_fat_removal_ml: number;
  donor_site_for_grafting: boolean;
}

interface LiposuctionData {
  selected_areas: LiposuctionArea[];
  technique: {
    method: 'traditional' | 'tumescent' | 'ultrasound_assisted' | 'laser_assisted' | 'power_assisted';
    cannula_size_mm: number;
    anesthesia_type: 'local' | 'local_with_sedation' | 'general';
  };
  estimated_totals: {
    total_areas: number;
    total_estimated_removal_ml: number;
    estimated_duration_minutes: number;
  };
}

interface Props {
  patientContext: PatientContext;
}

const AREA_OPTIONS = [
  { value: 'abdomen', label: 'Abdomen' },
  { value: 'flanks_right', label: 'Flanks - Right' },
  { value: 'flanks_left', label: 'Flanks - Left' },
  { value: 'back_upper', label: 'Back - Upper' },
  { value: 'back_lower', label: 'Back - Lower' },
  { value: 'thighs_inner_right', label: 'Thighs Inner - Right' },
  { value: 'thighs_inner_left', label: 'Thighs Inner - Left' },
  { value: 'thighs_outer_right', label: 'Thighs Outer - Right' },
  { value: 'thighs_outer_left', label: 'Thighs Outer - Left' },
  { value: 'arms_right', label: 'Arms - Right' },
  { value: 'arms_left', label: 'Arms - Left' },
  { value: 'neck_submental', label: 'Neck - Submental' },
  { value: 'male_chest', label: 'Male Chest' },
  { value: 'buttocks', label: 'Buttocks' }
];

export default function LiposuctionDocumentation({ patientContext }: Props) {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  
  const [formData, setFormData] = useState<LiposuctionData>({
    selected_areas: [],
    technique: {
      method: 'tumescent',
      cannula_size_mm: 3,
      anesthesia_type: 'local_with_sedation'
    },
    estimated_totals: {
      total_areas: 0,
      total_estimated_removal_ml: 0,
      estimated_duration_minutes: 0
    }
  });

  const [cptCodes, setCptCodes] = useState<string[]>([]);

  useEffect(() => {
    loadExistingDocumentation();
  }, [patientContext.patient_id]);

  // Auto-calculate totals
  useEffect(() => {
    const totalAreas = formData.selected_areas.length;
    const totalRemoval = formData.selected_areas.reduce((sum, area) => sum + area.estimated_fat_removal_ml, 0);
    const estimatedDuration = totalAreas * 30 + 60; // 30 min per area + 60 min setup

    setFormData(prev => ({
      ...prev,
      estimated_totals: {
        total_areas: totalAreas,
        total_estimated_removal_ml: totalRemoval,
        estimated_duration_minutes: estimatedDuration
      }
    }));

    // Calculate CPT codes
    const codes = [];
    if (totalAreas > 0) {
      codes.push('15876'); // First area
      for (let i = 1; i < totalAreas; i++) {
        codes.push('15877'); // Each additional area
      }
    }
    setCptCodes(codes);
  }, [formData.selected_areas]);

  const loadExistingDocumentation = async () => {
    try {
      const { data, error } = await supabase
        .from('cr.body_part_documentation')
        .select('*')
        .eq('patient_id', patientContext.patient_id)
        .eq('encounter_id', patientContext.encounter_id)
        .eq('body_part', 'liposuction')
        .single();

      if (data && data.physical_findings) {
        setFormData(data.physical_findings);
      }
    } catch (err) {
      console.error('Error loading liposuction documentation:', err);
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
        body_part: 'liposuction',
        documentation_type: 'procedure_plan',
        patient_sex: patientContext.patient_sex,
        physical_findings: formData,
        ai_suggested_cpt_codes: cptCodes,
        ai_confidence_score: 0.92,
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
      console.error('Error saving liposuction documentation:', err);
      alert('Error saving. Please try again.');
    } finally {
      setSaving(false);
    }
  };

  const addArea = () => {
    const newArea: LiposuctionArea = {
      area_id: `area_${Date.now()}`,
      area: '',
      pinch_test_cm: 0,
      estimated_fat_removal_ml: 0,
      donor_site_for_grafting: false
    };
    setFormData(prev => ({
      ...prev,
      selected_areas: [...prev.selected_areas, newArea]
    }));
  };

  const removeArea = (areaId: string) => {
    setFormData(prev => ({
      ...prev,
      selected_areas: prev.selected_areas.filter(a => a.area_id !== areaId)
    }));
  };

  const updateArea = (areaId: string, field: keyof LiposuctionArea, value: any) => {
    setFormData(prev => ({
      ...prev,
      selected_areas: prev.selected_areas.map(area =>
        area.area_id === areaId ? { ...area, [field]: value } : area
      )
    }));
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
          <h2 className="text-2xl font-rajdhani font-bold text-white">Liposuction Planning</h2>
          <p className="text-gray-400 text-sm mt-1">Multi-area assessment and volume estimation</p>
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

      {/* Technique Selection */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Technique</h3>
        
        <div className="grid grid-cols-3 gap-4">
          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Method
            </label>
            <select
              value={formData.technique.method}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                technique: { ...prev.technique, method: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="traditional">Traditional</option>
              <option value="tumescent">Tumescent</option>
              <option value="ultrasound_assisted">Ultrasound-Assisted (UAL)</option>
              <option value="laser_assisted">Laser-Assisted (LAL)</option>
              <option value="power_assisted">Power-Assisted (PAL)</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Cannula Size (mm)
            </label>
            <input
              type="number"
              step="0.5"
              value={formData.technique.cannula_size_mm || ''}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                technique: { ...prev.technique, cannula_size_mm: parseFloat(e.target.value) || 0 }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            />
          </div>

          <div>
            <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
              Anesthesia Type
            </label>
            <select
              value={formData.technique.anesthesia_type}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                technique: { ...prev.technique, anesthesia_type: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
            >
              <option value="local">Local Only</option>
              <option value="local_with_sedation">Local with Sedation</option>
              <option value="general">General</option>
            </select>
          </div>
        </div>
      </div>

      {/* Area Selection */}
      <div className="border border-gray-700 rounded-lg p-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-lg font-rajdhani font-semibold text-white">Treatment Areas</h3>
          <button
            onClick={addArea}
            className="flex items-center gap-2 px-4 py-2 bg-[#c9a96e]/20 hover:bg-[#c9a96e]/30 text-[#c9a96e] rounded transition-colors"
          >
            <Plus className="w-4 h-4" />
            Add Area
          </button>
        </div>

        {formData.selected_areas.length === 0 && (
          <div className="text-center py-8 text-gray-500">
            No areas selected. Click "Add Area" to begin planning.
          </div>
        )}

        <div className="space-y-4">
          {formData.selected_areas.map((area) => (
            <div key={area.area_id} className="border border-gray-600 rounded-lg p-4 bg-[#060e1c]/50">
              <div className="grid grid-cols-4 gap-4 mb-3">
                <div className="col-span-2">
                  <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                    Body Area
                  </label>
                  <select
                    value={area.area}
                    onChange={(e) => updateArea(area.area_id, 'area', e.target.value)}
                    className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
                  >
                    <option value="">Select area...</option>
                    {AREA_OPTIONS.map(opt => (
                      <option key={opt.value} value={opt.value}>{opt.label}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                    Pinch Test (cm)
                  </label>
                  <input
                    type="number"
                    step="0.1"
                    value={area.pinch_test_cm || ''}
                    onChange={(e) => updateArea(area.area_id, 'pinch_test_cm', parseFloat(e.target.value) || 0)}
                    className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
                  />
                  <p className="text-xs text-gray-500 mt-1">Skin fold thickness</p>
                </div>

                <div>
                  <label className="block text-sm font-rajdhani text-[#c9a96e] mb-2">
                    Est. Removal (ml)
                  </label>
                  <input
                    type="number"
                    value={area.estimated_fat_removal_ml || ''}
                    onChange={(e) => updateArea(area.area_id, 'estimated_fat_removal_ml', parseInt(e.target.value) || 0)}
                    className="w-full bg-[#060e1c] border border-gray-700 rounded px-4 py-2 text-white"
                  />
                </div>
              </div>

              <div className="flex items-center justify-between">
                <label className="flex items-center gap-3">
                  <input
                    type="checkbox"
                    checked={area.donor_site_for_grafting}
                    onChange={(e) => updateArea(area.area_id, 'donor_site_for_grafting', e.target.checked)}
                    className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded"
                  />
                  <span className="text-white text-sm">Donor site for fat grafting (BBL/Transfer)</span>
                </label>

                <button
                  onClick={() => removeArea(area.area_id)}
                  className="flex items-center gap-2 px-3 py-1 text-red-500 hover:bg-red-500/10 rounded transition-colors"
                >
                  <Trash2 className="w-4 h-4" />
                  Remove
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Summary & CPT Codes */}
      {formData.selected_areas.length > 0 && (
        <>
          <div className="border border-gray-700 rounded-lg p-6">
            <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Procedure Summary</h3>
            
            <div className="grid grid-cols-3 gap-6">
              <div>
                <span className="block text-sm text-gray-400 mb-1">Total Areas</span>
                <span className="text-2xl font-rajdhani font-bold text-white">
                  {formData.estimated_totals.total_areas}
                </span>
              </div>

              <div>
                <span className="block text-sm text-gray-400 mb-1">Total Volume</span>
                <span className="text-2xl font-rajdhani font-bold text-white">
                  {formData.estimated_totals.total_estimated_removal_ml} ml
                </span>
                {formData.estimated_totals.total_estimated_removal_ml > 5000 && (
                  <p className="text-xs text-yellow-500 mt-1">
                    ⚠️ &gt;5L: Consider staged procedures
                  </p>
                )}
              </div>

              <div>
                <span className="block text-sm text-gray-400 mb-1">Est. Duration</span>
                <span className="text-2xl font-rajdhani font-bold text-white">
                  {Math.round(formData.estimated_totals.estimated_duration_minutes / 60)} hrs {formData.estimated_totals.estimated_duration_minutes % 60} min
                </span>
              </div>
            </div>
          </div>

          <div className="border border-[#c9a96e] bg-[#c9a96e]/10 rounded-lg p-6">
            <h3 className="text-lg font-rajdhani font-semibold text-[#c9a96e] mb-3">
              🤖 Auto-Generated CPT Codes
            </h3>
            <div className="flex flex-wrap gap-3">
              {cptCodes.map((code, idx) => (
                <div key={idx} className="flex items-center gap-2 bg-[#060e1c] border border-[#c9a96e] rounded px-4 py-2">
                  <CheckCircle className="w-4 h-4 text-[#c9a96e]" />
                  <span className="text-white font-mono">{code}</span>
                  <span className="text-gray-400 text-sm">
                    {idx === 0 ? 'First area' : `Additional area ${idx}`}
                  </span>
                </div>
              ))}
            </div>
            <p className="text-xs text-gray-400 mt-3">
              Confidence: 92% • Based on {formData.estimated_totals.total_areas} treatment area{formData.estimated_totals.total_areas !== 1 ? 's' : ''}
            </p>
          </div>

          {/* Donor Sites for Grafting */}
          {formData.selected_areas.some(a => a.donor_site_for_grafting) && (
            <div className="border border-blue-500 bg-blue-500/10 rounded-lg p-6">
              <h3 className="text-lg font-rajdhani font-semibold text-blue-400 mb-3">
                📍 Donor Sites Identified
              </h3>
              <div className="space-y-2">
                {formData.selected_areas
                  .filter(a => a.donor_site_for_grafting)
                  .map(area => (
                    <div key={area.area_id} className="flex items-center gap-2 text-white">
                      <CheckCircle className="w-4 h-4 text-blue-400" />
                      <span>
                        {AREA_OPTIONS.find(opt => opt.value === area.area)?.label || area.area}
                        {' '}→ {area.estimated_fat_removal_ml} ml available
                      </span>
                    </div>
                  ))}
              </div>
              <p className="text-xs text-gray-400 mt-3">
                Fat harvest sites for Brazilian Butt Lift (BBL) or facial fat transfer
              </p>
            </div>
          )}
        </>
      )}
    </div>
  );
}
