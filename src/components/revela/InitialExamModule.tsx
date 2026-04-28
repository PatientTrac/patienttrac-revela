import React, { useState, useEffect } from 'react';
import { supabase } from '@/lib/supabase';
import { Save, AlertCircle, CheckCircle, Upload } from 'lucide-react';

interface PatientContext {
  patient_id: number;
  encounter_id: number;
  patient_name: string;
  patient_sex: 'male' | 'female' | 'non_binary';
  org_id: string;
  provider_id: number;
}

interface InitialExamData {
  consultation_id?: number;
  chief_complaint: string;
  goals_desires: string;
  medical_history: string[];
  surgical_history: string[];
  medications: string[];
  allergies: string[];
  social_history: {
    smoking_status: 'never' | 'former' | 'current';
    packs_per_day?: number;
    quit_date?: string;
    alcohol_use: 'none' | 'occasional' | 'moderate' | 'heavy';
    drug_use: string;
  };
  physical_exam: {
    height_cm: number;
    weight_kg: number;
    bmi: number;
    blood_pressure: string;
    heart_rate: number;
  };
  body_dysmorphia_screening: {
    preoccupation_with_appearance: boolean;
    significant_distress: boolean;
    impaired_functioning: boolean;
    score: number;
  };
  surgical_candidacy: 'excellent' | 'good' | 'fair' | 'poor';
  recommended_procedures: string[];
  contraindications: string[];
  photos_taken: boolean;
  photo_consent: boolean;
}

interface Props {
  patientContext: PatientContext;
}

export default function InitialExamModule({ patientContext }: Props) {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  const [formData, setFormData] = useState<InitialExamData>({
    chief_complaint: '',
    goals_desires: '',
    medical_history: [],
    surgical_history: [],
    medications: [],
    allergies: [],
    social_history: {
      smoking_status: 'never',
      alcohol_use: 'none',
      drug_use: 'none'
    },
    physical_exam: {
      height_cm: 0,
      weight_kg: 0,
      bmi: 0,
      blood_pressure: '',
      heart_rate: 0
    },
    body_dysmorphia_screening: {
      preoccupation_with_appearance: false,
      significant_distress: false,
      impaired_functioning: false,
      score: 0
    },
    surgical_candidacy: 'good',
    recommended_procedures: [],
    contraindications: [],
    photos_taken: false,
    photo_consent: false
  });

  useEffect(() => {
    loadExistingConsultation();
  }, [patientContext.patient_id]);

  useEffect(() => {
    // Auto-calculate BMI
    if (formData.physical_exam.height_cm > 0 && formData.physical_exam.weight_kg > 0) {
      const heightM = formData.physical_exam.height_cm / 100;
      const bmi = formData.physical_exam.weight_kg / (heightM * heightM);
      setFormData(prev => ({
        ...prev,
        physical_exam: {
          ...prev.physical_exam,
          bmi: parseFloat(bmi.toFixed(1))
        }
      }));
    }
  }, [formData.physical_exam.height_cm, formData.physical_exam.weight_kg]);

  const loadExistingConsultation = async () => {
    try {
      const { data, error } = await supabase
        .from('cr.surgical_consultation')
        .select('*')
        .eq('patient_id', patientContext.patient_id)
        .eq('encounter_id', patientContext.encounter_id)
        .single();

      if (data) {
        setFormData({
          consultation_id: data.consultation_id,
          chief_complaint: data.chief_complaint || '',
          goals_desires: data.goals_desires || '',
          medical_history: data.medical_history || [],
          surgical_history: data.surgical_history || [],
          medications: data.current_medications || [],
          allergies: data.allergies || [],
          social_history: data.social_history || formData.social_history,
          physical_exam: data.physical_exam || formData.physical_exam,
          body_dysmorphia_screening: data.body_dysmorphia_screening || formData.body_dysmorphia_screening,
          surgical_candidacy: data.surgical_candidacy || 'good',
          recommended_procedures: data.recommended_procedures || [],
          contraindications: data.contraindications || [],
          photos_taken: data.photos_taken || false,
          photo_consent: data.photo_consent || false
        });
      }
    } catch (err) {
      console.error('Error loading consultation:', err);
    } finally {
      setLoading(false);
    }
  };

  const handleSave = async () => {
    setSaving(true);
    try {
      const consultationData = {
        org_id: patientContext.org_id,
        patient_id: patientContext.patient_id,
        encounter_id: patientContext.encounter_id,
        provider_id: patientContext.provider_id,
        consultation_date: new Date().toISOString(),
        chief_complaint: formData.chief_complaint,
        goals_desires: formData.goals_desires,
        medical_history: formData.medical_history,
        surgical_history: formData.surgical_history,
        current_medications: formData.medications,
        allergies: formData.allergies,
        social_history: formData.social_history,
        physical_exam: formData.physical_exam,
        body_dysmorphia_screening: formData.body_dysmorphia_screening,
        surgical_candidacy: formData.surgical_candidacy,
        recommended_procedures: formData.recommended_procedures,
        contraindications: formData.contraindications,
        photos_taken: formData.photos_taken,
        photo_consent: formData.photo_consent
      };

      if (formData.consultation_id) {
        // Update existing
        const { error } = await supabase
          .from('cr.surgical_consultation')
          .update(consultationData)
          .eq('consultation_id', formData.consultation_id);
        
        if (error) throw error;
      } else {
        // Insert new
        const { data, error } = await supabase
          .from('cr.surgical_consultation')
          .insert(consultationData)
          .select('consultation_id')
          .single();
        
        if (error) throw error;
        setFormData(prev => ({ ...prev, consultation_id: data.consultation_id }));
      }

      setSaved(true);
      setTimeout(() => setSaved(false), 3000);
    } catch (err) {
      console.error('Error saving consultation:', err);
      alert('Error saving consultation. Please try again.');
    } finally {
      setSaving(false);
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
        <h2 className="text-2xl font-rajdhani font-bold text-white">Initial Consultation</h2>
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

      {/* Chief Complaint */}
      <div>
        <label className="block text-sm font-rajdhani font-semibold text-[#c9a96e] mb-2">
          Chief Complaint *
        </label>
        <textarea
          value={formData.chief_complaint}
          onChange={(e) => setFormData(prev => ({ ...prev, chief_complaint: e.target.value }))}
          placeholder="What brought the patient in today?"
          className="w-full bg-[#060e1c] border border-gray-700 rounded-lg px-4 py-3 text-white placeholder-gray-500 focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e] transition-colors"
          rows={3}
        />
      </div>

      {/* Patient Goals */}
      <div>
        <label className="block text-sm font-rajdhani font-semibold text-[#c9a96e] mb-2">
          Patient Goals & Desires *
        </label>
        <textarea
          value={formData.goals_desires}
          onChange={(e) => setFormData(prev => ({ ...prev, goals_desires: e.target.value }))}
          placeholder="What does the patient want to achieve? What are their expectations?"
          className="w-full bg-[#060e1c] border border-gray-700 rounded-lg px-4 py-3 text-white placeholder-gray-500 focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e] transition-colors"
          rows={3}
        />
      </div>

      {/* Physical Exam */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div>
          <label className="block text-sm font-rajdhani font-semibold text-[#c9a96e] mb-2">
            Height (cm) *
          </label>
          <input
            type="number"
            value={formData.physical_exam.height_cm || ''}
            onChange={(e) => setFormData(prev => ({
              ...prev,
              physical_exam: { ...prev.physical_exam, height_cm: parseFloat(e.target.value) || 0 }
            }))}
            className="w-full bg-[#060e1c] border border-gray-700 rounded-lg px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e] transition-colors"
          />
        </div>

        <div>
          <label className="block text-sm font-rajdhani font-semibold text-[#c9a96e] mb-2">
            Weight (kg) *
          </label>
          <input
            type="number"
            value={formData.physical_exam.weight_kg || ''}
            onChange={(e) => setFormData(prev => ({
              ...prev,
              physical_exam: { ...prev.physical_exam, weight_kg: parseFloat(e.target.value) || 0 }
            }))}
            className="w-full bg-[#060e1c] border border-gray-700 rounded-lg px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e] transition-colors"
          />
        </div>

        <div>
          <label className="block text-sm font-rajdhani font-semibold text-[#c9a96e] mb-2">
            BMI (auto-calculated)
          </label>
          <input
            type="text"
            value={formData.physical_exam.bmi || ''}
            readOnly
            className="w-full bg-gray-800 border border-gray-700 rounded-lg px-4 py-2 text-white cursor-not-allowed"
          />
          {formData.physical_exam.bmi > 40 && (
            <p className="text-xs text-yellow-500 mt-1">
              ⚠️ BMI &gt;40 may increase surgical risks
            </p>
          )}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-rajdhani font-semibold text-[#c9a96e] mb-2">
            Blood Pressure
          </label>
          <input
            type="text"
            value={formData.physical_exam.blood_pressure}
            onChange={(e) => setFormData(prev => ({
              ...prev,
              physical_exam: { ...prev.physical_exam, blood_pressure: e.target.value }
            }))}
            placeholder="120/80"
            className="w-full bg-[#060e1c] border border-gray-700 rounded-lg px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e] transition-colors"
          />
        </div>

        <div>
          <label className="block text-sm font-rajdhani font-semibold text-[#c9a96e] mb-2">
            Heart Rate (bpm)
          </label>
          <input
            type="number"
            value={formData.physical_exam.heart_rate || ''}
            onChange={(e) => setFormData(prev => ({
              ...prev,
              physical_exam: { ...prev.physical_exam, heart_rate: parseInt(e.target.value) || 0 }
            }))}
            className="w-full bg-[#060e1c] border border-gray-700 rounded-lg px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e] transition-colors"
          />
        </div>
      </div>

      {/* Social History */}
      <div className="border border-gray-700 rounded-lg p-4">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Social History</h3>
        
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-rajdhani font-semibold text-[#c9a96e] mb-2">
              Smoking Status *
            </label>
            <select
              value={formData.social_history.smoking_status}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                social_history: { ...prev.social_history, smoking_status: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded-lg px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e] transition-colors"
            >
              <option value="never">Never</option>
              <option value="former">Former Smoker</option>
              <option value="current">Current Smoker</option>
            </select>
            {formData.social_history.smoking_status === 'current' && (
              <p className="text-xs text-red-500 mt-1">
                ⚠️ Active smoking contraindicated for abdominoplasty and some facial procedures
              </p>
            )}
          </div>

          <div>
            <label className="block text-sm font-rajdhani font-semibold text-[#c9a96e] mb-2">
              Alcohol Use
            </label>
            <select
              value={formData.social_history.alcohol_use}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                social_history: { ...prev.social_history, alcohol_use: e.target.value as any }
              }))}
              className="w-full bg-[#060e1c] border border-gray-700 rounded-lg px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e] transition-colors"
            >
              <option value="none">None</option>
              <option value="occasional">Occasional</option>
              <option value="moderate">Moderate</option>
              <option value="heavy">Heavy</option>
            </select>
          </div>
        </div>
      </div>

      {/* Body Dysmorphia Screening */}
      <div className="border border-gray-700 rounded-lg p-4">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">
          Body Dysmorphia Screening
        </h3>
        
        <div className="space-y-3">
          <label className="flex items-center gap-3">
            <input
              type="checkbox"
              checked={formData.body_dysmorphia_screening.preoccupation_with_appearance}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                body_dysmorphia_screening: {
                  ...prev.body_dysmorphia_screening,
                  preoccupation_with_appearance: e.target.checked,
                  score: (e.target.checked ? 1 : 0) + 
                         (prev.body_dysmorphia_screening.significant_distress ? 1 : 0) +
                         (prev.body_dysmorphia_screening.impaired_functioning ? 1 : 0)
                }
              }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded focus:ring-[#c9a96e]"
            />
            <span className="text-gray-300 text-sm">
              Excessive preoccupation with perceived appearance flaw
            </span>
          </label>

          <label className="flex items-center gap-3">
            <input
              type="checkbox"
              checked={formData.body_dysmorphia_screening.significant_distress}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                body_dysmorphia_screening: {
                  ...prev.body_dysmorphia_screening,
                  significant_distress: e.target.checked,
                  score: (prev.body_dysmorphia_screening.preoccupation_with_appearance ? 1 : 0) +
                         (e.target.checked ? 1 : 0) +
                         (prev.body_dysmorphia_screening.impaired_functioning ? 1 : 0)
                }
              }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded focus:ring-[#c9a96e]"
            />
            <span className="text-gray-300 text-sm">
              Causes clinically significant distress or impairment
            </span>
          </label>

          <label className="flex items-center gap-3">
            <input
              type="checkbox"
              checked={formData.body_dysmorphia_screening.impaired_functioning}
              onChange={(e) => setFormData(prev => ({
                ...prev,
                body_dysmorphia_screening: {
                  ...prev.body_dysmorphia_screening,
                  impaired_functioning: e.target.checked,
                  score: (prev.body_dysmorphia_screening.preoccupation_with_appearance ? 1 : 0) +
                         (prev.body_dysmorphia_screening.significant_distress ? 1 : 0) +
                         (e.target.checked ? 1 : 0)
                }
              }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded focus:ring-[#c9a96e]"
            />
            <span className="text-gray-300 text-sm">
              Interferes with social, occupational, or other functioning
            </span>
          </label>
        </div>

        {formData.body_dysmorphia_screening.score >= 2 && (
          <div className="mt-4 p-3 bg-yellow-900/20 border border-yellow-500/50 rounded">
            <AlertCircle className="w-5 h-5 text-yellow-500 inline mr-2" />
            <span className="text-yellow-500 text-sm font-semibold">
              Consider psychology evaluation before proceeding with surgery
            </span>
          </div>
        )}
      </div>

      {/* Photo Documentation */}
      <div className="border border-gray-700 rounded-lg p-4">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">Photo Documentation</h3>
        
        <div className="space-y-3">
          <label className="flex items-center gap-3">
            <input
              type="checkbox"
              checked={formData.photo_consent}
              onChange={(e) => setFormData(prev => ({ ...prev, photo_consent: e.target.checked }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded focus:ring-[#c9a96e]"
            />
            <span className="text-gray-300 text-sm">
              Photo consent obtained
            </span>
          </label>

          <label className="flex items-center gap-3">
            <input
              type="checkbox"
              checked={formData.photos_taken}
              onChange={(e) => setFormData(prev => ({ ...prev, photos_taken: e.target.checked }))}
              className="w-4 h-4 text-[#c9a96e] bg-[#060e1c] border-gray-700 rounded focus:ring-[#c9a96e]"
            />
            <span className="text-gray-300 text-sm">
              Pre-operative photos taken
            </span>
          </label>

          <button className="flex items-center gap-2 px-4 py-2 bg-[#c9a96e]/20 hover:bg-[#c9a96e]/30 text-[#c9a96e] rounded transition-colors">
            <Upload className="w-4 h-4" />
            Upload Photos
          </button>
        </div>
      </div>

      {/* Surgical Candidacy */}
      <div>
        <label className="block text-sm font-rajdhani font-semibold text-[#c9a96e] mb-2">
          Surgical Candidacy Assessment
        </label>
        <select
          value={formData.surgical_candidacy}
          onChange={(e) => setFormData(prev => ({ ...prev, surgical_candidacy: e.target.value as any }))}
          className="w-full bg-[#060e1c] border border-gray-700 rounded-lg px-4 py-2 text-white focus:border-[#c9a96e] focus:ring-1 focus:ring-[#c9a96e] transition-colors"
        >
          <option value="excellent">Excellent - No concerns</option>
          <option value="good">Good - Minor concerns managed</option>
          <option value="fair">Fair - Moderate concerns, proceed with caution</option>
          <option value="poor">Poor - Significant concerns, not recommended</option>
        </select>
      </div>
    </div>
  );
}
