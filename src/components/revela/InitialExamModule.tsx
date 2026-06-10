import React, { useState, useEffect } from 'react';
import { supabase } from '../../lib/supabase';
import { Save, AlertCircle, CheckCircle, Upload } from 'lucide-react';
import { generateClinicalFlags, calculateRiskScore, determineORClearance, type PatientRiskInput } from '../../lib/riskEngine';

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
  const [asaClass, setAsaClass] = useState<'I'|'II'|'III'|'IV'|'V'|null>(null);
  const [smokingCessation, setSmokingCessation] = useState({
    counseled: false,
    referral_given: false,
    quit_date_set: false,
  });
  const [lmpDate, setLmpDate] = useState('');
  const [pregnancyTest, setPregnancyTest] = useState<'positive'|'negative'|'declined'|'not_applicable'|''>('');
  const [psychClearance, setPsychClearance] = useState<'cleared'|'deferred'|'referred'|'not_required'|''>('');
  const [psychNotes, setPsychNotes] = useState('');
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
        if (data.asa_class) setAsaClass(data.asa_class);
        if (data.smoking_cessation_counseled) setSmokingCessation((prev: typeof smokingCessation) => ({ ...prev, counseled: data.smoking_cessation_counseled }));
        if (data.lmp_date) setLmpDate(data.lmp_date);
        if (data.pregnancy_test_result) setPregnancyTest(data.pregnancy_test_result);
        if (data.psych_clearance) setPsychClearance(data.psych_clearance);
        if (data.psych_clearance_notes) setPsychNotes(data.psych_clearance_notes);
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
        photo_consent: formData.photo_consent,
        asa_class: asaClass,
        smoking_cessation_counseled: smokingCessation.counseled,
        lmp_date: lmpDate || null,
        pregnancy_test_result: pregnancyTest || null,
        psych_clearance: psychClearance || null,
        psych_clearance_notes: psychNotes || null,
      };

      // ── Risk engine computation ────────────────────────────────
      const bpParts = formData.physical_exam.blood_pressure.split('/').map(Number)
      const riskInput: PatientRiskInput = {
        bpSystolic: bpParts[0] || undefined,
        bpDiastolic: bpParts[1] || undefined,
        bmi: formData.physical_exam.bmi || undefined,
        medications: formData.medications,
        smokingStatus: formData.social_history.smoking_status,
        alcoholUse: formData.social_history.alcohol_use === 'occasional' ? 'light' : formData.social_history.alcohol_use,
        priorAbdominalSurgeries: formData.surgical_history,
        psychiatricStability: formData.body_dysmorphia_screening.significant_distress ? 'unstable' : 'stable',
      }
      const riskFlags = generateClinicalFlags(riskInput)
      const riskScore = calculateRiskScore(riskInput, riskFlags)
      const orClearance = determineORClearance(riskFlags, riskScore)

      const consultationWithRisk = {
        ...consultationData,
        risk_score: riskScore,
        or_clearance: orClearance,
        clinical_flags: riskFlags,
      }

      if (formData.consultation_id) {
        // Update existing
        const { error } = await supabase
          .from('cr.surgical_consultation')
          .update(consultationWithRisk)
          .eq('consultation_id', formData.consultation_id);

        if (error) throw error;
      } else {
        // Insert new
        const { data, error } = await supabase
          .from('cr.surgical_consultation')
          .insert(consultationWithRisk)
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

      {/* ASA Physical Status */}
      <div className="border border-gray-700 rounded-lg p-6">
        <h3 className="text-lg font-rajdhani font-semibold text-white mb-4">ASA Physical Status Classification</h3>
        <div className="grid grid-cols-1 gap-2">
          {([
            ['I',  'Normal healthy patient. No organic, physiologic, or psychiatric disturbance.'],
            ['II', 'Mild systemic disease without substantive functional limitations (e.g., controlled DM/HTN, smoker, BMI 30-40).'],
            ['III','Severe systemic disease with substantive functional limitations (e.g., poorly controlled DM/HTN, COPD, morbid obesity ≥40 BMI, EF < 40%).'],
            ['IV', 'Severe systemic disease that is a constant threat to life (e.g., recent MI/CVA/TIA, severe valvular dysfunction, sepsis).'],
            ['V',  'Moribund patient not expected to survive without the operation.'],
          ] as [string,string][]).map(([cls, desc]) => (
            <label key={cls} className={`flex items-start gap-3 p-3 rounded-lg cursor-pointer border transition-colors ${
              asaClass === cls ? 'border-[#c9a96e] bg-[#c9a96e]/10' : 'border-gray-700 hover:border-gray-600'
            }`}>
              <input type="radio" name="asa_class" value={cls}
                checked={asaClass === cls}
                onChange={() => setAsaClass(cls as 'I'|'II'|'III'|'IV'|'V')}
                className="mt-1 accent-[#c9a96e]" />
              <div>
                <span className="text-[#c9a96e] font-rajdhani font-bold">ASA {cls}</span>
                <span className="text-gray-300 text-sm ml-2">{desc}</span>
              </div>
            </label>
          ))}
        </div>
      </div>

      {/* Smoking Cessation — show only for current/former smokers */}
      {(formData.social_history?.smoking_status === 'current' || formData.social_history?.smoking_status === 'former') && (
        <div className="border border-amber-700/40 rounded-lg p-5 bg-amber-900/10">
          <h3 className="text-base font-rajdhani font-semibold text-amber-400 mb-3">Smoking Cessation Counseling <span className="text-xs font-normal text-gray-500">CMS Tobacco Use / SCIP</span></h3>
          <div className="space-y-2">
            {([
              ['counseled', 'Cessation counseling provided (risks of smoking on wound healing, anesthesia, outcomes)'],
              ['referral_given', 'Referral to cessation program given (nicotine replacement, varenicline, bupropion)'],
              ['quit_date_set', 'Quit date established (minimum 4 weeks pre-op recommended)'],
            ] as [keyof typeof smokingCessation, string][]).map(([key, label]) => (
              <label key={key} className="flex items-center gap-3 cursor-pointer text-sm text-gray-300">
                <input type="checkbox"
                  checked={smokingCessation[key]}
                  onChange={e => setSmokingCessation(p => ({ ...p, [key]: e.target.checked }))}
                  className="w-4 h-4 accent-amber-400" />
                {label}
              </label>
            ))}
          </div>
        </div>
      )}

      {/* Pregnancy Screening — female / non-binary patients */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-3">Pregnancy Screening</h3>
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="block text-xs text-[#c9a96e] mb-1">Last Menstrual Period</label>
            <input type="date" value={lmpDate}
              onChange={e => setLmpDate(e.target.value)}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm" />
          </div>
          <div>
            <label className="block text-xs text-[#c9a96e] mb-1">Pregnancy Test Result</label>
            <select value={pregnancyTest}
              onChange={e => setPregnancyTest(e.target.value as typeof pregnancyTest)}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm">
              <option value="">Select...</option>
              <option value="negative">Negative</option>
              <option value="positive">Positive — procedure deferred</option>
              <option value="declined">Patient declined testing</option>
              <option value="not_applicable">Not applicable (post-menopausal / surgical)</option>
            </select>
          </div>
        </div>
      </div>

      {/* Psychological / Psychiatric Clearance */}
      <div className="border border-gray-700 rounded-lg p-5">
        <h3 className="text-base font-rajdhani font-semibold text-white mb-3">Psychological Clearance</h3>
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="block text-xs text-[#c9a96e] mb-1">Clearance Status</label>
            <select value={psychClearance}
              onChange={e => setPsychClearance(e.target.value as typeof psychClearance)}
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm">
              <option value="">Select...</option>
              <option value="not_required">Not required (BDD score &lt; 2, no concerns)</option>
              <option value="cleared">Cleared by surgeon / evaluator</option>
              <option value="deferred">Deferred — additional evaluation needed</option>
              <option value="referred">Referred to mental health professional</option>
            </select>
          </div>
          <div>
            <label className="block text-xs text-[#c9a96e] mb-1">Notes</label>
            <input type="text" value={psychNotes}
              onChange={e => setPsychNotes(e.target.value)}
              placeholder="e.g., BDD concerns, unrealistic expectations..."
              className="w-full bg-[#060e1c] border border-gray-700 rounded px-3 py-2 text-white text-sm" />
          </div>
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
