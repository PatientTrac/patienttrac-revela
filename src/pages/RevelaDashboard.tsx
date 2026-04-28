import React, { useState, useEffect } from 'react';
import { useSearchParams, useNavigate } from 'react-router-dom';
import { supabase } from '@/lib/supabase';
import { User, Calendar, AlertCircle, CheckCircle } from 'lucide-react';

// Import body region modules
import InitialExamModule from '@/components/revela/InitialExamModule';
import BreastDocumentation from '@/components/revela/BreastDocumentation';
import BodyDocumentation from '@/components/revela/BodyDocumentation';
import LiposuctionDocumentation from '@/components/revela/LiposuctionDocumentation';
import FaceDocumentation from '@/components/revela/FaceDocumentation';
import BasicOperativeNote from '@/components/revela/BasicOperativeNote';

interface PatientContext {
  patient_id: number;
  encounter_id: number;
  patient_name: string;
  patient_sex: 'male' | 'female' | 'non_binary';
  patient_dob: string;
  patient_photo_url?: string;
  chief_complaint?: string;
  provider_id: number;
  provider_name: string;
  org_id: string;
}

type BodyRegion = 'initial_exam' | 'breast' | 'body' | 'liposuction' | 'face' | 'operative_note';

export default function RevelaDashboard() {
  const [searchParams] = useSearchParams();
  const navigate = useNavigate();
  
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [patientContext, setPatientContext] = useState<PatientContext | null>(null);
  const [selectedRegion, setSelectedRegion] = useState<BodyRegion>('initial_exam');
  
  // Get parameters from URL (cross-app bridge from PatientTracForge)
  const encounterId = searchParams.get('encounter_id');
  const patientId = searchParams.get('patient_id');
  const token = searchParams.get('token');

  useEffect(() => {
    if (!encounterId || !patientId || !token) {
      setError('Missing required parameters. Please access Revela through PatientTracForge check-in.');
      setLoading(false);
      return;
    }

    loadPatientContext();
  }, [encounterId, patientId, token]);

  const loadPatientContext = async () => {
    try {
      // Validate cross-app token
      const { data: tokenValidation, error: tokenError } = await supabase
        .rpc('validate_cross_app_token', { token_string: token });

      if (tokenError || !tokenValidation) {
        throw new Error('Invalid or expired session token');
      }

      // Load patient and encounter context
      const { data: encounterData, error: encounterError } = await supabase
        .from('cr.revela_encounter_view')
        .select('*')
        .eq('encounter_id', encounterId)
        .eq('patient_id', patientId)
        .single();

      if (encounterError) throw encounterError;

      setPatientContext({
        patient_id: encounterData.patient_id,
        encounter_id: encounterData.encounter_id,
        patient_name: `${encounterData.first_name} ${encounterData.last_name}`,
        patient_sex: encounterData.sex,
        patient_dob: encounterData.dob,
        patient_photo_url: encounterData.photo_url,
        chief_complaint: encounterData.chief_complaint,
        provider_id: encounterData.provider_id,
        provider_name: encounterData.provider_name,
        org_id: encounterData.org_id
      });

      setLoading(false);
    } catch (err) {
      console.error('Error loading patient context:', err);
      setError(err.message || 'Failed to load patient information');
      setLoading(false);
    }
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-[#060e1c] flex items-center justify-center">
        <div className="text-center">
          <div className="animate-spin rounded-full h-16 w-16 border-t-2 border-b-2 border-[#c9a96e] mx-auto mb-4"></div>
          <p className="text-[#c9a96e] font-rajdhani text-lg">Loading patient context...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="min-h-screen bg-[#060e1c] flex items-center justify-center">
        <div className="max-w-md mx-auto bg-red-900/20 border border-red-500 rounded-lg p-6">
          <AlertCircle className="w-12 h-12 text-red-500 mx-auto mb-4" />
          <h2 className="text-xl font-rajdhani text-red-500 text-center mb-2">Access Error</h2>
          <p className="text-gray-300 text-center mb-4">{error}</p>
          <button
            onClick={() => window.location.href = 'https://patienttracforge.com/schedule'}
            className="w-full bg-[#c9a96e] hover:bg-[#b39960] text-[#060e1c] font-rajdhani py-2 rounded transition-colors"
          >
            Return to PatientTracForge
          </button>
        </div>
      </div>
    );
  }

  const regions: { id: BodyRegion; label: string; icon: string }[] = [
    { id: 'initial_exam', label: 'Initial Exam', icon: '📋' },
    { id: 'breast', label: 'Breast', icon: '🫁' },
    { id: 'body', label: 'Body Contour', icon: '🧍' },
    { id: 'liposuction', label: 'Liposuction', icon: '💉' },
    { id: 'face', label: 'Face/Neck', icon: '👤' },
    { id: 'operative_note', label: 'Op Note', icon: '⚕️' }
  ];

  return (
    <div className="min-h-screen bg-[#060e1c]">
      {/* Patient Context Header */}
      <div className="bg-gradient-to-r from-[#060e1c] via-[#0a1628] to-[#060e1c] border-b border-[#c9a96e]/20 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            {/* Patient Info */}
            <div className="flex items-center gap-4">
              {patientContext?.patient_photo_url ? (
                <img
                  src={patientContext.patient_photo_url}
                  alt={patientContext.patient_name}
                  className="w-16 h-16 rounded-full border-2 border-[#c9a96e]"
                />
              ) : (
                <div className="w-16 h-16 rounded-full bg-[#c9a96e]/20 border-2 border-[#c9a96e] flex items-center justify-center">
                  <User className="w-8 h-8 text-[#c9a96e]" />
                </div>
              )}
              <div>
                <h1 className="text-2xl font-rajdhani font-bold text-white">
                  {patientContext?.patient_name}
                </h1>
                <div className="flex items-center gap-4 text-sm text-gray-400">
                  <span>MRN: {patientContext?.patient_id}</span>
                  <span>DOB: {new Date(patientContext?.patient_dob).toLocaleDateString()}</span>
                  <span className="capitalize">Sex: {patientContext?.patient_sex}</span>
                </div>
                {patientContext?.chief_complaint && (
                  <p className="text-sm text-[#c9a96e] mt-1">
                    Chief Complaint: {patientContext.chief_complaint}
                  </p>
                )}
              </div>
            </div>

            {/* Provider Info */}
            <div className="text-right">
              <p className="text-sm text-gray-400">Provider</p>
              <p className="text-white font-rajdhani font-semibold">
                {patientContext?.provider_name}
              </p>
              <p className="text-xs text-gray-500">Encounter ID: {patientContext?.encounter_id}</p>
            </div>
          </div>
        </div>
      </div>

      {/* Body Region Navigator */}
      <div className="max-w-7xl mx-auto px-4 py-6">
        <div className="bg-[#0a1628] rounded-lg border border-[#c9a96e]/30 p-6 mb-6">
          <h2 className="text-lg font-rajdhani font-semibold text-[#c9a96e] mb-4">
            SELECT DOCUMENTATION AREA
          </h2>
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4">
            {regions.map((region) => (
              <button
                key={region.id}
                onClick={() => setSelectedRegion(region.id)}
                className={`
                  relative p-6 rounded-lg border-2 transition-all duration-200
                  ${selectedRegion === region.id
                    ? 'bg-[#c9a96e]/20 border-[#c9a96e] shadow-lg shadow-[#c9a96e]/50'
                    : 'bg-[#060e1c]/50 border-gray-700 hover:border-[#c9a96e]/50 hover:bg-[#060e1c]'
                  }
                `}
              >
                <div className="text-4xl mb-2">{region.icon}</div>
                <div className={`text-sm font-rajdhani font-semibold ${
                  selectedRegion === region.id ? 'text-[#c9a96e]' : 'text-white'
                }`}>
                  {region.label}
                </div>
                {selectedRegion === region.id && (
                  <CheckCircle className="absolute top-2 right-2 w-5 h-5 text-[#c9a96e]" />
                )}
              </button>
            ))}
          </div>
        </div>

        {/* Dynamic Documentation Module */}
        <div className="bg-[#0a1628] rounded-lg border border-[#c9a96e]/30 p-6">
          {selectedRegion === 'initial_exam' && (
            <InitialExamModule patientContext={patientContext!} />
          )}
          {selectedRegion === 'breast' && (
            <BreastDocumentation patientContext={patientContext!} />
          )}
          {selectedRegion === 'body' && (
            <BodyDocumentation patientContext={patientContext!} />
          )}
          {selectedRegion === 'liposuction' && (
            <LiposuctionDocumentation patientContext={patientContext!} />
          )}
          {selectedRegion === 'face' && (
            <FaceDocumentation patientContext={patientContext!} />
          )}
          {selectedRegion === 'operative_note' && (
            <BasicOperativeNote patientContext={patientContext!} />
          )}
        </div>
      </div>

      {/* Action Bar */}
      <div className="fixed bottom-0 left-0 right-0 bg-[#0a1628] border-t border-[#c9a96e]/30 px-4 py-4">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <button
            onClick={() => window.location.href = `https://patienttracforge.com/schedule/checkin/${encounterId}`}
            className="px-6 py-2 bg-gray-700 hover:bg-gray-600 text-white font-rajdhani rounded transition-colors"
          >
            ← Back to Check-In
          </button>
          
          <div className="flex items-center gap-4">
            <button
              className="px-6 py-2 bg-[#c9a96e]/20 hover:bg-[#c9a96e]/30 text-[#c9a96e] font-rajdhani rounded transition-colors"
            >
              Save Draft
            </button>
            <button
              className="px-6 py-2 bg-[#c9a96e] hover:bg-[#b39960] text-[#060e1c] font-rajdhani font-bold rounded transition-colors"
            >
              Complete & Close Encounter
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
