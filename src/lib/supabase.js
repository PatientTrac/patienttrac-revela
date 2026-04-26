import { createClient } from '@supabase/supabase-js'

const url  = import.meta.env.VITE_SUPABASE_URL  || 'https://mskormozwekezjmtcylv.supabase.co'
const anon = import.meta.env.VITE_SUPABASE_ANON_KEY || ''

export const supabase = createClient(url, anon, {
  auth: { persistSession: false },
})

export const BRIDGE_URL = `${url}/functions/v1/cross-app-bridge`

// ── Bridge session validation ─────────────────────────────────────────────────
export const validateBridgeSession = async (token, encounterId, patientId, providerId, orgId) => {
  const res = await fetch(BRIDGE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action:'validate', token, encounter_id:encounterId, patient_id:patientId, provider_id:providerId, org_id:orgId }),
  })
  if (!res.ok) throw new Error(`Bridge error: ${res.status}`)
  return res.json()
}

export const parseSessionFromUrl = () => {
  const p = new URLSearchParams(window.location.search)
  const token=p.get('token'), encounter_id=p.get('encounter_id'), patient_id=p.get('patient_id')
  const provider_id=p.get('provider_id'), org_id=p.get('org_id')
  if (!token||!encounter_id||!patient_id||!provider_id||!org_id) return null
  return { token, encounter_id, patient_id, provider_id, org_id }
}

// ── Tab 1: Patient Chart ──────────────────────────────────────────────────────
export const chartApi = {
  getPatient: (patientId) =>
    supabase.from('patient').select('*').eq('patient_id', patientId).single(),

  getEncounter: (encounterId) =>
    supabase.from('encounter').select('*, provider:provider_id(first_name,last_name,specialty,credentials,npi)').eq('encounter_id', encounterId).single(),

  getMedications: (patientId) =>
    supabase.from('patient_medications').select('*').eq('patient_id', patientId).eq('purge',false).order('insert_date',{ascending:false}),

  getAllergies: (patientId) =>
    supabase.from('patient_allergies').select('*').eq('patient_id', patientId).eq('purge',false),

  getDiagnoses: (patientId, encounterId) =>
    supabase.from('patient_diagnosis').select('*').eq('patient_id', patientId).eq('encounter_id', encounterId).eq('purge',false),

  getVitals: (patientId) =>
    supabase.from('patient_vitals').select('*').eq('patient_id', patientId).order('insert_date',{ascending:false}).limit(5),

  getSurgicalHistory: (patientId) =>
    supabase.from('patient_surghist').select('*').eq('patient_id', patientId).order('insert_date',{ascending:false}),
}

// ── Tab 2: Progress Note (Prognote) ──────────────────────────────────────────
export const prognoteApi = {
  get: (patientId, encounterId) =>
    supabase.from('surgical_prognote').select('*').eq('patient_id', patientId).eq('encounter_id', encounterId).maybeSingle(),

  save: (data) =>
    supabase.from('surgical_prognote').upsert({
      ...data,
      insert_date: data.prog_id ? undefined : new Date().toISOString(),
      update_date: new Date().toISOString(),
    }).select().single(),

  getPhysicalExam: (patientId, encounterId) =>
    supabase.from('surgical_physical').select('*').eq('patient_id', patientId).eq('encounter_id', encounterId).maybeSingle(),

  savePhysicalExam: (data) =>
    supabase.from('surgical_physical').upsert({
      ...data, update_date: new Date().toISOString(),
    }).select().single(),
}

// ── Tab 3: Operative Note ─────────────────────────────────────────────────────
export const opnoteApi = {
  get: (patientId, encounterId) =>
    supabase.from('operative_notes').select('*').eq('patient_id', patientId).eq('encounter_id', encounterId).maybeSingle(),

  save: (data) =>
    supabase.from('operative_notes').upsert({
      ...data, update_date: new Date().toISOString(),
    }).select().single(),

  getTemplate: (surgeryType) =>
    supabase.from('operative_template').select('*').eq('surgery_type', surgeryType).maybeSingle(),

  getSurgCodes: (term) =>
    supabase.from('surg_procedure_codes').select('proc_code, proc_name, description, rvu')
      .or(`proc_name.ilike.%${term}%,proc_code.ilike.%${term}%`).limit(20),
}

// ── Tab 4: Post-Op Plan ───────────────────────────────────────────────────────
export const postopApi = {
  get: (patientId, encounterId) =>
    supabase.from('postop_plan').select('*').eq('patient_id', patientId).eq('encounter_id', encounterId).maybeSingle(),

  save: (data) =>
    supabase.from('postop_plan').upsert({
      ...data, update_date: new Date().toISOString(),
    }).select().single(),

  closeEncounter: (encounterId, closingData) =>
    supabase.from('encounter').update({
      ...closingData,
      encounter_status: 'closed',
      close_date: new Date().toISOString(),
      update_date: new Date().toISOString(),
    }).eq('encounter_id', encounterId),
}
