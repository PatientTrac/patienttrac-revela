import { createClient } from '@supabase/supabase-js'

const supabaseUrl  = import.meta.env.VITE_SUPABASE_URL  as string
const supabaseKey  = import.meta.env.VITE_SUPABASE_ANON_KEY as string

export const supabase = createClient(supabaseUrl, supabaseKey)

// ── Typed helpers ─────────────────────────────────────────────

export async function getPatients() {
  const { data, error } = await supabase
    .schema('cr')
    .from('patient')
    .select('patient_id, first_name, last_name, birth, status, facility_id')
    .order('last_name')
  if (error) throw error
  return data
}

export async function getPatientFlags(patientId: number) {
  const { data, error } = await supabase
    .schema('cr')
    .from('patient_flag')
    .select('*')
    .eq('patient_id', patientId)
    .eq('dismissed', false)
  if (error) throw error
  return data
}

export async function savePatientNote(note: {
  patient_id: number
  encounter_id?: number
  provider_id?: number
  note_type: string
  note_text: string
}) {
  const { data, error } = await supabase
    .schema('cr')
    .from('patient_notes')
    .insert(note)
    .select()
    .single()
  if (error) throw error
  return data
}

export async function saveAIFlags(flags: {
  patient_id: number
  flag_type: string
  flag_description: string
  severity: string
}[]) {
  const { data, error } = await supabase
    .schema('cr')
    .from('patient_flag')
    .insert(flags)
    .select()
  if (error) throw error
  return data
}

export async function getProposalTerms() {
  const { data, error } = await supabase
    .schema('terms')
    .from('proposal_terms')
    .select('*')
    .eq('is_active', true)
  if (error) throw error
  return data
}

export async function getMSETerms() {
  const { data, error } = await supabase
    .schema('terms')
    .from('mse_header_format')
    .select('*')
    .order('display_order')
  if (error) throw error
  return data
}

export async function getSurgicalSpecialties() {
  const { data, error } = await supabase
    .schema('terms')
    .from('surgical_specialties')
    .select('*')
    .eq('is_active', true)
  if (error) throw error
  return data
}
