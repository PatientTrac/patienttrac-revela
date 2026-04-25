// Medications Module — PatientTrac Mind + Revela + Scheduling
// Drop into any app that has access to the medication-lookup edge function
// Add as a tab: { id: 'medications', label: 'Medications', icon: '💊' }

import { useState, useEffect, useCallback } from 'react'
import MedSearch from './MedSearch'
import MedList from './MedList'
import InteractionChecker from './InteractionChecker'
import { useAppStore } from '../../lib/store'

export const MEDS_URL = `${import.meta.env.VITE_SUPABASE_URL}/functions/v1/medication-lookup`

export async function callMeds(body: Record<string, unknown>) {
  const res = await fetch(MEDS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  return res.json()
}

export interface Medication {
  medication_id: number
  drug_name: string
  rxnorm_cui: string
  rxcui_name: string
  dose_form: string | null
  strength: string | null
  dose: string | null
  dose_unit: string | null
  route: string | null
  frequency: string | null
  drug_class: string | null
  start_date: string | null
  stop_date: string | null
  is_active: boolean
  notes: string | null
}

export interface DrugResult {
  rxcui: string
  drug_name: string
  base_name?: string
  brand_names?: string[]
  dose_form?: string
  strength?: string
  route?: string
  drug_class?: string[]
  is_controlled?: boolean
  dea_schedule?: string | null
  tty?: string
}

type SubTab = 'medications' | 'add' | 'interactions'

export default function MedicationsModule() {
  const { context, session } = useAppStore()
  const [subTab, setSubTab] = useState<SubTab>('medications')
  const [meds, setMeds] = useState<Medication[]>([])
  const [loading, setLoading] = useState(true)
  const [interactionCount, setInteractionCount] = useState(0)
  const [hasHighSeverity, setHasHighSeverity] = useState(false)

  const patient = context?.patient as Record<string, unknown> | undefined
  const patientId = patient?.patient_id ? Number(patient.patient_id) : 0
  const orgId = session?.org_id || '00000000-0000-0000-0000-000000000001'
  const providerId = Number(session?.provider_id || 1)
  const encounterId = session?.encounter_id ? Number(session.encounter_id) : undefined

  const loadMeds = useCallback(async () => {
    if (!patientId) return
    setLoading(true)
    try {
      const d = await callMeds({ action: 'get_patient_meds', patient_id: patientId, org_id: orgId })
      setMeds(d.medications || [])
    } finally {
      setLoading(false)
    }
  }, [patientId, orgId])

  useEffect(() => { loadMeds() }, [loadMeds])

  // Check interactions whenever meds change
  useEffect(() => {
    const cuis = meds.filter(m => m.rxnorm_cui).map(m => m.rxnorm_cui)
    if (cuis.length < 2) { setInteractionCount(0); setHasHighSeverity(false); return }
    callMeds({ action: 'check_interactions', rxcuis: cuis })
      .then(d => { setInteractionCount(d.count || 0); setHasHighSeverity(d.has_high_severity || false) })
      .catch(() => {})
  }, [meds])

  const handleMedAdded = () => {
    loadMeds()
    setSubTab('medications')
  }

  const handleDiscontinue = async (medicationId: number) => {
    await callMeds({ action: 'discontinue', medication_id: medicationId, org_id: orgId })
    loadMeds()
  }

  const SUBTABS: { id: SubTab; label: string; icon: string; badge?: number; badgeColor?: string }[] = [
    { id: 'medications', label: 'Current Meds', icon: '💊' },
    { id: 'add',         label: 'Add Medication', icon: '＋' },
    { id: 'interactions', label: 'Interactions',  icon: '⚡',
      badge: interactionCount || undefined,
      badgeColor: hasHighSeverity ? '#ff3b3b' : '#e67e22' },
  ]

  return (
    <div className="animate-fadeIn">
      {/* Sub-tab bar */}
      <div className="flex border-b px-4"
        style={{ background: 'rgba(6,14,28,0.5)', borderColor: 'rgba(0,212,255,0.1)' }}>
        {SUBTABS.map(t => (
          <button key={t.id} onClick={() => setSubTab(t.id)}
            className={`flex items-center gap-1.5 px-4 py-2.5 text-xs transition-all ${subTab === t.id ? 'tab-active' : 'tab-inactive'}`}
            style={{ fontFamily: 'Rajdhani', fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' }}>
            <span style={{ fontSize: 14 }}>{t.icon}</span>
            {t.label}
            {t.badge ? (
              <span className="ml-1 px-1.5 py-0.5 rounded text-xs font-bold"
                style={{ background: t.badgeColor, color: 'white', fontSize: '0.6rem' }}>
                {t.badge}
              </span>
            ) : null}
          </button>
        ))}
      </div>

      {/* High-severity interaction banner */}
      {hasHighSeverity && subTab !== 'interactions' && (
        <div className="mx-5 mt-4 p-3 rounded flex items-center justify-between"
          style={{ background: 'rgba(255,59,59,0.1)', border: '1px solid rgba(255,59,59,0.4)' }}>
          <div className="flex items-center gap-2">
            <span>⚠️</span>
            <span className="text-sm font-semibold" style={{ color: '#ff3b3b', fontFamily: 'Rajdhani' }}>
              High-severity drug interaction detected — review before prescribing
            </span>
          </div>
          <button onClick={() => setSubTab('interactions')}
            className="text-xs px-3 py-1 rounded"
            style={{ background: '#ff3b3b', color: 'white', fontFamily: 'Rajdhani', fontWeight: 700 }}>
            Review
          </button>
        </div>
      )}

      <div className="p-5">
        {subTab === 'medications' && (
          <MedList
            meds={meds}
            loading={loading}
            onDiscontinue={handleDiscontinue}
            onAddClick={() => setSubTab('add')}
          />
        )}
        {subTab === 'add' && (
          <MedSearch
            patientId={patientId}
            encounterId={encounterId}
            providerId={providerId}
            orgId={orgId}
            currentMeds={meds}
            onAdded={handleMedAdded}
            onCancel={() => setSubTab('medications')}
          />
        )}
        {subTab === 'interactions' && (
          <InteractionChecker
            meds={meds}
            orgId={orgId}
          />
        )}
      </div>
    </div>
  )
}
