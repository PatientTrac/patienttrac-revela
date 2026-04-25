// DiagnosisModule — PatientTrac Mind / Scheduling
// §170.315(a)(6) SNOMED-coded diagnoses + full ICD-10-CM search
// Add as tab: { id: 'diagnoses', label: 'Diagnoses', icon: '🩺' }

import { useState, useEffect, useRef, useCallback } from 'react'
import { useAppStore } from '../../lib/store'

const ICD_URL = `${import.meta.env.VITE_SUPABASE_URL}/functions/v1/icd10-lookup`

async function callICD(body: Record<string, unknown>) {
  const res = await fetch(ICD_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  return res.json()
}

interface Diagnosis {
  diagnosis_id: number
  icd_code: string
  icd_description: string
  is_primary: boolean
  is_active: boolean
  problem_status: string | null
  clinical_status: string | null
  onset_date: string | null
  is_chronic: boolean | null
  diagnosis_source: string | null
  icd_chapter: string | null
}

interface ICDResult {
  icd_code: string
  description: string
  category: string
  chapter: string
  is_billable: boolean
  is_dsm5: boolean
  specialty: string[]
}

const STATUS_COLORS: Record<string, { bg: string; text: string }> = {
  active:   { bg: 'rgba(39,174,96,0.1)',   text: '#27ae60' },
  chronic:  { bg: 'rgba(230,126,34,0.1)',  text: '#e67e22' },
  resolved: { bg: 'rgba(149,165,166,0.1)', text: '#95a5a6' },
  inactive: { bg: 'rgba(149,165,166,0.1)', text: '#95a5a6' },
}

const PROBLEM_STATUSES = ['active','chronic','resolved','inactive','recurrent','relapse']

export default function DiagnosisModule() {
  const { context, session } = useAppStore()
  const [subTab, setSubTab] = useState<'list' | 'add'>('list')
  const [diagnoses, setDiagnoses] = useState<Diagnosis[]>([])
  const [loading, setLoading] = useState(true)

  const patient = context?.patient as Record<string, unknown> | undefined
  const patientId = Number(patient?.patient_id || 0)
  const orgId = session?.org_id || '00000000-0000-0000-0000-000000000001'
  const providerId = Number(session?.provider_id || 1)
  const encounterId = session?.encounter_id

  const loadDiagnoses = useCallback(async () => {
    if (!patientId) return
    setLoading(true)
    try {
      const d = await callICD({ action: 'get_problem_list', patient_id: patientId })
      setDiagnoses(d.diagnoses || [])
    } finally { setLoading(false) }
  }, [patientId])

  useEffect(() => { loadDiagnoses() }, [loadDiagnoses])

  return (
    <div className="animate-fadeIn">
      <div className="flex border-b px-4" style={{ background: 'rgba(6,14,28,0.5)', borderColor: 'rgba(0,212,255,0.1)' }}>
        {[{ id: 'list' as const, label: 'Problem List', icon: '🩺' }, { id: 'add' as const, label: 'Add Diagnosis', icon: '＋' }].map(t => (
          <button key={t.id} onClick={() => setSubTab(t.id)}
            className={`flex items-center gap-1.5 px-4 py-2.5 text-xs transition-all ${subTab === t.id ? 'tab-active' : 'tab-inactive'}`}
            style={{ fontFamily: 'Rajdhani', fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' }}>
            <span style={{ fontSize: 14 }}>{t.icon}</span>{t.label}
          </button>
        ))}
      </div>
      <div className="p-5">
        {subTab === 'list'
          ? <ProblemList diagnoses={diagnoses} loading={loading} onAddClick={() => setSubTab('add')} onUpdated={loadDiagnoses} />
          : <DiagnosisSearch patientId={patientId} orgId={orgId} providerId={providerId}
              encounterId={String(encounterId || '')} onAdded={() => { loadDiagnoses(); setSubTab('list') }}
              onCancel={() => setSubTab('list')} />
        }
      </div>
    </div>
  )
}

function ProblemList({ diagnoses, loading, onAddClick, onUpdated }: {
  diagnoses: Diagnosis[]; loading: boolean; onAddClick: () => void; onUpdated: () => void
}) {
  const [expandedId, setExpandedId] = useState<number | null>(null)
  const [updatingId, setUpdatingId] = useState<number | null>(null)

  const handleStatusChange = async (id: number, status: string) => {
    setUpdatingId(id)
    await callICD({ action: 'update_diagnosis', diagnosis_id: id,
      updates: { problem_status: status, is_active: status !== 'resolved' && status !== 'inactive' } })
    onUpdated(); setUpdatingId(null)
  }

  if (loading) return <div className="py-12 text-center hud-label animate-pulse">Loading problem list...</div>
  if (diagnoses.length === 0) return (
    <div className="text-center py-16">
      <div style={{ fontSize: '3rem', marginBottom: 12, opacity: 0.3 }}>🩺</div>
      <div className="text-sm opacity-40 mb-4">No active diagnoses on record.</div>
      <button onClick={onAddClick} className="px-6 py-2.5 rounded text-sm font-bold"
        style={{ background: 'rgba(201,169,110,0.12)', border: '1px solid rgba(201,169,110,0.4)', color: '#c9a96e', fontFamily: 'Rajdhani', letterSpacing: '0.06em' }}>
        + Add Diagnosis
      </button>
    </div>
  )

  const primary = diagnoses.filter(d => d.is_primary)
  const secondary = diagnoses.filter(d => !d.is_primary)

  const renderDx = (d: Diagnosis) => {
    const statusCol = STATUS_COLORS[d.problem_status?.toLowerCase() || 'active'] || STATUS_COLORS.active
    const isExpanded = expandedId === d.diagnosis_id
    return (
      <div key={d.diagnosis_id} className="rounded overflow-hidden mb-2"
        style={{ border: '1px solid rgba(255,255,255,0.06)', background: 'rgba(10,22,40,0.7)' }}>
        <div className="flex items-center gap-3 px-4 py-3 cursor-pointer hover:bg-white/3"
          onClick={() => setExpandedId(isExpanded ? null : d.diagnosis_id)}>
          <div className="flex-1 min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="font-mono text-xs px-1.5 py-0.5 rounded"
                style={{ background: 'rgba(0,212,255,0.08)', color: '#00d4ff', border: '1px solid rgba(0,212,255,0.15)' }}>
                {d.icd_code}
              </span>
              <span className="text-sm" style={{ color: 'rgba(200,215,235,0.88)' }}>{d.icd_description}</span>
              {d.is_primary && (
                <span className="text-xs px-1.5 py-0.5 rounded" style={{ background: 'rgba(201,169,110,0.12)', color: '#c9a96e', fontFamily: 'Rajdhani' }}>PRIMARY</span>
              )}
              {d.is_chronic && (
                <span className="text-xs px-1.5 py-0.5 rounded" style={{ background: 'rgba(230,126,34,0.1)', color: '#e67e22', fontFamily: 'Rajdhani' }}>CHRONIC</span>
              )}
            </div>
            <div className="flex items-center gap-2 mt-0.5">
              <span className="text-xs px-1.5 py-0.5 rounded"
                style={{ background: statusCol.bg, color: statusCol.text }}>
                {d.problem_status || 'active'}
              </span>
              {d.onset_date && <span className="text-xs opacity-30">Onset: {new Date(d.onset_date).toLocaleDateString()}</span>}
              {d.icd_chapter && <span className="text-xs opacity-25">{d.icd_chapter}</span>}
            </div>
          </div>
          <span className="text-xs opacity-20">{isExpanded ? '▲' : '▼'}</span>
        </div>

        {isExpanded && (
          <div className="px-4 pb-4 border-t" style={{ borderColor: 'rgba(255,255,255,0.05)' }}>
            <div className="flex items-center gap-2 mt-3 mb-3 flex-wrap">
              <span className="text-xs opacity-40">Change status:</span>
              {PROBLEM_STATUSES.map(s => (
                <button key={s} onClick={() => handleStatusChange(d.diagnosis_id, s)}
                  disabled={updatingId === d.diagnosis_id}
                  className="text-xs px-2 py-1 rounded transition-all disabled:opacity-40"
                  style={{
                    background: d.problem_status === s ? 'rgba(0,212,255,0.12)' : 'rgba(255,255,255,0.03)',
                    border: d.problem_status === s ? '1px solid rgba(0,212,255,0.4)' : '1px solid rgba(255,255,255,0.08)',
                    color: d.problem_status === s ? '#00d4ff' : 'rgba(200,215,235,0.5)',
                    fontFamily: 'Rajdhani',
                  }}>
                  {s}
                </button>
              ))}
            </div>
            <div className="text-xs opacity-30">Source: {d.diagnosis_source || 'provider'}</div>
          </div>
        )}
      </div>
    )
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-2">
          <span style={{ fontFamily: 'Rajdhani', fontWeight: 700, color: '#c9a96e' }}>Active Problem List</span>
          <span className="px-2 py-0.5 rounded text-xs" style={{ background: 'rgba(0,212,255,0.08)', color: '#00d4ff', border: '1px solid rgba(0,212,255,0.15)', fontFamily: 'Rajdhani' }}>
            {diagnoses.length}
          </span>
        </div>
        <button onClick={onAddClick} className="px-4 py-1.5 rounded text-xs font-bold"
          style={{ background: 'rgba(201,169,110,0.1)', border: '1px solid rgba(201,169,110,0.3)', color: '#c9a96e', fontFamily: 'Rajdhani', letterSpacing: '0.06em' }}>
          + Add Diagnosis
        </button>
      </div>

      {primary.length > 0 && (
        <div className="mb-4">
          <div className="hud-label text-xs mb-2">Primary Diagnosis</div>
          {primary.map(renderDx)}
        </div>
      )}
      {secondary.length > 0 && (
        <div>
          <div className="hud-label text-xs mb-2">Secondary Diagnoses</div>
          {secondary.map(renderDx)}
        </div>
      )}
      <div className="mt-4 text-xs opacity-25 text-right">§170.315(a)(6) · ICD-10-CM · NLM Clinical Tables</div>
    </div>
  )
}

function DiagnosisSearch({ patientId, orgId, providerId, encounterId, onAdded, onCancel }: {
  patientId: number; orgId: string; providerId: number; encounterId: string;
  onAdded: () => void; onCancel: () => void
}) {
  const [query, setQuery] = useState('')
  const [results, setResults] = useState<ICDResult[]>([])
  const [searching, setSearching] = useState(false)
  const [selected, setSelected] = useState<ICDResult | null>(null)
  const [showDrop, setShowDrop] = useState(false)
  const [isPrimary, setIsPrimary] = useState(false)
  const [isChronic, setIsChronic] = useState(false)
  const [problemStatus, setProblemStatus] = useState('active')
  const [onsetDate, setOnsetDate] = useState('')
  const [dsm5Only, setDsm5Only] = useState(false)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [error, setError] = useState('')
  const debRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const dropRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (query.length < 2) { setResults([]); setShowDrop(false); return }
    if (debRef.current) clearTimeout(debRef.current)
    debRef.current = setTimeout(async () => {
      setSearching(true)
      try {
        const d = await callICD({ action: 'search', query, dsm5_only: dsm5Only, max_results: 15 })
        setResults(d.results || [])
        setShowDrop(true)
      } finally { setSearching(false) }
    }, 300)
  }, [query, dsm5Only])

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (dropRef.current && !dropRef.current.contains(e.target as Node)) setShowDrop(false)
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [])

  const handleSave = async () => {
    if (!selected) return
    setSaving(true); setError('')
    try {
      const d = await callICD({
        action: 'save_diagnosis', patient_id: patientId, org_id: orgId,
        provider_id: providerId, encounter_id: encounterId || undefined,
        diagnosis: {
          icd_code: selected.icd_code,
          description: selected.description,
          is_primary: isPrimary,
          problem_status: problemStatus,
          clinical_status: 'active',
          onset_date: onsetDate || null,
          is_chronic: isChronic,
          diagnosis_source: 'provider',
        }
      })
      if (d.duplicate) { setError(d.message); setSaving(false); return }
      if (!d.success) throw new Error(d.error || 'Save failed')
      setSaved(true)
      setTimeout(() => onAdded(), 900)
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Failed to save')
    } finally { setSaving(false) }
  }

  if (saved) return (
    <div className="text-center py-16 animate-fadeIn">
      <div style={{ fontSize: '3rem', marginBottom: 12 }}>✅</div>
      <div style={{ fontFamily: 'Rajdhani', fontSize: '1.25rem', color: '#27ae60', fontWeight: 700 }}>Diagnosis added</div>
    </div>
  )

  return (
    <div className="max-w-2xl mx-auto">
      <div className="flex items-center justify-between mb-4">
        <span style={{ fontFamily: 'Rajdhani', fontSize: '1rem', fontWeight: 700, color: '#c9a96e' }}>Add Diagnosis</span>
        <button onClick={onCancel} className="text-xs opacity-40 hover:opacity-70">✕ Cancel</button>
      </div>

      {/* DSM-5 filter toggle */}
      <div className="flex items-center gap-3 mb-4">
        <label className="flex items-center gap-2 cursor-pointer">
          <input type="checkbox" checked={dsm5Only} onChange={e => setDsm5Only(e.target.checked)} />
          <span className="text-xs opacity-60">DSM-5 / Psychiatric codes only</span>
        </label>
      </div>

      <div ref={dropRef} className="relative mb-5">
        <input type="text" placeholder="Search ICD-10-CM (e.g. generalized anxiety, F41.1, major depressive...)"
          value={query} onChange={e => { setQuery(e.target.value); setSelected(null) }}
          onFocus={() => results.length > 0 && setShowDrop(true)}
          className="w-full px-4 py-3 rounded text-sm"
          style={{ background: 'rgba(0,212,255,0.05)', border: '1px solid rgba(0,212,255,0.25)', color: 'rgba(200,215,235,0.9)', outline: 'none' }} />
        {searching && <div className="absolute right-3 top-3 text-sm opacity-40 animate-spin">⟳</div>}
        {!searching && query.length >= 2 && <div className="absolute right-3 top-3 text-xs opacity-30">ICD-10</div>}

        {showDrop && results.length > 0 && (
          <div className="absolute z-50 w-full mt-1 rounded shadow-lg"
            style={{ background: 'rgba(8,18,34,0.98)', border: '1px solid rgba(0,212,255,0.2)', maxHeight: 320, overflowY: 'auto' }}>
            {results.map((r, i) => (
              <button key={i} onClick={() => { setSelected(r); setQuery(`${r.icd_code} — ${r.description}`); setShowDrop(false) }}
                className="w-full text-left px-4 py-2.5 transition-all hover:bg-white/5 border-b last:border-b-0"
                style={{ borderColor: 'rgba(255,255,255,0.04)' }}>
                <div className="flex items-center gap-3">
                  <span className="font-mono text-xs px-1.5 py-0.5 rounded flex-shrink-0"
                    style={{ background: 'rgba(0,212,255,0.08)', color: '#00d4ff', border: '1px solid rgba(0,212,255,0.15)' }}>
                    {r.icd_code}
                  </span>
                  <div className="min-w-0">
                    <div className="text-sm truncate" style={{ color: 'rgba(200,215,235,0.9)' }}>{r.description}</div>
                    <div className="text-xs opacity-35">{r.chapter}</div>
                  </div>
                  {r.is_dsm5 && (
                    <span className="text-xs px-1.5 py-0.5 rounded flex-shrink-0"
                      style={{ background: 'rgba(142,68,173,0.12)', color: '#8e44ad', border: '1px solid rgba(142,68,173,0.25)', fontFamily: 'Rajdhani' }}>
                      DSM-5
                    </span>
                  )}
                </div>
              </button>
            ))}
          </div>
        )}
      </div>

      {selected && (
        <div className="animate-fadeIn">
          <div className="p-4 rounded mb-4" style={{ background: 'rgba(10,22,40,0.8)', border: '1px solid rgba(0,212,255,0.15)' }}>
            <div className="flex items-center gap-3 mb-1">
              <span className="font-mono text-sm px-2 py-0.5 rounded"
                style={{ background: 'rgba(0,212,255,0.1)', color: '#00d4ff', border: '1px solid rgba(0,212,255,0.2)' }}>
                {selected.icd_code}
              </span>
              {selected.is_dsm5 && (
                <span className="text-xs px-2 py-0.5 rounded"
                  style={{ background: 'rgba(142,68,173,0.12)', color: '#8e44ad', fontFamily: 'Rajdhani' }}>DSM-5</span>
              )}
            </div>
            <div className="font-semibold" style={{ color: 'rgba(200,215,235,0.92)' }}>{selected.description}</div>
            <div className="text-xs opacity-40 mt-1">{selected.chapter}</div>
          </div>

          <div className="grid grid-cols-2 gap-3 mb-4">
            <div>
              <div className="hud-label text-xs mb-1.5">Problem Status</div>
              <select value={problemStatus} onChange={e => setProblemStatus(e.target.value)}
                className="w-full px-3 py-2 rounded text-xs"
                style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.85)', outline: 'none' }}>
                {PROBLEM_STATUSES.map(s => <option key={s} value={s}>{s.charAt(0).toUpperCase() + s.slice(1)}</option>)}
              </select>
            </div>
            <div>
              <div className="hud-label text-xs mb-1.5">Onset Date</div>
              <input type="date" value={onsetDate} onChange={e => setOnsetDate(e.target.value)}
                className="w-full px-3 py-2 rounded text-xs"
                style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.85)', outline: 'none' }} />
            </div>
          </div>

          <div className="flex gap-6 mb-4">
            <label className="flex items-center gap-2 cursor-pointer text-xs opacity-70">
              <input type="checkbox" checked={isPrimary} onChange={e => setIsPrimary(e.target.checked)} />
              Primary diagnosis for this encounter
            </label>
            <label className="flex items-center gap-2 cursor-pointer text-xs opacity-70">
              <input type="checkbox" checked={isChronic} onChange={e => setIsChronic(e.target.checked)} />
              Chronic condition
            </label>
          </div>

          {error && (
            <div className="p-3 rounded mb-3 text-sm" style={{ background: 'rgba(255,59,59,0.1)', color: '#ff3b3b', border: '1px solid rgba(255,59,59,0.3)' }}>
              {error}
            </div>
          )}

          <div className="flex gap-3">
            <button onClick={onCancel} className="px-5 py-3 rounded text-sm"
              style={{ background: 'rgba(255,255,255,0.03)', border: '1px solid rgba(255,255,255,0.1)', color: 'rgba(200,215,235,0.4)', fontFamily: 'Rajdhani' }}>
              Cancel
            </button>
            <button onClick={handleSave} disabled={!selected || saving}
              className="flex-1 py-3 rounded text-sm font-bold disabled:opacity-30 transition-all"
              style={{ background: 'rgba(201,169,110,0.15)', border: '1px solid rgba(201,169,110,0.5)', color: '#c9a96e', fontFamily: 'Rajdhani', letterSpacing: '0.08em' }}>
              {saving ? 'Saving...' : `Add ${selected.icd_code} to Problem List →`}
            </button>
          </div>
        </div>
      )}
    </div>
  )
}
