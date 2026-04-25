// AllergyModule — PatientTrac Mind / Revela / Scheduling
// §170.315(a)(8) Drug allergy with RxNorm coding
// Add as tab: { id: 'allergies', label: 'Allergies', icon: '⚠️' }

import { useState, useEffect, useRef, useCallback } from 'react'
import { useAppStore } from '../../lib/store'

const ALLERGY_URL = `${import.meta.env.VITE_SUPABASE_URL}/functions/v1/allergy-lookup`

async function callAllergy(body: Record<string, unknown>) {
  const res = await fetch(ALLERGY_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  return res.json()
}

interface Allergy {
  allergy_id: number
  allergen: string
  substance_name: string | null
  allergen_type: string | null
  substance_type: string | null
  reaction: string | null
  severity: string | null
  clinical_status: string | null
  verification_status: string | null
  rxnorm_cui: string | null
  fhir_category: string | null
  drug_class: string | null
  ige_mediated: boolean | null
  cross_reactivity_note: string | null
  onset_date: string | null
}

interface AllergenResult {
  rxnorm_cui: string | null
  substance_name: string
  display_name: string | null
  substance_type: string
  fhir_category: string
  drug_class: string[] | null
  cross_reactive_classes: string[] | null
  common_reactions: string[] | null
  severity_potential: string | null
  ige_mediated: boolean | null
  synonyms: string[] | null
}

const SEVERITY_COLORS: Record<string, { bg: string; text: string; border: string }> = {
  anaphylaxis: { bg: 'rgba(255,59,59,0.12)',  text: '#ff3b3b', border: 'rgba(255,59,59,0.4)' },
  severe:      { bg: 'rgba(231,76,60,0.1)',   text: '#e74c3c', border: 'rgba(231,76,60,0.3)' },
  moderate:    { bg: 'rgba(230,126,34,0.1)',  text: '#e67e22', border: 'rgba(230,126,34,0.3)' },
  mild:        { bg: 'rgba(241,196,15,0.08)', text: '#f1c40f', border: 'rgba(241,196,15,0.25)' },
}

const TYPE_ICONS: Record<string, string> = {
  drug: '💊', drug_class: '💊', ingredient: '💊',
  food: '🍽️', environmental: '🌿', latex: '🧤',
  insect: '🐝', contrast_media: '💉', biological: '🧬', other: '⚕️',
}

const REACTIONS = [
  'Anaphylaxis','Urticaria (hives)','Angioedema','Rash','Pruritus (itching)',
  'Bronchospasm','Dyspnea','Nausea/vomiting','Diarrhea','Dizziness',
  'Hypotension','Tachycardia','Stevens-Johnson syndrome','Fixed drug eruption',
  'Contact dermatitis','Serum sickness','Photosensitivity','Hepatotoxicity',
  'Nephrotoxicity','Thrombocytopenia','Other',
]

const SEVERITY_OPTIONS = ['mild','moderate','severe','anaphylaxis']

type SubTab = 'list' | 'add'

export default function AllergyModule() {
  const { context, session } = useAppStore()
  const [subTab, setSubTab] = useState<SubTab>('list')
  const [allergies, setAllergies] = useState<Allergy[]>([])
  const [loading, setLoading] = useState(true)

  const patient = context?.patient as Record<string, unknown> | undefined
  const patientId = Number(patient?.patient_id || 0)
  const orgId = session?.org_id || '00000000-0000-0000-0000-000000000001'
  const providerId = Number(session?.provider_id || 1)
  const encounterId = session?.encounter_id

  const loadAllergies = useCallback(async () => {
    if (!patientId) return
    setLoading(true)
    try {
      const d = await callAllergy({ action: 'get_patient_allergies', patient_id: patientId })
      setAllergies(d.allergies || [])
    } finally { setLoading(false) }
  }, [patientId])

  useEffect(() => { loadAllergies() }, [loadAllergies])

  const handleAdded = () => { loadAllergies(); setSubTab('list') }

  return (
    <div className="animate-fadeIn">
      <div className="flex border-b px-4" style={{ background: 'rgba(6,14,28,0.5)', borderColor: 'rgba(0,212,255,0.1)' }}>
        {[{ id: 'list' as SubTab, label: 'Allergies', icon: '⚠️' }, { id: 'add' as SubTab, label: 'Add Allergy', icon: '＋' }].map(t => (
          <button key={t.id} onClick={() => setSubTab(t.id)}
            className={`flex items-center gap-1.5 px-4 py-2.5 text-xs transition-all ${subTab === t.id ? 'tab-active' : 'tab-inactive'}`}
            style={{ fontFamily: 'Rajdhani', fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' }}>
            <span style={{ fontSize: 14 }}>{t.icon}</span>{t.label}
          </button>
        ))}
      </div>
      <div className="p-5">
        {subTab === 'list'
          ? <AllergyList allergies={allergies} loading={loading} onAddClick={() => setSubTab('add')} onUpdated={loadAllergies} />
          : <AllergySearch patientId={patientId} orgId={orgId} providerId={providerId} encounterId={String(encounterId || '')} onAdded={handleAdded} onCancel={() => setSubTab('list')} />
        }
      </div>
    </div>
  )
}

function AllergyList({ allergies, loading, onAddClick, onUpdated }: {
  allergies: Allergy[]; loading: boolean; onAddClick: () => void; onUpdated: () => void
}) {
  const [expandedId, setExpandedId] = useState<number | null>(null)
  const [removing, setRemoving] = useState<number | null>(null)

  const handleRemove = async (id: number) => {
    setRemoving(id)
    await callAllergy({ action: 'update_allergy', allergy_id: id, updates: { is_active: false, clinical_status: 'resolved' } })
    onUpdated()
    setRemoving(null)
  }

  if (loading) return <div className="py-12 text-center hud-label animate-pulse">Loading allergies...</div>
  if (allergies.length === 0) return (
    <div className="text-center py-16">
      <div style={{ fontSize: '3rem', marginBottom: 12, opacity: 0.3 }}>⚠️</div>
      <div className="text-sm opacity-40 mb-4">No known drug allergies on record.</div>
      <div className="text-xs opacity-25 mb-6">NKDA — No Known Drug Allergies</div>
      <button onClick={onAddClick} className="px-6 py-2.5 rounded text-sm font-bold"
        style={{ background: 'rgba(201,169,110,0.12)', border: '1px solid rgba(201,169,110,0.4)', color: '#c9a96e', fontFamily: 'Rajdhani', letterSpacing: '0.06em' }}>
        + Add Allergy
      </button>
    </div>
  )

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-2">
          <span style={{ fontFamily: 'Rajdhani', fontWeight: 700, color: '#c9a96e' }}>Active Allergies</span>
          <span className="px-2 py-0.5 rounded text-xs" style={{ background: 'rgba(255,59,59,0.12)', color: '#ff3b3b', border: '1px solid rgba(255,59,59,0.25)', fontFamily: 'Rajdhani' }}>
            {allergies.length}
          </span>
        </div>
        <button onClick={onAddClick} className="px-4 py-1.5 rounded text-xs font-bold"
          style={{ background: 'rgba(201,169,110,0.1)', border: '1px solid rgba(201,169,110,0.3)', color: '#c9a96e', fontFamily: 'Rajdhani', letterSpacing: '0.06em' }}>
          + Add Allergy
        </button>
      </div>

      <div className="space-y-2">
        {allergies.map(a => {
          const sev = SEVERITY_COLORS[a.severity?.toLowerCase() || ''] || SEVERITY_COLORS.mild
          const icon = TYPE_ICONS[a.substance_type || ''] || '⚕️'
          const isExpanded = expandedId === a.allergy_id
          return (
            <div key={a.allergy_id} className="rounded overflow-hidden" style={{ border: `1px solid ${sev.border}` }}>
              <div className="flex items-center gap-3 px-4 py-3 cursor-pointer" style={{ background: sev.bg }}
                onClick={() => setExpandedId(isExpanded ? null : a.allergy_id)}>
                <span style={{ fontSize: '1.2rem', flexShrink: 0 }}>{icon}</span>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="font-bold text-sm" style={{ color: 'rgba(200,215,235,0.92)' }}>{a.allergen}</span>
                    {a.severity && (
                      <span className="text-xs px-2 py-0.5 rounded font-bold"
                        style={{ background: `${sev.text}20`, color: sev.text, border: `1px solid ${sev.text}40`, fontFamily: 'Rajdhani' }}>
                        {a.severity.toUpperCase()}
                      </span>
                    )}
                    {a.fhir_category && (
                      <span className="text-xs opacity-40">{a.fhir_category}</span>
                    )}
                  </div>
                  {a.reaction && <div className="text-xs opacity-50 mt-0.5">{a.reaction}</div>}
                </div>
                <div className="text-xs opacity-20">{isExpanded ? '▲' : '▼'}</div>
              </div>

              {isExpanded && (
                <div className="px-4 pb-4 border-t" style={{ background: 'rgba(6,14,28,0.6)', borderColor: `${sev.border}` }}>
                  <div className="grid grid-cols-3 gap-2 mt-3 mb-3">
                    {[
                      { label: 'Substance type', value: a.substance_type || '—' },
                      { label: 'Drug class',      value: a.drug_class || '—' },
                      { label: 'IgE-mediated',    value: a.ige_mediated == null ? '—' : a.ige_mediated ? 'Yes' : 'No' },
                      { label: 'RxNorm CUI',      value: a.rxnorm_cui || '—' },
                      { label: 'Clinical status',  value: a.clinical_status || '—' },
                      { label: 'Verification',     value: a.verification_status || '—' },
                    ].map(({ label, value }) => (
                      <div key={label} className="p-2 rounded" style={{ background: 'rgba(0,0,0,0.2)' }}>
                        <div className="text-xs opacity-35 mb-0.5">{label}</div>
                        <div className="text-xs font-mono" style={{ color: 'rgba(200,215,235,0.7)' }}>{value}</div>
                      </div>
                    ))}
                  </div>
                  {a.cross_reactivity_note && (
                    <div className="p-2 rounded mb-3 text-xs" style={{ background: 'rgba(230,126,34,0.08)', border: '1px solid rgba(230,126,34,0.2)', color: 'rgba(230,126,34,0.9)' }}>
                      ⚡ Cross-reactivity: {a.cross_reactivity_note}
                    </div>
                  )}
                  <button onClick={() => handleRemove(a.allergy_id)} disabled={removing === a.allergy_id}
                    className="text-xs px-3 py-1.5 rounded disabled:opacity-40"
                    style={{ background: 'rgba(231,76,60,0.08)', border: '1px solid rgba(231,76,60,0.25)', color: '#e74c3c', fontFamily: 'Rajdhani' }}>
                    {removing === a.allergy_id ? 'Removing...' : 'Mark Resolved / Remove'}
                  </button>
                </div>
              )}
            </div>
          )
        })}
      </div>
      <div className="mt-4 text-xs opacity-25 text-right">§170.315(a)(8) · RxNorm coded · NLM RxNav</div>
    </div>
  )
}

function AllergySearch({ patientId, orgId, providerId, encounterId, onAdded, onCancel }: {
  patientId: number; orgId: string; providerId: number; encounterId: string;
  onAdded: () => void; onCancel: () => void;
}) {
  const [query, setQuery] = useState('')
  const [results, setResults] = useState<AllergenResult[]>([])
  const [searching, setSearching] = useState(false)
  const [selected, setSelected] = useState<AllergenResult | null>(null)
  const [showDrop, setShowDrop] = useState(false)
  const [reaction, setReaction] = useState('')
  const [severity, setSeverity] = useState('')
  const [onsetDate, setOnsetDate] = useState('')
  const [notes, setNotes] = useState('')
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
        const d = await callAllergy({ action: 'search', query })
        setResults(d.results || [])
        setShowDrop(true)
      } finally { setSearching(false) }
    }, 300)
  }, [query])

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (dropRef.current && !dropRef.current.contains(e.target as Node)) setShowDrop(false)
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [])

  const handleSave = async () => {
    if (!selected || !reaction || !severity) return
    setSaving(true); setError('')
    try {
      const d = await callAllergy({
        action: 'save_allergy', patient_id: patientId, org_id: orgId,
        provider_id: providerId, encounter_id: encounterId || undefined,
        allergy: {
          allergen: selected.display_name || selected.substance_name,
          substance_name: selected.substance_name,
          substance_type: selected.substance_type,
          fhir_category: selected.fhir_category,
          rxnorm_cui: selected.rxnorm_cui,
          drug_class: selected.drug_class,
          ige_mediated: selected.ige_mediated,
          cross_reactivity_note: selected.cross_reactive_classes?.join('; ') || null,
          reaction,
          severity,
          onset_date: onsetDate || null,
          clinical_status: 'active',
          verification_status: 'confirmed',
          note: notes || null,
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
      <div style={{ fontFamily: 'Rajdhani', fontSize: '1.25rem', color: '#27ae60', fontWeight: 700 }}>Allergy recorded</div>
      <div className="text-sm opacity-50 mt-1">Returning to allergy list...</div>
    </div>
  )

  return (
    <div className="max-w-2xl mx-auto">
      <div className="flex items-center justify-between mb-5">
        <span style={{ fontFamily: 'Rajdhani', fontSize: '1rem', fontWeight: 700, color: '#c9a96e' }}>Add Allergy / Intolerance</span>
        <button onClick={onCancel} className="text-xs opacity-40 hover:opacity-70">✕ Cancel</button>
      </div>

      {/* Search */}
      <div ref={dropRef} className="relative mb-5">
        <input type="text" placeholder="Search allergen (drug name, food, environmental...)"
          value={query} onChange={e => { setQuery(e.target.value); setSelected(null) }}
          onFocus={() => results.length > 0 && setShowDrop(true)}
          className="w-full px-4 py-3 rounded text-sm"
          style={{ background: 'rgba(0,212,255,0.05)', border: '1px solid rgba(0,212,255,0.25)', color: 'rgba(200,215,235,0.9)', outline: 'none' }} />
        {searching && <div className="absolute right-3 top-3 text-sm opacity-40 animate-spin">⟳</div>}

        {showDrop && results.length > 0 && (
          <div className="absolute z-50 w-full mt-1 rounded shadow-lg"
            style={{ background: 'rgba(8,18,34,0.98)', border: '1px solid rgba(0,212,255,0.2)', maxHeight: 300, overflowY: 'auto' }}>
            {results.map((r, i) => {
              const sevCol = SEVERITY_COLORS[r.severity_potential || '']
              return (
                <button key={i} onClick={() => { setSelected(r); setQuery(r.display_name || r.substance_name); setShowDrop(false) }}
                  className="w-full text-left px-4 py-2.5 transition-all hover:bg-white/5 border-b last:border-b-0"
                  style={{ borderColor: 'rgba(255,255,255,0.04)' }}>
                  <div className="flex items-center justify-between gap-2">
                    <div className="min-w-0">
                      <div className="flex items-center gap-2">
                        <span>{TYPE_ICONS[r.substance_type] || '⚕️'}</span>
                        <span className="text-sm" style={{ color: 'rgba(200,215,235,0.9)' }}>{r.display_name || r.substance_name}</span>
                      </div>
                      {r.drug_class && r.drug_class.length > 0 && (
                        <div className="text-xs opacity-40 ml-6">{r.drug_class[0]}</div>
                      )}
                    </div>
                    <div className="flex items-center gap-2 flex-shrink-0">
                      {r.severity_potential && sevCol && (
                        <span className="text-xs px-1.5 py-0.5 rounded font-bold"
                          style={{ background: `${sevCol.text}18`, color: sevCol.text, fontFamily: 'Rajdhani' }}>
                          {r.severity_potential}
                        </span>
                      )}
                      <span className="text-xs opacity-25">{r.fhir_category}</span>
                    </div>
                  </div>
                </button>
              )
            })}
          </div>
        )}
      </div>

      {/* Selected allergen card */}
      {selected && (
        <div className="animate-fadeIn">
          <div className="p-4 rounded mb-4" style={{ background: 'rgba(10,22,40,0.8)', border: '1px solid rgba(0,212,255,0.15)' }}>
            <div className="flex items-start justify-between gap-3">
              <div>
                <div className="font-bold mb-1" style={{ fontFamily: 'Rajdhani', fontSize: '1.1rem', color: '#c9a96e' }}>
                  {TYPE_ICONS[selected.substance_type] || '⚕️'} {selected.display_name || selected.substance_name}
                </div>
                {selected.drug_class && <div className="text-xs opacity-50 mb-1">{selected.drug_class[0]}</div>}
                {selected.ige_mediated != null && (
                  <div className="text-xs opacity-40">{selected.ige_mediated ? 'IgE-mediated (true allergy)' : 'Non-IgE (intolerance/pharmacological)'}</div>
                )}
              </div>
              {selected.rxnorm_cui && (
                <div className="text-xs font-mono opacity-30">RxCUI {selected.rxnorm_cui}</div>
              )}
            </div>

            {selected.cross_reactive_classes && selected.cross_reactive_classes.length > 0 && (
              <div className="mt-3 p-2 rounded text-xs"
                style={{ background: 'rgba(230,126,34,0.08)', border: '1px solid rgba(230,126,34,0.2)', color: 'rgba(230,126,34,0.85)' }}>
                ⚡ Cross-reactivity: {selected.cross_reactive_classes.join(' · ')}
              </div>
            )}

            {selected.common_reactions && selected.common_reactions.length > 0 && (
              <div className="mt-2 text-xs opacity-40">
                Typical reactions: {selected.common_reactions.slice(0, 3).join(', ')}
              </div>
            )}
          </div>

          {/* Reaction + severity fields */}
          <div className="grid grid-cols-2 gap-3 mb-4">
            <div>
              <div className="hud-label text-xs mb-1.5">Reaction <span style={{ color: '#ff3b3b' }}>*</span></div>
              <select value={reaction} onChange={e => setReaction(e.target.value)}
                className="w-full px-3 py-2 rounded text-xs"
                style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.85)', outline: 'none' }}>
                <option value="">Select reaction...</option>
                {REACTIONS.map(r => <option key={r} value={r}>{r}</option>)}
              </select>
            </div>
            <div>
              <div className="hud-label text-xs mb-1.5">Severity <span style={{ color: '#ff3b3b' }}>*</span></div>
              <select value={severity} onChange={e => setSeverity(e.target.value)}
                className="w-full px-3 py-2 rounded text-xs"
                style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.85)', outline: 'none' }}>
                <option value="">Select severity...</option>
                {SEVERITY_OPTIONS.map(s => <option key={s} value={s}>{s.charAt(0).toUpperCase() + s.slice(1)}</option>)}
              </select>
            </div>
            <div>
              <div className="hud-label text-xs mb-1.5">Onset Date</div>
              <input type="date" value={onsetDate} onChange={e => setOnsetDate(e.target.value)}
                className="w-full px-3 py-2 rounded text-xs"
                style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.85)', outline: 'none' }} />
            </div>
            <div>
              <div className="hud-label text-xs mb-1.5">Notes</div>
              <input type="text" value={notes} onChange={e => setNotes(e.target.value)}
                placeholder="Additional clinical notes..."
                className="w-full px-3 py-2 rounded text-xs"
                style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.85)', outline: 'none' }} />
            </div>
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
            <button onClick={handleSave} disabled={!selected || !reaction || !severity || saving}
              className="flex-1 py-3 rounded text-sm font-bold disabled:opacity-30 transition-all"
              style={{ background: 'rgba(201,169,110,0.15)', border: '1px solid rgba(201,169,110,0.5)', color: '#c9a96e', fontFamily: 'Rajdhani', letterSpacing: '0.08em' }}>
              {saving ? 'Recording...' : `Record Allergy to ${selected.display_name || selected.substance_name} →`}
            </button>
          </div>
        </div>
      )}
    </div>
  )
}
