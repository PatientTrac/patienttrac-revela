import { useState } from 'react'
import { Medication } from './MedicationsModule'

interface Props {
  meds: Medication[]
  loading: boolean
  onDiscontinue: (medicationId: number) => Promise<void>
  onAddClick: () => void
}

const CLASS_COLORS: Record<string, { bg: string; text: string }> = {
  'selective serotonin':  { bg: 'rgba(52,152,219,0.12)',  text: '#3498db' },
  'serotonin-norepineph': { bg: 'rgba(52,152,219,0.12)',  text: '#3498db' },
  'norepinephrine-dopam': { bg: 'rgba(52,152,219,0.12)',  text: '#3498db' },
  'benzodiazepine':       { bg: 'rgba(231,76,60,0.12)',   text: '#e74c3c' },
  'antipsychotic':        { bg: 'rgba(142,68,173,0.12)',  text: '#8e44ad' },
  'antimanic':            { bg: 'rgba(142,68,173,0.12)',  text: '#8e44ad' },
  'stimulant':            { bg: 'rgba(230,126,34,0.12)',  text: '#e67e22' },
  'opioid':               { bg: 'rgba(231,76,60,0.12)',   text: '#e74c3c' },
  'ace inhibitor':        { bg: 'rgba(39,174,96,0.12)',   text: '#27ae60' },
  'statin':               { bg: 'rgba(39,174,96,0.12)',   text: '#27ae60' },
  'default':              { bg: 'rgba(0,212,255,0.08)',   text: '#00d4ff' },
}

function getClassColor(cls: string | null) {
  if (!cls) return CLASS_COLORS.default
  const lower = cls.toLowerCase()
  for (const [key, val] of Object.entries(CLASS_COLORS)) {
    if (key !== 'default' && lower.includes(key)) return val
  }
  return CLASS_COLORS.default
}

const CONTROLLED_COLORS: Record<string, string> = {
  'II': '#ff3b3b', 'III': '#e67e22', 'IV': '#f39c12', 'V': '#f1c40f',
}

function isControlledFromName(drugName: string): { controlled: boolean; schedule: string | null } {
  const lower = drugName.toLowerCase()
  const controlled_keywords = ['alprazolam','lorazepam','clonazepam','diazepam',
    'methylphenidate','amphetamine','adderall','ritalin','xanax','ativan','klonopin',
    'oxycodone','hydrocodone','morphine','fentanyl','buprenorphine','methadone',
    'zolpidem','tramadol','pregabalin']
  const schedule2 = ['methylphenidate','amphetamine','oxycodone','hydrocodone','fentanyl','morphine']
  const schedule3 = ['buprenorphine']
  const schedule4 = ['alprazolam','lorazepam','clonazepam','diazepam','zolpidem','tramadol']

  if (schedule2.some(k => lower.includes(k))) return { controlled: true, schedule: 'II' }
  if (schedule3.some(k => lower.includes(k))) return { controlled: true, schedule: 'III' }
  if (schedule4.some(k => lower.includes(k))) return { controlled: true, schedule: 'IV' }
  if (controlled_keywords.some(k => lower.includes(k))) return { controlled: true, schedule: 'IV' }
  return { controlled: false, schedule: null }
}

export default function MedList({ meds, loading, onDiscontinue, onAddClick }: Props) {
  const [confirmDisc, setConfirmDisc] = useState<number | null>(null)
  const [discReason, setDiscReason] = useState('')
  const [discarding, setDiscarding] = useState(false)
  const [expandedId, setExpandedId] = useState<number | null>(null)

  const handleDiscontinue = async (id: number) => {
    setDiscarding(true)
    await onDiscontinue(id)
    setConfirmDisc(null)
    setDiscReason('')
    setDiscarding(false)
  }

  if (loading) return (
    <div className="py-12 text-center hud-label animate-pulse">Loading medications...</div>
  )

  if (meds.length === 0) return (
    <div className="text-center py-16">
      <div style={{ fontSize: '3rem', marginBottom: 12, opacity: 0.3 }}>💊</div>
      <div className="text-sm opacity-40 mb-4">No active medications on record.</div>
      <button onClick={onAddClick}
        className="px-6 py-2.5 rounded text-sm font-bold"
        style={{ background: 'rgba(201,169,110,0.12)', border: '1px solid rgba(201,169,110,0.4)', color: '#c9a96e', fontFamily: 'Rajdhani', letterSpacing: '0.06em' }}>
        + Add First Medication
      </button>
    </div>
  )

  return (
    <div>
      {/* Header */}
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-2">
          <span style={{ fontFamily: 'Rajdhani', fontWeight: 700, color: '#c9a96e' }}>
            Active Medications
          </span>
          <span className="px-2 py-0.5 rounded text-xs"
            style={{ background: 'rgba(0,212,255,0.08)', color: '#00d4ff', border: '1px solid rgba(0,212,255,0.15)', fontFamily: 'Rajdhani' }}>
            {meds.length}
          </span>
        </div>
        <button onClick={onAddClick}
          className="px-4 py-1.5 rounded text-xs font-bold"
          style={{ background: 'rgba(201,169,110,0.1)', border: '1px solid rgba(201,169,110,0.3)', color: '#c9a96e', fontFamily: 'Rajdhani', letterSpacing: '0.06em' }}>
          + Add Medication
        </button>
      </div>

      {/* Medication rows */}
      <div className="space-y-2">
        {meds.map(med => {
          const classColor = getClassColor(med.drug_class)
          const isExpanded = expandedId === med.medication_id
          const isConfirmingDisc = confirmDisc === med.medication_id
          const { controlled, schedule } = isControlledFromName(med.drug_name)

          return (
            <div key={med.medication_id} className="rounded overflow-hidden transition-all"
              style={{ border: '1px solid rgba(255,255,255,0.06)', background: 'rgba(10,22,40,0.7)' }}>

              {/* Main row */}
              <div className="flex items-center gap-3 px-4 py-3 cursor-pointer hover:bg-white/3"
                onClick={() => setExpandedId(isExpanded ? null : med.medication_id)}>

                {/* Drug class color strip */}
                <div className="w-1 h-10 rounded-full flex-shrink-0"
                  style={{ background: classColor.text, opacity: 0.6 }} />

                {/* Drug info */}
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="font-semibold text-sm"
                      style={{ color: 'rgba(200,215,235,0.92)' }}>
                      {med.drug_name}
                    </span>
                    {controlled && schedule && (
                      <span className="text-xs px-1.5 py-0.5 rounded font-bold"
                        style={{
                          background: `${CONTROLLED_COLORS[schedule]}18`,
                          color: CONTROLLED_COLORS[schedule],
                          border: `1px solid ${CONTROLLED_COLORS[schedule]}40`,
                          fontFamily: 'Rajdhani',
                        }}>
                        C-{schedule}
                      </span>
                    )}
                    {med.drug_class && (
                      <span className="text-xs px-2 py-0.5 rounded"
                        style={{ background: classColor.bg, color: classColor.text, border: `1px solid ${classColor.text}25` }}>
                        {med.drug_class}
                      </span>
                    )}
                  </div>
                  <div className="text-xs mt-0.5 opacity-45">
                    {[med.strength, med.dose_form, med.route, med.frequency].filter(Boolean).join(' · ')}
                  </div>
                </div>

                {/* Date + expand */}
                <div className="text-right flex-shrink-0">
                  <div className="text-xs opacity-30">
                    {med.start_date ? new Date(med.start_date).toLocaleDateString() : '—'}
                  </div>
                  <div className="text-xs opacity-20 mt-0.5">{isExpanded ? '▲' : '▼'}</div>
                </div>
              </div>

              {/* Expanded detail */}
              {isExpanded && (
                <div className="px-4 pb-4 border-t" style={{ borderColor: 'rgba(255,255,255,0.05)' }}>
                  <div className="grid grid-cols-3 gap-3 mt-3 mb-3">
                    {[
                      { label: 'Strength',  value: med.strength || '—' },
                      { label: 'Dose form', value: med.dose_form || '—' },
                      { label: 'Route',     value: med.route || '—' },
                      { label: 'Frequency', value: med.frequency || '—' },
                      { label: 'Start date',value: med.start_date ? new Date(med.start_date).toLocaleDateString() : '—' },
                      { label: 'RxNorm CUI',value: med.rxnorm_cui || '—' },
                    ].map(({ label, value }) => (
                      <div key={label} className="p-2 rounded"
                        style={{ background: 'rgba(0,0,0,0.2)' }}>
                        <div className="text-xs opacity-35 mb-0.5">{label}</div>
                        <div className="text-xs font-mono" style={{ color: 'rgba(200,215,235,0.7)' }}>{value}</div>
                      </div>
                    ))}
                  </div>

                  {med.notes && (
                    <div className="text-xs opacity-50 mb-3 px-1">{med.notes}</div>
                  )}

                  {/* Discontinue */}
                  {!isConfirmingDisc ? (
                    <button
                      onClick={() => setConfirmDisc(med.medication_id)}
                      className="text-xs px-3 py-1.5 rounded"
                      style={{ background: 'rgba(231,76,60,0.08)', border: '1px solid rgba(231,76,60,0.25)', color: '#e74c3c', fontFamily: 'Rajdhani' }}>
                      Discontinue
                    </button>
                  ) : (
                    <div className="p-3 rounded"
                      style={{ background: 'rgba(231,76,60,0.07)', border: '1px solid rgba(231,76,60,0.25)' }}>
                      <div className="text-xs font-bold mb-2" style={{ color: '#e74c3c', fontFamily: 'Rajdhani' }}>
                        Confirm Discontinuation
                      </div>
                      <input type="text" value={discReason} onChange={e => setDiscReason(e.target.value)}
                        placeholder="Reason (optional)..."
                        className="w-full px-2 py-1.5 rounded text-xs mb-2"
                        style={{ background: 'rgba(0,0,0,0.3)', border: '1px solid rgba(231,76,60,0.3)', color: 'rgba(200,215,235,0.7)', outline: 'none' }} />
                      <div className="flex gap-2">
                        <button onClick={() => setConfirmDisc(null)}
                          className="px-3 py-1.5 rounded text-xs"
                          style={{ background: 'rgba(255,255,255,0.04)', border: '1px solid rgba(255,255,255,0.1)', color: 'rgba(200,215,235,0.5)', fontFamily: 'Rajdhani' }}>
                          Cancel
                        </button>
                        <button onClick={() => handleDiscontinue(med.medication_id)} disabled={discarding}
                          className="flex-1 py-1.5 rounded text-xs font-bold disabled:opacity-40"
                          style={{ background: 'rgba(231,76,60,0.15)', border: '1px solid rgba(231,76,60,0.4)', color: '#e74c3c', fontFamily: 'Rajdhani' }}>
                          {discarding ? 'Discontinuing...' : 'Confirm Discontinue'}
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          )
        })}
      </div>

      {/* RxNorm attribution */}
      <div className="mt-4 text-xs opacity-25 text-right">
        Drug data via NLM RxNav · RxNorm · NDF-RT
      </div>
    </div>
  )
}
