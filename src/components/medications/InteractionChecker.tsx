import { useState, useEffect, useCallback } from 'react'
import { callMeds, Medication } from './MedicationsModule'

interface Interaction {
  rxcui_1: string
  rxcui_2: string
  drug_name_1: string
  drug_name_2: string
  severity: 'high' | 'moderate' | 'low'
  description: string
  source?: string
}

interface Props {
  meds: Medication[]
  orgId: string
}

const SEVERITY_CONFIG = {
  high:     { color: '#ff3b3b', bg: 'rgba(255,59,59,0.1)',  border: 'rgba(255,59,59,0.35)',  icon: '🚨', label: 'High' },
  moderate: { color: '#e67e22', bg: 'rgba(230,126,34,0.1)', border: 'rgba(230,126,34,0.35)', icon: '⚠️', label: 'Moderate' },
  low:      { color: '#f1c40f', bg: 'rgba(241,196,15,0.08)', border: 'rgba(241,196,15,0.25)', icon: 'ℹ️', label: 'Low' },
}

export default function InteractionChecker({ meds, orgId }: Props) {
  const [interactions, setInteractions] = useState<Interaction[]>([])
  const [loading, setLoading] = useState(false)
  const [lastChecked, setLastChecked] = useState<Date | null>(null)
  const [filter, setFilter] = useState<'all' | 'high' | 'moderate' | 'low'>('all')
  const [expandedIdx, setExpandedIdx] = useState<number | null>(null)

  const activeMedsWithCUI = meds.filter(m => m.rxnorm_cui && m.is_active)

  const runCheck = useCallback(async () => {
    const cuis = activeMedsWithCUI.map(m => m.rxnorm_cui)
    if (cuis.length < 2) return
    setLoading(true)
    try {
      const d = await callMeds({ action: 'check_interactions', rxcuis: cuis })
      setInteractions(d.interactions || [])
      setLastChecked(new Date())
    } finally {
      setLoading(false)
    }
  }, [activeMedsWithCUI.map(m => m.rxnorm_cui).join(',')])

  useEffect(() => { runCheck() }, [runCheck])

  const filtered = interactions.filter(i => filter === 'all' || i.severity === filter)
  const highCount = interactions.filter(i => i.severity === 'high').length
  const modCount  = interactions.filter(i => i.severity === 'moderate').length
  const lowCount  = interactions.filter(i => i.severity === 'low').length

  // Find the medication name from our list (fallback to API name)
  const getMedName = (rxcui: string, apiName: string) => {
    const med = meds.find(m => m.rxnorm_cui === rxcui)
    return med?.drug_name || apiName || rxcui
  }

  if (activeMedsWithCUI.length < 2) return (
    <div className="text-center py-16">
      <div style={{ fontSize: '3rem', marginBottom: 12, opacity: 0.3 }}>⚡</div>
      <div className="text-sm opacity-40 mb-2">Need at least 2 active medications with RxNorm codes to check interactions.</div>
      <div className="text-xs opacity-25">
        {activeMedsWithCUI.length === 0 ? 'No medications with RxNorm CUI found.' :
         `${activeMedsWithCUI.length} medication(s) found — need at least 2.`}
      </div>
    </div>
  )

  return (
    <div>
      {/* Header */}
      <div className="flex items-center justify-between mb-4">
        <div>
          <div style={{ fontFamily: 'Rajdhani', fontWeight: 700, color: '#c9a96e' }}>
            Drug Interaction Check
          </div>
          <div className="text-xs opacity-30 mt-0.5">
            Checking {activeMedsWithCUI.length} medications · via NLM RxNav
            {lastChecked && ` · ${lastChecked.toLocaleTimeString()}`}
          </div>
        </div>
        <button onClick={runCheck} disabled={loading}
          className="px-3 py-1.5 rounded text-xs disabled:opacity-40"
          style={{ background: 'rgba(0,212,255,0.07)', border: '1px solid rgba(0,212,255,0.2)', color: '#00d4ff', fontFamily: 'Rajdhani' }}>
          {loading ? '⟳ Checking...' : '↻ Re-check'}
        </button>
      </div>

      {/* Medications being checked */}
      <div className="p-3 rounded mb-4"
        style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.1)' }}>
        <div className="hud-label text-xs mb-2">Medications in Check</div>
        <div className="flex flex-wrap gap-1.5">
          {activeMedsWithCUI.map(m => (
            <span key={m.medication_id} className="px-2 py-1 rounded text-xs"
              style={{ background: 'rgba(0,212,255,0.08)', color: 'rgba(200,215,235,0.7)', border: '1px solid rgba(0,212,255,0.12)' }}>
              {m.drug_name.split(' ').slice(0, 2).join(' ')}
              <span className="opacity-30 ml-1 font-mono text-xs">{m.rxnorm_cui}</span>
            </span>
          ))}
        </div>
      </div>

      {loading ? (
        <div className="py-10 text-center hud-label animate-pulse">Checking interactions via NLM...</div>
      ) : (
        <>
          {/* Summary counts */}
          <div className="grid grid-cols-4 gap-2 mb-4">
            {[
              { key: 'all',      label: 'Total',    count: interactions.length, color: '#00d4ff' },
              { key: 'high',     label: 'High',     count: highCount,           color: '#ff3b3b' },
              { key: 'moderate', label: 'Moderate', count: modCount,            color: '#e67e22' },
              { key: 'low',      label: 'Low',      count: lowCount,            color: '#f1c40f' },
            ].map(({ key, label, count, color }) => (
              <button key={key}
                onClick={() => setFilter(key as typeof filter)}
                className="p-2.5 rounded text-center transition-all"
                style={{
                  background: filter === key ? `${color}15` : 'rgba(10,22,40,0.7)',
                  border: filter === key ? `1px solid ${color}50` : '1px solid rgba(255,255,255,0.06)',
                }}>
                <div className="text-xl font-bold" style={{ fontFamily: 'Rajdhani', color }}>{count}</div>
                <div className="text-xs opacity-50">{label}</div>
              </button>
            ))}
          </div>

          {/* No interactions */}
          {interactions.length === 0 && (
            <div className="text-center py-12">
              <div style={{ fontSize: '3rem', marginBottom: 12 }}>✅</div>
              <div className="text-sm font-semibold" style={{ color: '#27ae60', fontFamily: 'Rajdhani' }}>
                No known interactions detected
              </div>
              <div className="text-xs opacity-40 mt-1">
                Checked {activeMedsWithCUI.length} medications via NLM RxNav DDI database
              </div>
            </div>
          )}

          {/* Interaction list */}
          {filtered.length > 0 && (
            <div className="space-y-2">
              {filtered.map((ix, idx) => {
                const sev = SEVERITY_CONFIG[ix.severity] || SEVERITY_CONFIG.moderate
                const isExpanded = expandedIdx === idx
                const name1 = getMedName(ix.rxcui_1, ix.drug_name_1)
                const name2 = getMedName(ix.rxcui_2, ix.drug_name_2)

                return (
                  <div key={idx} className="rounded overflow-hidden"
                    style={{ border: `1px solid ${sev.border}`, background: sev.bg }}>

                    {/* Header row */}
                    <button className="w-full flex items-start gap-3 px-4 py-3 text-left"
                      onClick={() => setExpandedIdx(isExpanded ? null : idx)}>
                      <span style={{ fontSize: '1.1rem', lineHeight: 1, flexShrink: 0, marginTop: 1 }}>{sev.icon}</span>
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="text-xs font-bold px-2 py-0.5 rounded"
                            style={{ background: `${sev.color}20`, color: sev.color, border: `1px solid ${sev.color}40`, fontFamily: 'Rajdhani' }}>
                            {sev.label}
                          </span>
                          <span className="text-sm font-semibold" style={{ color: 'rgba(200,215,235,0.9)' }}>
                            {name1.split(' ').slice(0, 2).join(' ')}
                            <span className="opacity-40 mx-1">↔</span>
                            {name2.split(' ').slice(0, 2).join(' ')}
                          </span>
                        </div>
                        {!isExpanded && (
                          <div className="text-xs mt-1 opacity-50 truncate">{ix.description}</div>
                        )}
                      </div>
                      <span className="text-xs opacity-20 flex-shrink-0">{isExpanded ? '▲' : '▼'}</span>
                    </button>

                    {/* Expanded detail */}
                    {isExpanded && (
                      <div className="px-4 pb-4 border-t" style={{ borderColor: `${sev.color}20` }}>
                        {/* Drug pair */}
                        <div className="grid grid-cols-2 gap-2 my-3">
                          {[{ name: name1, cui: ix.rxcui_1 }, { name: name2, cui: ix.rxcui_2 }].map(({ name, cui }) => (
                            <div key={cui} className="p-2 rounded"
                              style={{ background: 'rgba(0,0,0,0.2)', border: '1px solid rgba(255,255,255,0.05)' }}>
                              <div className="text-xs font-semibold" style={{ color: 'rgba(200,215,235,0.8)' }}>{name}</div>
                              <div className="text-xs font-mono opacity-30 mt-0.5">RxCUI {cui}</div>
                            </div>
                          ))}
                        </div>

                        {/* Description */}
                        <div className="p-3 rounded mb-3"
                          style={{ background: 'rgba(0,0,0,0.2)', border: '1px solid rgba(255,255,255,0.05)' }}>
                          <div className="hud-label text-xs mb-1">Clinical Description</div>
                          <p className="text-sm leading-relaxed" style={{ color: 'rgba(200,215,235,0.75)' }}>
                            {ix.description}
                          </p>
                        </div>

                        {/* Source */}
                        {ix.source && (
                          <div className="text-xs opacity-30 mb-3">Source: {ix.source}</div>
                        )}

                        {/* Guidance for high severity */}
                        {ix.severity === 'high' && (
                          <div className="p-3 rounded mb-3"
                            style={{ background: 'rgba(255,59,59,0.08)', border: '1px solid rgba(255,59,59,0.2)' }}>
                            <div className="text-xs font-bold mb-1" style={{ color: '#ff3b3b', fontFamily: 'Rajdhani' }}>
                              Recommended Actions
                            </div>
                            <ul className="text-xs space-y-1" style={{ color: 'rgba(255,150,150,0.7)' }}>
                              <li>• Review clinical necessity of both medications</li>
                              <li>• Consider alternative agents from different drug classes</li>
                              <li>• If both required: increased monitoring, reduced doses, or staggered timing</li>
                              <li>• Document clinical rationale in encounter note</li>
                            </ul>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )
              })}
            </div>
          )}

          {/* NLM attribution */}
          <div className="mt-5 p-3 rounded text-xs"
            style={{ background: 'rgba(0,212,255,0.03)', border: '1px solid rgba(0,212,255,0.08)', color: 'rgba(200,215,235,0.3)' }}>
            Interaction data sourced from the National Library of Medicine (NLM) RxNav Drug Interaction API.
            This is a screening tool — always apply clinical judgment. Not a substitute for pharmacist review.
          </div>
        </>
      )}
    </div>
  )
}
