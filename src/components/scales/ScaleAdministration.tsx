import { useState, useEffect, useRef } from 'react'
import { callScales, ScaleMeta, ScaleAdminResult } from './RatingScalesModule'

interface ScaleItem {
  item_id: number
  item_number: number
  item_code: string
  subscale: string | null
  item_text: string
  item_short: string
  response_type: string
  min_value: number
  max_value: number
  anchor_min: string
  anchor_max: string
  anchor_labels: Record<string, string> | null
  is_required: boolean
}

interface Props {
  orgId: string
  patientId: string
  encounterId: string
  providerId: string
  initialScale: ScaleMeta | null
  onComplete: (result: ScaleAdminResult) => void
}

const SEVERITY_COLORS: Record<string, string> = {
  green:  '#27ae60',
  yellow: '#f39c12',
  orange: '#e67e22',
  red:    '#e74c3c',
  purple: '#8e44ad',
}

export default function ScaleAdministration({ orgId, patientId, encounterId, providerId, initialScale, onComplete }: Props) {
  const [scale, setScale] = useState<ScaleMeta | null>(initialScale)
  const [items, setItems] = useState<ScaleItem[]>([])
  const [responses, setResponses] = useState<Record<number, number>>({})
  const [loading, setLoading] = useState(false)
  const [submitting, setSubmitting] = useState(false)
  const [result, setResult] = useState<ScaleAdminResult | null>(null)
  const [error, setError] = useState('')
  const topRef = useRef<HTMLDivElement>(null)

  // Load items when scale changes
  useEffect(() => {
    if (!scale) return
    setLoading(true)
    setResponses({})
    setResult(null)
    setError('')
    callScales({ action: 'get_items', org_id: orgId, scale_code: scale.scale_code })
      .then(d => setItems(d.items || []))
      .finally(() => setLoading(false))
  }, [scale?.scale_code])

  // Update scale if prop changes
  useEffect(() => {
    if (initialScale) setScale(initialScale)
  }, [initialScale?.scale_code])

  // Group items by subscale
  const grouped = items.reduce((acc, item) => {
    const key = item.subscale || '__main__'
    if (!acc[key]) acc[key] = []
    acc[key].push(item)
    return acc
  }, {} as Record<string, ScaleItem[]>)

  const answeredCount = Object.keys(responses).length
  const totalRequired = items.filter(i => i.is_required).length
  const progress = totalRequired > 0 ? Math.round((answeredCount / totalRequired) * 100) : 0
  const allAnswered = answeredCount >= totalRequired

  // Live running score (sum of responses so far)
  const runningScore = Object.values(responses).reduce((s, v) => s + v, 0)

  const handleSubmit = async () => {
    if (!scale || !allAnswered) return
    setSubmitting(true)
    setError('')
    try {
      const data = await callScales({
        action: 'score',
        org_id: orgId,
        patient_id: Number(patientId),
        encounter_id: encounterId ? Number(encounterId) : undefined,
        provider_id: Number(providerId),
        scale_code: scale.scale_code,
        responses,
        administered_by: 'provider',
        administration_mode: 'in_session',
      })
      if (!data.success) throw new Error(data.error || 'Scoring failed')
      setResult(data)
      onComplete(data)
      topRef.current?.scrollIntoView({ behavior: 'smooth' })
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Unknown error')
    } finally {
      setSubmitting(false)
    }
  }

  if (!scale) return (
    <div className="text-center py-12 opacity-40">
      <div className="text-4xl mb-3">📋</div>
      <div className="text-sm">Select a scale from the catalog to administer it.</div>
    </div>
  )

  // Result view
  if (result) return (
    <div ref={topRef} className="animate-fadeIn max-w-2xl mx-auto">
      {/* Score card */}
      <div className="p-6 rounded-lg mb-5"
        style={{ background: 'rgba(10,22,40,0.9)', border: `2px solid ${SEVERITY_COLORS[result.severity?.color_code || 'green']}44` }}>

        <div className="flex items-start justify-between mb-4">
          <div>
            <div className="hud-label mb-1">{result.scale_name} — Result</div>
            <div className="flex items-baseline gap-3">
              <span style={{ fontFamily: 'Rajdhani', fontSize: '3rem', fontWeight: 700, color: SEVERITY_COLORS[result.severity?.color_code || 'green'], lineHeight: 1 }}>
                {result.total_score}
              </span>
              <span className="opacity-40 text-lg">/ {result.max_score}</span>
            </div>
          </div>
          <div className="text-right">
            <div className="px-4 py-2 rounded-lg text-sm font-bold"
              style={{
                background: `${SEVERITY_COLORS[result.severity?.color_code || 'green']}18`,
                border: `1px solid ${SEVERITY_COLORS[result.severity?.color_code || 'green']}50`,
                color: SEVERITY_COLORS[result.severity?.color_code || 'green'],
                fontFamily: 'Rajdhani', letterSpacing: '0.05em',
              }}>
              {result.severity?.severity_label || 'No range match'}
            </div>
            {result.score_delta !== null && result.delta_direction !== 'first_admin' && (
              <div className="mt-2 text-xs" style={{ color: 'rgba(200,215,235,0.5)' }}>
                {result.score_delta > 0 ? '▲' : result.score_delta < 0 ? '▼' : '='} {Math.abs(result.score_delta)} pts vs last ({result.delta_direction})
              </div>
            )}
          </div>
        </div>

        {/* Safety alert */}
        {result.item_9_flag && (
          <div className="p-3 rounded mb-4 flex items-center gap-2"
            style={{ background: 'rgba(255,59,59,0.12)', border: '1px solid rgba(255,59,59,0.5)' }}>
            <span style={{ fontSize: '1.2rem' }}>🚨</span>
            <div>
              <div className="text-sm font-bold" style={{ color: '#ff3b3b', fontFamily: 'Rajdhani' }}>
                SAFETY ALERT — Suicidal ideation indicated
              </div>
              <div className="text-xs opacity-70 mt-0.5">Complete C-SSRS assessment. Document safety plan. Contact supervisor if needed.</div>
            </div>
          </div>
        )}

        {/* Clinical guidance */}
        {result.clinical_guidance && (
          <div className="grid grid-cols-2 gap-3">
            <div className="p-3 rounded" style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.1)' }}>
              <div className="hud-label text-xs mb-1">Clinical Note</div>
              <div className="text-sm opacity-80">{result.clinical_guidance}</div>
            </div>
            <div className="p-3 rounded" style={{ background: 'rgba(201,169,110,0.04)', border: '1px solid rgba(201,169,110,0.1)' }}>
              <div className="hud-label text-xs mb-1" style={{ color: '#c9a96e' }}>Recommended Action</div>
              <div className="text-sm opacity-80">{result.recommended_action}</div>
            </div>
          </div>
        )}

        {/* Subscale breakdown */}
        {Object.keys(result.subscale_scores || {}).length > 0 && (
          <div className="mt-4">
            <div className="hud-label text-xs mb-2">Subscale Scores</div>
            <div className="grid grid-cols-3 gap-2">
              {Object.entries(result.subscale_scores).map(([sub, score]) => (
                <div key={sub} className="p-2 rounded text-center"
                  style={{ background: 'rgba(0,0,0,0.3)', border: '1px solid rgba(255,255,255,0.06)' }}>
                  <div className="text-lg font-bold" style={{ fontFamily: 'Rajdhani', color: '#00d4ff' }}>{score}</div>
                  <div className="text-xs opacity-50 capitalize">{sub.replace('_', ' ')}</div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      <div className="flex gap-3">
        <button onClick={() => { setResult(null); setResponses({}) }}
          className="flex-1 py-2.5 rounded text-sm"
          style={{ background: 'rgba(201,169,110,0.1)', border: '1px solid rgba(201,169,110,0.3)', color: '#c9a96e', fontFamily: 'Rajdhani', fontWeight: 700 }}>
          Administer Again
        </button>
        <button onClick={() => { setScale(null); setResult(null) }}
          className="flex-1 py-2.5 rounded text-sm"
          style={{ background: 'rgba(0,212,255,0.08)', border: '1px solid rgba(0,212,255,0.25)', color: '#00d4ff', fontFamily: 'Rajdhani', fontWeight: 700 }}>
          Different Scale
        </button>
      </div>
    </div>
  )

  return (
    <div ref={topRef} className="max-w-2xl mx-auto">
      {/* Scale header */}
      <div className="flex items-center justify-between mb-4">
        <div>
          <div className="flex items-center gap-2">
            <span style={{ fontFamily: 'Rajdhani', fontSize: '1.25rem', fontWeight: 700, color: '#c9a96e' }}>
              {scale.scale_code}
            </span>
            <span className="text-sm opacity-50">{scale.scale_full_name || scale.scale_name}</span>
          </div>
          <div className="text-xs opacity-40 mt-0.5">
            {scale.total_items} items · ~{scale.estimated_minutes} min · CPT {scale.cpt_code}
          </div>
        </div>
        <button onClick={() => setScale(null)} className="text-xs opacity-40 hover:opacity-70">✕ Change scale</button>
      </div>

      {/* Progress bar */}
      <div className="mb-5">
        <div className="flex justify-between text-xs mb-1.5 opacity-50">
          <span>{answeredCount} of {totalRequired} answered</span>
          <span>{progress}%</span>
        </div>
        <div className="h-1 rounded" style={{ background: 'rgba(255,255,255,0.06)' }}>
          <div className="h-1 rounded transition-all duration-300"
            style={{ width: `${progress}%`, background: progress === 100 ? '#27ae60' : '#00d4ff' }} />
        </div>
      </div>

      {/* Running score */}
      {answeredCount > 0 && (
        <div className="flex items-center gap-2 mb-4 px-3 py-2 rounded"
          style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.1)' }}>
          <span className="text-xs opacity-50">Running score:</span>
          <span style={{ fontFamily: 'Rajdhani', fontWeight: 700, color: '#00d4ff' }}>{runningScore}</span>
          <span className="text-xs opacity-30">/ {scale.max_score}</span>
        </div>
      )}

      {loading ? (
        <div className="text-center py-10 hud-label animate-pulse">Loading items...</div>
      ) : (
        <div className="space-y-6">
          {Object.entries(grouped).map(([subscale, groupItems]) => (
            <div key={subscale}>
              {/* Subscale header */}
              {subscale !== '__main__' && (
                <div className="mb-3 pb-1 border-b" style={{ borderColor: 'rgba(0,212,255,0.1)' }}>
                  <span className="text-xs font-bold uppercase tracking-wider"
                    style={{ color: '#00d4ff', fontFamily: 'Rajdhani' }}>
                    {subscale.replace('_', ' ')}
                  </span>
                </div>
              )}

              {groupItems.map(item => (
                <div key={item.item_id} className="p-4 rounded mb-3 transition-all"
                  style={{
                    background: responses[item.item_number] !== undefined
                      ? 'rgba(0,212,255,0.04)' : 'rgba(10,22,40,0.6)',
                    border: responses[item.item_number] !== undefined
                      ? '1px solid rgba(0,212,255,0.2)' : '1px solid rgba(255,255,255,0.05)',
                  }}>
                  {/* Question */}
                  <div className="flex gap-3 mb-3">
                    <span className="flex-shrink-0 text-xs font-mono opacity-30 mt-0.5 w-5 text-right">
                      {item.item_number}
                    </span>
                    <p className="text-sm leading-relaxed" style={{ color: 'rgba(200,215,235,0.88)' }}>
                      {item.item_text}
                    </p>
                  </div>

                  {/* Likert response buttons */}
                  {item.response_type === 'likert' && item.anchor_labels && (
                    <div className="ml-8">
                      <div className="flex justify-between text-xs mb-2 opacity-35">
                        <span>{item.anchor_min}</span>
                        <span>{item.anchor_max}</span>
                      </div>
                      <div className="flex gap-2 flex-wrap">
                        {Object.entries(item.anchor_labels).map(([val, label]) => {
                          const numVal = Number(val)
                          const selected = responses[item.item_number] === numVal
                          return (
                            <button key={val}
                              onClick={() => setResponses(r => ({ ...r, [item.item_number]: numVal }))}
                              className="flex-1 min-w-0 py-2 px-1 rounded text-xs transition-all"
                              style={{
                                background: selected ? 'rgba(201,169,110,0.2)' : 'rgba(255,255,255,0.03)',
                                border: selected ? '1px solid rgba(201,169,110,0.6)' : '1px solid rgba(255,255,255,0.08)',
                                color: selected ? '#c9a96e' : 'rgba(200,215,235,0.5)',
                                fontWeight: selected ? 700 : 400,
                                fontFamily: selected ? 'Rajdhani' : undefined,
                              }}>
                              <div className="font-bold mb-0.5">{val}</div>
                              <div className="leading-tight" style={{ fontSize: '0.6rem' }}>{label}</div>
                            </button>
                          )
                        })}
                      </div>
                    </div>
                  )}

                  {/* Yes/No response */}
                  {item.response_type === 'yes_no' && (
                    <div className="ml-8 flex gap-3">
                      {[{ val: 0, label: 'No' }, { val: 1, label: 'Yes' }].map(({ val, label }) => {
                        const selected = responses[item.item_number] === val
                        return (
                          <button key={val}
                            onClick={() => setResponses(r => ({ ...r, [item.item_number]: val }))}
                            className="px-8 py-2 rounded text-sm font-bold transition-all"
                            style={{
                              fontFamily: 'Rajdhani', letterSpacing: '0.08em',
                              background: selected
                                ? (val === 1 ? 'rgba(255,59,59,0.15)' : 'rgba(0,212,255,0.1)')
                                : 'rgba(255,255,255,0.03)',
                              border: selected
                                ? (val === 1 ? '1px solid rgba(255,59,59,0.5)' : '1px solid rgba(0,212,255,0.4)')
                                : '1px solid rgba(255,255,255,0.08)',
                              color: selected
                                ? (val === 1 ? '#ff3b3b' : '#00d4ff')
                                : 'rgba(200,215,235,0.4)',
                            }}>
                            {label}
                          </button>
                        )
                      })}
                    </div>
                  )}
                </div>
              ))}
            </div>
          ))}
        </div>
      )}

      {/* Error */}
      {error && (
        <div className="mt-4 p-3 rounded text-sm" style={{ background: 'rgba(255,59,59,0.1)', color: '#ff3b3b', border: '1px solid rgba(255,59,59,0.3)' }}>
          {error}
        </div>
      )}

      {/* Submit */}
      {!loading && (
        <div className="mt-6 flex gap-3">
          <button onClick={() => { setScale(null); setResponses({}) }}
            className="px-5 py-3 rounded text-sm"
            style={{ background: 'rgba(255,255,255,0.03)', border: '1px solid rgba(255,255,255,0.1)', color: 'rgba(200,215,235,0.5)', fontFamily: 'Rajdhani' }}>
            Cancel
          </button>
          <button onClick={handleSubmit} disabled={!allAnswered || submitting}
            className="flex-1 py-3 rounded text-sm font-bold transition-all disabled:opacity-30"
            style={{ background: allAnswered ? 'rgba(201,169,110,0.15)' : 'rgba(255,255,255,0.03)', border: allAnswered ? '1px solid rgba(201,169,110,0.5)' : '1px solid rgba(255,255,255,0.08)', color: allAnswered ? '#c9a96e' : 'rgba(200,215,235,0.3)', fontFamily: 'Rajdhani', letterSpacing: '0.08em', fontSize: '0.9rem' }}>
            {submitting ? 'Scoring...' : !allAnswered ? `Answer all items (${totalRequired - answeredCount} remaining)` : `Score ${scale.scale_code} →`}
          </button>
        </div>
      )}
    </div>
  )
}
