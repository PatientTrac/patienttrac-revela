import { useState, useEffect, useRef } from 'react'
import { callScales, ScaleAdminResult } from './RatingScalesModule'

interface TrendEntry {
  admin_id: number
  scale_code: string
  scale_name: string
  administered_at: string
  total_score: number
  severity_label: string
  severity_code: string
  color_code: string
  score_delta: number | null
  delta_direction: string
  alert_triggered: boolean
  item_9_flag: boolean
}

interface PatientScaleSummary {
  scale_code: string
  total_score: number
  severity_label: string
  severity_code: string
  color_code: string
  score_delta: number | null
  delta_direction: string
  alert_triggered: boolean
  administered_at: string
  terms_rating_scale_catalog: { scale_name: string; category: string; max_score: number }
}

interface Props {
  orgId: string
  patientId: string
  lastResult: ScaleAdminResult | null
  onAdminister: (scale: { scale_code: string; scale_name: string; [k: string]: unknown }) => void
}

const SEVERITY_COLORS: Record<string, string> = {
  green: '#27ae60', yellow: '#f39c12', orange: '#e67e22', red: '#e74c3c', purple: '#8e44ad',
}

const DELTA_ICONS: Record<string, string> = {
  improved: '▼', worsened: '▲', stable: '=', first_admin: '●',
}
const DELTA_COLORS: Record<string, string> = {
  improved: '#27ae60', worsened: '#e74c3c', stable: '#95a5a6', first_admin: '#00d4ff',
}

export default function ScaleTrend({ orgId, patientId, lastResult, onAdminister }: Props) {
  const [summary, setSummary] = useState<PatientScaleSummary[]>([])
  const [selected, setSelected] = useState<string | null>(null)
  const [trend, setTrend] = useState<TrendEntry[]>([])
  const [loading, setLoading] = useState(false)
  const [loadingTrend, setLoadingTrend] = useState(false)
  const canvasRef = useRef<HTMLCanvasElement>(null)

  // Load patient summary on mount / when new result comes in
  useEffect(() => {
    if (!patientId) return
    setLoading(true)
    callScales({ action: 'patient_summary', org_id: orgId, patient_id: Number(patientId) })
      .then(d => setSummary(d.scales || []))
      .finally(() => setLoading(false))
  }, [patientId, orgId, lastResult?.admin_id])

  // Auto-select last result's scale
  useEffect(() => {
    if (lastResult?.scale_code) setSelected(lastResult.scale_code)
  }, [lastResult?.scale_code])

  // Load trend when scale selected
  useEffect(() => {
    if (!selected || !patientId) return
    setLoadingTrend(true)
    callScales({ action: 'get_trend', org_id: orgId, patient_id: Number(patientId), scale_code: selected, limit: 12 })
      .then(d => setTrend((d.trend || []).reverse()))
      .finally(() => setLoadingTrend(false))
  }, [selected, patientId, orgId])

  // Draw chart when trend data loads
  useEffect(() => {
    if (!canvasRef.current || trend.length === 0) return
    const canvas = canvasRef.current
    const ctx = canvas.getContext('2d')
    if (!ctx) return

    const W = canvas.width = canvas.offsetWidth * window.devicePixelRatio
    const H = canvas.height = 160 * window.devicePixelRatio
    ctx.scale(window.devicePixelRatio, window.devicePixelRatio)
    const w = canvas.offsetWidth
    const h = 160
    ctx.clearRect(0, 0, w, h)

    const pad = { top: 16, right: 16, bottom: 32, left: 40 }
    const chartW = w - pad.left - pad.right
    const chartH = h - pad.top - pad.bottom

    // Find scale max
    const maxScore = summary.find(s => s.scale_code === selected)?.terms_rating_scale_catalog?.max_score || 100
    const scores = trend.map(t => t.total_score)
    const minY = 0
    const maxY = maxScore

    const xStep = trend.length > 1 ? chartW / (trend.length - 1) : chartW
    const toX = (i: number) => pad.left + i * xStep
    const toY = (v: number) => pad.top + chartH - ((v - minY) / (maxY - minY)) * chartH

    // Grid lines
    ctx.strokeStyle = 'rgba(255,255,255,0.05)'
    ctx.lineWidth = 0.5
    for (let i = 0; i <= 4; i++) {
      const y = pad.top + (chartH / 4) * i
      ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(pad.left + chartW, y); ctx.stroke()
      ctx.fillStyle = 'rgba(200,215,235,0.25)'
      ctx.font = '10px DM Mono, monospace'
      ctx.textAlign = 'right'
      ctx.fillText(String(Math.round(maxY - (maxY / 4) * i)), pad.left - 6, y + 4)
    }

    // Gradient fill under line
    const grad = ctx.createLinearGradient(0, pad.top, 0, pad.top + chartH)
    grad.addColorStop(0, 'rgba(0,212,255,0.18)')
    grad.addColorStop(1, 'rgba(0,212,255,0)')
    ctx.beginPath()
    ctx.moveTo(toX(0), toY(scores[0]))
    scores.forEach((s, i) => { if (i > 0) ctx.lineTo(toX(i), toY(s)) })
    ctx.lineTo(toX(scores.length - 1), pad.top + chartH)
    ctx.lineTo(toX(0), pad.top + chartH)
    ctx.closePath()
    ctx.fillStyle = grad
    ctx.fill()

    // Line
    ctx.beginPath()
    ctx.strokeStyle = '#00d4ff'
    ctx.lineWidth = 2
    ctx.lineJoin = 'round'
    scores.forEach((s, i) => { i === 0 ? ctx.moveTo(toX(i), toY(s)) : ctx.lineTo(toX(i), toY(s)) })
    ctx.stroke()

    // Dots + labels
    trend.forEach((t, i) => {
      const x = toX(i)
      const y = toY(t.total_score)
      const color = SEVERITY_COLORS[t.color_code] || '#00d4ff'

      // Alert ring
      if (t.alert_triggered) {
        ctx.beginPath(); ctx.arc(x, y, 8, 0, Math.PI * 2)
        ctx.strokeStyle = 'rgba(255,59,59,0.5)'; ctx.lineWidth = 1.5; ctx.stroke()
      }

      // Dot
      ctx.beginPath(); ctx.arc(x, y, 4, 0, Math.PI * 2)
      ctx.fillStyle = color; ctx.fill()

      // Score label
      ctx.fillStyle = 'rgba(200,215,235,0.7)'
      ctx.font = '10px DM Mono, monospace'
      ctx.textAlign = 'center'
      ctx.fillText(String(t.total_score), x, y - 10)

      // Date label
      const date = new Date(t.administered_at)
      ctx.fillStyle = 'rgba(200,215,235,0.3)'
      ctx.font = '9px DM Mono, monospace'
      ctx.fillText(`${date.getMonth() + 1}/${date.getDate()}`, x, h - 6)
    })

  }, [trend, selected, summary])

  const selectedSummary = summary.find(s => s.scale_code === selected)

  return (
    <div>
      {/* Patient scale summary cards */}
      <div className="mb-5">
        <div className="hud-label mb-3">Active Scales — Latest Scores</div>
        {loading ? (
          <div className="hud-label animate-pulse">Loading...</div>
        ) : summary.length === 0 ? (
          <div className="text-sm opacity-40 py-4 text-center">No scales administered yet for this patient.</div>
        ) : (
          <div className="grid grid-cols-2 gap-2">
            {summary.map((s) => {
              const color = SEVERITY_COLORS[s.color_code] || '#00d4ff'
              const isSelected = selected === s.scale_code
              return (
                <button key={s.scale_code}
                  onClick={() => setSelected(s.scale_code)}
                  className="p-3 rounded text-left transition-all"
                  style={{
                    background: isSelected ? `${color}12` : 'rgba(10,22,40,0.6)',
                    border: isSelected ? `1px solid ${color}50` : '1px solid rgba(255,255,255,0.06)',
                  }}>
                  <div className="flex items-center justify-between mb-1">
                    <span className="text-xs font-bold" style={{ fontFamily: 'Rajdhani', color: '#c9a96e' }}>
                      {s.scale_code}
                    </span>
                    <div className="flex items-center gap-1">
                      {s.alert_triggered && <span style={{ color: '#ff3b3b', fontSize: '0.7rem' }}>⚠️</span>}
                      <span style={{ fontSize: '1.25rem', fontFamily: 'Rajdhani', fontWeight: 700, color }}>
                        {s.total_score}
                      </span>
                    </div>
                  </div>
                  <div className="text-xs" style={{ color }}>{s.severity_label}</div>
                  <div className="flex items-center gap-1 mt-1">
                    {s.delta_direction && s.delta_direction !== 'first_admin' && (
                      <span style={{ color: DELTA_COLORS[s.delta_direction], fontSize: '0.65rem', fontWeight: 700 }}>
                        {DELTA_ICONS[s.delta_direction]} {Math.abs(s.score_delta || 0)}
                      </span>
                    )}
                    <span className="text-xs opacity-30">
                      {new Date(s.administered_at).toLocaleDateString()}
                    </span>
                  </div>
                </button>
              )
            })}
          </div>
        )}
      </div>

      {/* Trend chart */}
      {selected && (
        <div className="p-4 rounded mb-4" style={{ background: 'rgba(10,22,40,0.7)', border: '1px solid rgba(0,212,255,0.12)' }}>
          <div className="flex items-center justify-between mb-3">
            <div>
              <span className="font-bold" style={{ fontFamily: 'Rajdhani', color: '#c9a96e' }}>{selected}</span>
              <span className="text-xs opacity-40 ml-2">Score over time</span>
            </div>
            <button
              onClick={() => onAdminister({ scale_code: selected, scale_name: selected } as Parameters<typeof onAdminister>[0])}
              className="text-xs px-3 py-1 rounded"
              style={{ background: 'rgba(201,169,110,0.1)', border: '1px solid rgba(201,169,110,0.3)', color: '#c9a96e', fontFamily: 'Rajdhani', fontWeight: 700 }}>
              + New Administration
            </button>
          </div>

          {loadingTrend ? (
            <div className="h-40 flex items-center justify-center hud-label animate-pulse">Loading trend...</div>
          ) : trend.length === 0 ? (
            <div className="h-40 flex items-center justify-center text-sm opacity-30">No history yet</div>
          ) : trend.length === 1 ? (
            <div className="flex items-center justify-center h-40">
              <div className="text-center">
                <div style={{ fontFamily: 'Rajdhani', fontSize: '3rem', fontWeight: 700, color: SEVERITY_COLORS[trend[0].color_code] || '#00d4ff' }}>
                  {trend[0].total_score}
                </div>
                <div className="text-xs opacity-40">{trend[0].severity_label}</div>
                <div className="text-xs opacity-30 mt-1">First administration — {new Date(trend[0].administered_at).toLocaleDateString()}</div>
              </div>
            </div>
          ) : (
            <canvas ref={canvasRef} style={{ width: '100%', height: 160, display: 'block' }} />
          )}

          {/* Legend */}
          <div className="flex gap-4 mt-2">
            <div className="flex items-center gap-1.5 text-xs opacity-40">
              <div className="w-2 h-2 rounded-full" style={{ background: '#00d4ff' }} />
              Score
            </div>
            <div className="flex items-center gap-1.5 text-xs opacity-40">
              <div className="w-2 h-2 rounded-full" style={{ background: '#ff3b3b', border: '1px solid #ff3b3b' }} />
              Alert triggered
            </div>
          </div>
        </div>
      )}

      {/* History list */}
      {selected && trend.length > 0 && (
        <div>
          <div className="hud-label mb-2">Administration History</div>
          <div className="space-y-2">
            {[...trend].reverse().map(t => (
              <div key={t.admin_id} className="flex items-center justify-between px-3 py-2 rounded"
                style={{ background: 'rgba(10,22,40,0.5)', border: '1px solid rgba(255,255,255,0.05)' }}>
                <div className="text-xs opacity-50">{new Date(t.administered_at).toLocaleDateString()}</div>
                <div className="flex items-center gap-3">
                  {t.alert_triggered && <span style={{ color: '#ff3b3b', fontSize: '0.7rem' }}>⚠️ Alert</span>}
                  <span className="text-xs" style={{ color: SEVERITY_COLORS[t.color_code] || '#00d4ff' }}>{t.severity_label}</span>
                  <span className="font-bold" style={{ fontFamily: 'Rajdhani', color: SEVERITY_COLORS[t.color_code] || '#00d4ff' }}>
                    {t.total_score}
                  </span>
                  {t.delta_direction && t.delta_direction !== 'first_admin' && (
                    <span style={{ color: DELTA_COLORS[t.delta_direction], fontSize: '0.7rem', fontWeight: 700 }}>
                      {DELTA_ICONS[t.delta_direction]} {Math.abs(t.score_delta || 0)}
                    </span>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}
