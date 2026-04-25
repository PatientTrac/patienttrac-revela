import { useState, useEffect } from 'react'
import { callScales } from './RatingScalesModule'

interface Alert {
  admin_id: number
  patient_id: number
  patient_name: string
  scale_code: string
  scale_name: string
  total_score: number
  severity_label: string
  alert_type: string
  item_9_flag: boolean
  administered_at: string
  provider_name: string | null
}

interface Props {
  orgId: string
  providerId: string
  onAck: () => void
}

const ALERT_DESCRIPTIONS: Record<string, { label: string; color: string; icon: string }> = {
  suicide_risk_high:       { label: 'Suicide Risk — High',            color: '#ff3b3b', icon: '🚨' },
  suicide_risk_any:        { label: 'Suicide Risk — Present',         color: '#ff3b3b', icon: '🚨' },
  active_si_with_plan:     { label: 'Active SI with Plan — URGENT',   color: '#ff3b3b', icon: '🚨' },
  suicidal_behavior:       { label: 'Suicidal Behavior Reported',     color: '#ff3b3b', icon: '🚨' },
  suicide_risk:            { label: 'Suicide Risk Detected',          color: '#ff3b3b', icon: '🚨' },
  severe_depression:       { label: 'Severe Depression',              color: '#e74c3c', icon: '⚠️' },
  severe_depression_madrs: { label: 'Severe Depression (MADRS)',      color: '#e74c3c', icon: '⚠️' },
  suicidal_ideation_madrs: { label: 'Suicidal Ideation (MADRS)',      color: '#ff3b3b', icon: '🚨' },
  severe_anxiety:          { label: 'Severe Anxiety',                 color: '#e67e22', icon: '⚠️' },
  severe_ptsd:             { label: 'Severe PTSD',                    color: '#8e44ad', icon: '⚠️' },
  severe_psychosis:        { label: 'Severe Psychopathology (PANSS)', color: '#e74c3c', icon: '⚠️' },
}

export default function ScaleAlerts({ orgId, providerId, onAck }: Props) {
  const [alerts, setAlerts] = useState<Alert[]>([])
  const [loading, setLoading] = useState(true)
  const [ackingId, setAckingId] = useState<number | null>(null)
  const [ackNote, setAckNote] = useState<Record<number, string>>({})
  const [expandedId, setExpandedId] = useState<number | null>(null)
  const [acknowledged, setAcknowledged] = useState<number[]>([])

  useEffect(() => {
    loadAlerts()
  }, [orgId])

  const loadAlerts = () => {
    setLoading(true)
    callScales({ action: 'get_alerts', org_id: orgId })
      .then(d => setAlerts(d.alerts || []))
      .finally(() => setLoading(false))
  }

  const handleAck = async (adminId: number) => {
    setAckingId(adminId)
    try {
      await callScales({
        action: 'ack_alert',
        org_id: orgId,
        admin_id: adminId,
        provider_id: Number(providerId),
        alert_note: ackNote[adminId] || '',
      })
      setAcknowledged(prev => [...prev, adminId])
      setExpandedId(null)
      onAck()
    } catch (e) {
      console.error(e)
    } finally {
      setAckingId(null)
    }
  }

  const visibleAlerts = alerts.filter(a => !acknowledged.includes(a.admin_id))
  const isSuicidalityAlert = (type: string) =>
    ['suicide_risk_high','suicide_risk_any','active_si_with_plan','suicidal_behavior','suicide_risk','suicidal_ideation_madrs'].includes(type)

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-2">
          <span style={{ fontFamily: 'Rajdhani', fontSize: '1rem', fontWeight: 700, color: '#c9a96e' }}>
            Pending Alerts
          </span>
          {visibleAlerts.length > 0 && (
            <span className="px-2 py-0.5 rounded text-xs font-bold"
              style={{ background: 'rgba(255,59,59,0.15)', color: '#ff3b3b', border: '1px solid rgba(255,59,59,0.3)', fontFamily: 'Rajdhani' }}>
              {visibleAlerts.length}
            </span>
          )}
        </div>
        <button onClick={loadAlerts} className="text-xs opacity-40 hover:opacity-70">↻ Refresh</button>
      </div>

      {/* HIPAA note */}
      <div className="p-3 rounded mb-4 text-xs leading-relaxed"
        style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.12)', color: 'rgba(200,215,235,0.5)' }}>
        §170.315(a)(9) Clinical Decision Support — All alerts acknowledged here are logged to the EPCS audit trail with provider ID, timestamp, and note.
      </div>

      {loading ? (
        <div className="hud-label animate-pulse py-8 text-center">Loading alerts...</div>
      ) : visibleAlerts.length === 0 ? (
        <div className="text-center py-12">
          <div className="text-4xl mb-3">✅</div>
          <div className="text-sm opacity-50">No pending alerts. All clear.</div>
        </div>
      ) : (
        <div className="space-y-3">
          {visibleAlerts.map(alert => {
            const meta = ALERT_DESCRIPTIONS[alert.alert_type] || { label: alert.alert_type, color: '#e74c3c', icon: '⚠️' }
            const isSuicidal = isSuicidalityAlert(alert.alert_type) || alert.item_9_flag
            const isExpanded = expandedId === alert.admin_id

            return (
              <div key={alert.admin_id} className="rounded overflow-hidden"
                style={{ border: `1px solid ${isSuicidal ? 'rgba(255,59,59,0.4)' : 'rgba(231,76,60,0.25)'}` }}>

                {/* Alert header */}
                <div className="p-4"
                  style={{ background: isSuicidal ? 'rgba(255,59,59,0.08)' : 'rgba(10,22,40,0.7)' }}>
                  <div className="flex items-start justify-between gap-3">
                    <div className="flex items-start gap-3">
                      <span style={{ fontSize: '1.4rem', lineHeight: 1 }}>{meta.icon}</span>
                      <div>
                        <div className="font-bold text-sm mb-0.5"
                          style={{ fontFamily: 'Rajdhani', color: meta.color, letterSpacing: '0.04em' }}>
                          {meta.label}
                        </div>
                        <div className="text-sm font-semibold" style={{ color: 'rgba(200,215,235,0.8)' }}>
                          {alert.patient_name}
                        </div>
                        <div className="text-xs mt-0.5" style={{ color: 'rgba(200,215,235,0.4)' }}>
                          {alert.scale_code} · Score {alert.total_score} ({alert.severity_label}) · {new Date(alert.administered_at).toLocaleDateString()} {new Date(alert.administered_at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                        </div>
                        {alert.provider_name && (
                          <div className="text-xs mt-0.5 opacity-40">Administered by {alert.provider_name}</div>
                        )}
                      </div>
                    </div>

                    <button onClick={() => setExpandedId(isExpanded ? null : alert.admin_id)}
                      className="text-xs px-3 py-1.5 rounded flex-shrink-0"
                      style={{ background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)', color: 'rgba(200,215,235,0.6)', fontFamily: 'Rajdhani' }}>
                      {isExpanded ? 'Collapse' : 'Acknowledge'}
                    </button>
                  </div>

                  {/* Suicide guidance */}
                  {isSuicidal && (
                    <div className="mt-3 p-2 rounded text-xs leading-relaxed"
                      style={{ background: 'rgba(255,59,59,0.08)', color: 'rgba(255,150,150,0.8)' }}>
                      <strong>Required actions:</strong> Complete C-SSRS structured interview · Document safety plan in care plan · Notify supervising physician · Consider involuntary hold if imminent risk
                    </div>
                  )}
                </div>

                {/* Acknowledge panel */}
                {isExpanded && (
                  <div className="p-4 border-t" style={{ background: 'rgba(6,14,28,0.8)', borderColor: 'rgba(255,255,255,0.06)' }}>
                    <div className="hud-label text-xs mb-2">Acknowledgement Note (required for suicidality alerts)</div>
                    <textarea
                      placeholder={isSuicidal
                        ? 'Document actions taken: C-SSRS completed, safety plan created, supervisor notified...'
                        : 'Optional clinical note...'}
                      value={ackNote[alert.admin_id] || ''}
                      onChange={e => setAckNote(n => ({ ...n, [alert.admin_id]: e.target.value }))}
                      className="w-full px-3 py-2 rounded text-sm mb-3 resize-none"
                      rows={3}
                      style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.15)', color: 'rgba(200,215,235,0.8)', outline: 'none', fontFamily: 'DM Sans' }}
                    />
                    <div className="flex gap-3">
                      <button onClick={() => setExpandedId(null)}
                        className="px-4 py-2 rounded text-xs"
                        style={{ background: 'rgba(255,255,255,0.03)', border: '1px solid rgba(255,255,255,0.1)', color: 'rgba(200,215,235,0.4)', fontFamily: 'Rajdhani' }}>
                        Cancel
                      </button>
                      <button
                        onClick={() => handleAck(alert.admin_id)}
                        disabled={ackingId === alert.admin_id || (isSuicidal && !(ackNote[alert.admin_id]?.trim()))}
                        className="flex-1 py-2 rounded text-xs font-bold transition-all disabled:opacity-30"
                        style={{ background: 'rgba(201,169,110,0.12)', border: '1px solid rgba(201,169,110,0.4)', color: '#c9a96e', fontFamily: 'Rajdhani', letterSpacing: '0.06em' }}>
                        {ackingId === alert.admin_id ? 'Saving...' : '✓ Acknowledge & Log'}
                      </button>
                    </div>
                    {isSuicidal && !(ackNote[alert.admin_id]?.trim()) && (
                      <p className="text-xs mt-2" style={{ color: 'rgba(255,59,59,0.6)' }}>
                        A note is required before acknowledging suicidality alerts.
                      </p>
                    )}
                  </div>
                )}
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}
