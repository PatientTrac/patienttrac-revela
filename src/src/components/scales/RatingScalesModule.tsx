// Rating Scales Module — PatientTrac Mind
// Plug into TabNav as the 5th tab: { id: 'scales', label: 'Scales', icon: '📊' }
// Covers: §170.315(a)(9) Clinical Decision Support alerts

import { useState, useEffect } from 'react'
import ScaleCatalog from './ScaleCatalog'
import ScaleAdministration from './ScaleAdministration'
import ScaleTrend from './ScaleTrend'
import ScaleAlerts from './ScaleAlerts'
import { useAppStore } from '../../lib/store'

const SCALES_URL = `${import.meta.env.VITE_SUPABASE_URL}/functions/v1/rating-scales`

export interface ScaleMeta {
  scale_id: number
  scale_code: string
  scale_name: string
  scale_full_name: string
  category: string
  total_items: number
  max_score: number
  min_score: number
  higher_is_worse: boolean
  estimated_minutes: number
  administration: string
  cpt_code: string
  license_required: boolean
  license_note: string | null
  description: string
}

export interface ScaleAdminResult {
  admin_id: number
  scale_code: string
  scale_name: string
  total_score: number
  max_score: number
  subscale_scores: Record<string, number>
  severity: SeverityRange | null
  item_9_flag: boolean
  alert_triggered: boolean
  alert_type: string | null
  score_delta: number | null
  delta_direction: string
  clinical_guidance: string | null
  recommended_action: string | null
}

export interface SeverityRange {
  severity_label: string
  severity_code: string
  color_code: string
  clinical_note: string
  recommended_action: string
}

export async function callScales(body: Record<string, unknown>) {
  const res = await fetch(SCALES_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  return res.json()
}

type SubTab = 'catalog' | 'administer' | 'trend' | 'alerts'

export default function RatingScalesModule() {
  const { context, session } = useAppStore()
  const [subTab, setSubTab] = useState<SubTab>('catalog')
  const [selectedScale, setSelectedScale] = useState<ScaleMeta | null>(null)
  const [lastResult, setLastResult] = useState<ScaleAdminResult | null>(null)
  const [alertCount, setAlertCount] = useState(0)

  const patient = context?.patient as Record<string, unknown> | undefined
  const patientId = patient?.patient_id ? String(patient.patient_id) : ''
  const orgId = session?.org_id || '00000000-0000-0000-0000-000000000001'
  const providerId = session?.provider_id || '1'
  const encounterId = session?.encounter_id || ''

  // Load alert count on mount
  useEffect(() => {
    callScales({ action: 'get_alerts', org_id: orgId })
      .then(d => setAlertCount(d.count || 0))
      .catch(console.error)
  }, [orgId])

  const handleSelectScale = (scale: ScaleMeta) => {
    setSelectedScale(scale)
    setSubTab('administer')
  }

  const handleScoreComplete = (result: ScaleAdminResult) => {
    setLastResult(result)
    if (result.alert_triggered) setAlertCount(c => c + 1)
  }

  const SUBTABS: { id: SubTab; label: string; icon: string; badge?: number }[] = [
    { id: 'catalog',    label: 'Scale Catalog',   icon: '📚' },
    { id: 'administer', label: 'Administer',       icon: '✏️' },
    { id: 'trend',      label: 'Trend',            icon: '📈' },
    { id: 'alerts',     label: 'Alerts',           icon: '⚠️', badge: alertCount },
  ]

  return (
    <div className="animate-fadeIn">
      {/* Sub-tab bar */}
      <div className="flex border-b px-4"
        style={{ background: 'rgba(6,14,28,0.5)', borderColor: 'rgba(0,212,255,0.1)' }}>
        {SUBTABS.map(t => (
          <button key={t.id} onClick={() => setSubTab(t.id)}
            className={`flex items-center gap-1.5 px-4 py-2.5 text-xs transition-all relative ${subTab === t.id ? 'tab-active' : 'tab-inactive'}`}
            style={{ fontFamily: 'Rajdhani', fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' }}>
            <span style={{ fontSize: 14 }}>{t.icon}</span>
            {t.label}
            {t.badge ? (
              <span className="ml-1 px-1.5 py-0.5 rounded text-xs font-bold"
                style={{ background: '#ff3b3b', color: 'white', fontSize: '0.6rem' }}>
                {t.badge}
              </span>
            ) : null}
          </button>
        ))}
      </div>

      {/* Alert banner if last result triggered */}
      {lastResult?.alert_triggered && subTab !== 'alerts' && (
        <div className="mx-6 mt-4 p-3 rounded flex items-center justify-between"
          style={{ background: 'rgba(255,59,59,0.1)', border: '1px solid rgba(255,59,59,0.4)' }}>
          <div className="flex items-center gap-2">
            <span>⚠️</span>
            <span className="text-sm font-semibold" style={{ color: '#ff3b3b', fontFamily: 'Rajdhani' }}>
              {lastResult.item_9_flag
                ? 'SAFETY ALERT — Suicidality indicated. Safety assessment required.'
                : `Alert: ${lastResult.severity?.severity_label} — ${lastResult.clinical_guidance}`}
            </span>
          </div>
          <button onClick={() => setSubTab('alerts')}
            className="text-xs px-3 py-1 rounded"
            style={{ background: '#ff3b3b', color: 'white', fontFamily: 'Rajdhani', fontWeight: 700 }}>
            View Alerts
          </button>
        </div>
      )}

      {/* Content */}
      <div className="p-5">
        {subTab === 'catalog' && (
          <ScaleCatalog orgId={orgId} onSelect={handleSelectScale} />
        )}
        {subTab === 'administer' && (
          <ScaleAdministration
            orgId={orgId}
            patientId={patientId}
            encounterId={encounterId}
            providerId={providerId}
            initialScale={selectedScale}
            onComplete={result => { handleScoreComplete(result); setSubTab('trend') }}
          />
        )}
        {subTab === 'trend' && (
          <ScaleTrend
            orgId={orgId}
            patientId={patientId}
            lastResult={lastResult}
            onAdminister={scale => { setSelectedScale(scale); setSubTab('administer') }}
          />
        )}
        {subTab === 'alerts' && (
          <ScaleAlerts
            orgId={orgId}
            providerId={providerId}
            onAck={() => setAlertCount(c => Math.max(0, c - 1))}
          />
        )}
      </div>
    </div>
  )
}
