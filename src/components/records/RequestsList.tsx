import { useEffect, useState } from 'react'
import { supabase } from '../../lib/supabase'
import { REPORT_TYPE_LABELS, REPORT_TYPE_ICONS } from './types'
import type { RecordsRequest, RequestStatus } from './types'

interface Props {
  patientId: string
  orgId: string
  onSelectRequest: (r: RecordsRequest) => void
}

const STATUS_COLORS: Record<RequestStatus, string> = {
  pending: '#ffa500',
  in_review: '#00d4ff',
  approved: '#00ff88',
  denied: '#ff3b3b',
  fulfilled: '#c9a96e',
  expired: '#666',
}

export default function RequestsList({ patientId, orgId, onSelectRequest }: Props) {
  const [requests, setRequests] = useState<RecordsRequest[]>([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    const load = async () => {
      const { data } = await supabase
        .from('cr.records_requests')
        .select('*')
        .eq('patient_id', patientId)
        .order('created_at', { ascending: false })
      setRequests((data || []) as RecordsRequest[])
      setLoading(false)
    }
    load()
  }, [patientId])

  if (loading) return <div className="hud-label animate-pulse p-4">Loading requests...</div>

  if (requests.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center py-16 opacity-40">
        <div className="text-4xl mb-3">📭</div>
        <div className="hud-label">No records requests yet for this patient</div>
      </div>
    )
  }

  return (
    <div className="space-y-3">
      <div className="hud-label mb-4">
        {requests.length} request{requests.length !== 1 ? 's' : ''} on file
      </div>
      {requests.map((r) => (
        <div key={r.request_id}
          className="hud-card p-4 flex items-center justify-between gap-4 cursor-pointer transition-all hover:border-cyan-mind/40"
          onClick={() => onSelectRequest(r)}>
          <div className="flex items-center gap-4">
            <span className="text-2xl">{REPORT_TYPE_ICONS[r.request_type]}</span>
            <div>
              <div className="font-semibold text-sm" style={{ fontFamily: 'Rajdhani', color: '#c9a96e' }}>
                {REPORT_TYPE_LABELS[r.request_type]}
              </div>
              <div className="text-xs opacity-60 mt-0.5">
                {r.requestor_type === 'provider_internal' ? 'Internal' :
                 r.requestor_type === 'patient' ? 'Patient Request' :
                 `Third Party: ${r.requestor_name || r.requestor_org || 'Unknown'}`}
                {' · '}
                {new Date(r.created_at).toLocaleDateString()}
              </div>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-xs px-2 py-0.5 rounded font-mono uppercase"
              style={{
                background: `${STATUS_COLORS[r.status]}18`,
                border: `1px solid ${STATUS_COLORS[r.status]}40`,
                color: STATUS_COLORS[r.status],
              }}>
              {r.status}
            </span>
            {r.urgency !== 'routine' && (
              <span className="text-xs px-2 py-0.5 rounded font-mono uppercase"
                style={{ background: r.urgency === 'stat' ? 'rgba(255,59,59,0.15)' : 'rgba(255,165,0,0.15)', color: r.urgency === 'stat' ? '#ff3b3b' : '#ffa500' }}>
                {r.urgency}
              </span>
            )}
            <span className="text-xs opacity-40">→ Build Report</span>
          </div>
        </div>
      ))}
    </div>
  )
}
