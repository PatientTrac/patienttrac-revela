import { useState } from 'react'
import type { GeneratedReport } from './types'

const BRIDGE_URL = 'https://mskormozwekezjmtcylv.supabase.co/functions/v1/generate-report'

interface Props {
  orgId: string
  patientName: string
  selectedReport?: GeneratedReport | null
}

export default function ThirdPartyPortal({ orgId, patientName, selectedReport }: Props) {
  const [form, setForm] = useState({
    party_name: '',
    party_org: '',
    party_email: '',
    party_type: 'attorney',
    max_views: '5',
  })
  const [accessUrl, setAccessUrl] = useState('')
  const [token, setToken] = useState('')
  const [expires, setExpires] = useState('')
  const [creating, setCreating] = useState(false)
  const [copied, setCopied] = useState(false)

  const up = (k: string, v: string) => setForm(p => ({ ...p, [k]: v }))

  const handleCreate = async () => {
    if (!selectedReport) return
    setCreating(true)
    try {
      const res = await fetch(BRIDGE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'third_party_access',
          org_id: orgId,
          request_id: selectedReport.request_id,
          party_name: form.party_name,
          party_org: form.party_org,
          party_email: form.party_email,
          party_type: form.party_type,
          report_ids: [selectedReport.report_id],
          max_views: parseInt(form.max_views),
        }),
      })
      const data = await res.json()
      if (data.access_url) {
        setAccessUrl(data.access_url)
        setToken(data.access_token)
        setExpires(new Date(data.expires_at).toLocaleDateString())
      }
    } catch (e) { console.error(e) }
    finally { setCreating(false) }
  }

  const handleCopy = () => {
    navigator.clipboard.writeText(accessUrl)
    setCopied(true)
    setTimeout(() => setCopied(false), 2000)
  }

  return (
    <div className="space-y-5 max-w-2xl">

      {/* Info */}
      <div className="p-4 rounded space-y-1"
        style={{ background: 'rgba(0,212,255,0.05)', border: '1px solid rgba(0,212,255,0.2)' }}>
        <div className="hud-label">Secure Third-Party Access Protocol</div>
        <p className="text-sm opacity-70 leading-relaxed mt-2">
          Generate a time-limited, view-limited secure access token for attorneys, insurance companies, or other physicians. The recipient must accept a HIPAA compliance agreement before viewing any records. All access is logged with IP and timestamp.
        </p>
      </div>

      {/* Selected report */}
      {selectedReport ? (
        <div className="p-3 rounded flex items-center gap-3"
          style={{ background: 'rgba(201,169,110,0.08)', border: '1px solid rgba(201,169,110,0.25)' }}>
          <span className="text-xl">📄</span>
          <div>
            <div className="text-sm font-semibold" style={{ color: '#c9a96e', fontFamily: 'Rajdhani' }}>
              {selectedReport.report_title}
            </div>
            <div className="text-xs opacity-60">
              Patient: {patientName} ·
              {selectedReport.signed ? ' ✓ Signed' : ' ⚠️ Not yet signed — sign before sharing'}
            </div>
          </div>
        </div>
      ) : (
        <div className="p-4 rounded text-center text-sm opacity-50"
          style={{ border: '1px dashed rgba(0,212,255,0.2)' }}>
          Generate and sign a report first, then come back here to create third-party access
        </div>
      )}

      {/* Third-party form */}
      <div className="hud-card p-5 space-y-4">
        <div className="hud-label">Third-Party Details</div>

        <div className="grid grid-cols-2 gap-3">
          <div>
            <div className="hud-label mb-1">Party Type</div>
            <select className="select-hud w-full rounded px-3 py-2 text-sm" value={form.party_type} onChange={e => up('party_type', e.target.value)}>
              <option value="attorney">Attorney / Law Firm</option>
              <option value="insurance">Insurance Company</option>
              <option value="physician">Treating Physician</option>
              <option value="employer">Employer (FMLA/Disability)</option>
              <option value="government">Government Agency</option>
              <option value="other">Other</option>
            </select>
          </div>
          <div>
            <div className="hud-label mb-1">Max Views Allowed</div>
            <select className="select-hud w-full rounded px-3 py-2 text-sm" value={form.max_views} onChange={e => up('max_views', e.target.value)}>
              <option value="1">1 view</option>
              <option value="3">3 views</option>
              <option value="5">5 views</option>
              <option value="10">10 views</option>
            </select>
          </div>
          <div>
            <div className="hud-label mb-1">Contact Name</div>
            <input type="text" value={form.party_name} onChange={e => up('party_name', e.target.value)}
              className="input-hud w-full rounded px-3 py-2 text-sm" />
          </div>
          <div>
            <div className="hud-label mb-1">Organization</div>
            <input type="text" value={form.party_org} onChange={e => up('party_org', e.target.value)}
              className="input-hud w-full rounded px-3 py-2 text-sm" />
          </div>
          <div className="col-span-2">
            <div className="hud-label mb-1">Email (receives secure link)</div>
            <input type="email" value={form.party_email} onChange={e => up('party_email', e.target.value)}
              className="input-hud w-full rounded px-3 py-2 text-sm" />
          </div>
        </div>

        {/* Access rules */}
        <div className="text-xs p-3 rounded space-y-1"
          style={{ background: 'rgba(6,14,28,0.6)', border: '1px solid rgba(0,212,255,0.1)' }}>
          <div className="hud-label mb-2">Access Rules Applied Automatically</div>
          <div className="opacity-70 space-y-1">
            <div>✓ Expires in 30 days</div>
            <div>✓ Max {form.max_views} view{form.max_views !== '1' ? 's' : ''}</div>
            <div>✓ HIPAA agreement required before access</div>
            <div>✓ IP address + timestamp logged on every view</div>
            <div>✓ Token revocable at any time</div>
          </div>
        </div>

        <button onClick={handleCreate}
          disabled={!selectedReport || !form.party_email || creating || !selectedReport?.signed}
          className="btn-primary w-full py-3 rounded text-sm disabled:opacity-40">
          {creating ? 'Creating Secure Link...' : '🔗 Generate Secure Access Link'}
        </button>
      </div>

      {/* Access link result */}
      {accessUrl && (
        <div className="hud-card p-5 space-y-3"
          style={{ border: '1px solid rgba(0,255,136,0.3)', background: 'rgba(0,255,136,0.04)' }}>
          <div className="flex items-center gap-2">
            <div className="pulse-dot" />
            <div className="hud-label" style={{ color: '#00ff88' }}>Secure Access Link Created</div>
          </div>
          <div className="space-y-1 text-xs opacity-60">
            <div>Expires: {expires} · Max views: {form.max_views} · Sent to: {form.party_email}</div>
          </div>
          <div className="flex gap-2">
            <input readOnly value={accessUrl}
              className="input-hud flex-1 rounded px-3 py-2 text-xs font-mono" />
            <button onClick={handleCopy}
              className="btn-cyan px-4 py-2 rounded text-xs flex-shrink-0">
              {copied ? '✓ Copied' : 'Copy Link'}
            </button>
          </div>
          <div className="text-xs font-mono opacity-40">Token: {token.slice(0, 16)}...</div>
        </div>
      )}
    </div>
  )
}
