import { useState } from 'react'
import { supabase } from '../../lib/supabase'
import { REPORT_TYPE_LABELS, REPORT_TYPE_ICONS, REPORT_APP_MAP } from './types'
import type { ReportType, RequestorType } from './types'

interface Props {
  patientId: string
  encounterId?: string
  providerId: string
  orgId: string
  appSource: 'mind' | 'revela'
  patientName: string
  onCreated: () => void
}

export default function NewRequestForm({ patientId, encounterId, providerId, orgId, appSource, patientName, onCreated }: Props) {
  const [requestorType, setRequestorType] = useState<RequestorType>('provider_internal')
  const [reportType, setReportType] = useState<ReportType | ''>('')
  const [urgency, setUrgency] = useState<'routine' | 'urgent' | 'stat'>('routine')
  const [roiSigned, setRoiSigned] = useState(false)
  const [agreementType, setAgreementType] = useState('hipaa_roi')
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)

  const [form, setForm] = useState({
    requestor_name: '',
    requestor_org: '',
    requestor_email: '',
    requestor_phone: '',
    requestor_fax: '',
    date_range_start: '',
    date_range_end: '',
    specific_notes: '',
  })

  // Filter report types by app
  const availableTypes = (Object.entries(REPORT_TYPE_LABELS) as [ReportType, string][])
    .filter(([type]) => {
      const app = REPORT_APP_MAP[type]
      return app === 'both' || app === appSource
    })

  const handleSubmit = async () => {
    if (!reportType) return
    setSaving(true)
    try {
      await supabase.schema('cr').from('records_requests').insert({
        org_id: orgId,
        patient_id: patientId,
        encounter_id: encounterId || null,
        provider_id: providerId,
        requestor_type: requestorType,
        requestor_name: form.requestor_name || null,
        requestor_org: form.requestor_org || null,
        requestor_email: form.requestor_email || null,
        requestor_phone: form.requestor_phone || null,
        requestor_fax: form.requestor_fax || null,
        request_type: reportType,
        date_range_start: form.date_range_start || null,
        date_range_end: form.date_range_end || null,
        specific_notes: form.specific_notes || null,
        urgency,
        has_signed_roi: roiSigned,
        agreement_type: requestorType === 'third_party' ? agreementType : null,
        status: 'pending',
      })
      setSaved(true)
      setTimeout(onCreated, 1000)
    } catch (e) { console.error(e) }
    finally { setSaving(false) }
  }

  const up = (k: string, v: string) => setForm(p => ({ ...p, [k]: v }))

  const urgencyColor = urgency === 'stat' ? '#ff3b3b' : urgency === 'urgent' ? '#ffa500' : '#00ff88'

  return (
    <div className="space-y-5 max-w-3xl">

      {/* Requestor Type */}
      <div className="hud-card p-5">
        <div className="hud-label mb-3">Who Is Requesting?</div>
        <div className="grid grid-cols-3 gap-3">
          {([
            { value: 'provider_internal', label: 'Provider Internal', icon: '👨‍⚕️', desc: 'Provider generates report' },
            { value: 'patient', label: 'Patient Request', icon: '🧑', desc: 'Patient requests own records' },
            { value: 'third_party', label: 'Third Party', icon: '🏛️', desc: 'Attorney / Insurance / Physician' },
          ] as const).map((opt) => (
            <button
              key={opt.value}
              onClick={() => setRequestorType(opt.value)}
              className="p-4 rounded text-left transition-all"
              style={{
                background: requestorType === opt.value ? 'rgba(0,212,255,0.1)' : 'rgba(6,14,28,0.6)',
                border: `1px solid ${requestorType === opt.value ? '#00d4ff' : 'rgba(0,212,255,0.15)'}`,
                boxShadow: requestorType === opt.value ? '0 0 16px rgba(0,212,255,0.15)' : 'none',
              }}
            >
              <div className="text-xl mb-1">{opt.icon}</div>
              <div className="font-semibold text-sm" style={{ fontFamily: 'Rajdhani', color: requestorType === opt.value ? '#00d4ff' : '#f0f4ff' }}>{opt.label}</div>
              <div className="text-xs mt-1 opacity-60">{opt.desc}</div>
            </button>
          ))}
        </div>
      </div>

      {/* Third-party fields */}
      {requestorType === 'third_party' && (
        <div className="hud-card p-5 space-y-4"
          style={{ border: '1px solid rgba(201,169,110,0.25)' }}>
          <div className="hud-label">Third-Party Details & Authorization</div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Requestor Name" value={form.requestor_name} onChange={v => up('requestor_name', v)} />
            <Field label="Organization" value={form.requestor_org} onChange={v => up('requestor_org', v)} />
            <Field label="Email" value={form.requestor_email} onChange={v => up('requestor_email', v)} type="email" />
            <Field label="Phone" value={form.requestor_phone} onChange={v => up('requestor_phone', v)} />
            <Field label="Fax" value={form.requestor_fax} onChange={v => up('requestor_fax', v)} />
            <div>
              <div className="hud-label mb-1">Agreement Type</div>
              <select className="select-hud w-full rounded px-3 py-2 text-sm" value={agreementType} onChange={e => setAgreementType(e.target.value)}>
                <option value="hipaa_roi">HIPAA Release of Information</option>
                <option value="baa">Business Associate Agreement (BAA)</option>
                <option value="court_order">Court Order</option>
                <option value="patient_consent">Patient Consent</option>
                <option value="emergency">Emergency Release</option>
              </select>
            </div>
          </div>
          <label className="flex items-center gap-3 cursor-pointer">
            <div onClick={() => setRoiSigned(!roiSigned)}
              className="w-5 h-5 rounded flex items-center justify-center flex-shrink-0"
              style={{ background: roiSigned ? '#00d4ff' : 'rgba(6,14,28,0.8)', border: `1px solid ${roiSigned ? '#00d4ff' : 'rgba(0,212,255,0.3)'}`, cursor: 'pointer' }}>
              {roiSigned && <span className="text-xs font-bold" style={{ color: '#060e1c' }}>✓</span>}
            </div>
            <span className="text-sm">Signed authorization / agreement on file</span>
          </label>
          {!roiSigned && (
            <div className="text-xs px-3 py-2 rounded" style={{ background: 'rgba(255,165,0,0.08)', border: '1px solid rgba(255,165,0,0.3)', color: '#ffa500' }}>
              ⚠️ Records cannot be released to third parties without signed authorization
            </div>
          )}
        </div>
      )}

      {/* Patient requestor name field */}
      {requestorType === 'patient' && (
        <div className="hud-card p-5">
          <div className="hud-label mb-3">Patient Contact for Delivery</div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Email" value={form.requestor_email} onChange={v => up('requestor_email', v)} type="email" />
            <Field label="Phone" value={form.requestor_phone} onChange={v => up('requestor_phone', v)} />
          </div>
        </div>
      )}

      {/* Report Type */}
      <div className="hud-card p-5">
        <div className="hud-label mb-3">Report / Record Type</div>
        <div className="grid grid-cols-2 gap-2">
          {availableTypes.map(([type, label]) => (
            <button
              key={type}
              onClick={() => setReportType(type)}
              className="flex items-center gap-3 p-3 rounded text-left transition-all"
              style={{
                background: reportType === type ? 'rgba(201,169,110,0.12)' : 'rgba(6,14,28,0.5)',
                border: `1px solid ${reportType === type ? '#c9a96e' : 'rgba(0,212,255,0.12)'}`,
              }}
            >
              <span className="text-xl flex-shrink-0">{REPORT_TYPE_ICONS[type]}</span>
              <span className="text-sm" style={{ color: reportType === type ? '#c9a96e' : '#e0e8f4', fontFamily: 'Rajdhani', fontWeight: 600 }}>{label}</span>
            </button>
          ))}
        </div>
      </div>

      {/* Date range + urgency + notes */}
      <div className="hud-card p-5 space-y-4">
        <div className="hud-label">Request Details</div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Date Range Start" value={form.date_range_start} onChange={v => up('date_range_start', v)} type="date" />
          <Field label="Date Range End" value={form.date_range_end} onChange={v => up('date_range_end', v)} type="date" />
        </div>
        <div>
          <div className="hud-label mb-2">Urgency</div>
          <div className="flex gap-2">
            {(['routine', 'urgent', 'stat'] as const).map(u => (
              <button key={u} onClick={() => setUrgency(u)}
                className="px-4 py-1.5 rounded text-xs font-mono uppercase transition-all"
                style={{
                  background: urgency === u ? urgencyColor : 'rgba(6,14,28,0.6)',
                  border: `1px solid ${urgencyColor}`,
                  color: urgency === u ? '#060e1c' : urgencyColor,
                  fontWeight: urgency === u ? 700 : 400,
                }}>
                {u}
              </button>
            ))}
          </div>
        </div>
        <div>
          <div className="hud-label mb-1">Specific Notes / Instructions</div>
          <textarea rows={3} value={form.specific_notes} onChange={e => up('specific_notes', e.target.value)}
            placeholder="Any specific records needed, time period, redaction requests..."
            className="input-hud w-full rounded px-3 py-2 text-sm resize-none" />
        </div>
      </div>

      {/* Submit */}
      <div className="flex items-center justify-end gap-4">
        {saved && <span className="hud-label" style={{ color: '#00ff88' }}>✓ Request Created</span>}
        <button onClick={handleSubmit} disabled={!reportType || saving}
          className="btn-primary px-6 py-2.5 rounded text-sm disabled:opacity-40">
          {saving ? 'Creating...' : 'Create Request'}
        </button>
      </div>
    </div>
  )
}

function Field({ label, value, onChange, type = 'text' }: { label: string; value: string; onChange: (v: string) => void; type?: string }) {
  return (
    <div>
      <div className="hud-label mb-1">{label}</div>
      <input type={type} value={value} onChange={e => onChange(e.target.value)}
        className="input-hud w-full rounded px-3 py-2 text-sm" />
    </div>
  )
}
