import { useState, useRef } from 'react'
import type { ReportType, RecordsRequest, GeneratedReport } from './types'
import { REPORT_TYPE_LABELS, REPORT_TYPE_ICONS } from './types'

const BRIDGE_URL = 'https://mskormozwekezjmtcylv.supabase.co/functions/v1/generate-report'

interface Props {
  patientId: string
  encounterId?: string
  providerId: string
  orgId: string
  appSource: 'mind' | 'revela'
  patientName: string
  providerName: string
  providerCreds: string
  providerNpi?: string
  initialRequest?: RecordsRequest | null
  onReportGenerated: (r: GeneratedReport) => void
}

// Dynamic form fields per report type
const REPORT_FIELDS: Record<ReportType, { key: string; label: string; type: 'text' | 'textarea' | 'date' | 'checkbox' }[]> = {
  medical_records_roi: [
    { key: 'requestor_name', label: 'Requestor Name', type: 'text' },
    { key: 'requestor_org', label: 'Requestor Organization', type: 'text' },
    { key: 'purpose', label: 'Purpose of Release', type: 'text' },
    { key: 'date_range_start', label: 'Date Range Start', type: 'date' },
    { key: 'date_range_end', label: 'Date Range End', type: 'date' },
    { key: 'records_description', label: 'Records to Release', type: 'textarea' },
  ],
  physician_letter: [
    { key: 'care_since', label: 'Patient Under Care Since', type: 'date' },
    { key: 'clinical_summary', label: 'Clinical Summary', type: 'textarea' },
    { key: 'diagnoses', label: 'Current Diagnoses', type: 'textarea' },
    { key: 'treatment', label: 'Current Treatment', type: 'textarea' },
    { key: 'recommendations', label: 'Recommendations', type: 'textarea' },
  ],
  psychiatric_evaluation: [
    { key: 'reason_for_eval', label: 'Reason for Evaluation', type: 'textarea' },
    { key: 'hpi', label: 'History of Present Illness', type: 'textarea' },
    { key: 'psych_history', label: 'Psychiatric History', type: 'textarea' },
    { key: 'medical_history', label: 'Medical History', type: 'textarea' },
    { key: 'medications', label: 'Current Medications', type: 'textarea' },
    { key: 'mse_appearance', label: 'MSE: Appearance', type: 'text' },
    { key: 'mse_behavior', label: 'MSE: Behavior', type: 'text' },
    { key: 'mse_speech', label: 'MSE: Speech', type: 'text' },
    { key: 'mse_mood', label: 'MSE: Mood', type: 'text' },
    { key: 'mse_affect', label: 'MSE: Affect', type: 'text' },
    { key: 'mse_thought_process', label: 'MSE: Thought Process', type: 'text' },
    { key: 'mse_thought_content', label: 'MSE: Thought Content', type: 'text' },
    { key: 'mse_perception', label: 'MSE: Perception', type: 'text' },
    { key: 'mse_cognition', label: 'MSE: Cognition', type: 'text' },
    { key: 'mse_insight', label: 'MSE: Insight', type: 'text' },
    { key: 'mse_judgment', label: 'MSE: Judgment', type: 'text' },
    { key: 'dsm5_diagnosis', label: 'DSM-5 Diagnosis', type: 'text' },
    { key: 'icd10_diagnosis', label: 'ICD-10 Diagnosis', type: 'text' },
    { key: 'assessment', label: 'Assessment', type: 'textarea' },
    { key: 'plan', label: 'Plan', type: 'textarea' },
  ],
  surgical_operative_report: [
    { key: 'procedure_date', label: 'Date of Procedure', type: 'date' },
    { key: 'facility', label: 'Facility Name', type: 'text' },
    { key: 'assistant', label: 'Surgical Assistant', type: 'text' },
    { key: 'anesthesiologist', label: 'Anesthesiologist', type: 'text' },
    { key: 'anesthesia_type', label: 'Type of Anesthesia', type: 'text' },
    { key: 'preop_diagnosis', label: 'Preoperative Diagnosis', type: 'textarea' },
    { key: 'postop_diagnosis', label: 'Postoperative Diagnosis', type: 'textarea' },
    { key: 'procedures_performed', label: 'Procedures Performed', type: 'textarea' },
    { key: 'operative_findings', label: 'Operative Findings', type: 'textarea' },
    { key: 'procedure_description', label: 'Description of Procedure', type: 'textarea' },
    { key: 'ebl', label: 'Estimated Blood Loss', type: 'text' },
    { key: 'specimens', label: 'Specimens', type: 'text' },
    { key: 'complications', label: 'Complications', type: 'text' },
    { key: 'disposition', label: 'Disposition', type: 'textarea' },
  ],
  progress_note: [
    { key: 'visit_date', label: 'Visit Date', type: 'date' },
    { key: 'visit_type', label: 'Visit Type', type: 'text' },
    { key: 'cpt_code', label: 'CPT Code', type: 'text' },
    { key: 'chief_complaint', label: 'Chief Complaint', type: 'textarea' },
    { key: 'subjective', label: 'Subjective', type: 'textarea' },
    { key: 'objective', label: 'Objective / MSE', type: 'textarea' },
    { key: 'assessment', label: 'Assessment', type: 'textarea' },
    { key: 'plan', label: 'Plan', type: 'textarea' },
    { key: 'follow_up', label: 'Follow-Up', type: 'text' },
  ],
  prior_auth_letter: [
    { key: 'insurance_name', label: 'Insurance Company Name', type: 'text' },
    { key: 'member_id', label: 'Member / Policy ID', type: 'text' },
    { key: 'requested_service', label: 'Requested Service or Medication', type: 'textarea' },
    { key: 'cpt_code', label: 'CPT/HCPCS Code', type: 'text' },
    { key: 'icd10_code', label: 'ICD-10 Code', type: 'text' },
    { key: 'medical_necessity', label: 'Medical Necessity Statement', type: 'textarea' },
    { key: 'clinical_justification', label: 'Clinical Justification', type: 'textarea' },
    { key: 'treatment_history', label: 'Treatment History / Failed Alternatives', type: 'textarea' },
  ],
  disability_fmla: [
    { key: 'diagnosis', label: 'Diagnosis / Condition', type: 'textarea' },
    { key: 'functional_limitations', label: 'Functional Limitations', type: 'textarea' },
    { key: 'leave_start', label: 'Estimated Leave Start', type: 'date' },
    { key: 'leave_end', label: 'Estimated Return Date', type: 'date' },
    { key: 'work_restrictions', label: 'Work Restrictions', type: 'textarea' },
    { key: 'intermittent', label: 'Intermittent Leave Required', type: 'checkbox' },
    { key: 'full_disability', label: 'Full Disability (Unable to Work)', type: 'checkbox' },
  ],
  court_legal_report: [
    { key: 'provider_qualifications', label: 'Provider Qualifications Statement', type: 'textarea' },
    { key: 'treatment_summary', label: 'Summary of Treatment', type: 'textarea' },
    { key: 'clinical_findings', label: 'Clinical Findings', type: 'textarea' },
    { key: 'diagnosis', label: 'Diagnosis', type: 'textarea' },
    { key: 'causation_opinion', label: 'Causation Opinion', type: 'textarea' },
    { key: 'prognosis', label: 'Prognosis', type: 'textarea' },
    { key: 'expert_opinion', label: 'Expert Opinion', type: 'textarea' },
  ],
}

export default function ReportBuilder({
  patientId, encounterId, providerId, orgId, appSource,
  patientName, providerName, providerCreds, providerNpi,
  initialRequest, onReportGenerated,
}: Props) {
  const availableTypes = Object.entries(REPORT_TYPE_LABELS) as [ReportType, string][]
  const [reportType, setReportType] = useState<ReportType | ''>(initialRequest?.request_type || '')
  const [formData, setFormData] = useState<Record<string, string | boolean>>({})
  const [html, setHtml] = useState('')
  const [reportId, setReportId] = useState('')
  const [generating, setGenerating] = useState(false)
  const [signing, setSigning] = useState(false)
  const [signed, setSigned] = useState(false)
  const [emailSending, setEmailSending] = useState(false)
  const [emailSent, setEmailSent] = useState(false)
  const [deliveryEmail, setDeliveryEmail] = useState('')
  const iframeRef = useRef<HTMLIFrameElement>(null)

  const fields = reportType ? REPORT_FIELDS[reportType] || [] : []

  const upd = (k: string, v: string | boolean) => setFormData(p => ({ ...p, [k]: v }))

  const handleGenerate = async () => {
    if (!reportType) return
    setGenerating(true)
    try {
      const res = await fetch(BRIDGE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'generate',
          report_type: reportType,
          patient_id: patientId,
          encounter_id: encounterId,
          provider_id: providerId,
          org_id: orgId,
          app_source: appSource,
          request_id: initialRequest?.request_id || null,
          report_data: formData,
          report_title: REPORT_TYPE_LABELS[reportType],
        }),
      })
      const data = await res.json()
      if (data.html) {
        setHtml(data.html)
        setReportId(data.report_id)
        onReportGenerated(data)
      }
    } catch (e) { console.error(e) }
    finally { setGenerating(false) }
  }

  const handleSign = async () => {
    if (!reportId) return
    setSigning(true)
    try {
      await fetch(BRIDGE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'sign',
          report_id: reportId,
          provider_id: providerId,
          org_id: orgId,
          provider_name: providerName,
          provider_creds: providerCreds,
          provider_npi: providerNpi,
          request_id: initialRequest?.request_id,
        }),
      })
      setSigned(true)
    } catch (e) { console.error(e) }
    finally { setSigning(false) }
  }

  const handlePrint = () => {
    const win = window.open('', '_blank')
    if (win) { win.document.write(html); win.document.close(); win.print() }
  }

  const handleEmailSend = async () => {
    if (!deliveryEmail || !reportId) return
    setEmailSending(true)
    try {
      await fetch(BRIDGE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'email',
          report_id: reportId,
          delivery_email: deliveryEmail,
          org_id: orgId,
          report_html: html,
        }),
      })
      setEmailSent(true)
    } catch (e) { console.error(e) }
    finally { setEmailSending(false) }
  }

  return (
    <div className="grid grid-cols-2 gap-5" style={{ minHeight: '70vh' }}>

      {/* Left: Form */}
      <div className="space-y-4 overflow-y-auto" style={{ maxHeight: '75vh' }}>

        {/* Report type selector */}
        <div className="hud-card p-4">
          <div className="hud-label mb-3">Report Type</div>
          <div className="grid grid-cols-2 gap-2">
            {availableTypes.map(([type, label]) => (
              <button key={type} onClick={() => { setReportType(type); setHtml(''); setSigned(false) }}
                className="flex items-center gap-2 p-2.5 rounded text-left text-xs transition-all"
                style={{
                  background: reportType === type ? 'rgba(201,169,110,0.12)' : 'rgba(6,14,28,0.5)',
                  border: `1px solid ${reportType === type ? '#c9a96e' : 'rgba(0,212,255,0.12)'}`,
                  color: reportType === type ? '#c9a96e' : '#e0e8f4',
                  fontFamily: 'Rajdhani', fontWeight: 600,
                }}>
                <span>{REPORT_TYPE_ICONS[type]}</span>
                {label}
              </button>
            ))}
          </div>
        </div>

        {/* Dynamic fields */}
        {fields.length > 0 && (
          <div className="hud-card p-4 space-y-3">
            <div className="hud-label">Report Content</div>
            {fields.map((f) => (
              <div key={f.key}>
                <div className="hud-label mb-1">{f.label}</div>
                {f.type === 'textarea' ? (
                  <textarea rows={3} value={(formData[f.key] as string) || ''}
                    onChange={e => upd(f.key, e.target.value)}
                    className="input-hud w-full rounded px-3 py-2 text-sm resize-none" />
                ) : f.type === 'checkbox' ? (
                  <label className="flex items-center gap-3 cursor-pointer">
                    <div onClick={() => upd(f.key, !formData[f.key])}
                      className="w-5 h-5 rounded flex items-center justify-center"
                      style={{ background: formData[f.key] ? '#00d4ff' : 'rgba(6,14,28,0.8)', border: `1px solid ${formData[f.key] ? '#00d4ff' : 'rgba(0,212,255,0.3)'}`, cursor: 'pointer' }}>
                      {formData[f.key] && <span className="text-xs font-bold" style={{ color: '#060e1c' }}>✓</span>}
                    </div>
                    <span className="text-sm">Yes</span>
                  </label>
                ) : (
                  <input type={f.type} value={(formData[f.key] as string) || ''}
                    onChange={e => upd(f.key, e.target.value)}
                    className="input-hud w-full rounded px-3 py-2 text-sm" />
                )}
              </div>
            ))}
          </div>
        )}

        <button onClick={handleGenerate} disabled={!reportType || generating}
          className="btn-primary w-full py-3 rounded text-sm disabled:opacity-40">
          {generating ? 'Generating...' : '⚡ Generate Report'}
        </button>
      </div>

      {/* Right: Preview + Actions */}
      <div className="flex flex-col gap-4">
        {html ? (
          <>
            {/* Preview iframe */}
            <div className="flex-1 rounded overflow-hidden"
              style={{ border: '1px solid rgba(0,212,255,0.2)', minHeight: '50vh' }}>
              <iframe
                ref={iframeRef}
                srcDoc={html}
                className="w-full h-full"
                style={{ minHeight: '50vh', background: 'white' }}
                title="Report Preview"
              />
            </div>

            {/* Action bar */}
            <div className="hud-card p-4 space-y-3">
              <div className="hud-label">Report Actions</div>

              <div className="flex gap-2 flex-wrap">
                <button onClick={handlePrint}
                  className="btn-cyan px-4 py-2 rounded text-xs">
                  🖨️ Print / Save PDF
                </button>

                {!signed ? (
                  <button onClick={handleSign} disabled={signing}
                    className="btn-primary px-4 py-2 rounded text-xs disabled:opacity-40">
                    {signing ? 'Signing...' : '✍️ Sign Report'}
                  </button>
                ) : (
                  <div className="px-4 py-2 rounded text-xs"
                    style={{ background: 'rgba(0,255,136,0.08)', border: '1px solid rgba(0,255,136,0.3)', color: '#00ff88' }}>
                    ✓ Signed by {providerCreds} {providerName}
                  </div>
                )}
              </div>

              {/* Email delivery */}
              <div className="space-y-2 pt-2 border-t" style={{ borderColor: 'rgba(0,212,255,0.1)' }}>
                <div className="hud-label">Encrypted Email Delivery</div>
                <div className="flex gap-2">
                  <input type="email" value={deliveryEmail} onChange={e => setDeliveryEmail(e.target.value)}
                    placeholder="recipient@example.com"
                    className="input-hud flex-1 rounded px-3 py-2 text-sm" />
                  <button onClick={handleEmailSend} disabled={!deliveryEmail || emailSending || !signed}
                    className="btn-cyan px-4 py-2 rounded text-xs disabled:opacity-40">
                    {emailSending ? 'Sending...' : '📧 Send'}
                  </button>
                </div>
                {!signed && (
                  <div className="text-xs" style={{ color: 'rgba(255,165,0,0.8)' }}>
                    ⚠️ Sign the report before sending
                  </div>
                )}
                {emailSent && (
                  <div className="text-xs" style={{ color: '#00ff88' }}>
                    ✓ Report sent securely to {deliveryEmail}
                  </div>
                )}
              </div>
            </div>
          </>
        ) : (
          <div className="flex-1 flex items-center justify-center hud-card"
            style={{ minHeight: '50vh', borderStyle: 'dashed' }}>
            <div className="text-center opacity-40">
              <div className="text-4xl mb-3">📄</div>
              <div className="hud-label">Select report type and fill fields,<br />then click Generate Report</div>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}
