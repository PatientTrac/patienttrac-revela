export type RequestorType = 'provider_internal' | 'patient' | 'third_party'
export type ReportType =
  | 'medical_records_roi'
  | 'physician_letter'
  | 'psychiatric_evaluation'
  | 'surgical_operative_report'
  | 'progress_note'
  | 'prior_auth_letter'
  | 'disability_fmla'
  | 'court_legal_report'

export type RequestStatus = 'pending' | 'in_review' | 'approved' | 'denied' | 'fulfilled' | 'expired'

export interface RecordsRequest {
  request_id: string
  patient_id: string
  provider_id: string
  requestor_type: RequestorType
  requestor_name?: string
  requestor_org?: string
  requestor_email?: string
  request_type: ReportType
  status: RequestStatus
  urgency: 'routine' | 'urgent' | 'stat'
  created_at: string
}

export interface GeneratedReport {
  report_id: string
  request_id?: string
  report_type: ReportType
  report_title: string
  app_source: string
  report_content?: { html: string; data: Record<string, unknown> }
  signed: boolean
  signed_by_name?: string
  signed_at?: string
  email_sent: boolean
  created_at: string
}

export const REPORT_TYPE_LABELS: Record<ReportType, string> = {
  medical_records_roi: 'Medical Records Request (ROI)',
  physician_letter: 'Physician Letter / Referral',
  psychiatric_evaluation: 'Psychiatric Evaluation Report',
  surgical_operative_report: 'Surgical Operative Report',
  progress_note: 'Progress Note / Visit Summary',
  prior_auth_letter: 'Prior Authorization Letter',
  disability_fmla: 'Disability / FMLA Letter',
  court_legal_report: 'Court / Legal Report',
}

export const REPORT_TYPE_ICONS: Record<ReportType, string> = {
  medical_records_roi: '📂',
  physician_letter: '✉️',
  psychiatric_evaluation: '🧠',
  surgical_operative_report: '🔬',
  progress_note: '📝',
  prior_auth_letter: '🏥',
  disability_fmla: '♿',
  court_legal_report: '⚖️',
}

// Which apps each report type applies to
export const REPORT_APP_MAP: Record<ReportType, ('mind' | 'revela' | 'both')> = {
  medical_records_roi: 'both',
  physician_letter: 'both',
  psychiatric_evaluation: 'mind',
  surgical_operative_report: 'revela',
  progress_note: 'both',
  prior_auth_letter: 'both',
  disability_fmla: 'both',
  court_legal_report: 'both',
}
