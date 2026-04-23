// ================================================================
// PatientTrac Revela — Nursing Home EMR & Appointment Rules
// Source: frmNursingHome.frm + frmInPatientAppt.frm
// ================================================================

import type {
  NursingHomeRecord,
  NursingHomeSection,
  InPatientAppointment,
  AppointmentStatus,
} from '../types/vb6-domain'

// ════════════════════════════════════════════════════════════════
// NURSING HOME EMR (frmNursingHome)
// ════════════════════════════════════════════════════════════════

// ── Section navigation labels (Button Caption values from VB6) ────────────────
export const NURSING_HOME_SECTION_LABELS: Record<NursingHomeSection, string> = {
  PatientRoster:           'Patient Roster',
  NewPatient:              'New Patient',
  AddToVisits:             'Add to Visits',
  TabletProgressNotes:     'PatientTrac Tablet Progress Notes',
  InpatientClinicalRecords:'PatientTrac Inpatient Clinical Records',
  ViewEncounters:          'View Encounters',
  UpdateStatus:            'Update Status',
  Coding:                  'Coding',
}

// ── Validate record before save ───────────────────────────────────────────────
export function validateNursingHomeRecord(record: Partial<NursingHomeRecord>): string[] {
  const errors: string[] = []
  if (!record.patientId || record.patientId === 0)
    errors.push('Patient must be selected.')
  if (!record.visitDate?.trim())
    errors.push('Visit date is required.')
  if (!record.providerId || record.providerId === 0)
    errors.push('Provider must be assigned.')
  return errors
}

// ── Search patient by last/first name (Search field pattern) ─────────────────
// VB6: Caption = "Search (Last Name, First Name)"
export function parsePatientSearchQuery(query: string): { lastName: string; firstName: string } {
  const parts = query.split(',').map((s) => s.trim())
  return {
    lastName: parts[0] ?? '',
    firstName: parts[1] ?? '',
  }
}

// ── Active roster count badge (lblActiveRosterCount) ─────────────────────────
export function formatRosterCount(count: number): string {
  if (count === 0) return 'No active patients'
  if (count === 1) return '1 active patient'
  return `${count} active patients`
}

// ── Coding completion check ───────────────────────────────────────────────────
export function isCodingRequired(record: NursingHomeRecord): boolean {
  return (
    !!record.encounterId &&
    record.encounterId > 0 &&
    !record.codingComplete
  )
}

// ── Section access guard (certain sections need patient selected first) ────────
const SECTIONS_REQUIRING_PATIENT: Set<NursingHomeSection> = new Set([
  'AddToVisits',
  'TabletProgressNotes',
  'InpatientClinicalRecords',
  'ViewEncounters',
  'UpdateStatus',
  'Coding',
])

export function canAccessSection(
  section: NursingHomeSection,
  patientId: number
): { allowed: boolean; reason?: string } {
  if (SECTIONS_REQUIRING_PATIENT.has(section) && (!patientId || patientId === 0)) {
    return { allowed: false, reason: 'Please select a patient first.' }
  }
  return { allowed: true }
}

// ── Encounter status options (cboUpdateStatus) ────────────────────────────────
export const ENCOUNTER_STATUS_OPTIONS = [
  'Open',
  'In Progress',
  'Pending Review',
  'Completed',
  'Billed',
  'Closed',
]

// ════════════════════════════════════════════════════════════════
// INPATIENT APPOINTMENTS (frmInPatientAppt)
// ════════════════════════════════════════════════════════════════

// ── Appointment status workflow ───────────────────────────────────────────────
export const APPOINTMENT_STATUS_TRANSITIONS: Record<AppointmentStatus, AppointmentStatus[]> = {
  scheduled:  ['confirmed', 'cancelled'],
  confirmed:  ['checked-in', 'cancelled', 'no-show'],
  'checked-in': ['completed'],
  completed:  [],
  cancelled:  ['scheduled'],
  'no-show':  ['scheduled'],
}

export function canTransitionTo(
  current: AppointmentStatus,
  next: AppointmentStatus
): boolean {
  return APPOINTMENT_STATUS_TRANSITIONS[current].includes(next)
}

// ── Validate appointment before saving ───────────────────────────────────────
export function validateAppointment(appt: Partial<InPatientAppointment>): string[] {
  const errors: string[] = []
  if (!appt.patientId || appt.patientId === 0)
    errors.push('Patient required.')
  if (!appt.visitDate?.trim())
    errors.push('Visit date required.')
  if (!appt.visitTime?.trim())
    errors.push('Visit time required.')
  if (!appt.providerId || appt.providerId === 0)
    errors.push('Provider required.')
  if (!appt.facilityId || appt.facilityId === 0)
    errors.push('Facility required.')
  if (!appt.appointmentType?.trim())
    errors.push('Appointment type required.')
  return errors
}

// ── Format appointment for calendar display ───────────────────────────────────
export function formatAppointmentTitle(appt: InPatientAppointment): string {
  return `${appt.visitTime} — ${appt.appointmentType} (${appt.status})`
}

// ── Previous visit dates grid (grdVisitDate FPSpread data) ────────────────────
export function formatVisitDateHistory(dates: string[]): string {
  if (!dates.length) return 'No previous visits on record.'
  return dates
    .slice()
    .sort((a, b) => new Date(b).getTime() - new Date(a).getTime())
    .map((d, i) => `${i + 1}. ${new Date(d).toLocaleDateString('en-US', { dateStyle: 'medium' })}`)
    .join('\n')
}

// ── Appointment type options ───────────────────────────────────────────────────
export const APPOINTMENT_TYPES = [
  'Initial Consultation',
  'Pre-Operative Evaluation',
  'Post-Operative Follow-Up',
  'Routine Visit',
  'Emergency',
  'Procedure',
  'Imaging Review',
  'Lab Review',
  'Telemedicine',
]
