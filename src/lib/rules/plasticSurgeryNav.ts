// ================================================================
// PatientTrac Revela — Plastic Surgery Navigation Rules
// Source: frmPlasticSurgery.frm + frmPlastiSurg.frm
// Preserves all Load_ChildForm routing logic from VB6
// ================================================================

import type {
  ClinicalContext,
  PlasticSurgeryModule,
  PlasticSurgeryPanel,
  SURGEON_CHILD_FORMS,
} from '../types/vb6-domain'

// ── Patient selection guard (cmdLabOrders_Click / cmdMedication_Click pattern)
// VB6: If piPatientID = 0 Then MsgBox "Please select patient first!"
export function requiresPatient(context: Partial<ClinicalContext>): string | null {
  if (!context.patientId || context.patientId === 0)
    return 'Please select a patient first.'
  return null
}

// ── Encounter guard ───────────────────────────────────────────────────────────
export function requiresEncounter(context: Partial<ClinicalContext>): string | null {
  if (!context.encounterId || context.encounterId === 0)
    return 'Select an encounter first.'
  return null
}

// ── Full context guard (Load_ChildForm pattern) ───────────────────────────────
// VB6: Validate_PiValues piPatientID, piEncounterID, piProviderID, piFacilityID
export function validateClinicalContext(context: Partial<ClinicalContext>): string[] {
  const errors: string[] = []
  if (!context.patientId || context.patientId === 0)
    errors.push('Patient ID is required.')
  if (!context.encounterId || context.encounterId === 0)
    errors.push('Encounter ID is required.')
  if (!context.providerId || context.providerId === 0)
    errors.push('Provider ID is required.')
  if (!context.facilityId || context.facilityId === 0)
    errors.push('Facility ID is required.')
  return errors
}

// ── Panel → modules mapping (fraConsult / fraOR / fraOrders / etc.) ───────────
export const PANEL_MODULES: Record<PlasticSurgeryPanel, PlasticSurgeryModule[]> = {
  consult: [
    'PhysicalExamination', 'BreastExam', 'Glutealplasty', 'BodyExam',
    'UpperEye', 'Face', 'LowerEyes', 'ForeheadHair', 'Nose', 'ProgressNote',
  ],
  or: [
    'MedicationInventory', 'PostOperativeOrders', 'OperativeReports',
    'PostOperativeExam', 'DischargeSheet', 'RecoveryRoom', 'AnesthesiaEvaluation',
    'PatientCheckIn', 'HistoryAndPhysical', 'PreOperativeChecklist',
    'Anesthesia', 'OperatingRoomRecord', 'OperativeLogs',
  ],
  orders: [
    'OperativePlan', 'FollowUpPlan', 'SurgeryOrders', 'MedicationOrders',
    'UtilizationReview', 'Results', 'SpecimenLogs', 'OtherReports',
    'ProgressNotes', 'Referrals',
  ],
  'patient-info': [
    'Schedule', 'Prescriptions', 'ConsentAndTreatmentForms',
    'HospitalRoomPhone', 'Administration', 'PatientContact',
  ],
  admin: ['Proposals', 'Scheduling', 'PracticeManagement', 'Reports'],
  logs: ['OperativeLogs'],
}

// ── Module display labels (Button Caption values from VB6) ────────────────────
export const MODULE_LABELS: Record<PlasticSurgeryModule, string> = {
  PhysicalExamination:      'Physical Examination',
  BreastExam:               'Breast Exam',
  Glutealplasty:            'Glutealplasty',
  BodyExam:                 'Body Exam',
  UpperEye:                 'Upper Eye',
  Face:                     'Face',
  LowerEyes:                'Lower Eyes',
  ForeheadHair:             'Forehead / Hair',
  Nose:                     'Nose',
  ProgressNote:             'Progress Note',
  MedicationInventory:      'Medication Inventory — Dispense',
  PostOperativeOrders:      'Post Operative Orders',
  OperativeReports:         'Operative Reports',
  PostOperativeExam:        'Post Operative Exam',
  DischargeSheet:           'Discharge Sheet',
  RecoveryRoom:             'Recovery Room',
  AnesthesiaEvaluation:     'Anesthesia Evaluation — Questionnaire',
  PatientCheckIn:           'Patient Check In',
  HistoryAndPhysical:       'History and Physical',
  PreOperativeChecklist:    'Pre-Operative Checklist',
  Anesthesia:               'Anesthesia',
  OperatingRoomRecord:      'Operating Room Record',
  OperativeLogs:            'Operative Logs',
  OperativePlan:            'Operative Plan',
  FollowUpPlan:             'Follow-Up Plan',
  SurgeryOrders:            'Surgery Orders',
  MedicationOrders:         'Medication Orders',
  UtilizationReview:        'Utilization Review',
  Results:                  'Results',
  SpecimenLogs:             'Specimen Logs',
  OtherReports:             'Other Reports',
  ProgressNotes:            'Progress Notes',
  Referrals:                'Referrals',
  Schedule:                 'Schedule',
  Prescriptions:            'Prescriptions',
  ConsentAndTreatmentForms: 'Consent and Treatment Forms',
  HospitalRoomPhone:        'Hospital Room Phone No',
  Administration:           'Administration',
  PatientContact:           'Patient Contact',
  Proposals:                'Proposals',
  Scheduling:               'Scheduling',
  PracticeManagement:       'Practice Management',
  Reports:                  'Reports',
}

// ── Modules that require encounter (guard before loading) ─────────────────────
export const REQUIRES_ENCOUNTER: Set<PlasticSurgeryModule> = new Set([
  'OperativeReports', 'PostOperativeOrders', 'PostOperativeExam',
  'DischargeSheet', 'RecoveryRoom', 'AnesthesiaEvaluation',
  'OperatingRoomRecord', 'OperativeLogs', 'OperativePlan',
  'FollowUpPlan', 'SurgeryOrders', 'Results', 'SpecimenLogs',
  'ProgressNotes',
])

// ── Modules that open as modal (vbModal in VB6) ───────────────────────────────
export const OPENS_AS_MODAL: Set<PlasticSurgeryModule> = new Set([
  'BreastExam', 'BodyExam', 'ConsentAndTreatmentForms',
])

// ── Resolve which route/module to load (Load_ChildForm equivalent) ────────────
export function resolveModuleRoute(module: PlasticSurgeryModule): string {
  const routeMap: Partial<Record<PlasticSurgeryModule, string>> = {
    BreastExam:            '/intake/breast-exam',
    PhysicalExamination:   '/intake/physical-exam',
    Proposals:             '/proposals',
    Scheduling:            '/or-scheduling',
    Administration:        '/staff',
    MedicationOrders:      '/proposals',
    OperativeReports:      '/or-scheduling/reports',
    PreOperativeChecklist: '/intake/pre-op-checklist',
    ConsentAndTreatmentForms: '/intake/consent',
    Prescriptions:         '/proposals',
  }
  return routeMap[module] ?? `/modules/${module.toLowerCase()}`
}

// ── Schedule combo data (cboSchedule in frmPlasticSurgery) ────────────────────
export function getScheduleOptions(): Array<{ value: string; label: string }> {
  return [
    { value: 'today',    label: "Today's Schedule" },
    { value: 'tomorrow', label: "Tomorrow's Schedule" },
    { value: 'week',     label: 'This Week' },
    { value: 'all',      label: 'All Appointments' },
  ]
}
