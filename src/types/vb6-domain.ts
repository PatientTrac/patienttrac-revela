// ================================================================
// PatientTrac Revela — Core Domain Types
// Source: 8 VB6 forms migrated to TypeScript
//   frmPlasticSurgery  — main navigation hub
//   frmPlastiSurg      — surgeon clinical data host
//   frmBreastExam      — breast examination module
//   frmDiagnosis       — ICD diagnosis search & selection
//   frmEmergContact    — patient emergency contact
//   fmCareIntrvn       — mental status assessment
//   frmInPatientAppt   — appointment scheduling
//   frmNursingHome     — full EMR / inpatient records
// ================================================================

// ── Shared clinical context (piPatientID/piEncounterID/piProviderID pattern) ─
export interface ClinicalContext {
  patientId: number
  encounterId: number
  providerId: number
  facilityId: number
  encounterDate?: string
}

// ── frmEmergContact ──────────────────────────────────────────────────────────
export interface EmergencyContact {
  id?: number
  patientId: number
  contactName: string
  relationship: string
  primaryPhone: string
  secondaryPhone?: string
  address?: string
  notes?: string
}

// ── frmDiagnosis → ICD_TEMPLATE / EVALUATION_MANAGEMENT tables ───────────────
export interface DiagnosisCode {
  icdCode: string
  icdDescription: string
  isPreferred?: boolean
  sequenceNo?: number
}

export interface PatientDiagnosis {
  patientId: number
  encounterId: number
  diagnoses: DiagnosisCode[]
  selectedDiagnosisText: string   // stored in EVALUATION_MANAGEMENT.DIAGNOSIS
}

// ── fmCareIntrvn — Mental Status Selection ────────────────────────────────────
export type MentalStatusGroup =
  | 'Orientation'
  | 'Memory'
  | 'Affect'
  | 'Behavior'
  | 'Cognition'
  | 'Speech'
  | 'Mood'

export interface MentalStatusTerm {
  group: MentalStatusGroup
  term: string
  selected: boolean
}

export interface MentalStatusAssessment {
  patientId: number
  encounterId: number
  providerId: number
  visitType: string               // cboGroup dropdown
  providerName: string            // cctProv text field
  terms: MentalStatusTerm[]
  notes?: string
}

// ── frmBreastExam → FN_BreastExam.BreastMeasSim OCX ─────────────────────────
export type BreastSide = 'left' | 'right' | 'bilateral'

export type BreastFindingType =
  | 'mass'
  | 'tenderness'
  | 'discharge'
  | 'skin-change'
  | 'lymph-node'
  | 'asymmetry'
  | 'implant'
  | 'normal'

export interface BreastFinding {
  side: BreastSide
  findingType: BreastFindingType
  location?: string               // clock position e.g. "2 o'clock"
  sizeCm?: number
  description?: string
  aiFlag?: boolean
  aiFlagReason?: string
}

export interface BreastExamRecord {
  patientId: number
  encounterId: number
  providerId: number
  examDate: string
  findings: BreastFinding[]
  measurementsData?: string       // BreastMeasSim serialized JSON
  impression?: string
  recommendedFollowUp?: string
  providerSignature?: string
}

// DEFAULT_CLINICAL table — Set_Default pattern (BreastExam + PlastiSurg forms)
export interface ClinicalDefault {
  providerId: number
  clinicalValue: string
  clinicalValueRef: string        // encounter/field reference key
  clinicalDesc: string            // form name context e.g. "frmPatients"
  sequenceNo: number
}

// ── frmPlasticSurgery — Main navigation hub ───────────────────────────────────
export type PlasticSurgeryPanel =
  | 'consult'
  | 'or'
  | 'orders'
  | 'patient-info'
  | 'admin'
  | 'logs'

export type PlasticSurgeryModule =
  // Consult panel
  | 'PhysicalExamination'
  | 'BreastExam'
  | 'Glutealplasty'
  | 'BodyExam'
  | 'UpperEye'
  | 'Face'
  | 'LowerEyes'
  | 'ForeheadHair'
  | 'Nose'
  | 'ProgressNote'
  // OR panel
  | 'MedicationInventory'
  | 'PostOperativeOrders'
  | 'OperativeReports'
  | 'PostOperativeExam'
  | 'DischargeSheet'
  | 'RecoveryRoom'
  | 'AnesthesiaEvaluation'
  | 'PatientCheckIn'
  | 'HistoryAndPhysical'
  | 'PreOperativeChecklist'
  | 'Anesthesia'
  | 'OperatingRoomRecord'
  | 'OperativeLogs'
  // Orders panel
  | 'OperativePlan'
  | 'FollowUpPlan'
  | 'SurgeryOrders'
  | 'MedicationOrders'
  | 'UtilizationReview'
  | 'Results'
  | 'SpecimenLogs'
  | 'OtherReports'
  | 'ProgressNotes'
  | 'Referrals'
  // Patient info panel
  | 'Schedule'
  | 'Prescriptions'
  | 'ConsentAndTreatmentForms'
  | 'HospitalRoomPhone'
  | 'Administration'
  | 'PatientContact'
  // Admin panel
  | 'Proposals'
  | 'Scheduling'
  | 'PracticeManagement'
  | 'Reports'

export interface PlasticSurgeryNavState {
  context: ClinicalContext
  activePanel: PlasticSurgeryPanel
  activeModule: PlasticSurgeryModule | null
  patientName: string
  encounterDate: string
  scheduleId?: string
}

// ── frmPlastiSurg — Surgeon clinical data child-form host ─────────────────────
export interface SurgeonClinicalData extends ClinicalContext {
  childFormOCX: string            // e.g. "FN_Orders"
  childFormControl: string        // e.g. "Anesthesia", "OpReport", "SurgSet"
  reportDate?: string
  reportDate2?: string
  defaultProvider?: string
  defaultFacility?: string
}

// Child form / OCX control name map (Load_ChildForm pattern)
export const SURGEON_CHILD_FORMS: Record<string, { ocx: string; control: string }> = {
  Administration:     { ocx: 'FN_Orders',      control: 'Administration' },
  AnesEvaluation:     { ocx: 'FN_Orders',      control: 'AnesEvaluation' },
  Anesthesia:         { ocx: 'FN_Orders',      control: 'Anesthesia' },
  OpRoomRecord:       { ocx: 'FN_Orders',      control: 'OpRoomRecord' },
  CosmeticOrders:     { ocx: 'FN_Orders',      control: 'CosmeticOrders' },
  OpReport:           { ocx: 'FN_Orders',      control: 'OpReport' },
  SurgSet:            { ocx: 'FN_Orders',      control: 'SurgSet' },
  Prescription:       { ocx: 'PT_Medication',  control: 'Prescription' },
}

// ── frmInPatientAppt — Appointment scheduling ─────────────────────────────────
export type AppointmentStatus =
  | 'scheduled'
  | 'confirmed'
  | 'checked-in'
  | 'completed'
  | 'cancelled'
  | 'no-show'

export interface InPatientAppointment {
  id?: number
  patientId: number
  providerId: number
  facilityId: number
  visitDate: string
  visitTime: string
  appointmentType: string
  status: AppointmentStatus
  previousVisitDates?: string[]   // grdVisitDate FPSpread grid data
  notes?: string
}

// ── frmNursingHome — Full EMR / Inpatient Clinical Records ────────────────────
export type NursingHomeSection =
  | 'PatientRoster'
  | 'NewPatient'
  | 'AddToVisits'
  | 'TabletProgressNotes'
  | 'InpatientClinicalRecords'
  | 'ViewEncounters'
  | 'UpdateStatus'
  | 'Coding'

export interface NursingHomeRecord {
  patientId: number
  encounterId: number
  providerId: number
  facilityId: number
  visitDate: string
  visitTime?: string
  activeRosterCount?: number
  progressNotes?: string
  internalNotes?: string
  diagnosisEntry?: string
  encounterStatus?: string
  codingComplete?: boolean
  activeSection?: NursingHomeSection
}
