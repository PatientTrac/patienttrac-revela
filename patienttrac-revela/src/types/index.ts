// ─── Patient ──────────────────────────────────────────────────────────────────
export interface Patient {
  id: string
  mrn: string
  firstName: string
  lastName: string
  dob: string
  phone: string
  email: string
  status: 'active' | 'pre-op' | 'post-op' | 'consult'
  flags: ClinicalFlag[]
  intakeComplete: boolean
  createdAt: string
  updatedAt: string
}

// ─── Clinical Flags ───────────────────────────────────────────────────────────
export type FlagSeverity = 'critical' | 'warning' | 'info'
export interface ClinicalFlag {
  id: string
  patientId: string
  severity: FlagSeverity
  category: 'cardiovascular' | 'medication' | 'surgical-history' | 'oncology' | 'anesthesia' | 'documentation'
  title: string
  description: string
  recommendedAction: string
  dismissed: boolean
  createdAt: string
}

// ─── OR Scheduling ────────────────────────────────────────────────────────────
export type ORSuite = 'OR-1' | 'OR-2' | 'OR-3'
export interface ORSlot {
  id: string
  patientId: string
  procedure: string
  surgeonId: string
  anesthesiologistId: string
  orSuite: ORSuite
  scheduledStart: string
  estimatedDurationMins: number
  status: 'scheduled' | 'in-progress' | 'completed' | 'cancelled'
  notes?: string
}

// ─── Staff ────────────────────────────────────────────────────────────────────
export type StaffRole = 'surgeon' | 'anesthesiologist' | 'crna' | 'scrub-nurse' | 'circulator' | 'coordinator'
export interface StaffMember {
  id: string
  firstName: string
  lastName: string
  role: StaffRole
  credentials: string[]
  credentialExpiry: Record<string, string>
  availability: 'available' | 'in-or' | 'away' | 'off'
}

// ─── Proposals & Costing ──────────────────────────────────────────────────────
export interface ProcedureLineItem {
  id: string
  label: string
  amount: number
  category: 'surgeon' | 'anesthesia' | 'facility' | 'implant' | 'post-op' | 'supplies'
}
export interface SurgicalProposal {
  id: string
  patientId: string
  procedureName: string
  lineItems: ProcedureLineItem[]
  total: number
  discountPercent: number
  financingOptions: FinancingOption[]
  status: 'draft' | 'pending-review' | 'approved' | 'sent' | 'accepted'
  createdAt: string
}
export interface FinancingOption {
  id: string
  provider: string
  termMonths: number
  apr: number
  monthlyPayment: number
  requiresDownPayment: boolean
  downPaymentAmount?: number
}

// ─── Inventory ────────────────────────────────────────────────────────────────
export interface InventoryItem {
  id: string
  name: string
  category: string
  currentQty: number
  parLevel: number
  reorderQty: number
  unitCost: number
  vendor: string
  lastOrderDate: string
}

// ─── AI Intake ────────────────────────────────────────────────────────────────
export interface IntakeSection {
  id: string
  title: string
  questions: IntakeQuestion[]
  completed: boolean
}
export interface IntakeQuestion {
  id: string
  text: string
  type: 'radio' | 'checkbox' | 'scale' | 'text' | 'date'
  options?: string[]
  aiPrefilled?: boolean
  aiPrefilledValue?: string
  value?: string | string[] | number
  flags?: Omit<ClinicalFlag, 'id' | 'patientId' | 'createdAt'>[]
}

// ─── Accounting ───────────────────────────────────────────────────────────────
export interface AccountingSummary {
  mtdRevenue: number
  mtdExpenses: number
  netMargin: number
  outstandingAR: number
  collectionRate: number
  revenueByProcedure: Record<string, number>
}
