// ============================================================
// PatientTrac Revela — Core Type Definitions
// ============================================================

// --- Auth & User Roles ---

export type UserRole = 'surgeon' | 'anesthesiologist' | 'nurse' | 'admin' | 'billing';

export interface User {
  id: string;
  email: string;
  firstName: string;
  lastName: string;
  role: UserRole;
  credentials?: string[];       // e.g. ['MD', 'FACS']
  npi?: string;
  lastLogin?: string;
  mfaEnabled: boolean;
}

// --- Clinical Flag ---

export type FlagSeverity = 'red' | 'amber' | 'blue' | 'green';

export interface ClinicalFlag {
  id: string;
  severity: FlagSeverity;
  category: 'vitals' | 'medication' | 'surgical' | 'anesthesia' | 'psychological' | 'planning';
  title: string;
  detail: string;
  suggestedAction?: string;
  createdAt: string;
  resolvedAt?: string;
  resolvedBy?: string;
}

// --- Patient Record ---

export type ReconstructionType =
  | 'diep-flap'
  | 'tram-flap'
  | 'lat-dorsi'
  | 'implant-expander'
  | 'implant-direct'
  | 'nar'                // nipple-areola reconstruction
  | 'rhinoplasty'
  | 'facelift'
  | 'blepharoplasty'
  | 'abdominoplasty'
  | 'breast-augmentation'
  | 'breast-reduction'
  | 'liposuction'
  | 'scar-revision'
  | 'other';

export interface SurgicalHistory {
  id: string;
  procedureType: ReconstructionType;
  procedureName: string;
  date: string;
  surgeon?: string;
  facility?: string;
  notes?: string;
}

export interface OncologyHistory {
  cancerType?: string;
  stage?: string;
  receptorStatus?: string;       // e.g. 'ER+/PR+/HER2-'
  mastectomyType?: string;
  mastectomyDate?: string;
  radiationCompleted: boolean;
  radiationDetails?: string;
  chemotherapyCompleted: boolean;
  chemotherapyRegimen?: string[];
}

export interface Vitals {
  bloodPressureSystolic?: number;
  bloodPressureDiastolic?: number;
  heartRate?: number;
  weight?: number;               // kg
  height?: number;               // cm
  bmi?: number;
  recordedAt: string;
}

export interface Medication {
  name: string;
  dose: string;
  frequency: string;
  indication?: string;
}

export interface Patient {
  id: string;
  mrn: string;
  firstName: string;
  lastName: string;
  dateOfBirth: string;
  email?: string;
  phone?: string;
  address?: string;
  insuranceProvider?: string;
  insuranceMemberId?: string;
  primarySurgeon?: string;
  oncologyHistory?: OncologyHistory;
  surgicalHistory: SurgicalHistory[];
  currentMedications: Medication[];
  allergies: string[];
  comorbidities: string[];
  vitals?: Vitals;
  latestFlags: ClinicalFlag[];
  proposalIds: string[];
  intakeComplete: boolean;
  createdAt: string;
  updatedAt: string;
}

// --- OR Scheduling ---

export type ORStatus = 'scheduled' | 'in-progress' | 'complete' | 'cancelled' | 'blocked';

export interface ORSuite {
  id: string;
  name: string;               // e.g. 'OR Suite 1'
  floor?: string;
  isActive: boolean;
}

export interface ORStaffAssignment {
  userId: string;
  role: 'surgeon' | 'anesthesiologist' | 'crna' | 'scrub-nurse' | 'circulator' | 'resident';
  confirmedAt?: string;
}

export interface ORCase {
  id: string;
  suiteId: string;
  patientId: string;
  surgeonId: string;
  procedureType: ReconstructionType;
  procedureName: string;
  scheduledStart: string;        // ISO datetime
  estimatedDurationMinutes: number;
  actualStart?: string;
  actualEnd?: string;
  status: ORStatus;
  staff: ORStaffAssignment[];
  preOpFlagsCleared: boolean;
  anesthesiaCleared: boolean;
  notes?: string;
  createdAt: string;
}

// --- Surgical Proposal ---

export type ProposalStatus = 'draft' | 'pending-approval' | 'sent' | 'accepted' | 'declined' | 'expired';

export interface CostLineItem {
  label: string;
  amount: number;
  category: 'surgeon' | 'anesthesia' | 'facility' | 'implants' | 'post-op' | 'supplies' | 'other';
  editable: boolean;
}

export type FinancingType = 'pay-in-full' | 'carecredit-24mo' | 'in-house';

export interface FinancingOption {
  type: FinancingType;
  label: string;
  totalAmount: number;
  downPayment?: number;
  monthlyPayment?: number;
  termMonths?: number;
  interestRate?: number;
  discountApplied?: number;
}

export interface SurgicalProposal {
  id: string;
  patientId: string;
  surgeonId: string;
  procedureType: ReconstructionType;
  procedureName: string;
  lineItems: CostLineItem[];
  totalEstimate: number;
  financingOptions: FinancingOption[];
  physicianNotes?: string;
  status: ProposalStatus;
  approvedBy?: string;
  approvedAt?: string;
  sentAt?: string;
  respondedAt?: string;
  pdfUrl?: string;
  createdAt: string;
  updatedAt: string;
}

// --- Inventory ---

export interface InventoryItem {
  id: string;
  name: string;
  category: 'implant' | 'suture' | 'drain' | 'instrument' | 'disposable' | 'medication' | 'other';
  sku?: string;
  supplier?: string;
  unitCost: number;
  quantityOnHand: number;
  parLevel: number;
  reorderPoint: number;
  expirationDate?: string;
  location?: string;
  lastUpdated: string;
}

// --- Audit Log ---

export interface AuditEntry {
  id: string;
  userId: string;
  action: 'read' | 'create' | 'update' | 'delete' | 'login' | 'logout' | 'export';
  resourceType: string;
  resourceId?: string;
  ipAddress: string;
  userAgent?: string;
  timestamp: string;
  details?: Record<string, unknown>;
}
