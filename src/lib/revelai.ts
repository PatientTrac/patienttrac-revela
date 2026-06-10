// PatientTrac Revela — AI Function Client
// All calls go through Netlify Functions (HIPAA — API key server-side only)

const BASE = '/.netlify/functions'

// ── Types ─────────────────────────────────────────────────────────────────────

export interface SurgicalFlag {
  severity: 'critical' | 'warning' | 'info'
  category: 'contraindication' | 'drug_interaction' | 'missing_preop' | 'risk_factor' | 'guideline'
  title: string
  detail: string
  action: string
}

export interface SurgicalFlagsRequest {
  procedureType: string
  procedureName: string
  asaClass?: string
  capriniScore?: number
  stopBangScore?: number
  rcriScore?: number
  bmi?: number
  smokingStatus?: string
  diabetic?: boolean
  anticoagulants?: boolean
  allergies?: string
  plannedImplants?: boolean
  estimatedDurationHours?: number
  orClearance?: string
  psychClearance?: string
  mammogramStatus?: string
  pregnancyTestResult?: string
  encounterId?: string
  orgId?: string
  providerId?: string
}

export interface RiskNarrativeResult {
  narrative: string
  orClearanceRecommendation: 'clear' | 'conditional' | 'hold'
  orClearanceRationale: string
  requiredConsults: string[]
  keyRecommendations: string[]
}

export interface RiskNarrativeRequest {
  procedureName: string
  capriniScore?: number
  capriniRisk?: string
  stopBangScore?: number
  stopBangRisk?: string
  rcriScore?: number
  rcriRisk?: string
  asaClass?: string
  bmi?: number
  smokingStatus?: string
  estimatedDurationHours?: number
  encounterId?: string
  orgId?: string
  providerId?: string
}

export interface OpNoteDraftResult {
  preoperativeDiagnosis: string
  postoperativeDiagnosis: string
  procedurePerformed: string
  operativeNarrative: string
  findings: string
  complications: string
  ebl: string
  disposition: string
}

export interface OpNoteDraftRequest {
  procedureName: string
  procedureType: string
  laterality?: string
  anesthesiaType?: string
  anesthesiaProvider?: string
  surgeonPrimary?: string
  positioning?: string
  prepSolution?: string
  estimatedBloodLossML?: number
  crystalloidGivenML?: number
  drainsPlaced?: boolean
  drainDetails?: string
  implantDetails?: string
  tissueRemovedDetails?: string
  fatHarvestedDetails?: string
  complications?: string
  condition?: string
  disposition?: string
  asaClass?: string
  capriniScore?: number
  stopBangScore?: number
  rcriScore?: number
  woundClass?: string
  universalProtocolComplete?: boolean
  ssiAntibiotic?: string
  encounterId?: string
  orgId?: string
  providerId?: string
}

export interface PromsInsightsResult {
  overallScore: number
  overallInterpretation: string
  domainInsights: { domain: string; score: number; interpretation: string; concern: boolean }[]
  clinicalConcerns: string[]
  providerRecommendations: string[]
  proceedWithSurgery: 'yes' | 'with_caution' | 'defer_pending_evaluation'
  proceedRationale: string
}

export interface PromsInsightsRequest {
  instrument: 'BREAST-Q' | 'FACE-Q' | 'BODY-Q' | 'BODY-Q-lipo'
  domains: Record<string, number>
  procedureName: string
  procedureType: string
  bodyDysmorphiaScreening?: string
  surgicalCandidacy?: string
  psychClearance?: string
  encounterId?: string
  orgId?: string
  providerId?: string
}

// ── Client functions ──────────────────────────────────────────────────────────

export async function getSurgicalFlags(req: SurgicalFlagsRequest): Promise<SurgicalFlag[]> {
  const res = await fetch(`${BASE}/ai-surgical-flags`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(req),
  })
  if (!res.ok) throw new Error(`ai-surgical-flags: ${res.status}`)
  const data = await res.json()
  return data.flags ?? []
}

export async function getRiskNarrative(req: RiskNarrativeRequest): Promise<RiskNarrativeResult> {
  const res = await fetch(`${BASE}/ai-risk-narrative`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(req),
  })
  if (!res.ok) throw new Error(`ai-risk-narrative: ${res.status}`)
  return res.json()
}

export async function getOpNoteDraft(req: OpNoteDraftRequest): Promise<OpNoteDraftResult> {
  const res = await fetch(`${BASE}/ai-op-note-draft`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(req),
  })
  if (!res.ok) throw new Error(`ai-op-note-draft: ${res.status}`)
  return res.json()
}

export async function getPromsInsights(req: PromsInsightsRequest): Promise<PromsInsightsResult> {
  const res = await fetch(`${BASE}/ai-proms-insights`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(req),
  })
  if (!res.ok) throw new Error(`ai-proms-insights: ${res.status}`)
  return res.json()
}
