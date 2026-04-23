// ================================================================
// PatientTrac Revela — Breast Exam & Mental Status Rules
// Source: frmBreastExam.frm + fmCareIntrvn.frm
// ================================================================

import type {
  BreastExamRecord,
  BreastFinding,
  MentalStatusAssessment,
  MentalStatusTerm,
  MentalStatusGroup,
  ClinicalDefault,
} from '../types/vb6-domain'

// ════════════════════════════════════════════════════════════════
// BREAST EXAM (frmBreastExam)
// ════════════════════════════════════════════════════════════════

// ── Validate before saving (Validate_PiValues pattern) ───────────────────────
export function validateBreastExam(record: Partial<BreastExamRecord>): string[] {
  const errors: string[] = []
  if (!record.patientId || record.patientId === 0)
    errors.push('Patient must be selected.')
  if (!record.encounterId || record.encounterId === 0)
    errors.push('Active encounter is required.')
  if (!record.providerId || record.providerId === 0)
    errors.push('Provider must be assigned.')
  if (!record.examDate)
    errors.push('Exam date is required.')
  return errors
}

// ── AI flag rule: findings that warrant clinical flag ────────────────────────
export function applyBreastExamAIFlags(findings: BreastFinding[]): BreastFinding[] {
  return findings.map((f) => {
    const shouldFlag =
      f.findingType === 'mass' ||
      f.findingType === 'discharge' ||
      f.findingType === 'skin-change' ||
      f.findingType === 'lymph-node' ||
      (f.sizeCm !== undefined && f.sizeCm > 2)

    if (shouldFlag) {
      return {
        ...f,
        aiFlag: true,
        aiFlagReason: buildFlagReason(f),
      }
    }
    return f
  })
}

function buildFlagReason(f: BreastFinding): string {
  switch (f.findingType) {
    case 'mass':
      return `Palpable mass found — ${f.side} ${f.location ?? ''}. Imaging recommended.`
    case 'discharge':
      return `Nipple discharge noted — ${f.side}. Clinical evaluation required.`
    case 'skin-change':
      return `Skin change observed — ${f.side}. Document and monitor.`
    case 'lymph-node':
      return `Lymph node finding — ${f.side}. Surgical review recommended.`
    default:
      if (f.sizeCm && f.sizeCm > 2)
        return `Finding >2cm — ${f.side} ${f.location ?? ''}. Requires further workup.`
      return 'Clinical finding requires physician review.'
  }
}

// ── Build DEFAULT_CLINICAL record (Set_Default pattern from VB6) ─────────────
// VB6: INSERT/UPDATE DEFAULT_CLINICAL where PROVIDER_ID = mintProviderID
//      and CLINICAL_VALUE_REF = vstrEncounter and CLINICAL_DESC = "frmPatients"
export function buildClinicalDefault(
  providerId: number,
  encounterRef: string,
  value: string,
  sequenceNo = 1
): ClinicalDefault {
  return {
    providerId,
    clinicalValue: value,
    clinicalValueRef: encounterRef,
    clinicalDesc: 'frmPatients',    // matches VB6 strOCX constant
    sequenceNo,
  }
}

// ── Encounter navigation (cmdRevEnctr — Previous/Next/Current encounter) ──────
export type EncounterDirection = 'previous' | 'next' | 'current'

export function resolveEncounterNavigation(
  currentEncounterId: number,
  encounterIds: number[],
  direction: EncounterDirection
): number {
  const sorted = [...encounterIds].sort((a, b) => a - b)
  const idx = sorted.indexOf(currentEncounterId)
  if (direction === 'current') return currentEncounterId
  if (direction === 'previous') return sorted[Math.max(0, idx - 1)]
  if (direction === 'next') return sorted[Math.min(sorted.length - 1, idx + 1)]
  return currentEncounterId
}

// ════════════════════════════════════════════════════════════════
// MENTAL STATUS (fmCareIntrvn — "Mental Status Selection")
// ════════════════════════════════════════════════════════════════

// ── Default mental status terms by group (from cboGroup + FPSpread grid) ─────
export const MENTAL_STATUS_TERMS: Record<MentalStatusGroup, string[]> = {
  Orientation: [
    'Oriented x1 (person)', 'Oriented x2 (person, place)',
    'Oriented x3 (person, place, time)', 'Oriented x4 (person, place, time, situation)',
    'Disoriented', 'Confused',
  ],
  Memory: [
    'Intact recent memory', 'Intact remote memory', 'Impaired short-term memory',
    'Impaired long-term memory', 'Amnesia', 'Confabulation',
  ],
  Affect: [
    'Flat', 'Blunted', 'Restricted', 'Full range', 'Labile',
    'Euphoric', 'Dysphoric', 'Anxious', 'Irritable',
  ],
  Behavior: [
    'Cooperative', 'Agitated', 'Combative', 'Restless', 'Withdrawn',
    'Impulsive', 'Appropriate', 'Inappropriate',
  ],
  Cognition: [
    'Intact judgment', 'Impaired judgment', 'Good insight', 'Poor insight',
    'Abstract thinking intact', 'Concrete thinking', 'Thought disorganized',
  ],
  Speech: [
    'Normal rate and rhythm', 'Pressured speech', 'Slow speech', 'Slurred speech',
    'Tangential', 'Circumstantial', 'Loose associations', 'Mute',
  ],
  Mood: [
    'Euthymic', 'Depressed', 'Elevated', 'Irritable', 'Anxious',
    'Hopeless', 'Stable', 'Labile',
  ],
}

// ── Default terms as full MentalStatusTerm array ──────────────────────────────
export function getDefaultMentalStatusTerms(): MentalStatusTerm[] {
  return (Object.entries(MENTAL_STATUS_TERMS) as [MentalStatusGroup, string[]][]).flatMap(
    ([group, terms]) => terms.map((term) => ({ group, term, selected: false }))
  )
}

// ── Visit type options (cboGroup dropdown — Show All toggle) ──────────────────
export const VISIT_TYPE_OPTIONS = [
  'Initial Evaluation',
  'Follow-Up',
  'Post-Operative',
  'Pre-Operative',
  'Emergency',
  'Consultation',
  'Routine Check-Up',
]

// ── Format mental status for clinical note (selected terms as prose) ──────────
export function formatMentalStatusNote(assessment: MentalStatusAssessment): string {
  const selected = assessment.terms.filter((t) => t.selected)
  if (!selected.length) return 'Mental status not documented.'

  const byGroup = selected.reduce<Partial<Record<MentalStatusGroup, string[]>>>(
    (acc, t) => {
      if (!acc[t.group]) acc[t.group] = []
      acc[t.group]!.push(t.term)
      return acc
    },
    {}
  )

  return Object.entries(byGroup)
    .map(([group, terms]) => `${group}: ${terms!.join(', ')}`)
    .join('. ') + '.'
}

// ── Validate mental status assessment ────────────────────────────────────────
export function validateMentalStatus(assessment: Partial<MentalStatusAssessment>): string[] {
  const errors: string[] = []
  if (!assessment.patientId || assessment.patientId === 0)
    errors.push('Patient required.')
  if (!assessment.providerId || assessment.providerId === 0)
    errors.push('Provider required.')
  if (!assessment.visitType?.trim())
    errors.push('Visit type must be selected.')
  return errors
}
