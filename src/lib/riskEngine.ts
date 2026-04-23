// ================================================================
// PatientTrac Revela — AI Risk Intelligence Engine
// Covers: family history, substance use, surgical history,
//         cardiovascular, medications, psychiatric, oncology,
//         anesthesia, social, and compliance risk categories
// ================================================================

export type RiskSeverity = 'critical' | 'warning' | 'info'
export type RiskCategory =
  | 'cardiovascular'
  | 'family-history'
  | 'substance-history'
  | 'surgical-history'
  | 'medication'
  | 'psychiatric'
  | 'oncology'
  | 'anesthesia'
  | 'social'
  | 'compliance'
  | 'wound-healing'
  | 'documentation'

export interface ClinicalFlag {
  id: string
  severity: RiskSeverity
  category: RiskCategory
  title: string
  description: string
  evidence: string[]           // what triggered this flag
  requiredActions: FlagAction[]
  inquiryQuestions?: string[]  // follow-up questions for physician
  blocksOR: boolean
  dismissed: boolean
  dismissedBy?: string
  dismissedAt?: string
  createdAt: string
}

export interface FlagAction {
  label: string
  type: 'order' | 'referral' | 'document' | 'protocol' | 'verify'
  jiraStory?: string
}

export interface RiskScore {
  composite: number            // 0-100
  cardiovascular: number
  substanceHistory: number
  surgicalHistory: number
  familyHistory: number
  medicationRisk: number
  psychiatricRisk: number
  anesthesiaRisk: number
  woundHealingRisk: number
  level: 'low' | 'moderate' | 'high' | 'critical'
}

export interface PatientRiskProfile {
  patientId: number
  encounterId: number
  assessedAt: string
  score: RiskScore
  flags: ClinicalFlag[]
  inquiriesNeeded: InquiryItem[]
  aiSummary: string
  orClearanceStatus: 'clear' | 'conditional' | 'hold' | 'pending'
}

export interface InquiryItem {
  id: string
  question: string
  context: string
  category: RiskCategory
  response?: string
  respondedAt?: string
}

// ── Risk scoring weights ──────────────────────────────────────────────────────
const RISK_WEIGHTS = {
  hypertension_uncontrolled: 25,
  prior_radiation: 20,
  cardiac_family_history: 15,
  alcohol_use_undocumented: 12,
  prior_opioid_use: 10,
  multiple_abdominal_surgeries: 10,
  ssri_opioid_interaction: 8,
  ace_inhibitor_surgical: 8,
  bmi_over_35: 15,
  smoking_active: 18,
  smoking_former: 8,
  diabetes_uncontrolled: 20,
  anticoagulant_therapy: 15,
  malignant_hyperthermia_family: 30,
  prior_anesthesia_complications: 25,
  psychiatric_instability: 10,
  poor_social_support: 8,
  consent_incomplete: 5,
}

// ── Clinical flag rule engine ─────────────────────────────────────────────────
export interface PatientRiskInput {
  // Vitals
  bpSystolic?: number
  bpDiastolic?: number
  bmi?: number

  // Medical history
  hasHypertension?: boolean
  hasDiabetes?: boolean
  diabetesA1C?: number
  hasCOPD?: boolean
  hasChronicKidneyDisease?: boolean

  // Surgical history
  priorAbdominalSurgeries?: string[]
  priorChestWallRadiation?: boolean
  radiationDose?: string
  multipleSurgeries?: number

  // Medications
  medications?: string[]
  onAnticoagulants?: boolean
  onSSRI?: boolean
  onACEInhibitor?: boolean
  onSteroids?: boolean

  // Substance history
  smokingStatus?: 'never' | 'former' | 'current'
  alcoholUse?: 'none' | 'light' | 'moderate' | 'heavy' | 'undocumented'
  priorOpioidUse?: boolean | 'undocumented'
  recreationalDrugUse?: boolean | 'undocumented'

  // Family history
  familyCardiacHistory?: boolean
  familyCardiacAge?: number
  familyMalignantHyperthermia?: boolean
  familyCancerHistory?: string[]
  familyBleedingDisorder?: boolean

  // Psychiatric
  hasAnxiety?: boolean
  hasDepression?: boolean
  hasActiveTherapy?: boolean
  psychiatricStability?: 'stable' | 'unstable' | 'unknown'

  // Oncology
  cancerHistory?: string
  cancerStage?: string
  chemotherapyHistory?: string[]
  radiationHistory?: boolean

  // Social
  livesAlone?: boolean
  homeSupportAvailable?: boolean
  employmentStatus?: string
  transportationAvailable?: boolean

  // Documentation
  consentComplete?: boolean
  labsComplete?: boolean
  anesthesiaEvalComplete?: boolean
}

// ── Main flag generation engine ───────────────────────────────────────────────
export function generateClinicalFlags(input: PatientRiskInput): ClinicalFlag[] {
  const flags: ClinicalFlag[] = []
  const now = new Date().toISOString()
  let id = 1
  const mkId = () => `flag-${id++}-${Date.now()}`

  // ── CARDIOVASCULAR ──────────────────────────────────────────────────────────
  if (input.bpSystolic && input.bpSystolic >= 160 ||
      input.bpDiastolic && input.bpDiastolic >= 100) {
    flags.push({
      id: mkId(), severity: 'critical', category: 'cardiovascular', blocksOR: true, dismissed: false, createdAt: now,
      title: 'Severe hypertension — immediate BP management required',
      description: `BP ${input.bpSystolic}/${input.bpDiastolic} indicates Stage 2+ hypertension. Anesthesia induction risk is significantly elevated. Do not schedule OR without cardiology clearance and documented BP optimization.`,
      evidence: [`BP reading: ${input.bpSystolic}/${input.bpDiastolic} mmHg`, 'Threshold: 160/100 = Stage 2 hypertension'],
      requiredActions: [
        { label: 'Cardiology referral', type: 'referral' },
        { label: 'Order echocardiogram', type: 'order' },
        { label: 'Document repeat readings ×3', type: 'document' },
      ],
      inquiryQuestions: [
        'How long has the patient had elevated blood pressure?',
        'Is the patient compliant with antihypertensive medications?',
        'Any prior cardiac events, chest pain, or shortness of breath?',
      ],
    })
  } else if (input.bpSystolic && input.bpSystolic >= 140 ||
             input.bpDiastolic && input.bpDiastolic >= 90) {
    flags.push({
      id: mkId(), severity: 'warning', category: 'cardiovascular', blocksOR: false, dismissed: false, createdAt: now,
      title: 'Elevated BP — above safe surgical threshold',
      description: `BP ${input.bpSystolic}/${input.bpDiastolic} exceeds the 140/90 surgical safety guideline. Optimize before OR scheduling. Anesthesia team must review and approve.`,
      evidence: [`BP: ${input.bpSystolic}/${input.bpDiastolic}`, 'Surgical threshold: <140/90'],
      requiredActions: [
        { label: 'Anesthesia note', type: 'protocol' },
        { label: 'Repeat BP check', type: 'document' },
      ],
      inquiryQuestions: ['Is this reading consistent with prior measurements?'],
    })
  }

  // ── FAMILY HISTORY ──────────────────────────────────────────────────────────
  if (input.familyMalignantHyperthermia) {
    flags.push({
      id: mkId(), severity: 'critical', category: 'anesthesia', blocksOR: true, dismissed: false, createdAt: now,
      title: 'Family history of malignant hyperthermia — modified anesthesia required',
      description: 'Malignant hyperthermia (MH) is an autosomal dominant inherited condition. First-degree relative with MH requires TIVA (total intravenous anesthesia) protocol. Triggering agents (halogenated vapors, succinylcholine) must be avoided. Dantrolene must be immediately available.',
      evidence: ['Family history: MH in first-degree relative'],
      requiredActions: [
        { label: 'TIVA protocol order', type: 'protocol' },
        { label: 'MH hotline consult', type: 'referral' },
        { label: 'Dantrolene availability confirm', type: 'verify' },
      ],
      inquiryQuestions: [
        'Which relative was affected (parent, sibling, child)?',
        'Was MH confirmed by muscle biopsy or genetic testing?',
        'Has patient ever had general anesthesia without incident?',
      ],
    })
  }

  if (input.familyCardiacHistory && input.familyCardiacAge && input.familyCardiacAge < 60) {
    flags.push({
      id: mkId(), severity: 'warning', category: 'family-history', blocksOR: false, dismissed: false, createdAt: now,
      title: `Premature family cardiac event — age ${input.familyCardiacAge}`,
      description: `First-degree relative with cardiac event at age ${input.familyCardiacAge} indicates elevated inherited cardiovascular risk. Compounded by patient's current cardiovascular status. Lipid panel and ECG required.`,
      evidence: [`Family cardiac event at age ${input.familyCardiacAge}`, 'Threshold: <60 years = premature cardiac event'],
      requiredActions: [
        { label: 'Fasting lipid panel', type: 'order' },
        { label: 'Resting 12-lead ECG', type: 'order' },
        { label: 'Cardiology consult', type: 'referral' },
      ],
      inquiryQuestions: [
        'What type of cardiac event? (MI, arrhythmia, sudden death)',
        'Are other siblings or children affected?',
        'Has the patient ever had a cardiac workup?',
      ],
    })
  }

  if (input.familyCancerHistory && input.familyCancerHistory.length > 0) {
    flags.push({
      id: mkId(), severity: 'info', category: 'family-history', blocksOR: false, dismissed: false, createdAt: now,
      title: 'Family cancer history — genetic counseling consideration',
      description: `Family history of ${input.familyCancerHistory.join(', ')} noted. Depending on cancer type and degree of relation, genetic counseling and BRCA/hereditary cancer screening may be appropriate, particularly for reconstructive surgery candidates.`,
      evidence: [`Family cancer: ${input.familyCancerHistory.join(', ')}`],
      requiredActions: [
        { label: 'Genetic counseling referral', type: 'referral' },
        { label: 'Document family pedigree', type: 'document' },
      ],
      inquiryQuestions: [
        'How many first-degree relatives affected?',
        'Has patient had genetic testing (BRCA1/2, Lynch)?',
        'What types of cancer and at what ages?',
      ],
    })
  }

  // ── SUBSTANCE HISTORY ───────────────────────────────────────────────────────
  if (input.smokingStatus === 'current') {
    flags.push({
      id: mkId(), severity: 'critical', category: 'substance-history', blocksOR: true, dismissed: false, createdAt: now,
      title: 'Active smoker — wound healing and flap viability severely compromised',
      description: 'Active smoking reduces tissue oxygenation and microvascular perfusion by up to 40%. For flap-based reconstruction, smoking is associated with a 3–4× increase in flap failure, necrosis, and infection. Standard protocol requires cessation ≥6 weeks before elective surgery.',
      evidence: ['Smoking status: current active smoker'],
      requiredActions: [
        { label: 'Cessation counseling', type: 'protocol' },
        { label: 'Nicotine replacement therapy', type: 'order' },
        { label: 'Reschedule — minimum 6 weeks post-cessation', type: 'protocol' },
      ],
      inquiryQuestions: [
        'How many cigarettes/day and for how many years?',
        'Has patient attempted cessation before?',
        'Is patient willing to commit to cessation program?',
      ],
    })
  } else if (input.smokingStatus === 'former') {
    flags.push({
      id: mkId(), severity: 'warning', category: 'substance-history', blocksOR: false, dismissed: false, createdAt: now,
      title: 'Former smoker — confirm cessation date and duration',
      description: 'Former smoking affects wound healing based on cessation duration. Confirm date of cessation. Risk returns toward baseline after 6–8 weeks. Document pack-year history for anesthesia risk stratification.',
      evidence: ['Smoking status: former smoker'],
      requiredActions: [
        { label: 'Document cessation date', type: 'document' },
        { label: 'Pack-year history', type: 'document' },
      ],
      inquiryQuestions: [
        'When exactly did the patient stop smoking?',
        'Total pack-year history?',
        'Any e-cigarette or vaping use currently?',
      ],
    })
  }

  if (input.alcoholUse === 'heavy') {
    flags.push({
      id: mkId(), severity: 'critical', category: 'substance-history', blocksOR: true, dismissed: false, createdAt: now,
      title: 'Heavy alcohol use — anesthesia dosing and withdrawal risk',
      description: 'Heavy alcohol use requires modified anesthesia dosing (increased tolerance) and poses alcohol withdrawal risk in the perioperative period. CIWA protocol may be needed. Liver function tests required. INR check for coagulopathy.',
      evidence: ['Alcohol use: heavy (15+ drinks/week)'],
      requiredActions: [
        { label: 'Liver function tests', type: 'order' },
        { label: 'INR / coagulation panel', type: 'order' },
        { label: 'CIWA protocol evaluation', type: 'protocol' },
        { label: 'Addiction medicine consult', type: 'referral' },
      ],
      inquiryQuestions: [
        'Exact number of drinks per day and last drink?',
        'Any prior alcohol withdrawal seizures or DTs?',
        'Current liver disease or jaundice?',
      ],
    })
  } else if (input.alcoholUse === 'undocumented') {
    flags.push({
      id: mkId(), severity: 'warning', category: 'substance-history', blocksOR: false, dismissed: false, createdAt: now,
      title: 'Alcohol use — frequency and quantity not documented',
      description: 'Patient reported alcohol use but quantity and frequency are not recorded. Chronic use affects anesthesia dosing, wound healing, and clotting. AUDIT-C screening required.',
      evidence: ['Alcohol use: reported but not quantified'],
      requiredActions: [
        { label: 'AUDIT-C screening tool', type: 'document' },
        { label: 'Document drinks per week', type: 'document' },
      ],
      inquiryQuestions: [
        'How many drinks per week on average?',
        'Any recent increase in alcohol use?',
        'Last alcoholic drink?',
      ],
    })
  }

  if (input.priorOpioidUse === true || input.priorOpioidUse === 'undocumented') {
    flags.push({
      id: mkId(), severity: 'warning', category: 'substance-history', blocksOR: false, dismissed: false, createdAt: now,
      title: 'Prior opioid use — tolerance and postoperative pain management',
      description: 'Prior opioid use (prescription or recreational) affects postoperative pain management planning. Opioid-tolerant patients require higher doses or alternative protocols. Risk of undertreated pain or opioid-induced hyperalgesia.',
      evidence: [`Prior opioid use: ${input.priorOpioidUse === 'undocumented' ? 'reported, not detailed' : 'documented'}`],
      requiredActions: [
        { label: 'Pain management consult', type: 'referral' },
        { label: 'Document opioid history', type: 'document' },
        { label: 'PDMP check', type: 'verify' },
      ],
      inquiryQuestions: [
        'What opioids, at what doses, and for how long?',
        'Currently on opioid therapy or in MAT program?',
        'Any history of opioid use disorder or overdose?',
      ],
    })
  }

  if (input.recreationalDrugUse === 'undocumented') {
    flags.push({
      id: mkId(), severity: 'warning', category: 'substance-history', blocksOR: false, dismissed: false, createdAt: now,
      title: 'Recreational drug use — not adequately screened',
      description: 'Standard substance use screening is incomplete. Cocaine, methamphetamine, and cannabis all have significant anesthetic implications. Cocaine and meth cause sympathomimetic effects dangerous under anesthesia. Cannabis affects anesthetic requirements.',
      evidence: ['Recreational drug use: not formally screened'],
      requiredActions: [
        { label: 'Urine drug screen', type: 'order' },
        { label: 'DAST-10 screening', type: 'document' },
      ],
      inquiryQuestions: [
        'Any use of cocaine, methamphetamine, or stimulants?',
        'Cannabis use — frequency and last use?',
        'Any IV drug use history?',
      ],
    })
  }

  // ── SURGICAL HISTORY ────────────────────────────────────────────────────────
  if (input.priorChestWallRadiation) {
    flags.push({
      id: mkId(), severity: 'critical', category: 'oncology', blocksOR: false, dismissed: false, createdAt: now,
      title: 'Prior chest wall radiation — implant failure risk 3× elevated',
      description: 'Chest wall radiation causes microvascular damage and tissue fibrosis. Implant-based reconstruction has a 3× failure rate in irradiated fields. DIEP or other autologous flap strongly preferred. Radiation field, dose, and timing records required.',
      evidence: ['Prior chest wall radiation documented'],
      requiredActions: [
        { label: 'Request XRT records + dosimetry', type: 'order' },
        { label: 'Order CTA angiography — perforators', type: 'order' },
        { label: 'Plan DIEP flap preferred approach', type: 'protocol' },
      ],
      inquiryQuestions: [
        'Total radiation dose received and treatment dates?',
        'Which specific field was irradiated?',
        'Any skin changes, fibrosis, or wound healing problems since radiation?',
      ],
    })
  }

  if (input.multipleSurgeries && input.multipleSurgeries >= 3) {
    flags.push({
      id: mkId(), severity: 'warning', category: 'surgical-history', blocksOR: false, dismissed: false, createdAt: now,
      title: `${input.multipleSurgeries} prior surgeries — anesthesia sensitivity and adhesion risk`,
      description: `Multiple prior surgeries (${input.multipleSurgeries} documented) increase risk of adhesions, altered anatomy, and cumulative anesthesia exposure. Review all operative reports. Anesthesia team should review prior drug reactions and intubation history.`,
      evidence: [`Prior surgical procedures: ${input.multipleSurgeries}`],
      requiredActions: [
        { label: 'Collect all operative reports', type: 'document' },
        { label: 'Anesthesia history review', type: 'protocol' },
      ],
      inquiryQuestions: [
        'Any prior anesthesia complications or reactions?',
        'Difficult intubation history?',
        'Any prior surgeries not listed in records?',
      ],
    })
  }

  // ── MEDICATIONS ─────────────────────────────────────────────────────────────
  if (input.onAnticoagulants) {
    flags.push({
      id: mkId(), severity: 'critical', category: 'medication', blocksOR: true, dismissed: false, createdAt: now,
      title: 'Anticoagulant therapy — bridge protocol required',
      description: 'Patient on anticoagulant therapy. Warfarin patients require INR ≤1.5 and may need bridge therapy with LMWH. NOAC patients require hold period specific to drug and renal function. Hematology consult may be needed for complex cases.',
      evidence: ['Anticoagulant therapy: active'],
      requiredActions: [
        { label: 'INR level (if warfarin)', type: 'order' },
        { label: 'Bridge therapy protocol', type: 'protocol' },
        { label: 'Hematology consult', type: 'referral' },
      ],
      inquiryQuestions: [
        'Which anticoagulant and at what dose?',
        'What is the indication for anticoagulation?',
        'History of clot or PE?',
      ],
    })
  }

  if (input.onSSRI) {
    flags.push({
      id: mkId(), severity: 'warning', category: 'medication', blocksOR: false, dismissed: false, createdAt: now,
      title: 'SSRI therapy — serotonin syndrome and bleeding risk',
      description: 'SSRI use is associated with platelet dysfunction and increased surgical bleeding. If opioids are planned postoperatively, serotonin syndrome protocol required. Discuss with anesthesia and consider dose timing.',
      evidence: ['SSRI: active prescription'],
      requiredActions: [
        { label: 'Anesthesia serotonin protocol', type: 'protocol' },
        { label: 'Platelet function assessment', type: 'order' },
      ],
      inquiryQuestions: ['Which SSRI and at what dose?', 'How long has patient been on SSRI?'],
    })
  }

  if (input.onACEInhibitor) {
    flags.push({
      id: mkId(), severity: 'warning', category: 'medication', blocksOR: false, dismissed: false, createdAt: now,
      title: 'ACE inhibitor — hold morning of surgery protocol',
      description: 'ACE inhibitors should be held the morning of surgery. Intraoperative hypotension risk is significantly elevated if taken day-of. Confirm hold instructions are given and documented.',
      evidence: ['ACE inhibitor: active prescription'],
      requiredActions: [
        { label: 'Document hold instructions', type: 'document' },
        { label: 'Anesthesia notification', type: 'protocol' },
      ],
      inquiryQuestions: ['Does patient understand medication hold instructions?'],
    })
  }

  // ── WOUND HEALING ───────────────────────────────────────────────────────────
  if (input.bmi && input.bmi > 35) {
    flags.push({
      id: mkId(), severity: 'warning', category: 'wound-healing', blocksOR: false, dismissed: false, createdAt: now,
      title: `BMI ${Math.round(input.bmi)} — elevated wound complication risk`,
      description: `BMI >35 is associated with significantly increased rates of wound dehiscence, seroma, infection, and flap necrosis. Nutritional status should be optimized preoperatively. Discuss realistic expectations with patient.`,
      evidence: [`BMI: ${Math.round(input.bmi)}`, 'Threshold: BMI >35 = elevated wound risk'],
      requiredActions: [
        { label: 'Nutritional assessment', type: 'order' },
        { label: 'Albumin / pre-albumin level', type: 'order' },
      ],
      inquiryQuestions: [
        'Has patient attempted weight loss prior to procedure?',
        'Any history of wound healing problems?',
        'Diabetic or nutritional deficiencies?',
      ],
    })
  }

  // ── PSYCHIATRIC ─────────────────────────────────────────────────────────────
  if (input.psychiatricStability === 'unstable') {
    flags.push({
      id: mkId(), severity: 'warning', category: 'psychiatric', blocksOR: false, dismissed: false, createdAt: now,
      title: 'Psychiatric instability — surgical readiness assessment needed',
      description: 'Unstable psychiatric status may affect surgical decision-making capacity, postoperative compliance, and recovery. Psychiatric clearance and psychological readiness assessment recommended before elective procedure.',
      evidence: ['Psychiatric status: documented as unstable'],
      requiredActions: [
        { label: 'Psychiatric clearance', type: 'referral' },
        { label: 'Psychological readiness assessment', type: 'referral' },
      ],
      inquiryQuestions: [
        'Current medications and compliance?',
        'Any recent hospitalizations or crisis events?',
        'Support system stability?',
      ],
    })
  }

  // ── DOCUMENTATION ───────────────────────────────────────────────────────────
  if (!input.consentComplete) {
    flags.push({
      id: mkId(), severity: 'warning', category: 'documentation', blocksOR: true, dismissed: false, createdAt: now,
      title: 'Informed consent incomplete — cannot proceed to OR',
      description: 'Signed informed consent is required before any surgical procedure. AI pre-fill is available for consent documents — physician review and patient signature required.',
      evidence: ['Consent status: incomplete'],
      requiredActions: [
        { label: 'Generate consent document', type: 'document' },
        { label: 'Patient signature required', type: 'verify' },
      ],
    })
  }

  return flags
}

// ── Calculate composite risk score ───────────────────────────────────────────
export function calculateRiskScore(input: PatientRiskInput, flags: ClinicalFlag[]): RiskScore {
  let cardiovascular = 0
  let substanceHistory = 0
  let surgicalHistory = 0
  let familyHistory = 0
  let medicationRisk = 0
  let psychiatricRisk = 0
  let anesthesiaRisk = 0
  let woundHealingRisk = 0

  if (input.bpSystolic && input.bpSystolic >= 160) cardiovascular += 35
  else if (input.bpSystolic && input.bpSystolic >= 140) cardiovascular += 20
  if (input.hasDiabetes) cardiovascular += 15
  if (input.familyCardiacHistory && input.familyCardiacAge && input.familyCardiacAge < 60) {
    familyHistory += 25; cardiovascular += 10
  }
  if (input.smokingStatus === 'current') { substanceHistory += 40; woundHealingRisk += 30 }
  else if (input.smokingStatus === 'former') substanceHistory += 15
  if (input.alcoholUse === 'heavy') substanceHistory += 35
  else if (input.alcoholUse === 'undocumented') substanceHistory += 15
  if (input.priorOpioidUse) substanceHistory += 20
  if (input.recreationalDrugUse === 'undocumented') substanceHistory += 10
  if (input.priorChestWallRadiation) { surgicalHistory += 30; woundHealingRisk += 20 }
  if (input.multipleSurgeries && input.multipleSurgeries >= 3) surgicalHistory += 20
  if (input.priorAbdominalSurgeries && input.priorAbdominalSurgeries.length >= 2) surgicalHistory += 15
  if (input.onAnticoagulants) medicationRisk += 30
  if (input.onSSRI) medicationRisk += 15
  if (input.onACEInhibitor) medicationRisk += 10
  if (input.familyMalignantHyperthermia) anesthesiaRisk += 50
  if (input.psychiatricStability === 'unstable') psychiatricRisk += 25
  if (input.bmi && input.bmi > 35) woundHealingRisk += 25

  const cap = (v: number) => Math.min(100, Math.round(v))
  const scores = {
    cardiovascular: cap(cardiovascular),
    substanceHistory: cap(substanceHistory),
    surgicalHistory: cap(surgicalHistory),
    familyHistory: cap(familyHistory),
    medicationRisk: cap(medicationRisk),
    psychiatricRisk: cap(psychiatricRisk),
    anesthesiaRisk: cap(anesthesiaRisk),
    woundHealingRisk: cap(woundHealingRisk),
  }

  const composite = cap(
    scores.cardiovascular * 0.25 +
    scores.substanceHistory * 0.20 +
    scores.surgicalHistory * 0.15 +
    scores.familyHistory * 0.10 +
    scores.medicationRisk * 0.10 +
    scores.anesthesiaRisk * 0.10 +
    scores.woundHealingRisk * 0.07 +
    scores.psychiatricRisk * 0.03
  )

  const level: RiskScore['level'] =
    composite >= 70 ? 'critical' :
    composite >= 50 ? 'high' :
    composite >= 30 ? 'moderate' : 'low'

  return { composite, level, ...scores }
}

// ── Generate inquiry items from flags ────────────────────────────────────────
export function generateInquiryItems(flags: ClinicalFlag[]): InquiryItem[] {
  const items: InquiryItem[] = []
  let id = 1
  flags.forEach(flag => {
    flag.inquiryQuestions?.forEach(q => {
      items.push({
        id: `inq-${id++}`,
        question: q,
        context: flag.title,
        category: flag.category,
      })
    })
  })
  return items
}

// ── OR clearance determination ────────────────────────────────────────────────
export function determineORClearance(
  flags: ClinicalFlag[],
  score: RiskScore
): PatientRiskProfile['orClearanceStatus'] {
  const hasBlockers = flags.some(f => f.blocksOR && !f.dismissed)
  if (hasBlockers) return 'hold'
  if (score.composite >= 70) return 'conditional'
  if (score.composite >= 40) return 'conditional'
  return 'clear'
}

// ── Futuristic design tokens ──────────────────────────────────────────────────
export const REVELA_DESIGN_TOKENS = {
  colors: {
    navy:     '#060e1c',
    navy2:    '#0a1628',
    navy3:    '#112240',
    navy4:    '#162d4f',
    gold:     '#c9a96e',
    gold2:    '#e8c98a',
    teal:     '#4a7fa8',
    teal2:    '#6fb3d9',
    red:      '#e11d48',
    red2:     '#ff4d7a',
    amber:    '#f59e0b',
    green:    '#10b981',
    ivory:    '#f0ece4',
    muted:    '#3a5470',
  },
  fonts: {
    display: "'Rajdhani', sans-serif",
    body:    "'DM Sans', system-ui, sans-serif",
    mono:    "'DM Mono', monospace",
  },
  severityColors: {
    critical: { bg: 'rgba(225,29,72,0.08)', border: 'rgba(225,29,72,0.25)', text: '#ff4d7a', accent: '#e11d48' },
    warning:  { bg: 'rgba(245,158,11,0.08)', border: 'rgba(245,158,11,0.25)', text: '#fcd34d', accent: '#f59e0b' },
    info:     { bg: 'rgba(74,127,168,0.1)',  border: 'rgba(74,127,168,0.25)', text: '#6fb3d9', accent: '#4a7fa8' },
  },
  categoryIcons: {
    'cardiovascular':   'heart',
    'family-history':   'users',
    'substance-history':'alert-triangle',
    'surgical-history': 'scissors',
    'medication':       'pill',
    'psychiatric':      'brain',
    'oncology':         'activity',
    'anesthesia':       'wind',
    'social':           'home',
    'compliance':       'file-check',
    'wound-healing':    'bandage',
    'documentation':    'file-text',
  },
} as const
