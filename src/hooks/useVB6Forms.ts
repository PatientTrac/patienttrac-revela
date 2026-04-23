// ================================================================
// PatientTrac Revela — VB6 Migration Hooks
// Wraps all extracted business rules into React hooks
// ================================================================

import { useState, useCallback, useMemo } from 'react'
import type {
  ClinicalContext,
  DiagnosisCode,
  PatientDiagnosis,
  BreastExamRecord,
  BreastFinding,
  MentalStatusAssessment,
  PlasticSurgeryModule,
  PlasticSurgeryPanel,
  PlasticSurgeryNavState,
  InPatientAppointment,
  NursingHomeRecord,
  NursingHomeSection,
} from '../types/vb6-domain'

import {
  searchDiagnoses,
  formatDiagnosisListItem,
  serializeDiagnoses,
  parseSavedDiagnoses,
  validateDiagnosis,
  autoSelectSavedDiagnoses,
} from '../lib/rules/diagnosis'

import {
  requiresPatient,
  requiresEncounter,
  validateClinicalContext,
  PANEL_MODULES,
  MODULE_LABELS,
  REQUIRES_ENCOUNTER,
  resolveModuleRoute,
} from '../lib/rules/plasticSurgeryNav'

import {
  validateBreastExam,
  applyBreastExamAIFlags,
  getDefaultMentalStatusTerms,
  formatMentalStatusNote,
  validateMentalStatus,
  resolveEncounterNavigation,
  VISIT_TYPE_OPTIONS,
} from '../lib/rules/breastExamAndMentalStatus'

import {
  validateNursingHomeRecord,
  canAccessSection,
  validateAppointment,
  canTransitionTo,
  APPOINTMENT_STATUS_TRANSITIONS,
} from '../lib/rules/nursingHomeAndAppointments'

// ── useDiagnosis — frmDiagnosis ───────────────────────────────────────────────
export function useDiagnosis(allCodes: DiagnosisCode[], savedText?: string) {
  const [searchText, setSearchText] = useState('')
  const [selected, setSelected] = useState<DiagnosisCode[]>(() =>
    savedText ? parseSavedDiagnoses(savedText) : []
  )

  const filtered = useMemo(
    () => searchDiagnoses(allCodes, searchText),
    [allCodes, searchText]
  )

  const codesWithAutoSelect = useMemo(
    () => (savedText ? autoSelectSavedDiagnoses(allCodes, savedText) : allCodes),
    [allCodes, savedText]
  )

  const toggleDiagnosis = useCallback((code: DiagnosisCode) => {
    setSelected((prev) =>
      prev.some((d) => d.icdCode === code.icdCode)
        ? prev.filter((d) => d.icdCode !== code.icdCode)
        : [...prev, code]
    )
  }, [])

  const serialized = useMemo(() => serializeDiagnoses(selected), [selected])

  const validate = useCallback(
    (context: Partial<ClinicalContext>) =>
      validateDiagnosis({ ...context, diagnoses: selected }),
    [selected]
  )

  return {
    searchText, setSearchText,
    filtered, selected, codesWithAutoSelect,
    toggleDiagnosis, serialized,
    validate,
    formatItem: formatDiagnosisListItem,
  }
}

// ── usePlasticSurgeryNav — frmPlasticSurgery ──────────────────────────────────
export function usePlasticSurgeryNav(context: ClinicalContext, patientName: string) {
  const [activePanel, setActivePanel] = useState<PlasticSurgeryPanel>('consult')
  const [activeModule, setActiveModule] = useState<PlasticSurgeryModule | null>(null)

  const modulesForPanel = useMemo(() => PANEL_MODULES[activePanel], [activePanel])

  const selectModule = useCallback(
    (module: PlasticSurgeryModule) => {
      const patientError = requiresPatient(context)
      if (patientError) return { error: patientError }

      if (REQUIRES_ENCOUNTER.has(module)) {
        const encounterError = requiresEncounter(context)
        if (encounterError) return { error: encounterError }
      }

      setActiveModule(module)
      return { route: resolveModuleRoute(module), error: null }
    },
    [context]
  )

  const navState: PlasticSurgeryNavState = {
    context, activePanel, activeModule, patientName,
    encounterDate: context.encounterDate ?? new Date().toLocaleDateString(),
  }

  return {
    navState, activePanel, setActivePanel,
    activeModule, selectModule,
    modulesForPanel, MODULE_LABELS,
    validateContext: () => validateClinicalContext(context),
  }
}

// ── useBreastExam — frmBreastExam ──────────────────────────────────────────────
export function useBreastExam(context: ClinicalContext) {
  const [findings, setFindings] = useState<BreastFinding[]>([])
  const [impression, setImpression] = useState('')
  const [followUp, setFollowUp] = useState('')

  const addFinding = useCallback((finding: BreastFinding) => {
    const withFlags = applyBreastExamAIFlags([finding])
    setFindings((prev) => [...prev, ...withFlags])
  }, [])

  const removeFinding = useCallback((index: number) => {
    setFindings((prev) => prev.filter((_, i) => i !== index))
  }, [])

  const flaggedFindings = useMemo(
    () => findings.filter((f) => f.aiFlag),
    [findings]
  )

  const record: BreastExamRecord = {
    ...context,
    examDate: context.encounterDate ?? new Date().toISOString(),
    findings,
    impression,
    recommendedFollowUp: followUp,
  }

  const validate = useCallback(
    () => validateBreastExam(record),
    [record]
  )

  const navigateEncounter = useCallback(
    (encounterIds: number[], direction: 'previous' | 'next' | 'current') =>
      resolveEncounterNavigation(context.encounterId, encounterIds, direction),
    [context.encounterId]
  )

  return {
    findings, addFinding, removeFinding,
    flaggedFindings, impression, setImpression,
    followUp, setFollowUp,
    record, validate,
    navigateEncounter,
  }
}

// ── useMentalStatus — fmCareIntrvn ────────────────────────────────────────────
export function useMentalStatus(providerId: number, providerName: string) {
  const [terms, setTerms] = useState(getDefaultMentalStatusTerms)
  const [visitType, setVisitType] = useState('')
  const [notes, setNotes] = useState('')

  const toggleTerm = useCallback((termText: string) => {
    setTerms((prev) =>
      prev.map((t) => (t.term === termText ? { ...t, selected: !t.selected } : t))
    )
  }, [])

  const showAll = useCallback(() => {
    setTerms(getDefaultMentalStatusTerms())
  }, [])

  const assessment: MentalStatusAssessment = useMemo(
    () => ({ patientId: 0, encounterId: 0, providerId, providerName, visitType, terms, notes }),
    [providerId, providerName, visitType, terms, notes]
  )

  const clinicalNote = useMemo(
    () => formatMentalStatusNote(assessment),
    [assessment]
  )

  const validate = useCallback(
    (patientId: number) =>
      validateMentalStatus({ ...assessment, patientId }),
    [assessment]
  )

  return {
    terms, toggleTerm, showAll,
    visitType, setVisitType, visitTypeOptions: VISIT_TYPE_OPTIONS,
    notes, setNotes,
    assessment, clinicalNote, validate,
  }
}

// ── useAppointment — frmInPatientAppt ─────────────────────────────────────────
export function useAppointment(initialAppt?: Partial<InPatientAppointment>) {
  const [appt, setAppt] = useState<Partial<InPatientAppointment>>(initialAppt ?? {})

  const update = useCallback((patch: Partial<InPatientAppointment>) => {
    setAppt((prev) => ({ ...prev, ...patch }))
  }, [])

  const transition = useCallback(
    (nextStatus: InPatientAppointment['status']) => {
      if (!appt.status || !canTransitionTo(appt.status, nextStatus)) {
        return { error: `Cannot move from ${appt.status} to ${nextStatus}.` }
      }
      setAppt((prev) => ({ ...prev, status: nextStatus }))
      return { error: null }
    },
    [appt.status]
  )

  const validate = useCallback(
    () => validateAppointment(appt),
    [appt]
  )

  const allowedTransitions = useMemo(
    () => (appt.status ? APPOINTMENT_STATUS_TRANSITIONS[appt.status] : []),
    [appt.status]
  )

  return { appt, update, transition, validate, allowedTransitions }
}

// ── useNursingHome — frmNursingHome ───────────────────────────────────────────
export function useNursingHome() {
  const [record, setRecord] = useState<Partial<NursingHomeRecord>>({})
  const [activeSection, setActiveSectionState] = useState<NursingHomeSection>('PatientRoster')
  const [searchQuery, setSearchQuery] = useState('')

  const selectSection = useCallback(
    (section: NursingHomeSection) => {
      const { allowed, reason } = canAccessSection(section, record.patientId ?? 0)
      if (!allowed) return { error: reason ?? 'Access denied.' }
      setActiveSectionState(section)
      return { error: null }
    },
    [record.patientId]
  )

  const update = useCallback((patch: Partial<NursingHomeRecord>) => {
    setRecord((prev) => ({ ...prev, ...patch }))
  }, [])

  const validate = useCallback(
    () => validateNursingHomeRecord(record),
    [record]
  )

  return {
    record, update,
    activeSection, selectSection,
    searchQuery, setSearchQuery,
    validate,
  }
}
