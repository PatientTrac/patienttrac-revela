// ================================================================
// PatientTrac Revela — Diagnosis Business Rules
// Source: frmDiagnosis.frm
// VB6 tables: ICD_TEMPLATE, EVALUATION_MANAGEMENT
// ================================================================

import type { DiagnosisCode, PatientDiagnosis } from '../types/vb6-domain'

// ── ICD search (from Load_PreferredDiag + cctSearch logic) ──────────────────
export function searchDiagnoses(
  terms: DiagnosisCode[],
  searchText: string
): DiagnosisCode[] {
  if (!searchText.trim()) return terms
  const q = searchText.toLowerCase()
  return terms.filter(
    (d) =>
      d.icdDescription.toLowerCase().includes(q) ||
      d.icdCode.toLowerCase().includes(q)
  )
}

// ── Format for display: "Description ~ Code" (lstData.AddItem pattern) ──────
export function formatDiagnosisListItem(d: DiagnosisCode): string {
  return `${toProperCase(d.icdDescription)} ~ ${d.icdCode}`
}

// ── Parse stored diagnosis string back to structured array ──────────────────
// VB6: Split(rsICD.Fields("DIAGNOSIS"), vbNewLine) then Split(line, " - ")
export function parseSavedDiagnoses(savedText: string): DiagnosisCode[] {
  if (!savedText?.trim()) return []
  return savedText
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const parts = line.split(' - ')
      return {
        icdDescription: parts[0]?.trim() ?? line,
        icdCode: parts[1]?.trim() ?? '',
      }
    })
}

// ── Serialize diagnoses for storage in EVALUATION_MANAGEMENT.DIAGNOSIS ───────
// VB6: stored as newline-separated "Desc - Code" strings
export function serializeDiagnoses(diagnoses: DiagnosisCode[]): string {
  return diagnoses
    .map((d) => `${toProperCase(d.icdDescription)} - ${d.icdCode}`)
    .join('\n')
}

// ── Validate diagnosis selection before save ─────────────────────────────────
export function validateDiagnosis(diagnosis: Partial<PatientDiagnosis>): string[] {
  const errors: string[] = []
  if (!diagnosis.patientId || diagnosis.patientId === 0)
    errors.push('Patient must be selected before saving diagnosis.')
  if (!diagnosis.encounterId || diagnosis.encounterId === 0)
    errors.push('An encounter must be active to record a diagnosis.')
  if (!diagnosis.diagnoses?.length)
    errors.push('At least one diagnosis must be selected.')
  return errors
}

// ── TextFound equivalent (case-insensitive prefix match) ────────────────────
// VB6: TextFound(strDiag, lstData.List(i), Len(strDiag), True)
export function textFound(
  searchStr: string,
  targetStr: string,
  length: number,
  caseInsensitive: boolean
): boolean {
  const s = caseInsensitive ? searchStr.toLowerCase() : searchStr
  const t = caseInsensitive ? targetStr.toLowerCase() : targetStr
  return t.substring(0, length) === s.substring(0, length)
}

// ── Auto-select previously saved diagnoses in list ───────────────────────────
// VB6: for each saved diag, find in lstData and set lstData.Selected(i) = True
export function autoSelectSavedDiagnoses(
  allDiagnoses: DiagnosisCode[],
  savedText: string
): DiagnosisCode[] {
  const saved = parseSavedDiagnoses(savedText)
  return allDiagnoses.map((d) => ({
    ...d,
    isPreferred: saved.some((s) =>
      textFound(s.icdDescription, d.icdDescription, s.icdDescription.length, true)
    ),
  }))
}

// ── Helper: proper case (VB6 StrConv(s, vbProperCase)) ─────────────────────
function toProperCase(str: string): string {
  return str.replace(/\w\S*/g, (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
}
