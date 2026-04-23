import { useState, useCallback } from 'react'
import {
  generateClinicalFlags,
  calculateRiskScore,
  determineORClearance,
  type PatientRiskInput,
  type ClinicalFlag,
  type RiskScore,
} from '@/lib/riskEngine'

// ── Design tokens ─────────────────────────────────────────────
const C = {
  navy: '#060e1c', navy2: '#0a1628', navy3: '#112240', navy4: '#162d4f',
  gold: '#c9a96e', gold2: '#e8c98a',
  teal: '#4a7fa8', teal2: '#6fb3d9',
  red: '#e11d48', red2: '#ff4d7a',
  amber: '#f59e0b', amber2: '#fcd34d',
  green: '#10b981', green2: '#6ee7b7',
  ivory: '#f0ece4', muted: '#3a5470',
}

const SEV_COLORS = {
  critical: { bg: 'rgba(225,29,72,0.08)', border: 'rgba(225,29,72,0.25)', text: '#ff4d7a', bar: '#e11d48' },
  warning:  { bg: 'rgba(245,158,11,0.08)', border: 'rgba(245,158,11,0.22)', text: '#fcd34d', bar: '#f59e0b' },
  info:     { bg: 'rgba(74,127,168,0.1)',  border: 'rgba(74,127,168,0.22)', text: '#6fb3d9', bar: '#4a7fa8' },
}

// ── Helpers ───────────────────────────────────────────────────
const Field = ({ label, children }: { label: string; children: React.ReactNode }) => (
  <div style={{ marginBottom: 14 }}>
    <label style={{ display: 'block', fontFamily: "'DM Mono',monospace", fontSize: 9, letterSpacing: '0.16em', textTransform: 'uppercase', color: C.muted, marginBottom: 5 }}>
      {label}
    </label>
    {children}
  </div>
)

const inputStyle: React.CSSProperties = {
  width: '100%', background: C.navy2, border: `1px solid rgba(74,127,168,0.2)`,
  borderRadius: 6, padding: '8px 11px', fontFamily: "'DM Sans',system-ui,sans-serif",
  fontSize: 13, color: C.ivory, outline: 'none',
}

const selectStyle: React.CSSProperties = { ...inputStyle, cursor: 'pointer' }

const Row = ({ children }: { children: React.ReactNode }) => (
  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>{children}</div>
)

// ── Risk meter bar ─────────────────────────────────────────────
const RiskBar = ({ label, value, color }: { label: string; value: number; color: string }) => (
  <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 7 }}>
    <span style={{ fontFamily: "'DM Mono',monospace", fontSize: 10, color: C.muted, minWidth: 110 }}>{label}</span>
    <div style={{ flex: 1, height: 5, background: 'rgba(255,255,255,0.06)', borderRadius: 3, overflow: 'hidden' }}>
      <div style={{ width: `${value}%`, height: '100%', background: color, borderRadius: 3, transition: 'width 0.7s ease' }} />
    </div>
    <span style={{ fontFamily: "'DM Mono',monospace", fontSize: 10, color, minWidth: 30, textAlign: 'right' }}>{value}%</span>
  </div>
)

// ── Flag card ──────────────────────────────────────────────────
const FlagCard = ({ flag, onDismiss }: { flag: ClinicalFlag; onDismiss: (id: string) => void }) => {
  const [expanded, setExpanded] = useState(false)
  const sc = SEV_COLORS[flag.severity]
  return (
    <div style={{ background: sc.bg, border: `1px solid ${sc.border}`, borderRadius: 8, overflow: 'hidden', marginBottom: 8, borderLeft: `3px solid ${sc.bar}` }}>
      <div style={{ padding: '10px 12px 10px 14px', cursor: 'pointer' }} onClick={() => setExpanded(e => !e)}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
          <span style={{ fontFamily: "'DM Mono',monospace", fontSize: 9, letterSpacing: '0.14em', textTransform: 'uppercase', background: `${sc.bar}22`, color: sc.text, padding: '2px 7px', borderRadius: 3 }}>
            {flag.category.replace('-', ' ')}
          </span>
          {flag.blocksOR && (
            <span style={{ fontFamily: "'DM Mono',monospace", fontSize: 9, background: 'rgba(225,29,72,0.15)', color: C.red2, padding: '2px 7px', borderRadius: 3, letterSpacing: '0.1em' }}>
              BLOCKS OR
            </span>
          )}
          <span style={{ marginLeft: 'auto', fontFamily: "'DM Mono',monospace", fontSize: 10, color: C.muted }}>{expanded ? '▲' : '▼'}</span>
        </div>
        <div style={{ fontFamily: "'DM Sans',system-ui,sans-serif", fontSize: 13, fontWeight: 500, color: C.ivory }}>{flag.title}</div>
      </div>
      {expanded && (
        <div style={{ padding: '0 14px 12px 14px', borderTop: '1px solid rgba(255,255,255,0.05)' }}>
          <p style={{ fontFamily: "'DM Sans',system-ui,sans-serif", fontSize: 12, color: 'rgba(240,236,228,0.55)', lineHeight: 1.6, margin: '10px 0 10px' }}>
            {flag.description}
          </p>
          {flag.inquiryQuestions && flag.inquiryQuestions.length > 0 && (
            <div style={{ background: 'rgba(201,169,110,0.07)', border: '1px solid rgba(201,169,110,0.15)', borderRadius: 6, padding: '8px 12px', marginBottom: 10 }}>
              <div style={{ fontFamily: "'DM Mono',monospace", fontSize: 9, color: C.gold, letterSpacing: '0.14em', textTransform: 'uppercase', marginBottom: 6 }}>
                Further inquiry required
              </div>
              {flag.inquiryQuestions.map((q, i) => (
                <div key={i} style={{ fontFamily: "'DM Sans',system-ui,sans-serif", fontSize: 12, color: 'rgba(240,236,228,0.6)', marginBottom: 4, paddingLeft: 8, borderLeft: `2px solid rgba(201,169,110,0.3)` }}>
                  {q}
                </div>
              ))}
            </div>
          )}
          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
            {flag.requiredActions.map((action, i) => (
              <button key={i} style={{ padding: '4px 10px', borderRadius: 4, background: `${sc.bar}22`, border: 'none', color: sc.text, fontFamily: "'DM Mono',monospace", fontSize: 10, cursor: 'pointer', letterSpacing: '0.06em' }}>
                {action.label}
              </button>
            ))}
            <button onClick={() => onDismiss(flag.id)} style={{ padding: '4px 10px', borderRadius: 4, background: 'rgba(255,255,255,0.05)', border: 'none', color: C.muted, fontFamily: "'DM Mono',monospace", fontSize: 10, cursor: 'pointer' }}>
              Dismiss
            </button>
          </div>
        </div>
      )}
    </div>
  )
}

// ── OR Clearance badge ─────────────────────────────────────────
const ClearanceBadge = ({ status }: { status: string }) => {
  const styles: Record<string, React.CSSProperties> = {
    hold:        { background: 'rgba(225,29,72,0.15)', color: C.red2,   border: `1px solid rgba(225,29,72,0.3)` },
    conditional: { background: 'rgba(245,158,11,0.12)', color: C.amber2, border: `1px solid rgba(245,158,11,0.3)` },
    clear:       { background: 'rgba(16,185,129,0.1)', color: C.green2, border: `1px solid rgba(16,185,129,0.25)` },
    pending:     { background: 'rgba(74,127,168,0.1)', color: C.teal2,  border: `1px solid rgba(74,127,168,0.25)` },
  }
  const labels: Record<string, string> = {
    hold: 'OR HOLD', conditional: 'CONDITIONAL', clear: 'CLEAR FOR OR', pending: 'PENDING'
  }
  return (
    <span style={{ ...styles[status], padding: '5px 14px', borderRadius: 5, fontFamily: "'DM Mono',monospace", fontSize: 11, fontWeight: 700, letterSpacing: '0.1em' }}>
      {labels[status] ?? status.toUpperCase()}
    </span>
  )
}

// ── Section header ─────────────────────────────────────────────
const SectionHdr = ({ title, color = C.teal2 }: { title: string; color?: string }) => (
  <div style={{ display: 'flex', alignItems: 'center', gap: 8, margin: '20px 0 12px', paddingBottom: 8, borderBottom: '1px solid rgba(255,255,255,0.07)' }}>
    <div style={{ width: 3, height: 16, borderRadius: 2, background: color }} />
    <span style={{ fontFamily: "'Rajdhani',sans-serif", fontSize: 12, fontWeight: 700, letterSpacing: '0.14em', textTransform: 'uppercase', color }} >
      {title}
    </span>
  </div>
)

// ── Main component ─────────────────────────────────────────────
export default function IntakeModule() {
  const [form, setForm] = useState<PatientRiskInput>({
    bpSystolic: undefined, bpDiastolic: undefined, bmi: undefined,
    smokingStatus: 'never', alcoholUse: 'none',
    priorOpioidUse: false, recreationalDrugUse: false,
    familyCardiacHistory: false, familyCardiacAge: undefined,
    familyMalignantHyperthermia: false,
    priorChestWallRadiation: false, multipleSurgeries: 0,
    onAnticoagulants: false, onSSRI: false, onACEInhibitor: false,
    psychiatricStability: 'stable', consentComplete: false,
  })

  const [flags, setFlags] = useState<ClinicalFlag[]>([])
  const [score, setScore] = useState<RiskScore | null>(null)
  const [clearance, setClearance] = useState<string | null>(null)
  const [assessed, setAssessed] = useState(false)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)

  const set = useCallback((key: keyof PatientRiskInput, value: any) => {
    setForm(f => ({ ...f, [key]: value }))
    setAssessed(false)
  }, [])

  const runAssessment = useCallback(() => {
    const f = generateClinicalFlags(form)
    const s = calculateRiskScore(form, f)
    const c = determineORClearance(f, s)
    setFlags(f)
    setScore(s)
    setClearance(c)
    setAssessed(true)
  }, [form])

  const dismissFlag = useCallback((id: string) => {
    setFlags(prev => prev.map(f => f.id === id ? { ...f, dismissed: true } : f))
  }, [])

  const handleSave = async () => {
    setSaving(true)
    await new Promise(r => setTimeout(r, 800))
    setSaving(false)
    setSaved(true)
    setTimeout(() => setSaved(false), 3000)
  }

  const activeFlags = flags.filter(f => !f.dismissed)
  const criticalCount = activeFlags.filter(f => f.severity === 'critical').length
  const warningCount = activeFlags.filter(f => f.severity === 'warning').length

  const scoreColor = !score ? C.muted :
    score.composite >= 70 ? C.red2 :
    score.composite >= 50 ? C.amber2 :
    score.composite >= 30 ? C.teal2 : C.green2

  return (
    <div style={{ fontFamily: "'DM Sans',system-ui,sans-serif", background: C.navy, minHeight: '100vh', padding: 24, color: C.ivory }}>

      {/* Page header */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 20 }}>
        <div>
          <div style={{ fontFamily: "'Rajdhani',sans-serif", fontSize: 22, fontWeight: 700, letterSpacing: '0.08em' }}>AI Patient Intake</div>
          <div style={{ fontFamily: "'DM Mono',monospace", fontSize: 10, color: C.teal, letterSpacing: '0.12em', marginTop: 2 }}>
            PRE-OPERATIVE RISK ASSESSMENT · PATIENTTRAC REVELA
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          {assessed && clearance && <ClearanceBadge status={clearance} />}
          <button onClick={runAssessment} style={{ padding: '9px 20px', borderRadius: 7, background: C.teal, border: 'none', color: '#fff', fontFamily: "'DM Mono',monospace", fontSize: 11, fontWeight: 700, cursor: 'pointer', letterSpacing: '0.08em' }}>
            Run AI assessment
          </button>
          <button onClick={handleSave} style={{ padding: '9px 20px', borderRadius: 7, background: saved ? C.green : C.gold, border: 'none', color: C.navy, fontFamily: "'DM Mono',monospace", fontSize: 11, fontWeight: 700, cursor: 'pointer', letterSpacing: '0.08em' }}>
            {saving ? 'Saving…' : saved ? 'Saved ✓' : 'Save to Supabase'}
          </button>
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 380px', gap: 16 }}>

        {/* LEFT — intake form */}
        <div style={{ background: C.navy3, border: '1px solid rgba(74,127,168,0.15)', borderRadius: 10, padding: 20, position: 'relative', overflow: 'hidden' }}>
          <div style={{ position: 'absolute', top: 0, left: 0, right: 0, height: 1, background: 'linear-gradient(90deg,transparent,rgba(74,127,168,0.5),transparent)' }} />

          <SectionHdr title="Vitals" color={C.teal2} />
          <Row>
            <Field label="BP systolic (mmHg)">
              <input style={inputStyle} type="number" placeholder="e.g. 142" value={form.bpSystolic ?? ''} onChange={e => set('bpSystolic', e.target.value ? +e.target.value : undefined)} />
            </Field>
            <Field label="BP diastolic (mmHg)">
              <input style={inputStyle} type="number" placeholder="e.g. 88" value={form.bpDiastolic ?? ''} onChange={e => set('bpDiastolic', e.target.value ? +e.target.value : undefined)} />
            </Field>
          </Row>
          <Field label="BMI">
            <input style={inputStyle} type="number" placeholder="e.g. 28.4" value={form.bmi ?? ''} onChange={e => set('bmi', e.target.value ? +e.target.value : undefined)} />
          </Field>

          <SectionHdr title="Substance history" color={C.amber2} />
          <Row>
            <Field label="Smoking status">
              <select style={selectStyle} value={form.smokingStatus} onChange={e => set('smokingStatus', e.target.value)}>
                <option value="never">Never</option>
                <option value="former">Former smoker</option>
                <option value="current">Current smoker</option>
              </select>
            </Field>
            <Field label="Alcohol use">
              <select style={selectStyle} value={form.alcoholUse} onChange={e => set('alcoholUse', e.target.value)}>
                <option value="none">None</option>
                <option value="light">Light (1-7/week)</option>
                <option value="moderate">Moderate (8-14/week)</option>
                <option value="heavy">Heavy (15+/week)</option>
                <option value="undocumented">Reported — not quantified</option>
              </select>
            </Field>
          </Row>
          <Row>
            <Field label="Prior opioid use">
              <select style={selectStyle} value={form.priorOpioidUse === 'undocumented' ? 'undocumented' : form.priorOpioidUse ? 'yes' : 'no'} onChange={e => set('priorOpioidUse', e.target.value === 'undocumented' ? 'undocumented' : e.target.value === 'yes')}>
                <option value="no">No</option>
                <option value="yes">Yes — documented</option>
                <option value="undocumented">Reported — not detailed</option>
              </select>
            </Field>
            <Field label="Recreational drug use">
              <select style={selectStyle} value={form.recreationalDrugUse === 'undocumented' ? 'undocumented' : form.recreationalDrugUse ? 'yes' : 'no'} onChange={e => set('recreationalDrugUse', e.target.value === 'undocumented' ? 'undocumented' : e.target.value === 'yes')}>
                <option value="no">No / not screened</option>
                <option value="yes">Yes — documented</option>
                <option value="undocumented">Screening incomplete</option>
              </select>
            </Field>
          </Row>

          <SectionHdr title="Family history" color={C.gold} />
          <Row>
            <Field label="Cardiac history">
              <select style={selectStyle} value={form.familyCardiacHistory ? 'yes' : 'no'} onChange={e => set('familyCardiacHistory', e.target.value === 'yes')}>
                <option value="no">No</option>
                <option value="yes">Yes — first-degree relative</option>
              </select>
            </Field>
            <Field label="Age of cardiac event">
              <input style={inputStyle} type="number" placeholder="e.g. 52" disabled={!form.familyCardiacHistory} value={form.familyCardiacAge ?? ''} onChange={e => set('familyCardiacAge', e.target.value ? +e.target.value : undefined)} />
            </Field>
          </Row>
          <Row>
            <Field label="Malignant hyperthermia">
              <select style={selectStyle} value={form.familyMalignantHyperthermia ? 'yes' : 'no'} onChange={e => set('familyMalignantHyperthermia', e.target.value === 'yes')}>
                <option value="no">No</option>
                <option value="yes">Yes — first-degree relative</option>
              </select>
            </Field>
            <Field label="Cancer history">
              <input style={inputStyle} placeholder="e.g. breast, ovarian" onChange={e => set('familyCancerHistory', e.target.value ? e.target.value.split(',').map(s => s.trim()) : [])} />
            </Field>
          </Row>

          <SectionHdr title="Surgical history" color={C.red2} />
          <Row>
            <Field label="Prior chest wall radiation">
              <select style={selectStyle} value={form.priorChestWallRadiation ? 'yes' : 'no'} onChange={e => set('priorChestWallRadiation', e.target.value === 'yes')}>
                <option value="no">No</option>
                <option value="yes">Yes — documented</option>
              </select>
            </Field>
            <Field label="Total prior surgeries">
              <input style={inputStyle} type="number" min={0} placeholder="e.g. 3" value={form.multipleSurgeries ?? ''} onChange={e => set('multipleSurgeries', +e.target.value)} />
            </Field>
          </Row>

          <SectionHdr title="Current medications" color={C.teal2} />
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
            {[
              { key: 'onAnticoagulants', label: 'Anticoagulants' },
              { key: 'onSSRI', label: 'SSRI' },
              { key: 'onACEInhibitor', label: 'ACE inhibitor' },
            ].map(({ key, label }) => (
              <label key={key} style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', padding: '8px 10px', background: C.navy2, border: `1px solid ${(form as any)[key] ? 'rgba(201,169,110,0.3)' : 'rgba(255,255,255,0.06)'}`, borderRadius: 6 }}>
                <input type="checkbox" checked={(form as any)[key] || false} onChange={e => set(key as any, e.target.checked)} />
                <span style={{ fontFamily: "'DM Mono',monospace", fontSize: 11, color: (form as any)[key] ? C.gold : C.muted }}>{label}</span>
              </label>
            ))}
          </div>

          <SectionHdr title="Psychiatric & documentation" color={C.teal2} />
          <Row>
            <Field label="Psychiatric stability">
              <select style={selectStyle} value={form.psychiatricStability} onChange={e => set('psychiatricStability', e.target.value)}>
                <option value="stable">Stable</option>
                <option value="unstable">Unstable — needs clearance</option>
                <option value="unknown">Unknown</option>
              </select>
            </Field>
            <Field label="Informed consent">
              <select style={selectStyle} value={form.consentComplete ? 'yes' : 'no'} onChange={e => set('consentComplete', e.target.value === 'yes')}>
                <option value="no">Incomplete</option>
                <option value="yes">Signed and complete</option>
              </select>
            </Field>
          </Row>
        </div>

        {/* RIGHT — risk panel */}
        <div>
          {!assessed ? (
            <div style={{ background: C.navy3, border: '1px solid rgba(74,127,168,0.15)', borderRadius: 10, padding: 24, textAlign: 'center' }}>
              <div style={{ fontFamily: "'DM Mono',monospace", fontSize: 10, color: C.muted, letterSpacing: '0.14em', textTransform: 'uppercase', marginBottom: 12 }}>
                AI risk engine
              </div>
              <div style={{ fontFamily: "'Rajdhani',sans-serif", fontSize: 40, fontWeight: 700, color: 'rgba(255,255,255,0.08)', marginBottom: 8 }}>—</div>
              <div style={{ fontSize: 13, color: C.muted, lineHeight: 1.6 }}>
                Fill in the intake form and click <span style={{ color: C.teal2 }}>Run AI assessment</span> to generate clinical flags and risk score.
              </div>
            </div>
          ) : (
            <>
              {/* Score card */}
              <div style={{ background: C.navy3, border: '1px solid rgba(74,127,168,0.15)', borderRadius: 10, padding: 16, marginBottom: 12, position: 'relative', overflow: 'hidden' }}>
                <div style={{ position: 'absolute', top: 0, left: 0, right: 0, height: 1, background: `linear-gradient(90deg,transparent,${scoreColor},transparent)` }} />
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 14 }}>
                  <div>
                    <div style={{ fontFamily: "'DM Mono',monospace", fontSize: 9, color: C.muted, letterSpacing: '0.14em', textTransform: 'uppercase', marginBottom: 4 }}>Composite risk score</div>
                    <div style={{ fontFamily: "'Rajdhani',sans-serif", fontSize: 36, fontWeight: 700, color: scoreColor, lineHeight: 1 }}>
                      {score?.composite}
                      <span style={{ fontSize: 13, color: C.muted, marginLeft: 6 }}>/100 · {score?.level.toUpperCase()}</span>
                    </div>
                  </div>
                  {clearance && <ClearanceBadge status={clearance} />}
                </div>
                {score && (
                  <>
                    <RiskBar label="Cardiovascular" value={score.cardiovascular} color={score.cardiovascular >= 50 ? C.red : C.amber} />
                    <RiskBar label="Substance history" value={score.substanceHistory} color={score.substanceHistory >= 50 ? C.red : C.amber} />
                    <RiskBar label="Surgical history" value={score.surgicalHistory} color={C.amber} />
                    <RiskBar label="Family history" value={score.familyHistory} color={C.teal2} />
                    <RiskBar label="Medications" value={score.medicationRisk} color={C.teal2} />
                    <RiskBar label="Anesthesia" value={score.anesthesiaRisk} color={score.anesthesiaRisk >= 40 ? C.red : C.teal2} />
                    <RiskBar label="Wound healing" value={score.woundHealingRisk} color={C.teal2} />
                  </>
                )}
              </div>

              {/* Flags summary */}
              <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
                <div style={{ flex: 1, background: 'rgba(225,29,72,0.08)', border: '1px solid rgba(225,29,72,0.2)', borderRadius: 7, padding: '10px 12px', textAlign: 'center' }}>
                  <div style={{ fontFamily: "'Rajdhani',sans-serif", fontSize: 24, fontWeight: 700, color: C.red2 }}>{criticalCount}</div>
                  <div style={{ fontFamily: "'DM Mono',monospace", fontSize: 9, color: C.red2, letterSpacing: '0.1em' }}>CRITICAL</div>
                </div>
                <div style={{ flex: 1, background: 'rgba(245,158,11,0.08)', border: '1px solid rgba(245,158,11,0.2)', borderRadius: 7, padding: '10px 12px', textAlign: 'center' }}>
                  <div style={{ fontFamily: "'Rajdhani',sans-serif", fontSize: 24, fontWeight: 700, color: C.amber2 }}>{warningCount}</div>
                  <div style={{ fontFamily: "'DM Mono',monospace", fontSize: 9, color: C.amber2, letterSpacing: '0.1em' }}>WARNING</div>
                </div>
                <div style={{ flex: 1, background: 'rgba(74,127,168,0.08)', border: '1px solid rgba(74,127,168,0.2)', borderRadius: 7, padding: '10px 12px', textAlign: 'center' }}>
                  <div style={{ fontFamily: "'Rajdhani',sans-serif", fontSize: 24, fontWeight: 700, color: C.teal2 }}>{activeFlags.filter(f => f.severity === 'info').length}</div>
                  <div style={{ fontFamily: "'DM Mono',monospace", fontSize: 9, color: C.teal2, letterSpacing: '0.1em' }}>INFO</div>
                </div>
              </div>

              {/* Flag list */}
              <div>
                {activeFlags.length === 0 ? (
                  <div style={{ background: 'rgba(16,185,129,0.08)', border: '1px solid rgba(16,185,129,0.2)', borderRadius: 8, padding: 16, textAlign: 'center', fontFamily: "'DM Mono',monospace", fontSize: 11, color: C.green2 }}>
                    No active flags — patient may be cleared for OR
                  </div>
                ) : (
                  activeFlags.map(flag => (
                    <FlagCard key={flag.id} flag={flag} onDismiss={dismissFlag} />
                  ))
                )}
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  )
}
