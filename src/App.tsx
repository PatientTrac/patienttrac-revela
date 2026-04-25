import { useEffect, useState } from 'react'
import { supabase } from './lib/supabase'
import { auditLogin, auditView, downloadPatientEHI, initAudit } from './lib/audit'

// Revela layout + tabs (existing components)
import Header from './components/layout/Header'
import TabNav from './components/layout/TabNav'

// Existing clinical tabs
import PatientChart from './components/chart/PatientChart'
import PrognoteTab from './components/prognote/PrognoteTab'
import OpNoteTab from './components/opnote/OpNoteTab'
import PostOpTab from './components/postop/PostOpTab'

// Records & Reports module (new — drop in from records-module zip)
import RecordsModule from './components/records/RecordsModule'

// Cross-app bridge
import { parseSessionFromUrl, getEncounterData, DEV_CONTEXT } from './lib/bridge'
import { useAppStore } from './lib/store'

const IS_DEV = import.meta.env.DEV

// Tab → table mapping for audit
const TAB_TABLE_MAP: Record<string, string> = {
  chart:    'patient',
  prognote: 'surgical_prognote',
  opnote:   'operative_notes',
  postop:   'postop_plan',
  records:  'generated_reports',
}

export default function App() {
  const { activeTab, setActiveTab, setSession, setContext, context, session } = useAppStore()
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [exportingEHI, setExportingEHI] = useState(false)

  useEffect(() => {
    // §170.315(d)(2) Google OAuth login audit
    const { data: { subscription } } = supabase.auth.onAuthStateChange((event, s) => {
      if (event === 'SIGNED_IN' && s?.user) {
        auditLogin(s.user.email || '', true, false).catch(console.error)
      }
    })
    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => {
    const init = async () => {
      try {
        let ctx = parseSessionFromUrl()
        if (!ctx && IS_DEV) ctx = DEV_CONTEXT

        if (!ctx) {
          setError('No session token. Launch Revela from PatientTracForge scheduling.')
          setLoading(false)
          return
        }

        setSession(ctx)

        initAudit({
          org_id: ctx.org_id,
          user_id: ctx.provider_id,
          user_role: 'provider',
          session_id: ctx.token.slice(-16),
          app_source: 'revela',
        })

        const data = await getEncounterData(ctx)
        setContext(data)

        // §170.315(d)(2) — log patient access on load
        const patient = data?.patient as Record<string, unknown>
        if (patient?.patient_id) {
          await auditView('patient', String(patient.patient_id), String(patient.patient_id), ctx.encounter_id)
          await auditView('encounter', ctx.encounter_id, String(patient.patient_id), ctx.encounter_id)
        }
      } catch (e) {
        if (IS_DEV) {
          setContext({
            patient: { patient_id: 1, first_name: 'Maria', last_name: 'Santos', date_of_birth: '1978-06-22', gender: 'Female', phone_primary: '(305) 555-0199', email: 'maria.santos@email.com' },
            encounter: { encounter_id: 1, encounter_date: new Date().toISOString(), appointment_type: 'Plastic Surgery Consult', chief_complaint: 'Rhinoplasty consultation', encounter_status: 'open', provider_id: 1 },
            provider: { provider_id: 1, first_name: 'Dr. James', last_name: 'Revell', specialty: 'Plastic Surgery', credentials: 'MD FACS', npi: '9876543210' },
            intake: null,
            medications: [],
            diagnoses: [],
          } as Parameters<typeof setContext>[0])
        } else {
          setError('Failed to load patient context. Return to scheduling and try again.')
        }
      } finally {
        setLoading(false)
      }
    }
    init()
  }, [setSession, setContext])

  // §170.315(d)(2) — audit every tab switch
  useEffect(() => {
    if (!session || !context) return
    const patient = context.patient as Record<string, unknown>
    if (!patient?.patient_id) return
    const tableName = TAB_TABLE_MAP[activeTab]
    if (tableName) {
      auditView(tableName, String(session.encounter_id), String(patient.patient_id), session.encounter_id).catch(console.error)
    }
  }, [activeTab, session, context])

  const handleEHIExport = async () => {
    if (!context || !session) return
    const patient = context.patient as Record<string, unknown>
    if (!patient?.patient_id) return
    setExportingEHI(true)
    try {
      await downloadPatientEHI(String(patient.patient_id), session.org_id, `${patient.first_name} ${patient.last_name}` as string)
    } catch (e) {
      console.error(e)
    } finally {
      setExportingEHI(false)
    }
  }

  if (loading) return (
    <div className="min-h-screen flex items-center justify-center" style={{ background: '#060e1c' }}>
      <div className="text-center">
        <div style={{ fontFamily: 'Rajdhani', fontSize: '1.5rem', fontWeight: 700, color: '#c9a96e', marginBottom: 12 }}>
          PatientTrac Revela
        </div>
        <div style={{ fontFamily: 'DM Mono', fontSize: '0.65rem', letterSpacing: '0.2em', color: 'rgba(0,212,255,0.6)', textTransform: 'uppercase' }} className="animate-pulse">
          Loading patient context...
        </div>
      </div>
    </div>
  )

  if (error) return (
    <div className="min-h-screen flex items-center justify-center p-8" style={{ background: '#060e1c' }}>
      <div className="max-w-md text-center p-8" style={{ background: 'rgba(10,22,40,0.9)', border: '1px solid rgba(0,212,255,0.2)', borderRadius: 12 }}>
        <div style={{ fontSize: '2rem', marginBottom: 16 }}>⚠️</div>
        <div style={{ fontFamily: 'Rajdhani', fontSize: '1.2rem', color: '#ff3b3b', marginBottom: 12 }}>Session Error</div>
        <div style={{ fontSize: '0.875rem', color: 'rgba(200,215,235,0.7)', marginBottom: 24 }}>{error}</div>
        <a href="https://patienttracforge.com/schedule"
          style={{ fontFamily: 'Rajdhani', fontWeight: 700, background: '#c9a96e', color: '#060e1c', padding: '10px 24px', borderRadius: 4, textDecoration: 'none', fontSize: '0.875rem' }}>
          ← Return to Scheduling
        </a>
      </div>
    </div>
  )

  const patient = context?.patient as Record<string, unknown>
  const provider = context?.provider as Record<string, unknown>

  return (
    <div className="min-h-screen flex flex-col" style={{ background: '#060e1c' }}>
      <Header />

      {/* Tab navigation — 5 tabs including Records & Reports */}
      <nav className="flex border-b px-4"
        style={{ background: 'rgba(10,22,40,0.9)', borderColor: 'rgba(0,212,255,0.12)' }}>
        {[
          { id: 'chart',    label: 'Patient Chart',      icon: '📋' },
          { id: 'prognote', label: 'Prognote',           icon: '📝' },
          { id: 'opnote',   label: 'Operative Note',     icon: '🔬' },
          { id: 'postop',   label: 'Post-Op Plan',       icon: '🩺' },
          { id: 'records',  label: 'Records & Reports',  icon: '🗂️' },
        ].map(tab => (
          <button key={tab.id}
            onClick={() => setActiveTab(tab.id as typeof activeTab)}
            className={`flex items-center gap-1.5 px-4 py-3 text-xs transition-all ${activeTab === tab.id ? 'tab-active' : 'tab-inactive'}`}
            style={{ fontFamily: 'Rajdhani', fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' }}>
            <span>{tab.icon}</span>{tab.label}
          </button>
        ))}
      </nav>

      <main className="flex-1 overflow-y-auto">
        {activeTab === 'chart'    && <PatientChart />}
        {activeTab === 'prognote' && <PrognoteTab />}
        {activeTab === 'opnote'   && <OpNoteTab />}
        {activeTab === 'postop'   && <PostOpTab />}
        {activeTab === 'records'  && session && patient && (
          <RecordsModule
            patientId={String(patient.patient_id)}
            encounterId={String(session.encounter_id)}
            providerId={String(session.provider_id)}
            orgId={session.org_id}
            appSource="revela"
            patientName={`${patient.first_name} ${patient.last_name}`}
            providerName={`${provider?.first_name || ''} ${provider?.last_name || ''}`}
            providerCreds={String(provider?.credentials || 'MD')}
            providerNpi={String(provider?.npi || '')}
          />
        )}
      </main>

      {/* Bottom bar */}
      <div className="flex items-center justify-between px-6 py-3 border-t"
        style={{ background: 'rgba(6,14,28,0.98)', borderColor: 'rgba(0,212,255,0.1)' }}>
        <div style={{ fontFamily: 'DM Mono', fontSize: '0.6rem', letterSpacing: '0.15em', textTransform: 'uppercase', color: 'rgba(0,212,255,0.4)' }}>
          PatientTrac Revela · Plastic Surgery EMR · §170.315(d)(2) Audit Active
        </div>
        <div className="flex items-center gap-3">
          {/* §170.315(b)(10) EHI Export */}
          <button onClick={handleEHIExport} disabled={exportingEHI}
            className="px-4 py-1.5 rounded text-xs transition-all disabled:opacity-40"
            style={{ background: 'rgba(0,212,255,0.08)', border: '1px solid rgba(0,212,255,0.3)', color: '#00d4ff', fontFamily: 'Rajdhani', fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' }}>
            {exportingEHI ? '⏳ Exporting...' : '⬇ EHI Export'}
          </button>
          {/* Close encounter */}
          <button
            onClick={async () => {
              if (!session) return
              const { sendCallback } = await import('./lib/bridge')
              await sendCallback(session, 'encounter_closed', { trigger_billing_ai: true })
              window.location.href = 'https://patienttracforge.com/schedule'
            }}
            className="px-4 py-1.5 rounded text-xs font-bold transition-all"
            style={{ background: 'rgba(255,59,59,0.1)', border: '1px solid rgba(255,59,59,0.4)', color: '#ff3b3b', fontFamily: 'Rajdhani', letterSpacing: '0.08em', textTransform: 'uppercase' }}>
            Close Encounter
          </button>
        </div>
      </div>
    </div>
  )
}
