import { useState, useEffect } from 'react'
import { parseSessionFromUrl, validateBridgeSession, supabase } from './lib/supabase.js'
import RevelaHeader from './components/shared/RevelaHeader.jsx'
import TabChart   from './components/chart/TabChart.jsx'
import TabProgNote from './components/prognote/TabProgNote.jsx'
import TabOpNote  from './components/opnote/TabOpNote.jsx'
import TabPostOp  from './components/postop/TabPostOp.jsx'
import RevelaLogin from './RevelaLogin.jsx'

const TABS = [
  { id:'chart',    label:'Patient Chart',  icon:'📋' },
  { id:'prognote', label:'Progress Note',  icon:'📝' },
  { id:'opnote',   label:'Operative Note', icon:'🔬' },
  { id:'postop',   label:'Post-Op Plan',   icon:'🩺' },
]

const DEV_SESSION = {
  token:'dev', encounter_id:'1', patient_id:'1', provider_id:'1',
  org_id:'00000000-0000-0000-0000-000000000001',
}

export default function App() {
  // 'loading' | 'login' | 'bridge_ready' | 'direct_ready' | 'error'
  const [status,  setStatus]  = useState('loading')
  const [session, setSession] = useState(null)
  const [ctx,     setCtx]     = useState(null)
  const [tab,     setTab]     = useState('chart')
  const [error,   setError]   = useState(null)
  const [directAuth, setDirectAuth] = useState(null)
  const IS_DEV = import.meta.env.DEV

  useEffect(() => {
    const init = async () => {
      // 1) Bridge mode — URL token from PatientTracForge
      let s = parseSessionFromUrl()
      if (!s && IS_DEV) s = DEV_SESSION

      if (s) {
        setSession(s)
        try {
          const data = await validateBridgeSession(s.token, s.encounter_id, s.patient_id, s.provider_id, s.org_id)
          setCtx(data)
          setStatus('bridge_ready')
        } catch {
          if (IS_DEV) {
            setCtx({
              patient: { patient_id:1, first_name:'Jane', last_name:'Doe', birth:'1975-06-15', gender:'F' },
              encounter: { encounter_id:1, encounter_date:new Date().toISOString(), appointment_type:'Breast Augmentation Consult', chief_complaint:'Augmentation consultation', encounter_status:'open' },
              provider: { provider_id:1, first_name:'Dr. Sarah', last_name:'Chen', specialty:'Plastic Surgery', credentials:'MD, FACS', npi:'1234567890' },
              medications:[], diagnoses:[],
            })
            setStatus('bridge_ready')
          } else {
            setStatus('error')
            setError('Session validation failed. Return to scheduling and try again.')
          }
        }
        return
      }

      // 2) No bridge token — check for an existing direct-login session
      const { data: { session: authSession } } = await supabase.auth.getSession()
      if (authSession?.user) {
        await loadDirectContext(authSession.user)
        return
      }

      // 3) Otherwise show login screen
      setStatus('login')
    }
    init()
  }, [])

  const loadDirectContext = async (user) => {
    try {
      const { data: member } = await supabase
        .from('org_members')
        .select('org_id, role, first_name, last_name')
        .eq('user_id', user.id)
        .eq('is_active', true)
        .order('created_at', { ascending: false })
        .limit(1)
        .maybeSingle()
      if (!member) {
        await supabase.auth.signOut()
        setStatus('login')
        return
      }
      const { data: providerData } = await supabase
        .from('providers')
        .select('provider_id, first_name, last_name, specialty, credentials, npi')
        .eq('email', user.email)
        .eq('active', true)
        .maybeSingle()

      setDirectAuth({ user_id: user.id, org_id: member.org_id, role: member.role, email: user.email })
      setCtx({
        patient: null, encounter: null,
        provider: providerData || {
          first_name: member.first_name, last_name: member.last_name,
          credentials: '', specialty: 'Plastic Surgery',
        },
        medications: [], diagnoses: [],
      })
      setStatus('direct_ready')
    } catch (e) {
      console.error('Direct context load failed:', e)
      setStatus('login')
    }
  }

  const handleAuthenticated = async ({ user_id }) => {
    const { data: { session: s } } = await supabase.auth.getSession()
    if (s?.user) await loadDirectContext(s.user)
  }

  const handleSignOut = async () => {
    await supabase.auth.signOut()
    setDirectAuth(null); setCtx(null); setStatus('login')
  }

  if (status === 'loading') return (
    <div style={{minHeight:'100vh',display:'flex',alignItems:'center',justifyContent:'center',background:'#060e1c',flexDirection:'column',gap:16}}>
      <svg width="48" height="48" viewBox="0 0 48 48" fill="none">
        <rect width="48" height="48" rx="12" fill="rgba(0,212,255,0.1)" stroke="rgba(0,212,255,0.3)" strokeWidth="1"/>
        <path d="M24 10L14 25C14 30.5 18.5 35 24 35C29.5 35 34 30.5 34 25L24 10Z" fill="#c9a96e" opacity="0.8"/>
        <path d="M20 24H28M24 20V28" stroke="white" strokeWidth="2" strokeLinecap="round"/>
      </svg>
      <div style={{fontFamily:'DM Sans',fontSize:20,fontWeight:700,color:'#c9a96e',letterSpacing:2}}>REVELA</div>
      <div style={{width:120,height:3,background:'rgba(0,212,255,0.1)',borderRadius:9999,overflow:'hidden'}}>
        <div style={{height:'100%',width:'40%',background:'#00d4ff',borderRadius:9999,animation:'slide 1.4s ease-in-out infinite'}}/>
      </div>
      <style>{`@keyframes slide{0%{margin-left:-40%}100%{margin-left:100%}}`}</style>
    </div>
  )

  if (status === 'login') {
    return <RevelaLogin onAuthenticated={handleAuthenticated} />
  }

  if (status === 'error') return (
    <div style={{minHeight:'100vh',display:'flex',alignItems:'center',justifyContent:'center',background:'#060e1c',padding:24}}>
      <div style={{background:'#0a1628',border:'1px solid rgba(0,212,255,0.15)',borderRadius:16,padding:'40px 32px',maxWidth:400,textAlign:'center'}}>
        <div style={{fontSize:48,marginBottom:16}}>⚠️</div>
        <div style={{color:'#c9a96e',fontSize:20,fontWeight:700,fontFamily:'DM Sans',letterSpacing:2,textTransform:'uppercase',marginBottom:12}}>Session Error</div>
        <div style={{color:'rgba(200,215,235,0.6)',fontSize:14,lineHeight:1.6,marginBottom:24}}>{error}</div>
        <a href={import.meta.env.VITE_FORGE_URL || 'https://patienttracforge.com'}
          style={{display:'inline-block',background:'rgba(0,212,255,0.1)',border:'1px solid rgba(0,212,255,0.3)',color:'#00d4ff',padding:'10px 24px',borderRadius:8,fontSize:13,fontWeight:600,textDecoration:'none'}}>
          ← Return to PatientTracForge
        </a>
      </div>
    </div>
  )

  const tabProps = { session, ctx, patient: ctx?.patient, encounter: ctx?.encounter, provider: ctx?.provider, directAuth }
  const isDirect = status === 'direct_ready'

  return (
    <div style={{minHeight:'100vh',display:'flex',flexDirection:'column',background:'#060e1c'}}>
      <RevelaHeader ctx={ctx} session={session} />

      <nav style={{display:'flex',borderBottom:'1px solid rgba(0,212,255,0.1)',background:'rgba(10,22,40,0.95)',flexShrink:0,overflowX:'auto',alignItems:'center'}}>
        {TABS.map(t => (
          <button key={t.id} onClick={() => setTab(t.id)}
            className={`${tab===t.id?'tab-active':'tab-inactive'}`}
            style={{display:'flex',alignItems:'center',gap:8,padding:'12px 20px',fontSize:12,fontFamily:'Rajdhani,DM Sans,sans-serif',fontWeight:600,letterSpacing:'0.06em',textTransform:'uppercase',whiteSpace:'nowrap',background:'none',border:'none',cursor:'pointer',transition:'all 0.15s',borderBottom:'2px solid',color: tab===t.id ? '#c9a96e' : 'rgba(200,215,235,0.4)',borderColor: tab===t.id ? '#c9a96e' : 'transparent'}}>
            <span style={{fontSize:14}}>{t.icon}</span>{t.label}
          </button>
        ))}
        {isDirect && (
          <button onClick={handleSignOut}
            style={{marginLeft:'auto',marginRight:16,padding:'6px 14px',fontSize:11,fontFamily:'DM Sans',fontWeight:600,letterSpacing:'0.05em',textTransform:'uppercase',background:'transparent',border:'1px solid rgba(201,169,110,0.3)',borderRadius:6,color:'#c9a96e',cursor:'pointer'}}>
            Sign Out
          </button>
        )}
      </nav>

      <main style={{flex:1,overflowY:'auto'}}>
        {tab === 'chart'    && <TabChart    {...tabProps} />}
        {tab === 'prognote' && <TabProgNote {...tabProps} />}
        {tab === 'opnote'   && <TabOpNote   {...tabProps} />}
        {tab === 'postop'   && <TabPostOp   {...tabProps} onClose={() => {}} />}
      </main>

      <div style={{padding:'8px 24px',borderTop:'1px solid rgba(0,212,255,0.08)',background:'rgba(6,14,28,0.98)',display:'flex',justifyContent:'space-between',alignItems:'center',fontSize:10,color:'rgba(200,215,235,0.3)',fontFamily:'DM Sans',letterSpacing:'0.04em'}}>
        <span>Revela · PatientTrac Corp · Plastic Surgery EMR · HIPAA Compliant</span>
        {ctx?.provider && <span>{ctx.provider.first_name} {ctx.provider.last_name}{ctx.provider.credentials ? `, ${ctx.provider.credentials}` : ''}</span>}
      </div>
    </div>
  )
}
