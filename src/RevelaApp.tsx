import { useState, useEffect } from 'react'

const SUPABASE_URL = 'https://mskormozwekezjmtcylv.supabase.co'
const ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1za29ybW96d2VrZXpqbXRjeWx2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY5NzMzMzAsImV4cCI6MjA5MjU0OTMzMH0.nO9Q31CZJWRSIieLdLYVMkdu5NDDWjf1z0TgwCjBpp0'
const BRIDGE = `${SUPABASE_URL}/functions/v1/cross-app-bridge`

async function bridge(action: string, body: Record<string, any> = {}) {
  const res = await fetch(BRIDGE, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', apikey: ANON_KEY },
    body: JSON.stringify({ action, ...body }),
  })
  return res.json()
}

const s: Record<string, React.CSSProperties> = {
  shell:   { minHeight:'100vh', background:'#020a14', color:'#e8eaf0', fontFamily:"'DM Sans',sans-serif" },
  center:  { display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', minHeight:'100vh' },
  spin:    { width:32, height:32, border:'3px solid #3a4a6a', borderTopColor:'#c9a96e', borderRadius:'50%', animation:'spin 1s linear infinite' },
  errBox:  { maxWidth:480, margin:'80px auto', padding:32, background:'#0a1628', border:'1px solid rgba(255,107,107,0.3)' },
  hdr:     { background:'#060e1c', borderBottom:'1px solid rgba(201,169,110,0.15)', padding:'12px 24px', display:'flex', alignItems:'center', justifyContent:'space-between' },
  badge:   { background:'#c9a96e', color:'#020a14', padding:'4px 10px', fontSize:11, fontWeight:700, letterSpacing:'0.12em', fontFamily:'DM Mono,monospace' },
  pname:   { fontSize:20, fontWeight:600, fontFamily:'Rajdhani,sans-serif', letterSpacing:'0.02em' },
  meta:    { fontSize:12, color:'#8a9bc0', marginTop:2 },
  tabs:    { display:'flex', borderBottom:'1px solid rgba(201,169,110,0.1)', background:'#060e1c' },
  tab:     { padding:'12px 22px', background:'transparent', border:'none', borderBottom:'2px solid transparent', color:'#3a4a6a', cursor:'pointer', fontSize:12, fontFamily:'DM Mono,monospace', letterSpacing:'0.06em', textTransform:'uppercase' },
  tabOn:   { color:'#c9a96e', borderBottomColor:'#c9a96e' },
  body:    { padding:24 },
  grid:    { display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(280px,1fr))', gap:16 },
  card:    { background:'rgba(10,22,40,0.8)', border:'1px solid rgba(201,169,110,0.1)', padding:20 },
  cttl:    { fontSize:10, fontFamily:'DM Mono,monospace', color:'#c9a96e', letterSpacing:'0.12em', textTransform:'uppercase', marginBottom:10 },
  cbdy:    { fontSize:14, color:'#8a9bc0', lineHeight:1.6 },
  form:    { maxWidth:720, display:'flex', flexDirection:'column', gap:18 },
  row:     { display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 },
  fg:      { display:'flex', flexDirection:'column', gap:6 },
  lbl:     { fontSize:10, fontFamily:'DM Mono,monospace', color:'#3a4a6a', letterSpacing:'0.1em', textTransform:'uppercase' },
  inp:     { background:'rgba(6,14,28,0.8)', border:'1px solid rgba(201,169,110,0.15)', color:'#e8eaf0', padding:'10px 14px', fontSize:14, fontFamily:"'DM Sans',sans-serif", outline:'none', width:'100%' },
  ta:      { background:'rgba(6,14,28,0.8)', border:'1px solid rgba(201,169,110,0.15)', color:'#e8eaf0', padding:'12px 14px', fontSize:14, fontFamily:"'DM Sans',sans-serif", outline:'none', resize:'vertical', lineHeight:1.6, width:'100%' },
  btn:     { background:'#c9a96e', color:'#020a14', border:'none', padding:'12px 28px', fontSize:12, fontFamily:'DM Mono,monospace', letterSpacing:'0.12em', textTransform:'uppercase', cursor:'pointer', fontWeight:500, alignSelf:'flex-start' },
  btnOff:  { background:'rgba(201,169,110,0.3)', color:'#020a14', border:'none', padding:'12px 28px', fontSize:12, fontFamily:'DM Mono,monospace', letterSpacing:'0.12em', textTransform:'uppercase', cursor:'not-allowed', fontWeight:500, alignSelf:'flex-start' },
  ok:      { background:'rgba(74,222,128,0.08)', border:'1px solid rgba(74,222,128,0.2)', color:'#4ade80', padding:'8px 16px', marginBottom:16, fontSize:13 },
  err:     { background:'rgba(255,107,107,0.08)', border:'1px solid rgba(255,107,107,0.2)', color:'#ff6b6b', padding:'8px 16px', marginBottom:16, fontSize:13 },
  closeBtn:{ background:'rgba(255,107,107,0.1)', border:'1px solid rgba(255,107,107,0.3)', color:'#ff6b6b', padding:'6px 16px', cursor:'pointer', fontSize:12, fontFamily:'DM Mono,monospace', letterSpacing:'0.08em' },
  muted:   { color:'#3a4a6a', fontSize:13 },
  encTag:  { fontSize:11, background:'rgba(0,212,255,0.1)', color:'#00d4ff', padding:'3px 8px', fontFamily:'DM Mono,monospace' },
}

export default function RevelaApp() {
  const [token]   = useState(() => new URLSearchParams(location.search).get('token') ?? '')
  const [sess,  setSess]  = useState<any>(null)
  const [enc,   setEnc]   = useState<any>(null)
  const [intake,setIntake]= useState<any>(null)
  const [meds,  setMeds]  = useState<any[]>([])
  const [tab,   setTab]   = useState<'chart'|'prognote'|'opnote'|'postop'>('chart')
  const [saving,setSaving]= useState(false)
  const [saved, setSaved] = useState(false)
  const [error, setError] = useState('')
  const [loading,setLoading]=useState(true)
  const [pn, setPn]= useState({ pain_level:0, progress_notes:'', plan:'' })
  const [on, setOn]= useState({ procedure_name:'', anesthesia:'General', duration_mins:60, complications:'None', notes:'' })
  const [po, setPo]= useState({ follow_up_date:'', wound_status:'', instructions:'' })

  useEffect(() => {
    if (!token) { setError('No session token. Access from PatientTracForge scheduling.'); setLoading(false); return }
    bridge('get_encounter', { token }).then(d => {
      if (d.error) { setError(d.error); setLoading(false); return }
      setSess(d.session); setEnc(d.encounter); setIntake(d.intake); setMeds(d.medications ?? [])
      if (d.encounter?.progress_notes) setPn(p => ({...p, progress_notes: d.encounter.progress_notes, pain_level: d.encounter.pain_level ?? 0}))
      if (d.encounter?.procedure_name) setOn(o => ({...o, procedure_name: d.encounter.procedure_name, notes: d.encounter.notes ?? ''}))
      setLoading(false)
    }).catch(e => { setError(String(e)); setLoading(false) })
  }, [token])

  async function save(type: string, data: any) {
    setSaving(true); setSaved(false); setError('')
    const r = await bridge('save_note', { token, note_type: type, note_data: data })
    setSaving(false)
    if (r.success) { setSaved(true); setTimeout(() => setSaved(false), 3000) }
    else setError(r.error ?? 'Save failed')
  }

  async function closeEnc() {
    if (!confirm('Close encounter? Billing AI will auto-run.')) return
    await bridge('callback', { token, event_type:'encounter_closed', payload:{closed_by:'revela'} })
    alert('Encounter closed. Return to scheduling.'); window.close()
  }

  if (loading) return <div style={s.shell}><div style={s.center}><div style={s.spin}/><div style={{...s.muted,marginTop:12}}>Loading encounter...</div></div></div>
  if (error && !sess) return <div style={s.shell}><div style={s.errBox}><div style={{fontSize:18,fontWeight:600,color:'#ff6b6b',marginBottom:8}}>⚠ Access Error</div><div style={s.cbdy}>{error}</div><div style={{...s.muted,marginTop:12,fontSize:12}}>Access this app from PatientTracForge → Check-In → Route to Revela.</div></div></div>

  const pat = sess?.patient; const prov = sess?.provider

  return (
    <div style={s.shell}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}} * { box-sizing:border-box } select,input,textarea{color:#e8eaf0 !important}`}</style>
      <div style={s.hdr}>
        <div style={{display:'flex',alignItems:'center',gap:16}}>
          <div style={s.badge}>✦ REVELA</div>
          <div>
            <div style={s.pname}>{pat?.last_name}, {pat?.first_name}</div>
            <div style={s.meta}>DOB: {pat?.dob ? new Date(pat.dob).toLocaleDateString() : '—'} · {pat?.gender} · #{sess?.patient_id}</div>
          </div>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:10}}>
          <div style={{fontSize:11,color:'#8a9bc0',fontFamily:'DM Mono,monospace'}}>Dr. {prov?.last_name} · {prov?.specialty}</div>
          <div style={s.encTag}>Enc #{sess?.encounter_id}</div>
          <div style={{width:8,height:8,borderRadius:'50%',background: enc?.status==='open'?'#4ade80':'#f59e0b'}}/>
          <button onClick={closeEnc} style={s.closeBtn}>Close Encounter</button>
        </div>
      </div>

      <div style={s.tabs}>
        {([['chart','Patient Chart'],['prognote','Progress Note'],['opnote','Operative Note'],['postop','Post-Op Plan']] as const).map(([k,l]) => (
          <button key={k} onClick={() => setTab(k)} style={{...s.tab,...(tab===k?s.tabOn:{})}}>{l}</button>
        ))}
      </div>

      <div style={s.body}>
        {error && <div style={s.err}>{error}</div>}
        {saved && <div style={s.ok}>✓ Saved successfully — synced to PatientTracForge</div>}

        {tab==='chart' && (
          <div style={s.grid}>
            <div style={s.card}><div style={s.cttl}>Chief Complaint</div><div style={s.cbdy}>{enc?.chief_complaint || intake?.chief_complaint || '—'}</div></div>
            {intake && <div style={s.card}><div style={s.cttl}>AI Intake Summary</div><div style={s.cbdy}>{intake.ai_summary||'—'}</div>{intake.triage_level&&<div style={{marginTop:8,display:'inline-block',padding:'3px 8px',fontSize:11,background:'rgba(251,191,36,0.15)',color:'#fbbf24'}}>Triage: {intake.triage_level}</div>}</div>}
            <div style={s.card}><div style={s.cttl}>Allergies</div><div style={s.cbdy}>{pat?.allergies||'NKDA'}</div></div>
            <div style={s.card}><div style={s.cttl}>Current Medications ({meds.length})</div>{meds.length===0?<div style={s.muted}>No current medications</div>:meds.map((m:any,i:number)=><div key={i} style={{fontSize:13,color:'#8a9bc0',padding:'4px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}}><strong style={{color:'#e8eaf0'}}>{m.medication_name}</strong> {m.dosage} {m.route} — {m.frequency}</div>)}</div>
            <div style={s.card}><div style={s.cttl}>Contact</div><div style={s.cbdy}>{pat?.phone}<br/>{pat?.email}</div></div>
            <div style={s.card}><div style={s.cttl}>Provider / NPI</div><div style={s.cbdy}>Dr. {prov?.first_name} {prov?.last_name}, {prov?.credential}<br/><span style={{fontFamily:'DM Mono,monospace',fontSize:12}}>NPI {prov?.npi}</span></div></div>
          </div>
        )}

        {tab==='prognote' && (
          <div style={s.form}>
            <div style={s.fg}>
              <label style={s.lbl}>Pain Level (0–10)</label>
              <div style={{display:'flex',alignItems:'center',gap:12}}>
                <input type="range" min={0} max={10} value={pn.pain_level} onChange={e=>setPn(p=>({...p,pain_level:+e.target.value}))} style={{flex:1}}/>
                <span style={{background:'rgba(201,169,110,0.15)',color:'#c9a96e',padding:'4px 14px',fontFamily:'DM Mono,monospace',fontSize:14,minWidth:40,textAlign:'center'}}>{pn.pain_level}</span>
              </div>
            </div>
            <div style={s.fg}><label style={s.lbl}>Progress Notes *</label><textarea rows={8} value={pn.progress_notes} onChange={e=>setPn(p=>({...p,progress_notes:e.target.value}))} placeholder="Subjective findings, wound status, healing progress, objective assessment..." style={s.ta}/></div>
            <div style={s.fg}><label style={s.lbl}>Plan / Next Steps</label><textarea rows={4} value={pn.plan} onChange={e=>setPn(p=>({...p,plan:e.target.value}))} placeholder="Follow-up plan, prescription changes, referrals..." style={s.ta}/></div>
            <button disabled={saving||!pn.progress_notes} onClick={()=>save('surgical_prognote',pn)} style={saving||!pn.progress_notes?s.btnOff:s.btn}>{saving?'Saving...':'Save Progress Note → Scheduling'}</button>
          </div>
        )}

        {tab==='opnote' && (
          <div style={s.form}>
            <div style={s.row}>
              <div style={s.fg}><label style={s.lbl}>Procedure Name *</label><input value={on.procedure_name} onChange={e=>setOn(o=>({...o,procedure_name:e.target.value}))} placeholder="e.g. Augmentation mammaplasty" style={s.inp}/></div>
              <div style={s.fg}><label style={s.lbl}>Anesthesia</label><select value={on.anesthesia} onChange={e=>setOn(o=>({...o,anesthesia:e.target.value}))} style={s.inp}><option>General</option><option>Regional</option><option>Local + Sedation</option><option>Local</option></select></div>
            </div>
            <div style={s.row}>
              <div style={s.fg}><label style={s.lbl}>Duration (minutes)</label><input type="number" value={on.duration_mins} onChange={e=>setOn(o=>({...o,duration_mins:+e.target.value}))} style={s.inp}/></div>
              <div style={s.fg}><label style={s.lbl}>Complications</label><input value={on.complications} onChange={e=>setOn(o=>({...o,complications:e.target.value}))} placeholder="None / describe" style={s.inp}/></div>
            </div>
            <div style={s.fg}><label style={s.lbl}>Operative Notes *</label><textarea rows={10} value={on.notes} onChange={e=>setOn(o=>({...o,notes:e.target.value}))} placeholder="Detailed intraoperative findings, technique, closure method, implant specifics if applicable..." style={s.ta}/></div>
            <button disabled={saving||!on.procedure_name||!on.notes} onClick={()=>save('operative_notes',on)} style={saving||!on.procedure_name||!on.notes?s.btnOff:s.btn}>{saving?'Saving...':'Save Operative Note → Scheduling'}</button>
          </div>
        )}

        {tab==='postop' && (
          <div style={s.form}>
            <div style={s.row}>
              <div style={s.fg}><label style={s.lbl}>Follow-up Date</label><input type="date" value={po.follow_up_date} onChange={e=>setPo(p=>({...p,follow_up_date:e.target.value}))} style={s.inp}/></div>
              <div style={s.fg}><label style={s.lbl}>Wound / Healing Status</label><select value={po.wound_status} onChange={e=>setPo(p=>({...p,wound_status:e.target.value}))} style={s.inp}><option value="">Select...</option><option>Healing well — no concerns</option><option>Minor swelling — expected</option><option>Signs of infection — monitor</option><option>Requires intervention</option></select></div>
            </div>
            <div style={s.fg}><label style={s.lbl}>Patient Instructions *</label><textarea rows={8} value={po.instructions} onChange={e=>setPo(p=>({...p,instructions:e.target.value}))} placeholder="Activity restrictions, wound care, medications, warning signs, when to return to ER..." style={s.ta}/></div>
            <button disabled={saving||!po.instructions} onClick={()=>save('operative_notes',{...on,postop_plan:po})} style={saving||!po.instructions?s.btnOff:s.btn}>{saving?'Saving...':'Save Post-Op Plan → Scheduling'}</button>
          </div>
        )}
      </div>
    </div>
  )
}
