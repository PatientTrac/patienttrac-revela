import { useState } from 'react'
import { supabase } from '../lib/supabase'

interface Props {
  orgId: string
  providerId: string
  onPatientCreated: (patient: any) => void
  onClose: () => void
}

export default function NewPatientRevela({ orgId, providerId, onPatientCreated, onClose }: Props) {
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')
  const [form, setForm] = useState({
    first_name: '', last_name: '', date_of_birth: '', sex: 'F',
    phone_mobile: '', email: '', address_line1: '', city: '', state: '', zip: '',
    emergency_contact_name: '', emergency_contact_phone: '',
    chief_complaint: '', allergies: '', current_medications: '',
    insurance_provider: '', insurance_id: '', is_self_pay: false,
  })

  const set = (k: string, v: any) => setForm(p => ({ ...p, [k]: v }))

  const handleSave = async () => {
    setError('')
    if (!form.first_name || !form.last_name || !form.date_of_birth) {
      setError('First name, last name, and date of birth are required'); return
    }
    setSaving(true)
    try {
      const { data, error: insertErr } = await supabase
        .schema('cr').from('patient')
        .insert({
          org_id: orgId,
          first_name: form.first_name,
          last_name: form.last_name,
          date_of_birth: form.date_of_birth,
          sex: form.sex,
          phone_mobile: form.phone_mobile || null,
          email: form.email || null,
          address_line1: form.address_line1 || null,
          city: form.city || null,
          state: form.state || null,
          zip: form.zip || null,
          emergency_contact_name: form.emergency_contact_name || null,
          emergency_contact_phone: form.emergency_contact_phone || null,
          chief_complaint: form.chief_complaint || null,
          allergies: form.allergies || null,
          current_medications: form.current_medications || null,
          primary_provider_id: providerId || null,
          is_active: true,
          created_source: 'revela',
        })
        .select().single()
      if (insertErr) throw new Error(insertErr.message)
      onPatientCreated(data)
    } catch (e: any) { setError(e.message) }
    finally { setSaving(false) }
  }

  const S = {
    overlay: { position:'fixed' as const, inset:0, background:'rgba(0,0,0,0.75)', display:'flex', alignItems:'center', justifyContent:'center', zIndex:1000, padding:20 },
    modal: { background:'#0a1628', border:'1px solid rgba(201,169,110,0.2)', borderRadius:14, width:'100%', maxWidth:700, maxHeight:'90vh', overflow:'auto', boxShadow:'0 32px 80px rgba(0,0,0,0.6)' },
    hdr: { background:'#060e1c', padding:'20px 28px', borderBottom:'1px solid rgba(201,169,110,0.12)', display:'flex', justifyContent:'space-between', alignItems:'center', position:'sticky' as const, top:0 },
    body: { padding:'24px 28px' },
    section: { marginBottom:28 },
    sectionTitle: { fontSize:11, color:'rgba(201,169,110,0.7)', letterSpacing:'1.5px', textTransform:'uppercase' as const, marginBottom:16, paddingBottom:8, borderBottom:'1px solid rgba(255,255,255,0.06)' },
    grid2: { display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 } as React.CSSProperties,
    grid3: { display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:16 } as React.CSSProperties,
    lbl: { display:'block', fontSize:10, color:'rgba(201,169,110,0.6)', letterSpacing:'1px', textTransform:'uppercase' as const, marginBottom:5 },
    inp: { width:'100%', background:'rgba(255,255,255,0.05)', border:'1px solid rgba(255,255,255,0.1)', borderRadius:8, padding:'10px 13px', color:'#fff', fontSize:13, outline:'none', boxSizing:'border-box' as const },
    sel: { width:'100%', background:'#0d1f35', border:'1px solid rgba(255,255,255,0.1)', borderRadius:8, padding:'10px 13px', color:'#fff', fontSize:13, outline:'none' },
    ta: { width:'100%', background:'rgba(255,255,255,0.05)', border:'1px solid rgba(255,255,255,0.1)', borderRadius:8, padding:'10px 13px', color:'#fff', fontSize:13, outline:'none', boxSizing:'border-box' as const, resize:'vertical' as const, minHeight:70 },
    err: { background:'rgba(239,68,68,0.1)', border:'1px solid rgba(239,68,68,0.3)', borderRadius:8, padding:'10px 14px', color:'#f87171', fontSize:13, marginBottom:16 },
  }

  const Field = ({ label, k, type='text', placeholder='' }: { label:string, k:string, type?:string, placeholder?:string }) => (
    <div>
      <label style={S.lbl}>{label}</label>
      <input type={type} value={(form as any)[k]} onChange={e=>set(k,e.target.value)} placeholder={placeholder} style={S.inp}/>
    </div>
  )

  return (
    <div style={S.overlay}>
      <div style={S.modal}>
        <div style={S.hdr}>
          <div style={{ display:'flex', alignItems:'center', gap:10 }}>
            <svg width="20" height="20" viewBox="0 0 32 32" fill="none"><polygon points="16,2 30,16 16,30 2,16" fill="none" stroke="#c9a96e" strokeWidth="1.5"/><polygon points="16,11 21,16 16,21 11,16" fill="#c9a96e"/></svg>
            <span style={{ color:'#c9a96e', fontSize:16, fontWeight:600 }}>New Patient — Revela</span>
          </div>
          <button onClick={onClose} style={{ background:'none', border:'none', color:'rgba(255,255,255,0.4)', fontSize:22, cursor:'pointer' }}>×</button>
        </div>
        <div style={S.body}>

          {/* Demographics */}
          <div style={S.section}>
            <div style={S.sectionTitle}>Demographics</div>
            <div style={{ ...S.grid3, marginBottom:16 }}>
              <Field label="First Name *" k="first_name" placeholder="Maria"/>
              <Field label="Last Name *" k="last_name" placeholder="Garcia"/>
              <Field label="Date of Birth *" k="date_of_birth" type="date"/>
            </div>
            <div style={{ ...S.grid2, marginBottom:0 }}>
              <div>
                <label style={S.lbl}>Sex</label>
                <select value={form.sex} onChange={e=>set('sex',e.target.value)} style={S.sel}>
                  <option value="F">Female</option>
                  <option value="M">Male</option>
                  <option value="O">Other</option>
                  <option value="U">Prefer not to say</option>
                </select>
              </div>
              <Field label="Email" k="email" type="email" placeholder="patient@email.com"/>
            </div>
          </div>

          {/* Contact */}
          <div style={S.section}>
            <div style={S.sectionTitle}>Contact</div>
            <div style={{ ...S.grid2, marginBottom:16 }}>
              <Field label="Mobile Phone" k="phone_mobile" placeholder="(305) 555-0100"/>
              <Field label="Address" k="address_line1" placeholder="123 Main St"/>
            </div>
            <div style={S.grid3}>
              <Field label="City" k="city" placeholder="Miami"/>
              <Field label="State" k="state" placeholder="FL"/>
              <Field label="ZIP" k="zip" placeholder="33101"/>
            </div>
          </div>

          {/* Emergency Contact */}
          <div style={S.section}>
            <div style={S.sectionTitle}>Emergency Contact</div>
            <div style={S.grid2}>
              <Field label="Name" k="emergency_contact_name" placeholder="John Garcia"/>
              <Field label="Phone" k="emergency_contact_phone" placeholder="(305) 555-0200"/>
            </div>
          </div>

          {/* Clinical */}
          <div style={S.section}>
            <div style={S.sectionTitle}>Clinical Notes</div>
            <div style={{ marginBottom:12 }}>
              <label style={S.lbl}>Chief Complaint / Reason for Visit</label>
              <textarea value={form.chief_complaint} onChange={e=>set('chief_complaint',e.target.value)} placeholder="Requesting breast augmentation consultation..." style={S.ta}/>
            </div>
            <div style={S.grid2}>
              <div>
                <label style={S.lbl}>Allergies</label>
                <textarea value={form.allergies} onChange={e=>set('allergies',e.target.value)} placeholder="NKDA or list allergies..." style={{ ...S.ta, minHeight:60 }}/>
              </div>
              <div>
                <label style={S.lbl}>Current Medications</label>
                <textarea value={form.current_medications} onChange={e=>set('current_medications',e.target.value)} placeholder="List medications or None..." style={{ ...S.ta, minHeight:60 }}/>
              </div>
            </div>
          </div>

          {/* Insurance */}
          <div style={S.section}>
            <div style={S.sectionTitle}>Insurance</div>
            <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:14 }}>
              <input type="checkbox" id="self_pay" checked={form.is_self_pay} onChange={e=>set('is_self_pay',e.target.checked)} style={{ width:16, height:16, accentColor:'#c9a96e' }}/>
              <label htmlFor="self_pay" style={{ fontSize:13, color:'rgba(255,255,255,0.7)', cursor:'pointer' }}>Self-pay / No insurance</label>
            </div>
            {!form.is_self_pay && (
              <div style={S.grid2}>
                <Field label="Insurance Provider" k="insurance_provider" placeholder="Aetna"/>
                <Field label="Member ID" k="insurance_id" placeholder="ABC123456"/>
              </div>
            )}
          </div>

          {error && <div style={S.err}>{error}</div>}

          <div style={{ display:'flex', justifyContent:'flex-end', gap:12, paddingTop:8, borderTop:'1px solid rgba(255,255,255,0.06)' }}>
            <button onClick={onClose} style={{ padding:'10px 24px', background:'transparent', border:'1px solid rgba(255,255,255,0.1)', borderRadius:8, color:'rgba(255,255,255,0.5)', fontSize:13, cursor:'pointer' }}>Cancel</button>
            <button onClick={handleSave} disabled={saving} style={{ padding:'10px 28px', background:'#c9a96e', border:'none', borderRadius:8, color:'#060e1c', fontSize:13, fontWeight:700, cursor: saving?'wait':'pointer', opacity: saving?0.7:1 }}>
              {saving ? 'Saving…' : 'Create Patient →'}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
