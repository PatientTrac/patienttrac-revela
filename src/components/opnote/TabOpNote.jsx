import { useState, useEffect } from 'react'
import { opnoteApi } from '../../lib/supabase.js'
import { HudField, HudInput, HudTextarea, HudSelect, HudRow, SaveBar } from '../shared/HudField.jsx'
import styles from './TabOpNote.module.css'

const ANESTHESIA = ['General','MAC','Local','Spinal','Epidural','Twilight/Sedation']
const POSITIONS  = ['Supine','Prone','Lateral decubitus','Lithotomy','Semi-Fowler','Reverse Trendelenburg']

export default function TabOpNote({ patient, session }) {
  const [note,   setNote]   = useState({})
  const [saving, setSaving] = useState(false)
  const [saved,  setSaved]  = useState(false)
  const [codeSearch, setCodeSearch] = useState('')
  const [codeResults, setCodeResults] = useState([])

  useEffect(() => {
    if (!patient?.patient_id || !session?.encounter_id) return
    opnoteApi.get(patient.patient_id, session.encounter_id).then(({data}) => {
      if (data) setNote(data)
    })
  }, [patient?.patient_id, session?.encounter_id])

  const setN = (f, v) => setNote(prev => ({ ...prev, [f]: v }))

  const searchCodes = async (term) => {
    if (term.length < 2) { setCodeResults([]); return }
    const { data } = await opnoteApi.getSurgCodes(term)
    setCodeResults(data || [])
  }

  const addCode = (code) => {
    const existing = note.procedure_codes ? note.procedure_codes.split(', ') : []
    setN('procedure_codes', [...existing, code.proc_code].join(', '))
    setCodeResults([]); setCodeSearch('')
  }

  const handleSave = async () => {
    setSaving(true); setSaved(false)
    try {
      await opnoteApi.save({
        ...note,
        patient_id:   patient.patient_id,
        encounter_id: session.encounter_id,
        provider_id:  session.provider_id,
      })
      setSaved(true); setTimeout(() => setSaved(false), 3000)
    } catch (err) { alert('Save failed: ' + err.message) }
    finally { setSaving(false) }
  }

  return (
    <div className={styles.page}>
      <div className={styles.sections}>

        <section className={styles.section}>
          <div className={styles.sectionTitle}>📋 Operative Information</div>
          <div className={styles.fields}>
            <HudRow>
              <HudField label="Date of Surgery" half>
                <HudInput type="date" value={note.surgery_date||''} onChange={e=>setN('surgery_date',e.target.value)}/>
              </HudField>
              <HudField label="Facility" half>
                <HudInput value={note.facility||''} onChange={e=>setN('facility',e.target.value)} placeholder="Operating facility"/>
              </HudField>
            </HudRow>
            <HudRow>
              <HudField label="Surgeon" half>
                <HudInput value={note.surgeon||''} onChange={e=>setN('surgeon',e.target.value)} placeholder="Primary surgeon"/>
              </HudField>
              <HudField label="Assistant Surgeon" half>
                <HudInput value={note.assistant||''} onChange={e=>setN('assistant',e.target.value)} placeholder="Assistant (if any)"/>
              </HudField>
            </HudRow>
            <HudRow>
              <HudField label="Anesthesia Type" half>
                <HudSelect options={ANESTHESIA} placeholder="Select…" value={note.anesthesia||''} onChange={e=>setN('anesthesia',e.target.value)}/>
              </HudField>
              <HudField label="Patient Position" half>
                <HudSelect options={POSITIONS} placeholder="Select…" value={note.patient_position||''} onChange={e=>setN('patient_position',e.target.value)}/>
              </HudField>
            </HudRow>
            <HudRow>
              <HudField label="Estimated Blood Loss (mL)" half>
                <HudInput type="number" value={note.ebl||''} onChange={e=>setN('ebl',e.target.value)} placeholder="mL"/>
              </HudField>
              <HudField label="Operative Time (minutes)" half>
                <HudInput type="number" value={note.op_time||''} onChange={e=>setN('op_time',e.target.value)} placeholder="min"/>
              </HudField>
            </HudRow>
          </div>
        </section>

        <section className={styles.section}>
          <div className={styles.sectionTitle}>🔬 Procedure Codes</div>
          <div className={styles.fields}>
            <HudField label="Search Surgical Codes">
              <HudInput value={codeSearch} onChange={e=>{setCodeSearch(e.target.value);searchCodes(e.target.value)}}
                placeholder="Search CPT or procedure name…"/>
            </HudField>
            {codeResults.length > 0 && (
              <div className={styles.codeResults}>
                {codeResults.map(c => (
                  <button key={c.proc_code} className={styles.codeResult} onClick={()=>addCode(c)}>
                    <span className={styles.codeNum}>{c.proc_code}</span>
                    <span className={styles.codeName}>{c.proc_name}</span>
                    {c.rvu && <span className={styles.codeRvu}>{c.rvu} RVU</span>}
                  </button>
                ))}
              </div>
            )}
            <HudField label="Selected Procedure Codes">
              <HudInput value={note.procedure_codes||''} onChange={e=>setN('procedure_codes',e.target.value)}
                placeholder="e.g. 19340, 19342, 15830"/>
            </HudField>
          </div>
        </section>

        <section className={styles.section}>
          <div className={styles.sectionTitle}>📝 Operative Report</div>
          <div className={styles.fields}>
            <HudField label="Preoperative Diagnosis">
              <HudInput value={note.preop_diagnosis||''} onChange={e=>setN('preop_diagnosis',e.target.value)} placeholder="Preoperative diagnosis"/>
            </HudField>
            <HudField label="Postoperative Diagnosis">
              <HudInput value={note.postop_diagnosis||''} onChange={e=>setN('postop_diagnosis',e.target.value)} placeholder="Postoperative diagnosis"/>
            </HudField>
            <HudField label="Procedure Performed">
              <HudTextarea rows={3} value={note.procedure_performed||''} onChange={e=>setN('procedure_performed',e.target.value)}
                placeholder="Full procedure name(s) as they will appear on operative report…"/>
            </HudField>
            <HudField label="Operative Description">
              <HudTextarea rows={8} value={note.operative_description||''} onChange={e=>setN('operative_description',e.target.value)}
                placeholder="Detailed operative narrative: patient positioning, prep, incisions, dissection, implant placement, closure technique, wound management, specimen handling…"/>
            </HudField>
            <HudField label="Implants / Devices Used">
              <HudTextarea rows={2} value={note.implants||''} onChange={e=>setN('implants',e.target.value)}
                placeholder="Manufacturer, model, size, lot number, serial number…"/>
            </HudField>
            <HudField label="Complications">
              <HudSelect options={['None — procedure completed without complication','Intraoperative bleeding','Wound dehiscence','Adverse anesthesia reaction','Other — see notes']}
                placeholder="Select…" value={note.complications||''} onChange={e=>setN('complications',e.target.value)}/>
            </HudField>
            <HudField label="Disposition">
              <HudSelect options={['Discharged to home','Admitted for observation','Transferred to ICU','Admitted — planned overnight','Discharged to care facility']}
                placeholder="Select…" value={note.disposition||''} onChange={e=>setN('disposition',e.target.value)}/>
            </HudField>
          </div>
        </section>
      </div>

      <SaveBar onSave={handleSave} saving={saving} saved={saved} label="Save Operative Note" />
    </div>
  )
}
