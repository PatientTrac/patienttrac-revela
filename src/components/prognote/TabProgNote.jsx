import { useState, useEffect } from 'react'
import { prognoteApi } from '../../lib/supabase.js'
import { HudField, HudInput, HudTextarea, HudSelect, HudRow, SaveBar } from '../shared/HudField.jsx'
import styles from './TabProgNote.module.css'

const MSE_OPTS = {
  appearance: ['Well-groomed','Appropriate dress','Disheveled','Unusual'],
  behavior:   ['Calm','Cooperative','Agitated','Anxious','Guarded'],
  mood:       ['Euthymic','Depressed','Anxious','Euphoric','Irritable','Labile'],
  affect:     ['Full range','Constricted','Flat','Blunted','Inappropriate','Labile'],
}

export default function TabProgNote({ patient, session }) {
  const [note,   setNote]   = useState({})
  const [exam,   setExam]   = useState({})
  const [saving, setSaving] = useState(false)
  const [saved,  setSaved]  = useState(false)
  const [loading,setLoading]= useState(true)

  useEffect(() => {
    if (!patient?.patient_id || !session?.encounter_id) return
    Promise.all([
      prognoteApi.get(patient.patient_id, session.encounter_id),
      prognoteApi.getPhysicalExam(patient.patient_id, session.encounter_id),
    ]).then(([n, e]) => {
      if (n.data) setNote(n.data)
      if (e.data) setExam(e.data)
      setLoading(false)
    })
  }, [patient?.patient_id, session?.encounter_id])

  const setN = (f, v) => setNote(prev => ({ ...prev, [f]: v }))
  const setE = (f, v) => setExam(prev => ({ ...prev, [f]: v }))

  const handleSave = async () => {
    setSaving(true); setSaved(false)
    try {
      await prognoteApi.save({
        ...note,
        patient_id:   patient.patient_id,
        encounter_id: session.encounter_id,
        provider_id:  session.provider_id,
      })
      await prognoteApi.savePhysicalExam({
        ...exam,
        patient_id:   patient.patient_id,
        encounter_id: session.encounter_id,
      })
      setSaved(true)
      setTimeout(() => setSaved(false), 3000)
    } catch (err) {
      alert('Save failed: ' + err.message)
    } finally { setSaving(false) }
  }

  return (
    <div className={styles.page}>
      <div className={styles.sections}>

        {/* Subjective */}
        <section className={styles.section}>
          <div className={styles.sectionTitle}>S — Subjective</div>
          <div className={styles.fields}>
            <HudField label="Chief Complaint / HPI">
              <HudTextarea rows={4} value={note.subjective||''} onChange={e=>setN('subjective',e.target.value)}
                placeholder="History of present illness, patient-reported symptoms, concerns…"/>
            </HudField>
            <HudRow>
              <HudField label="Pain Level (0–10)" half>
                <HudInput type="number" min={0} max={10} value={note.pain_scale||''} onChange={e=>setN('pain_scale',e.target.value)}/>
              </HudField>
              <HudField label="Pain Location" half>
                <HudInput value={note.pain_location||''} onChange={e=>setN('pain_location',e.target.value)} placeholder="Location"/>
              </HudField>
            </HudRow>
          </div>
        </section>

        {/* Objective */}
        <section className={styles.section}>
          <div className={styles.sectionTitle}>O — Objective / Physical Exam</div>
          <div className={styles.fields}>
            <HudRow>
              <HudField label="General Appearance" half>
                <HudSelect options={['Appears stated age','Appears younger','Appears older','Healthy','Ill-appearing']}
                  placeholder="Select…" value={exam.general_appearance||''} onChange={e=>setE('general_appearance',e.target.value)}/>
              </HudField>
              <HudField label="Skin" half>
                <HudSelect options={['Normal','Dry','Oily','Scarring present','Previous surgery scars','Sun damage','Hyperpigmentation']}
                  placeholder="Select…" value={exam.skin||''} onChange={e=>setE('skin',e.target.value)}/>
              </HudField>
            </HudRow>
            <HudField label="Breast / Chest Exam">
              <HudTextarea rows={3} value={exam.breast_exam||''} onChange={e=>setE('breast_exam',e.target.value)}
                placeholder="Symmetry, ptosis grade, implant assessment, nipple-areolar complex, skin quality…"/>
            </HudField>
            <HudField label="Facial Exam">
              <HudTextarea rows={3} value={exam.facial_exam||''} onChange={e=>setE('facial_exam',e.target.value)}
                placeholder="Rhytids, volume loss, skin laxity, asymmetry, prior procedures…"/>
            </HudField>
            <HudField label="Body Exam">
              <HudTextarea rows={3} value={exam.body_exam||''} onChange={e=>setE('body_exam',e.target.value)}
                placeholder="Contour, adiposity distribution, skin excess, scars…"/>
            </HudField>
            <HudField label="Photographs Obtained">
              <HudSelect options={['Yes — standard series','Yes — custom series','No — patient declined','No — scheduled separately']}
                placeholder="Select…" value={exam.photos||''} onChange={e=>setE('photos',e.target.value)}/>
            </HudField>
          </div>
        </section>

        {/* Assessment */}
        <section className={styles.section}>
          <div className={styles.sectionTitle}>A — Assessment</div>
          <div className={styles.fields}>
            <HudField label="Clinical Assessment / Impression">
              <HudTextarea rows={4} value={note.assessment||''} onChange={e=>setN('assessment',e.target.value)}
                placeholder="Clinical impression, diagnosis, candidacy assessment for requested procedure(s)…"/>
            </HudField>
            <HudField label="Patient Goals / Expectations">
              <HudTextarea rows={3} value={note.patient_goals||''} onChange={e=>setN('patient_goals',e.target.value)}
                placeholder="Patient's stated goals, motivation, expectations discussed…"/>
            </HudField>
            <HudRow>
              <HudField label="Surgical Candidacy" half>
                <HudSelect options={['Excellent candidate','Good candidate','Acceptable candidate','Borderline — further evaluation','Not a candidate at this time']}
                  placeholder="Select…" value={note.candidacy||''} onChange={e=>setN('candidacy',e.target.value)}/>
              </HudField>
              <HudField label="ASA Classification" half>
                <HudSelect options={['ASA I — Healthy','ASA II — Mild systemic disease','ASA III — Severe systemic disease','ASA IV — Life-threatening']}
                  placeholder="Select…" value={note.asa_class||''} onChange={e=>setN('asa_class',e.target.value)}/>
              </HudField>
            </HudRow>
          </div>
        </section>

        {/* Plan */}
        <section className={styles.section}>
          <div className={styles.sectionTitle}>P — Plan</div>
          <div className={styles.fields}>
            <HudField label="Proposed Procedure(s)">
              <HudTextarea rows={3} value={note.plan_procedures||''} onChange={e=>setN('plan_procedures',e.target.value)}
                placeholder="e.g. Bilateral breast augmentation with silicone implants, submuscular placement…"/>
            </HudField>
            <HudField label="Pre-operative Orders / Labs">
              <HudTextarea rows={2} value={note.preop_orders||''} onChange={e=>setN('preop_orders',e.target.value)}
                placeholder="CBC, metabolic panel, EKG if >45y, mammogram if indicated…"/>
            </HudField>
            <HudField label="Informed Consent Discussion">
              <HudTextarea rows={3} value={note.consent_discussion||''} onChange={e=>setN('consent_discussion',e.target.value)}
                placeholder="Risks, benefits, alternatives discussed. Patient verbalizes understanding. Questions addressed…"/>
            </HudField>
            <HudRow>
              <HudField label="Follow-up" half>
                <HudSelect options={['1 week','2 weeks','1 month','3 months','6 months','As needed','Pre-operative appointment']}
                  placeholder="Select…" value={note.follow_up||''} onChange={e=>setN('follow_up',e.target.value)}/>
              </HudField>
              <HudField label="Referrals" half>
                <HudInput value={note.referrals||''} onChange={e=>setN('referrals',e.target.value)} placeholder="e.g. Anesthesia, Primary care clearance"/>
              </HudField>
            </HudRow>
            <HudField label="Additional Notes">
              <HudTextarea rows={3} value={note.additional_notes||''} onChange={e=>setN('additional_notes',e.target.value)}
                placeholder="Additional provider notes, patient communications, special considerations…"/>
            </HudField>
          </div>
        </section>
      </div>

      <SaveBar onSave={handleSave} saving={saving} saved={saved} label="Save Progress Note" />
    </div>
  )
}
