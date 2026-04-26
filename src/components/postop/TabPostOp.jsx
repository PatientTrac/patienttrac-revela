import { useState, useEffect } from 'react'
import { postopApi } from '../../lib/supabase.js'
import { HudField, HudInput, HudTextarea, HudSelect, HudRow, SaveBar } from '../shared/HudField.jsx'
import styles from './TabPostOp.module.css'

const ACTIVITY = ['Bedrest','Light activity only','No lifting >10 lbs × 6 weeks','No strenuous activity × 4 weeks','Resume normal activity gradually','Cleared for full activity']
const DIET     = ['Regular diet','Clear liquids','Soft diet','High protein','Low sodium','No dietary restrictions']

export default function TabPostOp({ patient, session, encounter, onClose }) {
  const [plan,   setPlan]   = useState({})
  const [saving, setSaving] = useState(false)
  const [saved,  setSaved]  = useState(false)
  const [closing,setClosing]= useState(false)
  const [closed, setClosed] = useState(encounter?.encounter_status === 'closed')

  useEffect(() => {
    if (!patient?.patient_id || !session?.encounter_id) return
    postopApi.get(patient.patient_id, session.encounter_id).then(({data}) => { if (data) setPlan(data) })
  }, [patient?.patient_id, session?.encounter_id])

  const setP = (f, v) => setPlan(prev => ({ ...prev, [f]: v }))

  const handleSave = async () => {
    setSaving(true); setSaved(false)
    try {
      await postopApi.save({ ...plan, patient_id: patient.patient_id, encounter_id: session.encounter_id })
      setSaved(true); setTimeout(() => setSaved(false), 3000)
    } catch (err) { alert('Save failed: ' + err.message) }
    finally { setSaving(false) }
  }

  const handleClose = async () => {
    if (!window.confirm('Close this encounter? This will mark the encounter as complete.')) return
    setClosing(true)
    try {
      await handleSave()
      await postopApi.closeEncounter(session.encounter_id, {
        close_reason: 'Completed', close_diagnosis: plan.follow_up_diagnosis || encounter?.chief_complaint,
      })
      setClosed(true)
      onClose?.()
    } catch (err) { alert('Failed to close encounter: ' + err.message) }
    finally { setClosing(false) }
  }

  return (
    <div className={styles.page}>
      {closed && (
        <div className={styles.closedBanner}>
          ✅ Encounter Closed — {new Date().toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'})}
        </div>
      )}

      <div className={styles.sections}>
        <section className={styles.section}>
          <div className={styles.sectionTitle}>🩺 Post-Operative Instructions</div>
          <div className={styles.fields}>
            <HudField label="Wound Care Instructions">
              <HudTextarea rows={4} value={plan.wound_care||''} onChange={e=>setP('wound_care',e.target.value)}
                placeholder="Dressing changes, drain care, suture/staple removal, steri-strip instructions…"/>
            </HudField>
            <HudRow>
              <HudField label="Activity Restrictions" half>
                <HudSelect options={ACTIVITY} placeholder="Select…" value={plan.activity||''} onChange={e=>setP('activity',e.target.value)}/>
              </HudField>
              <HudField label="Diet" half>
                <HudSelect options={DIET} placeholder="Select…" value={plan.diet||''} onChange={e=>setP('diet',e.target.value)}/>
              </HudField>
            </HudRow>
            <HudField label="Shower / Bathing">
              <HudSelect options={['No bathing × 48 hours','No bathing × 1 week','Shower only — no submerging × 4 weeks','No restrictions']}
                placeholder="Select…" value={plan.bathing||''} onChange={e=>setP('bathing',e.target.value)}/>
            </HudField>
            <HudField label="Garment / Compression">
              <HudTextarea rows={2} value={plan.compression||''} onChange={e=>setP('compression',e.target.value)}
                placeholder="Compression garment instructions, surgical bra requirements, duration…"/>
            </HudField>
          </div>
        </section>

        <section className={styles.section}>
          <div className={styles.sectionTitle}>💊 Post-Op Medications</div>
          <div className={styles.fields}>
            <HudField label="Pain Management">
              <HudTextarea rows={2} value={plan.pain_meds||''} onChange={e=>setP('pain_meds',e.target.value)}
                placeholder="e.g. Hydrocodone/APAP 5/325mg q6h PRN × 5 days, Ibuprofen 600mg q8h × 7 days…"/>
            </HudField>
            <HudField label="Antibiotics">
              <HudTextarea rows={2} value={plan.antibiotics||''} onChange={e=>setP('antibiotics',e.target.value)}
                placeholder="e.g. Cephalexin 500mg QID × 7 days, Allergies verified before prescribing…"/>
            </HudField>
            <HudField label="Other Medications / Supplements">
              <HudTextarea rows={2} value={plan.other_meds||''} onChange={e=>setP('other_meds',e.target.value)}
                placeholder="Arnica Montana, bromelain, vitamins, antiemetics…"/>
            </HudField>
          </div>
        </section>

        <section className={styles.section}>
          <div className={styles.sectionTitle}>📅 Follow-Up Schedule</div>
          <div className={styles.fields}>
            <HudRow>
              <HudField label="First Post-Op Visit" half>
                <HudSelect options={['1–2 days','3–5 days','1 week','2 weeks','4 weeks','6 weeks']}
                  placeholder="Select…" value={plan.followup_1||''} onChange={e=>setP('followup_1',e.target.value)}/>
              </HudField>
              <HudField label="Subsequent Visits" half>
                <HudSelect options={['1 month','3 months','6 months','1 year','Per patient need','No further follow-up required']}
                  placeholder="Select…" value={plan.followup_2||''} onChange={e=>setP('followup_2',e.target.value)}/>
              </HudField>
            </HudRow>
            <HudField label="Return Precautions — Call or Go to ER If:">
              <HudTextarea rows={3} value={plan.return_precautions||''} onChange={e=>setP('return_precautions',e.target.value)}
                placeholder="Fever >101.5°F, excessive bleeding, wound dehiscence, signs of infection, severe pain not controlled by prescribed medications, difficulty breathing…"/>
            </HudField>
            <HudField label="Follow-Up Diagnosis / Status">
              <HudInput value={plan.follow_up_diagnosis||''} onChange={e=>setP('follow_up_diagnosis',e.target.value)}
                placeholder="Post-operative status, procedure performed…"/>
            </HudField>
            <HudField label="Patient Education Provided">
              <HudTextarea rows={2} value={plan.education||''} onChange={e=>setP('education',e.target.value)}
                placeholder="Written post-op instructions provided, questions answered, patient verbalizes understanding…"/>
            </HudField>
          </div>
        </section>

        {/* Close encounter */}
        {!closed && (
          <section className={styles.closeSection}>
            <div className={styles.sectionTitle}>🔒 Close Encounter</div>
            <p className={styles.closeNote}>
              Closing the encounter will mark it complete and lock clinical documentation. Ensure all notes, operative report, and post-op plan are finalized before closing.
            </p>
            <button className={styles.closeBtn} onClick={handleClose} disabled={closing || closed}>
              {closing ? '⏳ Closing…' : '🔒 Close & Complete Encounter'}
            </button>
          </section>
        )}
      </div>

      <SaveBar onSave={handleSave} saving={saving} saved={saved} label="Save Post-Op Plan" />
    </div>
  )
}
