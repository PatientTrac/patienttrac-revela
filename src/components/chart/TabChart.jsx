import { useState, useEffect } from 'react'
import { chartApi } from '../../lib/supabase.js'
import styles from './TabChart.module.css'

function InfoBlock({ label, value, mono }) {
  return (
    <div className={styles.infoBlock}>
      <div className={styles.infoLabel}>{label}</div>
      <div className={`${styles.infoVal} ${mono ? styles.mono : ''}`}>{value || '—'}</div>
    </div>
  )
}

export default function TabChart({ patient, encounter, provider, session }) {
  const [meds,     setMeds]     = useState([])
  const [allergies,setAllergies]= useState([])
  const [diagnoses,setDiagnoses]= useState([])
  const [vitals,   setVitals]   = useState([])
  const [surgHist, setSurgHist] = useState([])
  const [loading,  setLoading]  = useState(true)

  useEffect(() => {
    if (!patient?.patient_id) return
    const pid = patient.patient_id
    const eid = session?.encounter_id
    Promise.all([
      chartApi.getMedications(pid),
      chartApi.getAllergies(pid),
      chartApi.getDiagnoses(pid, eid),
      chartApi.getVitals(pid),
      chartApi.getSurgicalHistory(pid),
    ]).then(([m,a,d,v,s]) => {
      setMeds(m.data||[]); setAllergies(a.data||[]);
      setDiagnoses(d.data||[]); setVitals(v.data||[]); setSurgHist(s.data||[])
      setLoading(false)
    })
  }, [patient?.patient_id])

  const age = patient?.birth ? Math.floor((Date.now()-new Date(patient.birth))/3.156e10) : null
  const latestVitals = vitals[0]

  if (loading && !patient) return (
    <div className={styles.loading}>Loading patient chart…</div>
  )

  return (
    <div className={styles.page}>
      {/* Demographics */}
      <section className={styles.section}>
        <div className={styles.sectionTitle}>👤 Patient Demographics</div>
        <div className={styles.infoGrid}>
          <InfoBlock label="Full Name" value={`${patient?.first_name} ${patient?.last_name}`} />
          <InfoBlock label="Date of Birth" value={patient?.birth ? new Date(patient.birth).toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'}) : null} />
          <InfoBlock label="Age" value={age ? `${age} years` : null} />
          <InfoBlock label="Gender" value={patient?.gender === 'F' ? 'Female' : patient?.gender === 'M' ? 'Male' : patient?.gender} />
          <InfoBlock label="Phone" value={patient?.cell_phone || patient?.phone} />
          <InfoBlock label="Email" value={patient?.email} />
          <InfoBlock label="Address" value={[patient?.address1, patient?.city, patient?.state, patient?.zipcode].filter(Boolean).join(', ')} />
          <InfoBlock label="Language" value={patient?.language} />
        </div>
      </section>

      {/* Encounter */}
      <section className={styles.section}>
        <div className={styles.sectionTitle}>📅 Current Encounter</div>
        <div className={styles.infoGrid}>
          <InfoBlock label="Encounter ID" value={`#${session?.encounter_id}`} mono />
          <InfoBlock label="Date" value={encounter?.encounter_date ? new Date(encounter.encounter_date).toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit'}) : null} />
          <InfoBlock label="Type" value={encounter?.appointment_type} />
          <InfoBlock label="Chief Complaint" value={encounter?.chief_complaint} />
          <InfoBlock label="Status" value={encounter?.encounter_status} />
          <InfoBlock label="Provider" value={provider ? `${provider.first_name} ${provider.last_name}, ${provider.credentials}` : null} />
        </div>
      </section>

      {/* Vitals */}
      {latestVitals && (
        <section className={styles.section}>
          <div className={styles.sectionTitle}>💗 Latest Vitals — {new Date(latestVitals.insert_date).toLocaleDateString()}</div>
          <div className={styles.vitalsGrid}>
            {[
              ['BP', latestVitals.bp_systolic && latestVitals.bp_diastolic ? `${latestVitals.bp_systolic}/${latestVitals.bp_diastolic}` : null, 'mmHg'],
              ['Heart Rate', latestVitals.pulse, 'bpm'],
              ['Temp', latestVitals.temp_fahrenheit, '°F'],
              ['Weight', latestVitals.weight_lbs, 'lbs'],
              ['Height', latestVitals.height_in ? `${Math.floor(latestVitals.height_in/12)}'${latestVitals.height_in%12}"` : null, ''],
              ['BMI', latestVitals.bmi, ''],
            ].filter(([,,]) => true).map(([label, val, unit]) => val ? (
              <div key={label} className={styles.vitalCard}>
                <div className={styles.vitalVal}>{val}<span className={styles.vitalUnit}>{unit}</span></div>
                <div className={styles.vitalLabel}>{label}</div>
              </div>
            ) : null)}
          </div>
        </section>
      )}

      <div className={styles.twoCol}>
        {/* Medications */}
        <section className={styles.section}>
          <div className={styles.sectionTitle}>💊 Current Medications
            <span className={styles.count}>{meds.length}</span>
          </div>
          {meds.length === 0
            ? <div className={styles.empty}>No active medications</div>
            : meds.map(m => (
              <div key={m.patient_med_id} className={styles.listItem}>
                <div className={styles.listPrimary}>{m.medication_name}</div>
                <div className={styles.listSub}>{[m.dosage, m.frequency, m.route].filter(Boolean).join(' · ')}</div>
              </div>
            ))
          }
        </section>

        {/* Allergies */}
        <section className={styles.section}>
          <div className={styles.sectionTitle}>⚠️ Allergies
            <span className={styles.count}>{allergies.length}</span>
          </div>
          {allergies.length === 0
            ? <div className={styles.empty}>NKDA — No known drug allergies</div>
            : allergies.map(a => (
              <div key={a.allergy_id} className={styles.listItem}>
                <div className={styles.listPrimary}>{a.allergen}</div>
                <div className={styles.listSub} style={{color:a.severity_level==='Severe'?'#C0392B':a.severity_level==='Moderate'?'#C97C2A':'#2D8A4E'}}>
                  {a.severity_level} {a.reaction_type ? `· ${a.reaction_type}` : ''}
                </div>
              </div>
            ))
          }
        </section>
      </div>

      {/* Surgical History */}
      <section className={styles.section}>
        <div className={styles.sectionTitle}>🔬 Surgical History
          <span className={styles.count}>{surgHist.length}</span>
        </div>
        {surgHist.length === 0
          ? <div className={styles.empty}>No prior surgical history on file</div>
          : surgHist.map(s => (
            <div key={s.past_surg_id} className={styles.listItem}>
              <div className={styles.listPrimary}>{s.surgery_name || s.procedure}</div>
              <div className={styles.listSub}>{[s.surgery_date, s.surgeon, s.facility].filter(Boolean).join(' · ')}</div>
            </div>
          ))
        }
      </section>

      {/* Diagnoses */}
      {diagnoses.length > 0 && (
        <section className={styles.section}>
          <div className={styles.sectionTitle}>🗂️ Diagnoses — This Encounter</div>
          {diagnoses.map(d => (
            <div key={d.diag_id} className={styles.listItem}>
              <div className={styles.listPrimary}>{d.diagnosis}</div>
              <div className={styles.listSub}>{d.icd10_code && `ICD-10: ${d.icd10_code}`}</div>
            </div>
          ))}
        </section>
      )}
    </div>
  )
}
