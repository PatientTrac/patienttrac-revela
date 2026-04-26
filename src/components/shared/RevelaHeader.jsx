import styles from './RevelaHeader.module.css'

export default function RevelaHeader({ ctx, session }) {
  const patient  = ctx?.patient
  const provider = ctx?.provider
  const encounter= ctx?.encounter

  const age = patient?.birth
    ? Math.floor((Date.now() - new Date(patient.birth)) / 3.156e10) : null

  return (
    <header className={styles.header}>
      <div className={styles.brand}>
        <svg width="28" height="28" viewBox="0 0 48 48" fill="none">
          <rect width="48" height="48" rx="10" fill="rgba(201,169,110,0.12)" stroke="rgba(201,169,110,0.3)" strokeWidth="1"/>
          <path d="M24 10L14 25C14 30.5 18.5 35 24 35C29.5 35 34 30.5 34 25L24 10Z" fill="#c9a96e" opacity="0.85"/>
          <path d="M20 24H28M24 20V28" stroke="white" strokeWidth="2" strokeLinecap="round"/>
        </svg>
        <div>
          <div className={styles.brandName}>REVELA</div>
          <div className={styles.brandSub}>Plastic Surgery EMR</div>
        </div>
      </div>

      {patient && (
        <div className={styles.patientBadge}>
          <div className={styles.patientAvatar}>
            {patient.first_name?.[0]}{patient.last_name?.[0]}
          </div>
          <div>
            <div className={styles.patientName}>
              {patient.first_name} {patient.last_name}
            </div>
            <div className={styles.patientMeta}>
              {age && `${age}y`} {patient.gender === 'F' ? '· Female' : patient.gender === 'M' ? '· Male' : ''} 
              {encounter?.appointment_type && ` · ${encounter.appointment_type}`}
            </div>
          </div>
          <div className={styles.encBadge}>
            ENC #{session?.encounter_id}
          </div>
        </div>
      )}

      <div className={styles.right}>
        {provider && (
          <div className={styles.providerTag}>
            {provider.first_name} {provider.last_name}, {provider.credentials}
          </div>
        )}
        <a href={import.meta.env.VITE_FORGE_URL || 'https://patienttracforge.com'}
          className={styles.returnBtn}>
          ← Scheduling
        </a>
      </div>
    </header>
  )
}
