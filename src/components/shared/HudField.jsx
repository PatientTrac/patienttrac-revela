import styles from './HudField.module.css'

export function HudField({ label, children, half, required }) {
  return (
    <div className={`${styles.field} ${half ? styles.half : ''}`}>
      <label className={styles.label}>{label}{required && <span style={{color:'#C0392B'}}> *</span>}</label>
      {children}
    </div>
  )
}

export function HudInput({ ...props }) {
  return <input className={`hud-input ${styles.input}`} {...props} />
}

export function HudTextarea({ rows = 4, ...props }) {
  return <textarea className={`hud-input ${styles.textarea}`} rows={rows} {...props} />
}

export function HudSelect({ options, placeholder, ...props }) {
  return (
    <select className={`hud-input ${styles.select}`} {...props}>
      {placeholder && <option value="">{placeholder}</option>}
      {options.map(o => (
        <option key={o.value??o} value={o.value??o}>{o.label??o}</option>
      ))}
    </select>
  )
}

export function HudRow({ children }) {
  return <div className={styles.row}>{children}</div>
}

export function SaveBar({ onSave, saving, saved, label = 'Save Note' }) {
  return (
    <div className={styles.saveBar}>
      {saved && <span className={styles.savedTag}>✓ Saved</span>}
      <button className={styles.saveBtn} onClick={onSave} disabled={saving}>
        {saving ? '⏳ Saving…' : label}
      </button>
    </div>
  )
}
