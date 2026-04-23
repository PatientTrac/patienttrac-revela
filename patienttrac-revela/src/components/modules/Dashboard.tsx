import { useNavigate } from 'react-router-dom'
import { useRevelaStore } from '../../store'

const s = {
  page: { maxWidth: 1100 },
  heading: { fontFamily: 'Georgia, serif', fontSize: 26, color: '#0a1628', letterSpacing: '0.02em', marginBottom: 4 } as React.CSSProperties,
  sub: { fontSize: 13, color: '#5a7a90', marginBottom: '1.5rem' } as React.CSSProperties,
  metrics: { display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: '1.5rem' } as React.CSSProperties,
  metric: { background: '#f5f3ef', borderRadius: 10, padding: '1rem', textAlign: 'center' } as React.CSSProperties,
  metricVal: { fontSize: 28, fontWeight: 500, color: '#0a1628', lineHeight: 1 } as React.CSSProperties,
  metricLabel: { fontSize: 12, color: '#5a7a90', marginTop: 4 } as React.CSSProperties,
  metricDelta: { fontSize: 11, color: '#0f766e', marginTop: 3 } as React.CSSProperties,
  twoCol: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1rem' } as React.CSSProperties,
  card: { background: '#fff', border: '0.5px solid #e5e0d8', borderRadius: 12, padding: '1rem 1.25rem' } as React.CSSProperties,
  cardTitle: { fontSize: 11, fontWeight: 500, letterSpacing: '0.1em', textTransform: 'uppercase' as const, color: '#5a7a90', marginBottom: 12 },
  flagAlert: { background: '#fff1f2', borderLeft: '3px solid #e11d48', padding: '8px 12px', fontSize: 12, color: '#9f1239', lineHeight: 1.5, marginBottom: 8 },
  flagWarn: { background: '#fffbeb', borderLeft: '3px solid #f59e0b', padding: '8px 12px', fontSize: 12, color: '#92400e', lineHeight: 1.5, marginBottom: 8 },
  flagInfo: { background: '#f0fdfa', borderLeft: '3px solid #0d9488', padding: '8px 12px', fontSize: 12, color: '#0f766e', lineHeight: 1.5, marginBottom: 8 },
  flagLabel: { fontSize: 10, fontWeight: 500, letterSpacing: '0.08em', textTransform: 'uppercase' as const, opacity: 0.7, marginBottom: 2 },
  patientRow: { display: 'flex', alignItems: 'center', gap: 10, padding: '8px 0', borderBottom: '0.5px solid #f0ece4' } as React.CSSProperties,
  avatar: { width: 32, height: 32, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 500, flexShrink: 0 } as React.CSSProperties,
  chip: { fontSize: 11, padding: '2px 8px', borderRadius: 20, fontWeight: 500 } as React.CSSProperties,
  btn: { padding: '7px 14px', borderRadius: 8, fontSize: 12, cursor: 'pointer', border: '0.5px solid #e5e0d8', background: '#fff', color: '#0a1628', fontFamily: 'inherit' } as React.CSSProperties,
  btnPrimary: { padding: '7px 14px', borderRadius: 8, fontSize: 12, cursor: 'pointer', border: 'none', background: '#0a1628', color: '#f0ece4', fontFamily: 'inherit' } as React.CSSProperties,
}

const PATIENTS = [
  { init: 'SL', name: 'Sarah L.', proc: 'Rhinoplasty consult', time: '9:00 AM', chipColor: '#f0fdfa', chipText: '#0f766e' },
  { init: 'MK', name: 'Marcus K.', proc: 'Post-op check — facelift', time: '10:30 AM', chipColor: '#fffbeb', chipText: '#92400e' },
  { init: 'JT', name: 'Jennifer T.', proc: 'Breast reconstruction', time: 'AI FLAG', chipColor: '#fff1f2', chipText: '#9f1239' },
  { init: 'RC', name: 'Robert C.', proc: 'Scar revision consult', time: '2:00 PM', chipColor: '#f0fdfa', chipText: '#0f766e' },
  { init: 'AM', name: 'Angela M.', proc: 'Abdominoplasty pre-op', time: '3:30 PM', chipColor: '#fffbeb', chipText: '#92400e' },
]

const AV_COLORS = [
  { bg: '#ccfbf1', color: '#0f766e' },
  { bg: '#fef3c7', color: '#92400e' },
  { bg: '#ffe4e6', color: '#9f1239' },
  { bg: '#dbeafe', color: '#1e3a8a' },
  { bg: '#ccfbf1', color: '#0f766e' },
]

export default function Dashboard() {
  const navigate = useNavigate()

  return (
    <div style={s.page}>
      <div style={s.heading}>Good morning, Dr. Williams</div>
      <div style={s.sub}>Thursday, April 23, 2026 · 3 OR suites active · 2 AI flags require review</div>

      <div style={s.metrics}>
        {[
          { val: '14', label: 'Patients today', delta: '+2 from yesterday' },
          { val: '3', label: 'Active ORs', delta: '2 completing today' },
          { val: '$284K', label: 'MTD revenue', delta: '+18% vs last month' },
          { val: '2', label: 'AI flags', delta: 'Requires review' },
        ].map((m) => (
          <div key={m.label} style={s.metric}>
            <div style={s.metricVal}>{m.val}</div>
            <div style={s.metricLabel}>{m.label}</div>
            <div style={s.metricDelta}>{m.delta}</div>
          </div>
        ))}
      </div>

      <div style={s.twoCol}>
        <div style={s.card}>
          <div style={s.cardTitle}>Today's patient queue</div>
          {PATIENTS.map((p, i) => (
            <div key={p.name} style={{ ...s.patientRow, borderBottom: i < PATIENTS.length - 1 ? '0.5px solid #f0ece4' : 'none' }}>
              <div style={{ ...s.avatar, background: AV_COLORS[i].bg, color: AV_COLORS[i].color }}>{p.init}</div>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 13, fontWeight: 500, color: '#0a1628' }}>{p.name}</div>
                <div style={{ fontSize: 11, color: '#5a7a90' }}>{p.proc}</div>
              </div>
              <span style={{ ...s.chip, background: p.chipColor, color: p.chipText }}>{p.time}</span>
            </div>
          ))}
        </div>

        <div style={s.card}>
          <div style={s.cardTitle}>AI clinical flags</div>
          <div style={s.flagAlert}>
            <div style={s.flagLabel}>High priority</div>
            Jennifer T. — BP 158/96. Anesthesia clearance required before OR scheduling.
          </div>
          <div style={s.flagWarn}>
            <div style={s.flagLabel}>Medication alert</div>
            Angela M. — On warfarin. Confirm INR &lt;1.5 before abdominoplasty.
          </div>
          <div style={s.flagInfo}>
            <div style={s.flagLabel}>Documentation</div>
            3 patient records have incomplete consent forms. AI pre-fill available.
          </div>
          <button style={{ ...s.btnPrimary, marginTop: 8 }} onClick={() => navigate('/intake')}>
            Review intake queue →
          </button>
        </div>
      </div>

      <div style={s.twoCol}>
        <div style={s.card}>
          <div style={s.cardTitle}>OR status — today</div>
          {[
            { suite: 'OR Suite 1', doctor: 'Dr. Chen', proc: 'Rhinoplasty + Chin augmentation', pct: 65, color: '#0d9488' },
            { suite: 'OR Suite 2', doctor: 'Dr. Patel', proc: 'TRAM flap reconstruction', pct: 35, color: '#d97706' },
            { suite: 'OR Suite 3', doctor: 'Dr. Williams', proc: 'Blepharoplasty bilateral', pct: 0, color: '#e5e0d8' },
          ].map((or) => (
            <div key={or.suite} style={{ marginBottom: 12 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 2 }}>
                <span style={{ fontSize: 13, fontWeight: 500, color: '#0a1628' }}>{or.suite}</span>
                <span style={{ fontSize: 11, color: '#5a7a90' }}>{or.doctor}</span>
              </div>
              <div style={{ fontSize: 12, color: '#5a7a90', marginBottom: 4 }}>{or.proc}</div>
              <div style={{ height: 5, background: '#f0ece4', borderRadius: 3, overflow: 'hidden' }}>
                <div style={{ height: '100%', width: `${or.pct}%`, background: or.color, borderRadius: 3 }} />
              </div>
            </div>
          ))}
          <button style={{ ...s.btn, marginTop: 4 }} onClick={() => navigate('/or-scheduling')}>
            View full OR schedule →
          </button>
        </div>

        <div style={s.card}>
          <div style={s.cardTitle}>Quick actions</div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {[
              { label: 'Start AI intake interview', path: '/intake', primary: true },
              { label: 'Book OR procedure', path: '/or-scheduling', primary: false },
              { label: 'Generate patient proposal', path: '/proposals', primary: false },
              { label: 'Check inventory levels', path: '/inventory', primary: false },
              { label: 'View financial reports', path: '/accounting', primary: false },
            ].map((action) => (
              <button
                key={action.label}
                style={action.primary ? s.btnPrimary : s.btn}
                onClick={() => navigate(action.path)}
              >
                {action.label}
              </button>
            ))}
          </div>
        </div>
      </div>
    </div>
  )
}
