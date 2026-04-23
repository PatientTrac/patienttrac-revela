import { Outlet, NavLink, useNavigate } from 'react-router-dom'
import { useRevelaStore } from '../../store'

const NAV = [
  { to: '/dashboard',     label: 'Dashboard',    icon: '▦' },
  { to: '/intake',        label: 'AI Intake',     icon: '◎' },
  { to: '/or-scheduling', label: 'OR Scheduling', icon: '⊞' },
  { to: '/proposals',     label: 'Proposals',     icon: '◈' },
  { to: '/inventory',     label: 'Inventory',     icon: '⊟' },
  { to: '/accounting',    label: 'Accounting',    icon: '◆' },
  { to: '/staff',         label: 'Staff',         icon: '◉' },
]

export default function Layout() {
  const { currentUser, flags, sidebarOpen, toggleSidebar, activeModule } = useRevelaStore()
  const criticalFlags = flags.filter((f) => !f.dismissed && f.severity === 'critical').length
  const navigate = useNavigate()

  return (
    <div style={{ display: 'flex', minHeight: '100vh', background: '#f5f3ef', fontFamily: 'Inter, system-ui, sans-serif' }}>

      {/* Sidebar */}
      <aside style={{
        width: sidebarOpen ? 220 : 64,
        background: '#0a1628',
        display: 'flex',
        flexDirection: 'column',
        transition: 'width 0.2s',
        flexShrink: 0,
      }}>
        {/* Logo */}
        <div style={{ padding: '1.5rem 1rem 1rem', borderBottom: '0.5px solid #1a3a5c', cursor: 'pointer' }}
          onClick={() => navigate('/dashboard')}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{
              width: 34, height: 34, background: '#142035', borderRadius: 8,
              display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0
            }}>
              <svg width="18" height="18" viewBox="0 0 44 44" fill="none">
                <circle cx="22" cy="22" r="14" stroke="#c9a96e" strokeWidth="1.5"/>
                <path d="M22 8 C28 14 30 18 30 22 C30 26 28 30 22 36 C16 30 14 26 14 22 C14 18 16 14 22 8Z"
                  fill="none" stroke="#f0ece4" strokeWidth="1.5"/>
                <circle cx="22" cy="22" r="3" fill="#c9a96e"/>
              </svg>
            </div>
            {sidebarOpen && (
              <div>
                <div style={{ fontSize: 9, letterSpacing: '0.2em', color: '#4a7fa8', textTransform: 'uppercase' }}>PatientTrac</div>
                <div style={{ fontFamily: 'Georgia, serif', fontSize: 18, color: '#f0ece4', letterSpacing: '0.04em', lineHeight: 1 }}>Revela</div>
              </div>
            )}
          </div>
        </div>

        {/* Nav */}
        <nav style={{ flex: 1, padding: '0.75rem 0' }}>
          {NAV.map((item) => (
            <NavLink key={item.to} to={item.to} style={({ isActive }) => ({
              display: 'flex', alignItems: 'center', gap: 10,
              padding: sidebarOpen ? '10px 1rem' : '10px',
              justifyContent: sidebarOpen ? 'flex-start' : 'center',
              color: isActive ? '#c9a96e' : '#4a7fa8',
              background: isActive ? '#142035' : 'transparent',
              borderLeft: isActive ? '2px solid #c9a96e' : '2px solid transparent',
              textDecoration: 'none', fontSize: 13, transition: 'all 0.15s',
            })}>
              <span style={{ fontSize: 16, flexShrink: 0 }}>{item.icon}</span>
              {sidebarOpen && <span style={{ letterSpacing: '0.03em' }}>{item.label}</span>}
            </NavLink>
          ))}
        </nav>

        {/* User */}
        <div style={{ padding: '1rem', borderTop: '0.5px solid #1a3a5c' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <div style={{
              width: 30, height: 30, borderRadius: '50%', background: '#1a3a5c',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              fontSize: 11, fontWeight: 500, color: '#c9a96e', flexShrink: 0
            }}>
              {currentUser?.name?.split(' ').map(n => n[0]).join('') || 'DW'}
            </div>
            {sidebarOpen && (
              <div>
                <div style={{ fontSize: 12, color: '#f0ece4', fontWeight: 500 }}>{currentUser?.name}</div>
                <div style={{ fontSize: 10, color: '#4a7fa8', textTransform: 'capitalize' }}>{currentUser?.role}</div>
              </div>
            )}
          </div>
        </div>
      </aside>

      {/* Main */}
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>

        {/* Topbar */}
        <header style={{
          background: '#fff', borderBottom: '0.5px solid #e5e0d8',
          padding: '0 1.5rem', height: 56,
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        }}>
          <button onClick={toggleSidebar} style={{
            background: 'none', border: 'none', cursor: 'pointer',
            fontSize: 18, color: '#0a1628', padding: '4px 8px',
          }}>☰</button>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            {criticalFlags > 0 && (
              <div style={{
                background: '#ffe4e6', color: '#9f1239',
                fontSize: 11, padding: '3px 10px', borderRadius: 20, fontWeight: 500,
              }}>
                {criticalFlags} critical flag{criticalFlags > 1 ? 's' : ''}
              </div>
            )}
            <div style={{ fontSize: 12, color: '#5a7a90' }}>
              {new Date().toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' })}
            </div>
          </div>
        </header>

        {/* Content */}
        <main style={{ flex: 1, overflow: 'auto', padding: '1.5rem' }}>
          <Outlet />
        </main>
      </div>
    </div>
  )
}
