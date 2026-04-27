// Revela Provider Login — HUD theme, matches PatientTracForge admin login.
// Email + password + Google Authenticator TOTP. No magic link, no Google OAuth.
// Uses the shared totp-setup edge function on Supabase mskormozwekezjmtcylv.

import { useState, useRef, useEffect } from 'react'
import { supabase } from './lib/supabase.js'
import * as OTPAuth from 'otpauth'

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || 'https://mskormozwekezjmtcylv.supabase.co'
const MFA_REQUIRED_ROLES = ['super_admin', 'admin', 'provider']

const C = {
  navy:     '#060e1c',
  navyDeep: '#0a1628',
  gold:     '#c9a96e',
  goldSoft: 'rgba(201,169,110,0.6)',
  goldDim:  'rgba(201,169,110,0.2)',
  goldLine: 'rgba(201,169,110,0.1)',
  text:     '#e2e8f0',
  textDim:  'rgba(226,232,240,0.6)',
  red:      '#fca5a5',
  redBg:    'rgba(239,68,68,0.1)',
  redLine:  'rgba(239,68,68,0.3)',
}

export default function RevelaLogin({ onAuthenticated }) {
  const [step, setStep] = useState('login') // login | totp | setup
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [showPw, setShowPw] = useState(false)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)

  // TOTP challenge
  const [code, setCode] = useState('')
  const [totpError, setTotpError] = useState(null)
  const [totpLoading, setTotpLoading] = useState(false)
  const [mfaSecret, setMfaSecret] = useState(null)
  const [authedUser, setAuthedUser] = useState(null)
  const codeRef = useRef(null)

  // First-time setup
  const [setupQrUrl, setSetupQrUrl] = useState(null)
  const [setupSecret, setSetupSecret] = useState(null)
  const [setupCode, setSetupCode] = useState('')
  const [setupError, setSetupError] = useState(null)
  const [setupLoading, setSetupLoading] = useState(false)

  async function callTotp(action, body) {
    const sess = (await supabase.auth.getSession()).data.session
    if (!sess) throw new Error('No session — please sign in again')
    const url = `${SUPABASE_URL}/functions/v1/totp-setup${action === 'setup' ? '?action=setup' : ''}`
    const res = await fetch(url, {
      method: action === 'setup' ? 'GET' : 'POST',
      headers: {
        Authorization: `Bearer ${sess.access_token}`,
        'Content-Type': 'application/json',
      },
      body: action !== 'setup' ? JSON.stringify({ action, ...(body || {}) }) : undefined,
    })
    const data = await res.json().catch(() => ({}))
    if (!res.ok) throw new Error(data.error || `${action} failed`)
    return data
  }

  const handleLogin = async (e) => {
    e.preventDefault()
    setLoading(true); setError(null)
    try {
      const { data, error: signInErr } = await supabase.auth.signInWithPassword({ email, password })
      if (signInErr) throw new Error(signInErr.message)
      const user = data?.user
      if (!user) throw new Error('Login failed')

      const { data: member, error: memberErr } = await supabase
        .from('org_members')
        .select('mfa_enabled, mfa_secret, role, is_active, org_id')
        .eq('user_id', user.id)
        .eq('is_active', true)
        .order('created_at', { ascending: false })
        .limit(1)
        .maybeSingle()

      if (memberErr) {
        await supabase.auth.signOut()
        throw new Error('Could not verify your account: ' + memberErr.message)
      }
      if (!member) {
        await supabase.auth.signOut()
        throw new Error('No active account found. Contact your administrator.')
      }
      if (!['provider', 'admin', 'super_admin'].includes(member.role)) {
        await supabase.auth.signOut()
        throw new Error('Provider account required to access Revela.')
      }

      setAuthedUser({ ...user, role: member.role, org_id: member.org_id })
      const requiresMfa = MFA_REQUIRED_ROLES.includes(member.role)

      if (requiresMfa && member.mfa_enabled && member.mfa_secret) {
        setMfaSecret(member.mfa_secret)
        setStep('totp')
      } else if (requiresMfa && !member.mfa_enabled) {
        try {
          const setupData = await callTotp('setup')
          setSetupQrUrl(setupData.qr_url)
          setSetupSecret(setupData.secret)
          setStep('setup')
        } catch (setupErr) {
          await supabase.auth.signOut()
          throw new Error('MFA setup failed: ' + setupErr.message)
        }
      } else {
        onAuthenticated?.({ user_id: user.id, org_id: member.org_id, role: member.role })
      }
    } catch (err) {
      setError(err.message || 'Invalid email or password')
    } finally { setLoading(false) }
  }

  const verifyTOTP = async (codeStr) => {
    setTotpLoading(true); setTotpError(null)
    try {
      const totp = new OTPAuth.TOTP({
        secret: OTPAuth.Secret.fromBase32(mfaSecret),
        algorithm: 'SHA1', digits: 6, period: 30,
      })
      const delta = totp.validate({ token: codeStr, window: 1 })
      if (delta === null) {
        setTotpError('Invalid code. Check Google Authenticator and try again.')
        setCode('')
        codeRef.current?.focus()
        return
      }
      callTotp('challenge', { token: codeStr }).catch(() => {})
      onAuthenticated?.({ user_id: authedUser.id, org_id: authedUser.org_id, role: authedUser.role })
    } catch (err) {
      setTotpError('Verification failed. Please try again.')
    } finally { setTotpLoading(false) }
  }

  const handleSetupVerify = async (e) => {
    e.preventDefault()
    if (setupCode.length !== 6) return
    setSetupLoading(true); setSetupError(null)
    try {
      await callTotp('verify', { token: setupCode })
      onAuthenticated?.({ user_id: authedUser.id, org_id: authedUser.org_id, role: authedUser.role })
    } catch (err) {
      setSetupError(err.message || 'Verification failed')
      if (err.message?.includes('No MFA setup in progress')) {
        try {
          await callTotp('disable')
          const fresh = await callTotp('setup')
          setSetupQrUrl(fresh.qr_url); setSetupSecret(fresh.secret); setSetupCode('')
        } catch {}
      }
    } finally { setSetupLoading(false) }
  }

  const handleResetSetup = async () => {
    setSetupLoading(true); setSetupError(null)
    try {
      await callTotp('disable')
      const fresh = await callTotp('setup')
      setSetupQrUrl(fresh.qr_url); setSetupSecret(fresh.secret); setSetupCode('')
    } catch (err) {
      setSetupError(err.message)
    } finally { setSetupLoading(false) }
  }

  useEffect(() => {
    if ((step === 'totp' || step === 'setup') && codeRef.current) {
      setTimeout(() => codeRef.current?.focus(), 100)
    }
  }, [step])

  const s = styles

  return (
    <div style={s.page}>
      <div style={s.card}>
        <div style={s.header}>
          <div style={s.logoLine}>
            <svg width="22" height="22" viewBox="0 0 32 32" fill="none">
              <polygon points="16,2 30,16 16,30 2,16" fill="none" stroke={C.gold} strokeWidth="1.5"/>
              <polygon points="16,7 25,16 16,25 7,16" fill={C.gold} opacity="0.3"/>
              <polygon points="16,11 21,16 16,21 11,16" fill={C.gold}/>
            </svg>
            <span style={s.logoText}>Revela</span>
          </div>
          <div style={s.tagline}>PLASTIC SURGERY EMR</div>
        </div>

        {step === 'login' && (
          <form onSubmit={handleLogin} style={s.form}>
            <div style={s.titleBlock}>
              <div style={s.title}>Provider Sign In</div>
              <div style={s.subtitle}>Email + password + Google Authenticator required</div>
            </div>

            <label style={s.label}>Email</label>
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)}
              required autoFocus style={s.input}
              placeholder="provider@practice.com" autoComplete="email" />

            <label style={s.label}>Password</label>
            <div style={{position: 'relative'}}>
              <input type={showPw ? 'text' : 'password'}
                value={password} onChange={(e) => setPassword(e.target.value)}
                required style={{...s.input, paddingRight: 40}}
                placeholder="••••••••" autoComplete="current-password" />
              <button type="button" onClick={() => setShowPw(!showPw)} style={s.eyeBtn}>
                {showPw ? '🔒' : '👁'}
              </button>
            </div>

            {error && <div style={s.error}>{error}</div>}

            <button type="submit" disabled={loading || !email || !password} style={s.btnPrimary}>
              {loading ? 'Signing in…' : 'Sign in securely →'}
            </button>

            <div style={s.hipaa}>
              HIPAA Compliant · AES-256 Encrypted · Access monitored · Licensed providers only
            </div>
          </form>
        )}

        {step === 'setup' && (
          <form onSubmit={handleSetupVerify} style={s.form}>
            <div style={s.setupBox}>
              <div style={s.setupTitle}>Set up Google Authenticator</div>
              <div style={s.setupSub}>Required for your role. One-time setup.</div>
            </div>
            <div style={s.qrLabel}>1. Install Google Authenticator · 2. Tap + → Scan QR</div>
            {setupQrUrl && <img src={setupQrUrl} alt="QR Code" style={s.qr}/>}
            {setupSecret && (
              <>
                <div style={s.manualEntry}>MANUAL ENTRY KEY</div>
                <div style={s.secret}>{setupSecret}</div>
              </>
            )}
            <label style={s.label}>3. Enter the 6-digit code to confirm</label>
            <input ref={codeRef} type="text" inputMode="numeric" pattern="[0-9]*" maxLength={6}
              value={setupCode}
              onChange={(e) => setSetupCode(e.target.value.replace(/\D/g, '').slice(0, 6))}
              required style={s.codeInput} placeholder="000000" />
            {setupError && <div style={s.error}>{setupError}</div>}
            <button type="submit" disabled={setupLoading || setupCode.length !== 6} style={s.btnPrimary}>
              {setupLoading ? 'Activating…' : 'Activate MFA & Sign In →'}
            </button>
            <div style={s.linkRow}>
              <button type="button" onClick={handleResetSetup} style={s.linkBtn}>↻ New QR code</button>
              <button type="button" onClick={async () => { await supabase.auth.signOut(); setStep('login'); }} style={s.linkBtn}>← Cancel</button>
            </div>
          </form>
        )}

        {step === 'totp' && (
          <form
            onSubmit={(e) => { e.preventDefault(); if (code.length === 6) verifyTOTP(code) }}
            style={s.form}
          >
            <div style={s.setupBox}>
              <div style={s.setupTitle}>Two-factor authentication</div>
              <div style={s.setupSub}>Open Google Authenticator and enter your 6-digit code</div>
            </div>
            <input ref={codeRef} type="text" inputMode="numeric" pattern="[0-9]*" maxLength={6}
              value={code}
              onChange={(e) => {
                const v = e.target.value.replace(/\D/g, '').slice(0, 6)
                setCode(v)
                setTotpError(null)
                if (v.length === 6) setTimeout(() => verifyTOTP(v), 50)
              }}
              required style={s.codeInput} placeholder="000000" autoFocus />
            {totpError && <div style={s.error}>{totpError}</div>}
            <button type="submit" disabled={totpLoading || code.length !== 6} style={s.btnPrimary}>
              {totpLoading ? 'Verifying…' : 'Verify →'}
            </button>
            <button type="button"
              onClick={async () => { await supabase.auth.signOut(); setStep('login'); setCode(''); setMfaSecret(null) }}
              style={s.linkBtn}>
              ← Back to login
            </button>
          </form>
        )}
      </div>
    </div>
  )
}

const styles = {
  page: {
    minHeight: '100vh',
    background: C.navy,
    backgroundImage: `radial-gradient(ellipse at top right, rgba(201,169,110,0.06) 0%, transparent 50%),
                      radial-gradient(ellipse at bottom left, rgba(0,212,255,0.04) 0%, transparent 50%)`,
    display: 'flex', alignItems: 'center', justifyContent: 'center',
    fontFamily: 'Rajdhani, "DM Sans", -apple-system, BlinkMacSystemFont, sans-serif',
    color: C.text, padding: '24px',
  },
  card: { width: '100%', maxWidth: 420, background: C.navyDeep,
    border: `1px solid ${C.goldDim}`, borderRadius: 12, overflow: 'hidden',
    boxShadow: '0 24px 80px rgba(6,14,28,0.6)' },
  header: { padding: '32px 28px 24px', textAlign: 'center', borderBottom: `1px solid ${C.goldLine}` },
  logoLine: { display: 'inline-flex', alignItems: 'center', gap: 8 },
  logoText: { fontSize: 22, fontWeight: 700, color: C.gold, letterSpacing: 0.5 },
  tagline: { fontSize: 11, color: C.goldSoft, letterSpacing: 2, marginTop: 6 },
  form: { padding: 28, display: 'flex', flexDirection: 'column', gap: 12 },
  titleBlock: { marginBottom: 4 },
  title: { fontSize: 20, fontWeight: 600, color: C.text },
  subtitle: { fontSize: 12, color: C.textDim, marginTop: 4 },
  label: { fontSize: 11, letterSpacing: 1.5, color: C.goldSoft, textTransform: 'uppercase', marginTop: 8 },
  input: { width: '100%', background: C.navy, border: `1px solid ${C.goldDim}`,
    borderRadius: 6, padding: '12px 14px', color: C.text, fontSize: 14,
    fontFamily: 'inherit', outline: 'none', boxSizing: 'border-box' },
  eyeBtn: { position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)',
    background: 'transparent', border: 'none', color: C.goldSoft,
    cursor: 'pointer', fontSize: 14, padding: 4 },
  codeInput: { width: '100%', background: C.navy, border: `1px solid rgba(201,169,110,0.3)`,
    borderRadius: 6, padding: '16px', color: C.gold, fontSize: 28,
    fontFamily: '"DM Mono", "Courier New", monospace', letterSpacing: 12,
    textAlign: 'center', fontWeight: 700, outline: 'none', boxSizing: 'border-box' },
  btnPrimary: { background: C.gold, color: C.navy, border: 'none', borderRadius: 6,
    padding: '14px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
    fontFamily: 'inherit', marginTop: 12, letterSpacing: 0.3 },
  linkBtn: { background: 'transparent', color: C.gold, border: 'none', fontSize: 12,
    cursor: 'pointer', padding: 8, fontFamily: 'inherit', textDecoration: 'underline' },
  linkRow: { display: 'flex', justifyContent: 'space-between', marginTop: 8 },
  error: { background: C.redBg, border: `1px solid ${C.redLine}`, color: C.red,
    padding: '10px 14px', borderRadius: 6, fontSize: 12 },
  hipaa: { fontSize: 10, color: C.goldSoft, textAlign: 'center', marginTop: 12,
    letterSpacing: 0.3, lineHeight: 1.5 },
  setupBox: { background: 'rgba(201,169,110,0.08)', border: `1px solid ${C.goldDim}`,
    borderRadius: 6, padding: '16px 20px' },
  setupTitle: { fontSize: 14, fontWeight: 600, color: C.gold, marginBottom: 4 },
  setupSub: { fontSize: 12, color: C.textDim },
  qrLabel: { fontSize: 11, color: C.goldSoft, textAlign: 'center', letterSpacing: 0.5 },
  qr: { display: 'block', margin: '8px auto', width: 200, height: 200,
    borderRadius: 8, background: '#fff', padding: 8, boxSizing: 'border-box' },
  manualEntry: { fontSize: 10, letterSpacing: 1.5, color: 'rgba(201,169,110,0.5)',
    textAlign: 'center', marginTop: 4 },
  secret: { fontFamily: '"DM Mono", monospace', fontSize: 13, color: C.gold,
    textAlign: 'center', padding: '10px', background: C.navy, borderRadius: 6,
    letterSpacing: 1, wordBreak: 'break-all' },
}
