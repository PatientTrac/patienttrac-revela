import { useState, useEffect, useRef, useCallback } from 'react'
import { supabase } from '../lib/supabase'

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL
const ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY

// ── Demo tour steps ───────────────────────────────────────────────────────────
const DEMO_STEPS = [
  {
    id: 'welcome',
    title: 'Welcome to Revela',
    text: 'Revela is PatientTrac\'s AI-powered plastic surgery EMR. I\'ll walk you through every feature. You can ask me anything at any time.',
    highlight: null,
    duration: 5000,
  },
  {
    id: 'login',
    title: 'Secure Login',
    text: 'Access is protected by email, password, and Google Authenticator two-factor authentication. Your account is HIPAA-compliant and fully encrypted.',
    highlight: 'login',
    duration: 5000,
  },
  {
    id: 'chart',
    title: 'Patient Chart',
    text: 'The Chart tab displays complete patient demographics, AI intake summary, current medications, and insurance information — all pre-populated from PatientTracForge.',
    highlight: 'chart',
    duration: 6000,
  },
  {
    id: 'prognote',
    title: 'Progress Notes',
    text: 'Document pain levels, clinical observations, and treatment plans. Notes save directly to the encounter and sync back to PatientTracForge in real time.',
    highlight: 'prognote',
    duration: 5000,
  },
  {
    id: 'opnote',
    title: 'Operative Notes',
    text: 'Record procedure name, anesthesia type, duration, complications, and full surgical narrative. Every operative note is timestamped and provider-signed.',
    highlight: 'opnote',
    duration: 5000,
  },
  {
    id: 'postop',
    title: 'Post-Op Plan',
    text: 'Set follow-up dates, document wound status, and provide patient instructions. Post-op plans are accessible to the patient via the Patient Portal.',
    highlight: 'postop',
    duration: 5000,
  },
  {
    id: 'drawing',
    title: 'Surgical Drawing Tool',
    text: 'Use the Fabric.js annotation tool to draw directly on Servier Medical Art body templates. Mark treatment zones, annotate measurements, and save diagrams per encounter.',
    highlight: 'drawing',
    duration: 6000,
  },
  {
    id: 'newpatient',
    title: 'New Patient Registration',
    text: 'Create new patients directly from Revela — no need to switch to PatientTracForge. Demographics, insurance, and chief complaint all captured in one form.',
    highlight: 'newpatient',
    duration: 5000,
  },
  {
    id: 'ai',
    title: 'AI Assistant',
    text: 'Ask me anything about the EMR by voice or text. I can answer clinical documentation questions, explain features, and route your suggestions directly to the PatientTrac team.',
    highlight: 'ai',
    duration: 5000,
  },
  {
    id: 'bridge',
    title: 'PatientTracForge Integration',
    text: 'Revela integrates seamlessly with PatientTracForge. Encounter IDs, patient records, AI intake data, and billing codes flow automatically between systems.',
    highlight: null,
    duration: 5000,
  },
]

interface Props {
  orgId: string
  providerId: string
  providerEmail: string
}

export default function RevelaAIAssistant({ orgId, providerId, providerEmail }: Props) {
  // Demo state
  const [demoActive, setDemoActive] = useState(false)
  const [demoStep, setDemoStep] = useState(0)
  const [demoProgress, setDemoProgress] = useState(0)
  const demoTimer = useRef<any>(null)
  const progressTimer = useRef<any>(null)

  // AI Assistant state
  const [assistantOpen, setAssistantOpen] = useState(false)
  const [messages, setMessages] = useState<{ role: 'user'|'assistant', content: string }[]>([
    { role: 'assistant', content: 'Hello! I\'m your Revela AI assistant. I can answer questions about the EMR, help with clinical documentation, or take your suggestions and route them to the PatientTrac team. How can I help you today?' }
  ])
  const [input, setInput] = useState('')
  const [aiLoading, setAiLoading] = useState(false)
  const [isRecording, setIsRecording] = useState(false)
  const [feedbackType, setFeedbackType] = useState<'question'|'suggestion'|'bug'|'feature_request'>('question')
  const messagesEndRef = useRef<HTMLDivElement>(null)
  const recognitionRef = useRef<any>(null)

  // Avatar animation state
  const [avatarPulse, setAvatarPulse] = useState(false)
  const [avatarSpeaking, setAvatarSpeaking] = useState(false)
  const speechRef = useRef<SpeechSynthesisUtterance | null>(null)

  // ── Text to Speech ────────────────────────────────────────────────────────
  const speak = useCallback((text: string) => {
    if (!window.speechSynthesis) return
    window.speechSynthesis.cancel()
    const utter = new SpeechSynthesisUtterance(text)
    utter.rate = 0.9
    utter.pitch = 1.0
    utter.volume = 0.8
    const voices = window.speechSynthesis.getVoices()
    const preferred = voices.find(v => v.name.includes('Google') && v.lang === 'en-US')
      || voices.find(v => v.lang === 'en-US')
    if (preferred) utter.voice = preferred
    utter.onstart = () => setAvatarSpeaking(true)
    utter.onend = () => setAvatarSpeaking(false)
    speechRef.current = utter
    window.speechSynthesis.speak(utter)
  }, [])

  // ── Demo Tour ─────────────────────────────────────────────────────────────
  const startDemo = () => {
    setDemoStep(0)
    setDemoProgress(0)
    setDemoActive(true)
    setAssistantOpen(false)
  }

  useEffect(() => {
    if (!demoActive) return
    const step = DEMO_STEPS[demoStep]
    speak(step.text)
    setAvatarPulse(true)
    setDemoProgress(0)

    const duration = step.duration
    const interval = 50
    let elapsed = 0
    progressTimer.current = setInterval(() => {
      elapsed += interval
      setDemoProgress(Math.min(100, (elapsed / duration) * 100))
      if (elapsed >= duration) {
        clearInterval(progressTimer.current)
        if (demoStep < DEMO_STEPS.length - 1) {
          setDemoStep(s => s + 1)
        } else {
          setDemoActive(false)
          setAvatarPulse(false)
          window.speechSynthesis?.cancel()
        }
      }
    }, interval)

    return () => {
      clearInterval(progressTimer.current)
    }
  }, [demoStep, demoActive])

  const stopDemo = () => {
    setDemoActive(false)
    setAvatarPulse(false)
    clearInterval(progressTimer.current)
    window.speechSynthesis?.cancel()
  }

  // ── Speech to Text ────────────────────────────────────────────────────────
  const startRecording = () => {
    const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition
    if (!SpeechRecognition) { alert('Speech recognition not supported in this browser. Use Chrome.'); return }
    const recognition = new SpeechRecognition()
    recognition.continuous = false
    recognition.interimResults = false
    recognition.lang = 'en-US'
    recognition.onresult = (e: any) => {
      const transcript = e.results[0][0].transcript
      setInput(prev => prev + (prev ? ' ' : '') + transcript)
    }
    recognition.onend = () => setIsRecording(false)
    recognition.onerror = () => setIsRecording(false)
    recognitionRef.current = recognition
    recognition.start()
    setIsRecording(true)
  }

  const stopRecording = () => {
    recognitionRef.current?.stop()
    setIsRecording(false)
  }

  // ── Send message to AI ────────────────────────────────────────────────────
  const sendMessage = async () => {
    if (!input.trim() || aiLoading) return
    const userMsg = input.trim()
    setInput('')
    setAiLoading(true)
    setMessages(prev => [...prev, { role: 'user', content: userMsg }])

    try {
      const response = await fetch('/.netlify/functions/ai-chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          messages: [...messages, { role: 'user', content: userMsg }],
          orgId,
          providerId,
        }),
      })

      const data = await response.json()
      const aiText = data.reply ?? 'I encountered an error. Please try again.'
      setMessages(prev => [...prev, { role: 'assistant', content: aiText }])
      speak(aiText)

      // Save to Supabase
      const isSuggestion = feedbackType !== 'question' ||
        userMsg.toLowerCase().includes('suggest') ||
        userMsg.toLowerCase().includes('feature') ||
        userMsg.toLowerCase().includes('would be nice') ||
        userMsg.toLowerCase().includes('add') ||
        userMsg.toLowerCase().includes('improve')

      await supabase.schema('cr').from('revela_ai_feedback').insert({
        org_id: orgId,
        provider_id: providerId || null,
        provider_email: providerEmail || null,
        feedback_type: isSuggestion ? feedbackType : 'question',
        input_method: isRecording ? 'speech' : 'text',
        content: userMsg,
        ai_response: aiText,
        status: 'new',
      })

      // Send email for suggestions
      if (isSuggestion) {
        await fetch(`${SUPABASE_URL}/functions/v1/send-notification`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', apikey: ANON_KEY },
          body: JSON.stringify({
            type: 'revela_feedback',
            to: 'wayne@patienttracforge.com',
            subject: `Revela ${feedbackType.replace('_', ' ')} from ${providerEmail || 'Provider'}`,
            content: userMsg,
            provider_email: providerEmail,
            ai_response: aiText,
          })
        }).catch(() => {}) // non-blocking
      }

    } catch (e: any) {
      setMessages(prev => [...prev, { role: 'assistant', content: 'Sorry, I encountered an error. Please try again.' }])
    } finally {
      setAiLoading(false)
    }
  }

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  // ── Styles ────────────────────────────────────────────────────────────────
  const S = {
    // HUD Avatar button
    avatarBtn: {
      position: 'fixed' as const, bottom: 24, right: 24, zIndex: 900,
      width: 64, height: 64, borderRadius: '50%',
      background: avatarSpeaking ? 'rgba(0,212,255,0.3)' : 'rgba(201,169,110,0.15)',
      border: `2px solid ${avatarSpeaking ? '#00d4ff' : '#c9a96e'}`,
      cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center',
      boxShadow: avatarSpeaking ? '0 0 24px rgba(0,212,255,0.4)' : '0 0 16px rgba(201,169,110,0.2)',
      transition: 'all 0.3s ease',
    },
    // Demo overlay
    demoOverlay: {
      position: 'fixed' as const, bottom: 100, right: 24, zIndex: 901,
      width: 380, background: '#0a1628', border: '1px solid rgba(0,212,255,0.3)',
      borderRadius: 14, padding: 20, boxShadow: '0 16px 48px rgba(0,0,0,0.5)',
    },
    // AI Assistant panel
    aiPanel: {
      position: 'fixed' as const, bottom: 100, right: 24, zIndex: 901,
      width: 400, height: 520, background: '#0a1628',
      border: '1px solid rgba(201,169,110,0.2)', borderRadius: 14,
      display: 'flex', flexDirection: 'column' as const,
      boxShadow: '0 16px 48px rgba(0,0,0,0.5)',
    },
  }

  return (
    <>
      {/* ── Demo Tour Overlay ── */}
      {demoActive && (
        <div style={S.demoOverlay}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <div style={{ width: 8, height: 8, borderRadius: '50%', background: '#00d4ff', boxShadow: '0 0 8px #00d4ff', animation: 'pulse 1s infinite' }}/>
              <span style={{ fontSize: 11, color: '#00d4ff', letterSpacing: '1px', textTransform: 'uppercase', fontFamily: 'sans-serif' }}>Live Demo</span>
            </div>
            <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.3)', fontFamily: 'sans-serif' }}>{demoStep + 1} / {DEMO_STEPS.length}</span>
          </div>

          {/* Avatar visual */}
          <div style={{ display: 'flex', gap: 14, marginBottom: 16 }}>
            <div style={{ flexShrink: 0 }}>
              <svg width="52" height="52" viewBox="0 0 52 52">
                <circle cx="26" cy="26" r="24" fill="rgba(6,14,28,0.9)" stroke={avatarSpeaking ? '#00d4ff' : '#c9a96e'} strokeWidth="1.5"/>
                <polygon points="26,8 40,26 26,44 12,26" fill="none" stroke="#c9a96e" strokeWidth="1"/>
                <polygon points="26,14 34,26 26,38 18,26" fill="#c9a96e" opacity="0.3"/>
                <polygon points="26,19 31,26 26,33 21,26" fill="#c9a96e"/>
                {avatarSpeaking && <>
                  <circle cx="26" cy="26" r="26" fill="none" stroke="#00d4ff" strokeWidth="0.5" opacity="0.5"/>
                  <circle cx="26" cy="26" r="28" fill="none" stroke="#00d4ff" strokeWidth="0.3" opacity="0.3"/>
                </>}
              </svg>
            </div>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: '#c9a96e', marginBottom: 6, fontFamily: 'sans-serif' }}>
                {DEMO_STEPS[demoStep].title}
              </div>
              <div style={{ fontSize: 13, color: 'rgba(255,255,255,0.7)', lineHeight: 1.6, fontFamily: 'sans-serif' }}>
                {DEMO_STEPS[demoStep].text}
              </div>
            </div>
          </div>

          {/* Progress bar */}
          <div style={{ height: 3, background: 'rgba(255,255,255,0.08)', borderRadius: 2, marginBottom: 14, overflow: 'hidden' }}>
            <div style={{ height: '100%', width: `${demoProgress}%`, background: 'linear-gradient(90deg, #c9a96e, #00d4ff)', borderRadius: 2, transition: 'width 0.05s linear' }}/>
          </div>

          {/* Step dots */}
          <div style={{ display: 'flex', justifyContent: 'center', gap: 5, marginBottom: 14 }}>
            {DEMO_STEPS.map((_, i) => (
              <div key={i} onClick={() => setDemoStep(i)}
                style={{ width: i === demoStep ? 16 : 6, height: 6, borderRadius: 3, background: i === demoStep ? '#c9a96e' : i < demoStep ? 'rgba(201,169,110,0.4)' : 'rgba(255,255,255,0.15)', cursor: 'pointer', transition: 'all 0.3s' }}/>
            ))}
          </div>

          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={stopDemo} style={{ flex: 1, padding: '8px', background: 'transparent', border: '1px solid rgba(255,255,255,0.1)', borderRadius: 8, color: 'rgba(255,255,255,0.4)', fontSize: 12, cursor: 'pointer', fontFamily: 'sans-serif' }}>
              Stop Demo
            </button>
            {demoStep < DEMO_STEPS.length - 1 && (
              <button onClick={() => { clearInterval(progressTimer.current); window.speechSynthesis?.cancel(); setDemoStep(s => s + 1) }}
                style={{ flex: 1, padding: '8px', background: 'rgba(201,169,110,0.15)', border: '1px solid rgba(201,169,110,0.3)', borderRadius: 8, color: '#c9a96e', fontSize: 12, cursor: 'pointer', fontFamily: 'sans-serif' }}>
                Next →
              </button>
            )}
          </div>
        </div>
      )}

      {/* ── AI Assistant Panel ── */}
      {assistantOpen && !demoActive && (
        <div style={S.aiPanel}>
          {/* Header */}
          <div style={{ padding: '14px 18px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexShrink: 0 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <svg width="22" height="22" viewBox="0 0 32 32" fill="none">
                <polygon points="16,2 30,16 16,30 2,16" fill="none" stroke="#c9a96e" strokeWidth="1.5"/>
                <polygon points="16,11 21,16 16,21 11,16" fill="#c9a96e"/>
              </svg>
              <span style={{ color: '#c9a96e', fontSize: 14, fontWeight: 600, fontFamily: 'sans-serif' }}>Revela AI Assistant</span>
            </div>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <button onClick={startDemo} style={{ fontSize: 11, padding: '4px 10px', background: 'rgba(0,212,255,0.1)', border: '1px solid rgba(0,212,255,0.3)', borderRadius: 6, color: '#00d4ff', cursor: 'pointer', fontFamily: 'sans-serif' }}>
                ▶ Demo
              </button>
              <button onClick={() => setAssistantOpen(false)} style={{ background: 'none', border: 'none', color: 'rgba(255,255,255,0.4)', fontSize: 18, cursor: 'pointer' }}>×</button>
            </div>
          </div>

          {/* Feedback type selector */}
          <div style={{ padding: '8px 18px', borderBottom: '1px solid rgba(255,255,255,0.04)', display: 'flex', gap: 6, flexShrink: 0 }}>
            {(['question', 'suggestion', 'feature_request', 'bug'] as const).map(t => (
              <button key={t} onClick={() => setFeedbackType(t)}
                style={{ fontSize: 10, padding: '3px 8px', borderRadius: 20, border: 'none', cursor: 'pointer', fontFamily: 'sans-serif', letterSpacing: '0.3px',
                  background: feedbackType === t ? 'rgba(201,169,110,0.2)' : 'rgba(255,255,255,0.05)',
                  color: feedbackType === t ? '#c9a96e' : 'rgba(255,255,255,0.35)',
                }}>
                {t.replace('_', ' ')}
              </button>
            ))}
          </div>

          {/* Messages */}
          <div style={{ flex: 1, overflow: 'auto', padding: '16px 18px', display: 'flex', flexDirection: 'column', gap: 12 }}>
            {messages.map((m, i) => (
              <div key={i} style={{ display: 'flex', justifyContent: m.role === 'user' ? 'flex-end' : 'flex-start' }}>
                <div style={{
                  maxWidth: '82%', padding: '10px 14px', borderRadius: m.role === 'user' ? '12px 12px 2px 12px' : '12px 12px 12px 2px',
                  background: m.role === 'user' ? 'rgba(201,169,110,0.15)' : 'rgba(255,255,255,0.06)',
                  border: m.role === 'user' ? '1px solid rgba(201,169,110,0.25)' : '1px solid rgba(255,255,255,0.08)',
                  fontSize: 13, color: m.role === 'user' ? '#f0e8d5' : 'rgba(255,255,255,0.8)',
                  lineHeight: 1.6, fontFamily: 'sans-serif',
                }}>
                  {m.content}
                </div>
              </div>
            ))}
            {aiLoading && (
              <div style={{ display: 'flex', gap: 4, padding: '10px 14px' }}>
                {[0,1,2].map(i => <div key={i} style={{ width: 6, height: 6, borderRadius: '50%', background: '#c9a96e', opacity: 0.6 }}/>)}
              </div>
            )}
            <div ref={messagesEndRef}/>
          </div>

          {/* Input */}
          <div style={{ padding: '12px 18px', borderTop: '1px solid rgba(255,255,255,0.06)', flexShrink: 0 }}>
            <div style={{ display: 'flex', gap: 8, alignItems: 'flex-end' }}>
              <textarea value={input} onChange={e => setInput(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendMessage() } }}
                placeholder="Ask a question or submit a suggestion…"
                rows={2}
                style={{ flex: 1, background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)', borderRadius: 10, padding: '9px 13px', color: '#fff', fontSize: 13, outline: 'none', resize: 'none', fontFamily: 'sans-serif', lineHeight: 1.5 }}
              />
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                <button onClick={isRecording ? stopRecording : startRecording}
                  style={{ width: 38, height: 38, borderRadius: 10, border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 16,
                    background: isRecording ? 'rgba(239,68,68,0.2)' : 'rgba(255,255,255,0.08)',
                    color: isRecording ? '#f87171' : 'rgba(255,255,255,0.5)',
                    boxShadow: isRecording ? '0 0 12px rgba(239,68,68,0.3)' : 'none',
                  }}>
                  🎤
                </button>
                <button onClick={sendMessage} disabled={!input.trim() || aiLoading}
                  style={{ width: 38, height: 38, borderRadius: 10, border: 'none', cursor: input.trim() ? 'pointer' : 'not-allowed', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 16,
                    background: input.trim() ? '#c9a96e' : 'rgba(255,255,255,0.05)',
                    color: input.trim() ? '#060e1c' : 'rgba(255,255,255,0.2)',
                  }}>
                  ↑
                </button>
              </div>
            </div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.2)', marginTop: 6, fontFamily: 'sans-serif' }}>
              Suggestions are sent to wayne@patienttracforge.com · Enter to send · Shift+Enter for new line
            </div>
          </div>
        </div>
      )}

      {/* ── Floating HUD Avatar Button ── */}
      {!demoActive && (
        <div style={{ position: 'fixed', bottom: 24, right: 24, zIndex: 902, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8 }}>
          {/* Demo button */}
          {!assistantOpen && (
            <button onClick={startDemo}
              style={{ padding: '6px 14px', background: 'rgba(0,212,255,0.1)', border: '1px solid rgba(0,212,255,0.3)', borderRadius: 20, color: '#00d4ff', fontSize: 11, cursor: 'pointer', fontFamily: 'sans-serif', letterSpacing: '0.5px', whiteSpace: 'nowrap' }}>
              ▶ Start Demo
            </button>
          )}
          {/* Main AI button */}
          <button onClick={() => { setAssistantOpen(o => !o); setDemoActive(false) }} style={S.avatarBtn}>
            <svg width="30" height="30" viewBox="0 0 32 32" fill="none">
              <polygon points="16,2 30,16 16,30 2,16" fill="none" stroke={avatarSpeaking ? '#00d4ff' : '#c9a96e'} strokeWidth="1.5"/>
              <polygon points="16,7 25,16 16,25 7,16" fill={avatarSpeaking ? '#00d4ff' : '#c9a96e'} opacity="0.3"/>
              <polygon points="16,11 21,16 16,21 11,16" fill={avatarSpeaking ? '#00d4ff' : '#c9a96e'}/>
            </svg>
          </button>
          <span style={{ fontSize: 10, color: 'rgba(255,255,255,0.3)', fontFamily: 'sans-serif', letterSpacing: '0.5px' }}>AI</span>
        </div>
      )}
    </>
  )
}
