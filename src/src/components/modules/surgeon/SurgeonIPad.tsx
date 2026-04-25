import { useState, useRef, useEffect, useCallback } from 'react'

// ── Types ────────────────────────────────────────────────────────────────────
type BodyRegion = 'full-body' | 'face' | 'breast' | 'stomach' | 'buttocks' | 'chest' | 'nose' | 'eyes' | 'arms' | 'legs'
type DrawTool = 'incision' | 'measurement' | 'annotation' | 'concern' | 'erase'
type VoiceField = 'chiefComplaint' | 'findings' | 'assessment' | 'plan' | 'followUp'
type NoteStatus = 'idle' | 'listening' | 'processing' | 'saved'

interface DrawStroke {
  tool: DrawTool
  points: [number, number][]
  label?: string
}

interface SurgicalNote {
  chiefComplaint: string
  findings: string
  assessment: string
  plan: string
  followUp: string
}

interface Patient {
  id: number
  name: string
  mrn: string
  procedure: string
  dob: string
  flags: number
}

// ── Tool colors ───────────────────────────────────────────────────────────────
const TOOL_COLORS: Record<DrawTool, string> = {
  incision:    '#e11d48',
  measurement: '#c9a96e',
  annotation:  '#4a7fa8',
  concern:     '#f59e0b',
  erase:       '#ffffff',
}

const TOOL_LABELS: Record<DrawTool, string> = {
  incision:    'Incision',
  measurement: 'Measurement',
  annotation:  'Annotation',
  concern:     'Concern',
  erase:       'Erase',
}

// ── Anatomical SVG body diagrams ──────────────────────────────────────────────
const BODY_DIAGRAMS: Record<BodyRegion, string> = {
  'full-body': `<g opacity="0.18" stroke="#0a1628" stroke-width="1.2" fill="none">
    <ellipse cx="150" cy="40" rx="28" ry="32"/>
    <rect x="118" y="72" width="64" height="80" rx="8"/>
    <line x1="118" y1="78" x2="82" y2="155"/><line x1="182" y1="78" x2="218" y2="155"/>
    <ellipse cx="82" cy="160" rx="10" ry="8"/><ellipse cx="218" cy="160" rx="10" ry="8"/>
    <rect x="126" y="152" width="22" height="90" rx="6"/>
    <rect x="152" y="152" width="22" height="90" rx="6"/>
    <ellipse cx="137" cy="246" rx="12" ry="8"/><ellipse cx="163" cy="246" rx="12" ry="8"/>
  </g>`,
  face: `<g opacity="0.18" stroke="#0a1628" stroke-width="1.2" fill="none">
    <ellipse cx="150" cy="120" rx="80" ry="100"/>
    <ellipse cx="115" cy="100" rx="18" ry="12"/>
    <ellipse cx="185" cy="100" rx="18" ry="12"/>
    <path d="M120 150 Q150 170 180 150"/>
    <ellipse cx="150" cy="130" rx="10" ry="14"/>
    <line x1="70" y1="100" x2="90" y2="95"/>
    <line x1="230" y1="100" x2="210" y2="95"/>
  </g>`,
  breast: `<g opacity="0.18" stroke="#0a1628" stroke-width="1.2" fill="none">
    <path d="M40 160 Q80 60 150 80 Q220 60 260 160"/>
    <ellipse cx="100" cy="140" rx="55" ry="55"/>
    <ellipse cx="200" cy="140" rx="55" ry="55"/>
    <circle cx="100" cy="138" r="8"/><circle cx="200" cy="138" r="8"/>
    <line x1="150" y1="60" x2="150" y2="200" stroke-dasharray="4 3"/>
    <text x="65" y="200" font-size="11" fill="#0a1628" opacity="0.4">Left</text>
    <text x="195" y="200" font-size="11" fill="#0a1628" opacity="0.4">Right</text>
  </g>`,
  stomach: `<g opacity="0.18" stroke="#0a1628" stroke-width="1.2" fill="none">
    <rect x="70" y="40" width="160" height="200" rx="20"/>
    <line x1="150" y1="40" x2="150" y2="240" stroke-dasharray="4 3"/>
    <ellipse cx="150" cy="140" rx="12" ry="12"/>
    <line x1="70" y1="100" x2="230" y2="100" stroke-dasharray="4 3"/>
    <line x1="70" y1="170" x2="230" y2="170" stroke-dasharray="4 3"/>
    <text x="80" y="85" font-size="10" fill="#0a1628" opacity="0.4">Epigastric</text>
    <text x="80" y="158" font-size="10" fill="#0a1628" opacity="0.4">Periumbilical</text>
    <text x="80" y="225" font-size="10" fill="#0a1628" opacity="0.4">Hypogastric</text>
  </g>`,
  buttocks: `<g opacity="0.18" stroke="#0a1628" stroke-width="1.2" fill="none">
    <ellipse cx="100" cy="130" rx="75" ry="90"/>
    <ellipse cx="200" cy="130" rx="75" ry="90"/>
    <line x1="150" y1="40" x2="150" y2="220" stroke-dasharray="4 3"/>
    <text x="65" y="210" font-size="11" fill="#0a1628" opacity="0.4">Left</text>
    <text x="190" y="210" font-size="11" fill="#0a1628" opacity="0.4">Right</text>
  </g>`,
  chest: `<g opacity="0.18" stroke="#0a1628" stroke-width="1.2" fill="none">
    <rect x="70" y="40" width="160" height="160" rx="12"/>
    <path d="M70 80 Q150 60 230 80"/>
    <line x1="150" y1="40" x2="150" y2="200" stroke-dasharray="4 3"/>
    <rect x="80" y="90" width="60" height="50" rx="6"/>
    <rect x="160" y="90" width="60" height="50" rx="6"/>
    <path d="M80 200 Q150 220 220 200"/>
  </g>`,
  nose: `<g opacity="0.18" stroke="#0a1628" stroke-width="1.2" fill="none">
    <path d="M150 30 L120 160 Q150 180 180 160 Z"/>
    <ellipse cx="125" cy="155" rx="22" ry="15"/>
    <ellipse cx="175" cy="155" rx="22" ry="15"/>
    <line x1="150" y1="30" x2="150" y2="170" stroke-dasharray="3 3"/>
    <ellipse cx="150" cy="80" rx="8" ry="6"/>
    <text x="60" y="155" font-size="10" fill="#0a1628" opacity="0.4">L ala</text>
    <text x="195" y="155" font-size="10" fill="#0a1628" opacity="0.4">R ala</text>
  </g>`,
  eyes: `<g opacity="0.18" stroke="#0a1628" stroke-width="1.2" fill="none">
    <path d="M40 100 Q90 60 140 100 Q90 140 40 100Z"/>
    <path d="M160 100 Q210 60 260 100 Q210 140 160 100Z"/>
    <circle cx="90" cy="100" r="18"/><circle cx="210" cy="100" r="18"/>
    <path d="M35 80 Q90 50 145 80" stroke-dasharray="3 3"/>
    <path d="M155 80 Q210 50 265 80" stroke-dasharray="3 3"/>
    <text x="55" y="160" font-size="10" fill="#0a1628" opacity="0.4">Upper lid</text>
    <text x="175" y="160" font-size="10" fill="#0a1628" opacity="0.4">Upper lid</text>
  </g>`,
  arms: `<g opacity="0.18" stroke="#0a1628" stroke-width="1.2" fill="none">
    <rect x="40" y="40" width="50" height="200" rx="12"/>
    <rect x="210" y="40" width="50" height="200" rx="12"/>
    <text x="45" y="260" font-size="10" fill="#0a1628" opacity="0.4">Left arm</text>
    <text x="210" y="260" font-size="10" fill="#0a1628" opacity="0.4">Right arm</text>
    <line x1="65" y1="40" x2="65" y2="240" stroke-dasharray="3 3"/>
    <line x1="235" y1="40" x2="235" y2="240" stroke-dasharray="3 3"/>
  </g>`,
  legs: `<g opacity="0.18" stroke="#0a1628" stroke-width="1.2" fill="none">
    <rect x="60" y="20" width="60" height="220" rx="12"/>
    <rect x="180" y="20" width="60" height="220" rx="12"/>
    <text x="65" y="260" font-size="10" fill="#0a1628" opacity="0.4">Left leg</text>
    <text x="185" y="260" font-size="10" fill="#0a1628" opacity="0.4">Right leg</text>
    <line x1="90" y1="20" x2="90" y2="240" stroke-dasharray="3 3"/>
    <line x1="210" y1="20" x2="210" y2="240" stroke-dasharray="3 3"/>
  </g>`,
}

const REGION_LABELS: Record<BodyRegion, string> = {
  'full-body': 'Full body',
  face: 'Face',
  breast: 'Breast',
  stomach: 'Abdomen',
  buttocks: 'Buttocks',
  chest: 'Chest',
  nose: 'Nose',
  eyes: 'Eyes',
  arms: 'Arms',
  legs: 'Legs',
}

const MOCK_PATIENTS: Patient[] = [
  { id: 1, name: 'Jennifer T.', mrn: '00482761', procedure: 'Breast reconstruction', dob: '03/14/1979', flags: 2 },
  { id: 2, name: 'Angela M.', mrn: '00391042', procedure: 'Abdominoplasty', dob: '07/22/1985', flags: 1 },
  { id: 3, name: 'Sarah L.', mrn: '00512398', procedure: 'Rhinoplasty', dob: '11/05/1990', flags: 0 },
  { id: 4, name: 'Marcus K.', mrn: '00298741', procedure: 'Facelift', dob: '04/18/1968', flags: 0 },
]

// ── Main Component ─────────────────────────────────────────────────────────────
export default function SurgeonIPad() {
  const [tabletMode, setTabletMode] = useState(false)
  const [selectedPatient, setSelectedPatient] = useState<Patient>(MOCK_PATIENTS[0])
  const [activeRegion, setActiveRegion] = useState<BodyRegion>('breast')
  const [activeTool, setActiveTool] = useState<DrawTool>('incision')
  const [strokes, setStrokes] = useState<DrawStroke[]>([])
  const [currentStroke, setCurrentStroke] = useState<[number, number][]>([])
  const [isDrawing, setIsDrawing] = useState(false)
  const [noteStatus, setNoteStatus] = useState<NoteStatus>('idle')
  const [activeVoiceField, setActiveVoiceField] = useState<VoiceField>('chiefComplaint')
  const [notes, setNotes] = useState<SurgicalNote>({
    chiefComplaint: '',
    findings: '',
    assessment: '',
    plan: '',
    followUp: '',
  })
  const [voiceTranscript, setVoiceTranscript] = useState('')
  const [showCamera, setShowCamera] = useState(false)
  const [savedPhotos, setSavedPhotos] = useState<string[]>([])
  const [activeTab, setActiveTab] = useState<'draw' | 'voice' | 'photos'>('draw')

  const canvasRef = useRef<HTMLCanvasElement>(null)
  const videoRef = useRef<HTMLVideoElement>(null)
  const recognitionRef = useRef<any>(null)

  // ── Canvas drawing ────────────────────────────────────────────────────────
  const getCanvasPoint = useCallback((e: React.PointerEvent<HTMLCanvasElement>): [number, number] => {
    const canvas = canvasRef.current!
    const rect = canvas.getBoundingClientRect()
    const scaleX = canvas.width / rect.width
    const scaleY = canvas.height / rect.height
    return [
      (e.clientX - rect.left) * scaleX,
      (e.clientY - rect.top) * scaleY,
    ]
  }, [])

  const redrawCanvas = useCallback((allStrokes: DrawStroke[], current: [number, number][] = []) => {
    const canvas = canvasRef.current
    if (!canvas) return
    const ctx = canvas.getContext('2d')!
    ctx.clearRect(0, 0, canvas.width, canvas.height)

    allStrokes.forEach(stroke => {
      if (stroke.points.length < 2) return
      ctx.beginPath()
      ctx.strokeStyle = TOOL_COLORS[stroke.tool]
      ctx.lineWidth = stroke.tool === 'erase' ? 20 : stroke.tool === 'measurement' ? 2 : 3
      ctx.lineCap = 'round'
      ctx.lineJoin = 'round'
      if (stroke.tool === 'erase') ctx.globalCompositeOperation = 'destination-out'
      else ctx.globalCompositeOperation = 'source-over'
      ctx.moveTo(stroke.points[0][0], stroke.points[0][1])
      stroke.points.forEach(pt => ctx.lineTo(pt[0], pt[1]))
      ctx.stroke()
      if (stroke.tool === 'measurement' && stroke.points.length >= 2) {
        const last = stroke.points[stroke.points.length - 1]
        ctx.font = '12px system-ui'
        ctx.fillStyle = TOOL_COLORS.measurement
        ctx.globalCompositeOperation = 'source-over'
        ctx.fillText('~', last[0] + 6, last[1] - 4)
      }
    })

    if (current.length >= 2) {
      ctx.beginPath()
      ctx.strokeStyle = TOOL_COLORS[activeTool]
      ctx.lineWidth = activeTool === 'erase' ? 20 : 3
      ctx.lineCap = 'round'
      ctx.lineJoin = 'round'
      ctx.globalCompositeOperation = activeTool === 'erase' ? 'destination-out' : 'source-over'
      ctx.moveTo(current[0][0], current[0][1])
      current.forEach(pt => ctx.lineTo(pt[0], pt[1]))
      ctx.stroke()
    }
    ctx.globalCompositeOperation = 'source-over'
  }, [activeTool])

  useEffect(() => { redrawCanvas(strokes) }, [strokes, redrawCanvas])

  const onPointerDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const pt = getCanvasPoint(e)
    setIsDrawing(true)
    setCurrentStroke([pt])
    canvasRef.current?.setPointerCapture(e.pointerId)
  }

  const onPointerMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (!isDrawing) return
    const pt = getCanvasPoint(e)
    const next = [...currentStroke, pt]
    setCurrentStroke(next)
    redrawCanvas(strokes, next)
  }

  const onPointerUp = () => {
    if (!isDrawing || currentStroke.length < 2) { setIsDrawing(false); return }
    const newStroke: DrawStroke = { tool: activeTool, points: currentStroke }
    setStrokes(prev => [...prev, newStroke])
    setCurrentStroke([])
    setIsDrawing(false)
  }

  const clearCanvas = () => { setStrokes([]); redrawCanvas([]) }
  const undoLast = () => {
    const next = strokes.slice(0, -1)
    setStrokes(next)
    redrawCanvas(next)
  }

  // ── Voice dictation ────────────────────────────────────────────────────────
  const startListening = () => {
    const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition
    if (!SpeechRecognition) { alert('Voice dictation requires Chrome or Safari on iPad.'); return }

    recognitionRef.current = new SpeechRecognition()
    recognitionRef.current.continuous = true
    recognitionRef.current.interimResults = true
    recognitionRef.current.lang = 'en-US'

    recognitionRef.current.onresult = (e: any) => {
      const transcript = Array.from(e.results)
        .map((r: any) => r[0].transcript)
        .join(' ')
      setVoiceTranscript(transcript)
    }

    recognitionRef.current.onend = () => setNoteStatus('processing')
    recognitionRef.current.start()
    setNoteStatus('listening')
  }

  const stopListening = () => {
    recognitionRef.current?.stop()
    setTimeout(() => {
      setNotes(prev => ({ ...prev, [activeVoiceField]: voiceTranscript }))
      setVoiceTranscript('')
      setNoteStatus('saved')
      setTimeout(() => setNoteStatus('idle'), 2000)
    }, 600)
  }

  // ── Camera ────────────────────────────────────────────────────────────────
  const openCamera = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } })
      if (videoRef.current) videoRef.current.srcObject = stream
      setShowCamera(true)
    } catch { alert('Camera access denied or not available.') }
  }

  const capturePhoto = () => {
    const video = videoRef.current
    if (!video) return
    const canvas = document.createElement('canvas')
    canvas.width = video.videoWidth
    canvas.height = video.videoHeight
    canvas.getContext('2d')!.drawImage(video, 0, 0)
    const dataUrl = canvas.toDataURL('image/jpeg', 0.85)
    setSavedPhotos(prev => [...prev, dataUrl])
    const stream = video.srcObject as MediaStream
    stream?.getTracks().forEach(t => t.stop())
    setShowCamera(false)
  }

  // ── Styles ────────────────────────────────────────────────────────────────
  const navy = '#0a1628'
  const gold = '#c9a96e'
  const teal = '#4a7fa8'
  const ivory = '#f0ece4'
  const s = {
    wrap: { fontFamily: "'DM Sans', system-ui, sans-serif", background: '#f5f3ef', minHeight: 400 } as React.CSSProperties,
    toggleBar: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px 20px', background: navy, borderRadius: 12 } as React.CSSProperties,
    brandRow: { display: 'flex', alignItems: 'center', gap: 10 } as React.CSSProperties,
    brandName: { fontFamily: 'Georgia, serif', fontSize: 18, color: ivory, letterSpacing: '0.04em' } as React.CSSProperties,
    brandSub: { fontSize: 10, color: teal, letterSpacing: '0.14em', textTransform: 'uppercase' as const },
    toggleBtn: { display: 'flex', alignItems: 'center', gap: 8, padding: '8px 16px', borderRadius: 8, border: `1px solid ${gold}`, background: 'transparent', color: gold, fontSize: 12, cursor: 'pointer', fontFamily: 'inherit', letterSpacing: '0.1em' } as React.CSSProperties,
    ipadFrame: { background: navy, borderRadius: 20, padding: 16, marginTop: 12, boxShadow: '0 8px 32px rgba(10,22,40,0.3)' } as React.CSSProperties,
    screen: { background: '#fff', borderRadius: 12, overflow: 'hidden' } as React.CSSProperties,
    topBar: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 16px', background: navy } as React.CSSProperties,
    patientChip: { display: 'flex', alignItems: 'center', gap: 8 } as React.CSSProperties,
    avatar: { width: 32, height: 32, borderRadius: '50%', background: '#ccfbf1', color: '#0f766e', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 500 } as React.CSSProperties,
    patientName: { fontSize: 14, fontWeight: 500, color: ivory } as React.CSSProperties,
    patientProcedure: { fontSize: 10, color: teal, letterSpacing: '0.08em' } as React.CSSProperties,
    flagBadge: { padding: '3px 8px', borderRadius: 10, background: '#ffe4e6', color: '#9f1239', fontSize: 10, fontWeight: 500 } as React.CSSProperties,
    tabBar: { display: 'flex', background: '#f8f7f5', borderBottom: '1px solid #e8e4de' } as React.CSSProperties,
    tab: (active: boolean) => ({ flex: 1, padding: '10px 4px', fontSize: 12, fontWeight: 500, textAlign: 'center' as const, cursor: 'pointer', borderBottom: active ? `2px solid ${navy}` : '2px solid transparent', color: active ? navy : '#888', background: 'none', border: 'none', borderBottom: active ? `2px solid ${gold}` : '2px solid transparent', fontFamily: 'inherit' }),
    body: { display: 'flex', gap: 0, minHeight: 360 } as React.CSSProperties,
    sidebar: { width: 110, background: '#f8f7f5', borderRight: '1px solid #e8e4de', padding: '8px 0' } as React.CSSProperties,
    regionBtn: (active: boolean) => ({ display: 'block', width: '100%', padding: '8px 6px', fontSize: 11, textAlign: 'center' as const, cursor: 'pointer', background: active ? navy : 'none', color: active ? gold : '#555', border: 'none', fontFamily: 'inherit', fontWeight: active ? 500 : 400, borderLeft: active ? `3px solid ${gold}` : '3px solid transparent' }),
    main: { flex: 1, display: 'flex', flexDirection: 'column' as const },
    canvasWrap: { flex: 1, position: 'relative' as const, background: '#fafaf8' },
    toolBar: { display: 'flex', alignItems: 'center', gap: 6, padding: '8px 10px', background: '#f8f7f5', borderTop: '1px solid #e8e4de', flexWrap: 'wrap' as const },
    toolBtn: (active: boolean, tool: DrawTool) => ({ padding: '6px 10px', borderRadius: 6, fontSize: 11, fontWeight: 500, cursor: 'pointer', border: active ? `2px solid ${TOOL_COLORS[tool]}` : '1px solid #ddd', background: active ? `${TOOL_COLORS[tool]}18` : '#fff', color: active ? TOOL_COLORS[tool] : '#555', fontFamily: 'inherit' }),
    iconBtn: { padding: '6px 10px', borderRadius: 6, fontSize: 11, cursor: 'pointer', border: '1px solid #ddd', background: '#fff', color: '#555', fontFamily: 'inherit' } as React.CSSProperties,
    voiceWrap: { padding: 16 } as React.CSSProperties,
    fieldBtn: (active: boolean) => ({ padding: '8px 12px', borderRadius: 8, fontSize: 12, cursor: 'pointer', border: active ? `2px solid ${navy}` : '1px solid #ddd', background: active ? navy : '#fff', color: active ? ivory : '#333', fontFamily: 'inherit', fontWeight: active ? 500 : 400 }),
    noteField: { width: '100%', padding: '8px 10px', borderRadius: 8, border: '1px solid #e0dbd4', background: '#fafaf8', fontSize: 13, color: '#0a1628', fontFamily: 'inherit', resize: 'vertical' as const, minHeight: 56 },
    micBtn: (listening: boolean) => ({ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, padding: '12px 24px', borderRadius: 10, border: 'none', background: listening ? '#e11d48' : navy, color: '#fff', fontSize: 13, fontWeight: 500, cursor: 'pointer', fontFamily: 'inherit', transition: 'all 0.2s' }),
    statusPill: (status: NoteStatus) => ({ padding: '4px 12px', borderRadius: 20, fontSize: 11, fontWeight: 500, background: status === 'listening' ? '#ffe4e6' : status === 'saved' ? '#ccfbf1' : status === 'processing' ? '#fef3c7' : '#f0ece4', color: status === 'listening' ? '#9f1239' : status === 'saved' ? '#0f766e' : status === 'processing' ? '#92400e' : '#666' }),
    photosWrap: { padding: 12, display: 'flex', flexDirection: 'column' as const, gap: 10 },
    photoGrid: { display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 6 },
    photoThumb: { width: '100%', aspectRatio: '1', objectFit: 'cover' as const, borderRadius: 6, border: '1px solid #e0dbd4' },
    camBtn: { display: 'flex', alignItems: 'center', gap: 8, padding: '10px 18px', borderRadius: 10, border: `1px solid ${navy}`, background: navy, color: ivory, fontSize: 13, cursor: 'pointer', fontFamily: 'inherit' } as React.CSSProperties,
    patientList: { display: 'flex', gap: 8, padding: '8px 12px', background: '#f8f7f5', borderBottom: '1px solid #e8e4de', overflowX: 'auto' as const },
    patientPill: (active: boolean) => ({ padding: '6px 12px', borderRadius: 20, fontSize: 11, cursor: 'pointer', border: active ? `2px solid ${navy}` : '1px solid #ddd', background: active ? navy : '#fff', color: active ? ivory : '#444', fontFamily: 'inherit', whiteSpace: 'nowrap' as const, flexShrink: 0 }),
    saveBtn: { padding: '10px 20px', borderRadius: 10, border: 'none', background: gold, color: navy, fontSize: 13, fontWeight: 500, cursor: 'pointer', fontFamily: 'inherit' } as React.CSSProperties,
  }

  const VOICE_FIELDS: { key: VoiceField; label: string; placeholder: string }[] = [
    { key: 'chiefComplaint', label: 'Chief complaint', placeholder: 'Tap mic and speak chief complaint...' },
    { key: 'findings', label: 'Findings', placeholder: 'Physical examination findings...' },
    { key: 'assessment', label: 'Assessment', placeholder: 'Clinical assessment and diagnosis...' },
    { key: 'plan', label: 'Plan', placeholder: 'Surgical and treatment plan...' },
    { key: 'followUp', label: 'Follow-up', placeholder: 'Post-op instructions and follow-up...' },
  ]

  const REGIONS: BodyRegion[] = ['full-body','face','breast','stomach','buttocks','chest','nose','eyes','arms','legs']

  return (
    <div style={s.wrap}>
      <div style={s.toggleBar}>
        <div style={s.brandRow}>
          <svg width="28" height="28" viewBox="0 0 44 44" fill="none">
            <circle cx="22" cy="22" r="14" stroke="#c9a96e" strokeWidth="1"/>
            <path d="M22 8 C28 14 30 18 30 22 C30 26 28 30 22 36 C16 30 14 26 14 22 C14 18 16 14 22 8Z" fill="none" stroke="#f0ece4" strokeWidth="1.2"/>
            <circle cx="22" cy="22" r="2.5" fill="#c9a96e"/>
          </svg>
          <div>
            <div style={s.brandSub}>PatientTrac</div>
            <div style={s.brandName}>Revela</div>
          </div>
        </div>
        <button style={s.toggleBtn} onClick={() => setTabletMode(t => !t)}>
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
            <rect x="3" y="4" width="14" height="18" rx="2"/>
            <circle cx="10" cy="18" r="1"/>
          </svg>
          {tabletMode ? 'Desktop mode' : 'Surgeon tablet mode'}
        </button>
      </div>

      {!tabletMode ? (
        <div style={{ padding: '24px 0', textAlign: 'center', color: '#666', fontSize: 14 }}>
          Toggle <strong style={{ color: navy }}>Surgeon tablet mode</strong> above to launch the iPad surgeon interface
        </div>
      ) : (
        <div style={s.ipadFrame}>
          <div style={s.screen}>
            <div style={s.topBar}>
              <div style={s.patientChip}>
                <div style={s.avatar}>{selectedPatient.name.split(' ').map(n => n[0]).join('')}</div>
                <div>
                  <div style={s.patientName}>{selectedPatient.name}</div>
                  <div style={s.patientProcedure}>{selectedPatient.procedure} · MRN {selectedPatient.mrn}</div>
                </div>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                {selectedPatient.flags > 0 && (
                  <span style={s.flagBadge}>{selectedPatient.flags} AI flag{selectedPatient.flags > 1 ? 's' : ''}</span>
                )}
                <button style={s.saveBtn} onClick={() => alert('Note saved to Supabase ✓')}>Save note</button>
              </div>
            </div>

            <div style={s.patientList}>
              {MOCK_PATIENTS.map(p => (
                <button key={p.id} style={s.patientPill(p.id === selectedPatient.id)} onClick={() => { setSelectedPatient(p); clearCanvas() }}>
                  {p.name} {p.flags > 0 ? '⚑' : ''}
                </button>
              ))}
            </div>

            <div style={s.tabBar}>
              {(['draw', 'voice', 'photos'] as const).map(tab => (
                <button key={tab} style={s.tab(activeTab === tab)} onClick={() => setActiveTab(tab)}>
                  {tab === 'draw' ? 'Diagram & draw' : tab === 'voice' ? 'Voice dictation' : 'Photos'}
                </button>
              ))}
            </div>

            {activeTab === 'draw' && (
              <div style={s.body}>
                <div style={s.sidebar}>
                  {REGIONS.map(r => (
                    <button key={r} style={s.regionBtn(r === activeRegion)} onClick={() => { setActiveRegion(r); clearCanvas() }}>
                      {REGION_LABELS[r]}
                    </button>
                  ))}
                </div>
                <div style={s.main}>
                  <div style={s.canvasWrap}>
                    <svg style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', pointerEvents: 'none' }}
                      viewBox="0 0 300 280" preserveAspectRatio="xMidYMid meet">
                      <g dangerouslySetInnerHTML={{ __html: BODY_DIAGRAMS[activeRegion] }} />
                    </svg>
                    <canvas
                      ref={canvasRef}
                      width={600} height={560}
                      style={{ width: '100%', height: 280, display: 'block', touchAction: 'none', cursor: 'crosshair' }}
                      onPointerDown={onPointerDown}
                      onPointerMove={onPointerMove}
                      onPointerUp={onPointerUp}
                      onPointerLeave={onPointerUp}
                    />
                  </div>
                  <div style={s.toolBar}>
                    {(['incision','measurement','annotation','concern','erase'] as DrawTool[]).map(tool => (
                      <button key={tool} style={s.toolBtn(activeTool === tool, tool)} onClick={() => setActiveTool(tool)}>
                        {TOOL_LABELS[tool]}
                      </button>
                    ))}
                    <button style={s.iconBtn} onClick={undoLast}>↩ Undo</button>
                    <button style={s.iconBtn} onClick={clearCanvas}>Clear</button>
                  </div>
                </div>
              </div>
            )}

            {activeTab === 'voice' && (
              <div style={s.voiceWrap}>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 12 }}>
                  {VOICE_FIELDS.map(f => (
                    <button key={f.key} style={s.fieldBtn(activeVoiceField === f.key)} onClick={() => setActiveVoiceField(f.key)}>
                      {f.label}
                    </button>
                  ))}
                </div>

                {VOICE_FIELDS.map(f => (
                  activeVoiceField === f.key && (
                    <div key={f.key}>
                      <div style={{ fontSize: 11, color: '#888', marginBottom: 6, textTransform: 'uppercase' as const, letterSpacing: '0.1em' }}>
                        {f.label}
                      </div>
                      <textarea
                        style={s.noteField}
                        placeholder={f.placeholder}
                        value={noteStatus === 'listening' ? voiceTranscript : notes[f.key]}
                        onChange={e => setNotes(prev => ({ ...prev, [f.key]: e.target.value }))}
                        rows={4}
                      />
                    </div>
                  )
                ))}

                <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 12 }}>
                  <button
                    style={s.micBtn(noteStatus === 'listening')}
                    onClick={noteStatus === 'listening' ? stopListening : startListening}
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                      <path d="M12 1a3 3 0 0 0-3 3v8a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/>
                      <path d="M19 10v2a7 7 0 0 1-14 0v-2" fill="none" stroke="currentColor" strokeWidth="2"/>
                      <line x1="12" y1="19" x2="12" y2="23" stroke="currentColor" strokeWidth="2"/>
                      <line x1="8" y1="23" x2="16" y2="23" stroke="currentColor" strokeWidth="2"/>
                    </svg>
                    {noteStatus === 'listening' ? 'Stop dictating' : 'Start dictating'}
                  </button>
                  {noteStatus !== 'idle' && (
                    <span style={s.statusPill(noteStatus)}>
                      {noteStatus === 'listening' ? 'Listening...' : noteStatus === 'processing' ? 'Processing...' : 'Saved ✓'}
                    </span>
                  )}
                </div>

                <div style={{ marginTop: 16, padding: 12, background: '#f8f7f5', borderRadius: 8, fontSize: 12, color: '#555' }}>
                  <div style={{ fontWeight: 500, color: navy, marginBottom: 6 }}>Structured note preview</div>
                  {VOICE_FIELDS.map(f => notes[f.key] && (
                    <div key={f.key} style={{ marginBottom: 6 }}>
                      <span style={{ color: teal, fontWeight: 500 }}>{f.label}: </span>
                      <span>{notes[f.key]}</span>
                    </div>
                  ))}
                  {!VOICE_FIELDS.some(f => notes[f.key]) && (
                    <span style={{ color: '#aaa' }}>Dictate into fields above to build structured note...</span>
                  )}
                </div>
              </div>
            )}

            {activeTab === 'photos' && (
              <div style={s.photosWrap}>
                {showCamera ? (
                  <div>
                    <video ref={videoRef} autoPlay playsInline style={{ width: '100%', borderRadius: 10, background: '#000' }} />
                    <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
                      <button style={{ ...s.saveBtn, flex: 1 }} onClick={capturePhoto}>Capture photo</button>
                      <button style={{ ...s.iconBtn, flex: 1 }} onClick={() => setShowCamera(false)}>Cancel</button>
                    </div>
                  </div>
                ) : (
                  <>
                    <button style={s.camBtn} onClick={openCamera}>
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                        <path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/>
                        <circle cx="12" cy="13" r="4"/>
                      </svg>
                      Open camera
                    </button>
                    {savedPhotos.length > 0 ? (
                      <div style={s.photoGrid}>
                        {savedPhotos.map((src, i) => (
                          <img key={i} src={src} style={s.photoThumb} alt={`Clinical photo ${i + 1}`} />
                        ))}
                      </div>
                    ) : (
                      <div style={{ textAlign: 'center', padding: '24px 0', color: '#aaa', fontSize: 13 }}>
                        No photos captured yet.<br />Use the camera to document surgical sites.
                      </div>
                    )}
                    <div style={{ fontSize: 11, color: '#aaa', textAlign: 'center' }}>
                      Photos save to Supabase storage · Linked to encounter {selectedPatient.mrn}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}
