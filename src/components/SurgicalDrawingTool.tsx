import { useEffect, useRef, useState, useCallback } from 'react'
import { supabase } from '../lib/supabaseClient'

// ─── Servier Medical Art template registry ───────────────────────────────────
// Templates downloaded from smart.servier.com (CC BY 4.0)
// Attribution: "Illustrations by Servier Medical Art (smart.servier.com), CC BY 4.0"
// Stored in Supabase Storage public bucket: revela-templates

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL

const TEMPLATES: Record<string, { label: string; views: { name: string; file: string }[] }> = {
  breast: {
    label: 'Breast',
    views: [
      { name: 'Front', file: 'female-breast-front.svg' },
      { name: 'Lateral', file: 'female-breast-lateral.svg' },
      { name: 'Oblique', file: 'female-breast-oblique.svg' },
    ],
  },
  body_female: {
    label: 'Body (Female)',
    views: [
      { name: 'Front', file: 'female-body-front.svg' },
      { name: 'Back', file: 'female-body-back.svg' },
      { name: 'Lateral', file: 'female-body-lateral.svg' },
    ],
  },
  face_female: {
    label: 'Face (Female)',
    views: [
      { name: 'Front', file: 'female-face-front.svg' },
      { name: 'Lateral', file: 'female-face-lateral.svg' },
    ],
  },
  pectoral: {
    label: 'Pectoral (Male)',
    views: [
      { name: 'Front', file: 'male-pectoral-front.svg' },
    ],
  },
  body_male: {
    label: 'Body (Male)',
    views: [
      { name: 'Front', file: 'male-body-front.svg' },
      { name: 'Back', file: 'male-body-back.svg' },
      { name: 'Lateral', file: 'male-body-lateral.svg' },
    ],
  },
  face_male: {
    label: 'Face (Male)',
    views: [
      { name: 'Front', file: 'male-face-front.svg' },
      { name: 'Lateral', file: 'male-face-lateral.svg' },
    ],
  },
}

type Tool = 'select' | 'draw' | 'line' | 'rect' | 'circle' | 'text' | 'eraser'

interface Props {
  encounterId: string
  patientId: string
  procedureType?: keyof typeof TEMPLATES
  onSave?: (drawingId: string) => void
}

export default function SurgicalDrawingTool({ encounterId, patientId, procedureType = 'breast', onSave }: Props) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const fabricRef = useRef<any>(null)
  const [fabricLoaded, setFabricLoaded] = useState(false)
  const [activeTool, setActiveTool] = useState<Tool>('draw')
  const [activeColor, setActiveColor] = useState('#e74c3c')
  const [brushSize, setBrushSize] = useState(3)
  const [activeTemplate, setActiveTemplate] = useState<keyof typeof TEMPLATES>(procedureType)
  const [activeView, setActiveView] = useState(0)
  const [saving, setSaving] = useState(false)
  const [saveStatus, setSaveStatus] = useState<'idle' | 'saved' | 'error'>('idle')
  const [history, setHistory] = useState<string[]>([])
  const [historyIndex, setHistoryIndex] = useState(-1)

  const colors = ['#e74c3c', '#e67e22', '#f1c40f', '#2ecc71', '#3498db', '#9b59b6', '#1abc9c', '#060e1c', '#ffffff']

  // ── Load Fabric.js from CDN ──────────────────────────────────────────────
  useEffect(() => {
    if ((window as any).fabric) { setFabricLoaded(true); return }
    const script = document.createElement('script')
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/fabric.js/5.3.1/fabric.min.js'
    script.onload = () => setFabricLoaded(true)
    document.head.appendChild(script)
  }, [])

  // ── Init canvas after Fabric loads ──────────────────────────────────────
  useEffect(() => {
    if (!fabricLoaded || !canvasRef.current) return
    const fabric = (window as any).fabric

    const canvas = new fabric.Canvas(canvasRef.current, {
      isDrawingMode: true,
      width: 720,
      height: 540,
      backgroundColor: '#ffffff',
    })

    canvas.freeDrawingBrush.color = activeColor
    canvas.freeDrawingBrush.width = brushSize

    // Save state on every change
    canvas.on('object:added', () => pushHistory(canvas))
    canvas.on('object:modified', () => pushHistory(canvas))
    canvas.on('object:removed', () => pushHistory(canvas))

    fabricRef.current = canvas
    loadTemplate(canvas, activeTemplate, activeView)

    return () => canvas.dispose()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fabricLoaded])

  // ── Load Servier template as background image ───────────────────────────
  const loadTemplate = useCallback((canvas: any, templateKey: string, viewIndex: number) => {
    if (!canvas) return
    const fabric = (window as any).fabric
    const template = TEMPLATES[templateKey]
    if (!template) return
    const file = template.views[viewIndex]?.file
    if (!file) return

    const url = `${SUPABASE_URL}/storage/v1/object/public/revela-templates/${file}`

    fabric.Image.fromURL(url, (img: any) => {
      if (!img) return
      // Scale to fit canvas maintaining aspect ratio
      const scale = Math.min(720 / (img.width || 720), 540 / (img.height || 540)) * 0.85
      img.set({
        scaleX: scale,
        scaleY: scale,
        left: (720 - (img.width || 0) * scale) / 2,
        top: (540 - (img.height || 0) * scale) / 2,
        selectable: false,
        evented: false,
        opacity: 0.85,
      })
      canvas.setBackgroundImage(img, canvas.renderAll.bind(canvas))
    }, { crossOrigin: 'anonymous' })
  }, [])

  // ── Template / view change ───────────────────────────────────────────────
  useEffect(() => {
    if (!fabricRef.current || !fabricLoaded) return
    fabricRef.current.clear()
    fabricRef.current.setBackgroundImage(null, fabricRef.current.renderAll.bind(fabricRef.current))
    loadTemplate(fabricRef.current, activeTemplate, activeView)
    setHistory([])
    setHistoryIndex(-1)
  }, [activeTemplate, activeView, fabricLoaded, loadTemplate])

  // ── Tool change ──────────────────────────────────────────────────────────
  useEffect(() => {
    const canvas = fabricRef.current
    if (!canvas) return
    const fabric = (window as any).fabric

    canvas.isDrawingMode = false
    canvas.selection = false

    switch (activeTool) {
      case 'draw':
        canvas.isDrawingMode = true
        canvas.freeDrawingBrush = new fabric.PencilBrush(canvas)
        canvas.freeDrawingBrush.color = activeColor
        canvas.freeDrawingBrush.width = brushSize
        break
      case 'eraser':
        canvas.isDrawingMode = true
        canvas.freeDrawingBrush = new fabric.PencilBrush(canvas)
        canvas.freeDrawingBrush.color = '#ffffff'
        canvas.freeDrawingBrush.width = brushSize * 4
        break
      case 'select':
        canvas.selection = true
        break
      case 'text': {
        canvas.on('mouse:down', handleTextAdd)
        return () => canvas.off('mouse:down', handleTextAdd)
      }
      case 'rect': {
        canvas.on('mouse:down', handleRectAdd)
        return () => canvas.off('mouse:down', handleRectAdd)
      }
      case 'circle': {
        canvas.on('mouse:down', handleCircleAdd)
        return () => canvas.off('mouse:down', handleCircleAdd)
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTool, activeColor, brushSize])

  const handleTextAdd = useCallback((opt: any) => {
    const fabric = (window as any).fabric
    const canvas = fabricRef.current
    const pointer = canvas.getPointer(opt.e)
    const text = new fabric.IText('Label', {
      left: pointer.x,
      top: pointer.y,
      fontSize: 14,
      fill: activeColor,
      fontFamily: 'Arial, sans-serif',
      fontWeight: 'bold',
      editable: true,
    })
    canvas.add(text)
    canvas.setActiveObject(text)
    text.enterEditing()
    canvas.off('mouse:down', handleTextAdd)
    setActiveTool('select')
  }, [activeColor])

  const handleRectAdd = useCallback((opt: any) => {
    const fabric = (window as any).fabric
    const canvas = fabricRef.current
    const pointer = canvas.getPointer(opt.e)
    const rect = new fabric.Rect({
      left: pointer.x - 40,
      top: pointer.y - 25,
      width: 80,
      height: 50,
      fill: 'transparent',
      stroke: activeColor,
      strokeWidth: 2,
      rx: 4,
    })
    canvas.add(rect)
    canvas.off('mouse:down', handleRectAdd)
    setActiveTool('select')
  }, [activeColor])

  const handleCircleAdd = useCallback((opt: any) => {
    const fabric = (window as any).fabric
    const canvas = fabricRef.current
    const pointer = canvas.getPointer(opt.e)
    const circle = new fabric.Circle({
      left: pointer.x - 30,
      top: pointer.y - 30,
      radius: 30,
      fill: 'transparent',
      stroke: activeColor,
      strokeWidth: 2,
    })
    canvas.add(circle)
    canvas.off('mouse:down', handleCircleAdd)
    setActiveTool('select')
  }, [activeColor])

  // ── History ──────────────────────────────────────────────────────────────
  const pushHistory = (canvas: any) => {
    const json = JSON.stringify(canvas.toJSON())
    setHistory(prev => {
      const next = prev.slice(0, historyIndex + 1)
      next.push(json)
      setHistoryIndex(next.length - 1)
      return next
    })
  }

  const undo = () => {
    const canvas = fabricRef.current
    if (!canvas || historyIndex <= 0) return
    const newIndex = historyIndex - 1
    canvas.loadFromJSON(history[newIndex], canvas.renderAll.bind(canvas))
    setHistoryIndex(newIndex)
  }

  const redo = () => {
    const canvas = fabricRef.current
    if (!canvas || historyIndex >= history.length - 1) return
    const newIndex = historyIndex + 1
    canvas.loadFromJSON(history[newIndex], canvas.renderAll.bind(canvas))
    setHistoryIndex(newIndex)
  }

  const clearDrawing = () => {
    const canvas = fabricRef.current
    if (!canvas) return
    canvas.getObjects().forEach((obj: any) => canvas.remove(obj))
    canvas.renderAll()
  }

  // ── Save drawing to Supabase ──────────────────────────────────────────────
  const saveDrawing = async () => {
    const canvas = fabricRef.current
    if (!canvas) return
    setSaving(true)
    setSaveStatus('idle')

    try {
      const svgData = canvas.toSVG()
      const jsonData = JSON.stringify(canvas.toJSON())
      const templateInfo = `${activeTemplate}:${TEMPLATES[activeTemplate].views[activeView].name}`

      const { data, error } = await supabase
        .from('surgical_drawings')
        .upsert({
          encounter_id: encounterId,
          patient_id: patientId,
          template_key: templateInfo,
          drawing_svg: svgData,
          drawing_json: jsonData,
          updated_at: new Date().toISOString(),
        }, { onConflict: 'encounter_id,template_key' })
        .select('id')
        .single()

      if (error) throw error
      setSaveStatus('saved')
      onSave?.(data.id)
      setTimeout(() => setSaveStatus('idle'), 3000)
    } catch (err) {
      console.error('Save error:', err)
      setSaveStatus('error')
    } finally {
      setSaving(false)
    }
  }

  // ── Export PNG ────────────────────────────────────────────────────────────
  const exportPng = () => {
    const canvas = fabricRef.current
    if (!canvas) return
    const dataUrl = canvas.toDataURL({ format: 'png', multiplier: 2 })
    const a = document.createElement('a')
    a.href = dataUrl
    a.download = `drawing-${encounterId}-${activeTemplate}-${Date.now()}.png`
    a.click()
  }

  const template = TEMPLATES[activeTemplate]

  return (
    <div style={{ fontFamily: 'var(--font-sans, sans-serif)', background: '#060e1c', borderRadius: 12, overflow: 'hidden', border: '1px solid rgba(201,169,110,0.2)' }}>

      {/* ── Header bar ── */}
      <div style={{ background: '#0a1628', padding: '12px 16px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid rgba(201,169,110,0.15)' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
            <polygon points="9,1 17,9 9,17 1,9" fill="none" stroke="#c9a96e" strokeWidth="1.2"/>
            <polygon points="9,5 13,9 9,13 5,9" fill="#c9a96e" opacity="0.4"/>
          </svg>
          <span style={{ color: '#c9a96e', fontSize: 14, fontWeight: 600, letterSpacing: '0.3px' }}>Surgical Drawing</span>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button onClick={exportPng} style={btnStyle('ghost')}>Export PNG</button>
          <button onClick={saveDrawing} disabled={saving} style={btnStyle('gold')}>
            {saving ? 'Saving…' : saveStatus === 'saved' ? '✓ Saved' : saveStatus === 'error' ? '✗ Error' : 'Save Drawing'}
          </button>
        </div>
      </div>

      <div style={{ display: 'flex' }}>

        {/* ── Left sidebar — templates ── */}
        <div style={{ width: 180, background: '#060e1c', borderRight: '1px solid rgba(255,255,255,0.06)', padding: '14px 10px', flexShrink: 0 }}>
          <div style={{ fontSize: 10, color: 'rgba(201,169,110,0.6)', letterSpacing: '1.2px', textTransform: 'uppercase', marginBottom: 10, paddingLeft: 4 }}>Template</div>
          {Object.entries(TEMPLATES).map(([key, tmpl]) => (
            <button key={key} onClick={() => { setActiveTemplate(key as keyof typeof TEMPLATES); setActiveView(0) }}
              style={{ width: '100%', textAlign: 'left', padding: '8px 10px', borderRadius: 6, border: 'none', cursor: 'pointer', marginBottom: 2, fontSize: 13,
                background: activeTemplate === key ? 'rgba(201,169,110,0.15)' : 'transparent',
                color: activeTemplate === key ? '#c9a96e' : 'rgba(255,255,255,0.55)',
                fontWeight: activeTemplate === key ? 600 : 400,
              }}>
              {tmpl.label}
            </button>
          ))}

          <div style={{ fontSize: 10, color: 'rgba(201,169,110,0.6)', letterSpacing: '1.2px', textTransform: 'uppercase', margin: '16px 0 10px', paddingLeft: 4 }}>View</div>
          {template.views.map((v, i) => (
            <button key={i} onClick={() => setActiveView(i)}
              style={{ width: '100%', textAlign: 'left', padding: '7px 10px', borderRadius: 6, border: 'none', cursor: 'pointer', marginBottom: 2, fontSize: 13,
                background: activeView === i ? 'rgba(0,212,255,0.12)' : 'transparent',
                color: activeView === i ? '#00d4ff' : 'rgba(255,255,255,0.45)',
                fontWeight: activeView === i ? 600 : 400,
              }}>
              {v.name}
            </button>
          ))}
        </div>

        {/* ── Canvas area ── */}
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column' }}>

          {/* toolbar */}
          <div style={{ background: '#0a1628', padding: '10px 14px', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', borderBottom: '1px solid rgba(255,255,255,0.06)' }}>

            {/* Tools */}
            <div style={{ display: 'flex', gap: 4 }}>
              {([
                { id: 'select', icon: '↖', label: 'Select' },
                { id: 'draw', icon: '✏', label: 'Draw' },
                { id: 'line', icon: '╱', label: 'Line' },
                { id: 'rect', icon: '▭', label: 'Rect' },
                { id: 'circle', icon: '○', label: 'Circle' },
                { id: 'text', icon: 'T', label: 'Text' },
                { id: 'eraser', icon: '⌫', label: 'Erase' },
              ] as { id: Tool; icon: string; label: string }[]).map(t => (
                <button key={t.id} title={t.label} onClick={() => setActiveTool(t.id)}
                  style={{ width: 32, height: 32, borderRadius: 6, border: 'none', cursor: 'pointer', fontSize: 14, fontWeight: 600,
                    background: activeTool === t.id ? '#c9a96e' : 'rgba(255,255,255,0.06)',
                    color: activeTool === t.id ? '#060e1c' : 'rgba(255,255,255,0.6)',
                  }}>
                  {t.icon}
                </button>
              ))}
            </div>

            {/* Divider */}
            <div style={{ width: 1, height: 28, background: 'rgba(255,255,255,0.08)' }}/>

            {/* Colors */}
            <div style={{ display: 'flex', gap: 4 }}>
              {colors.map(c => (
                <button key={c} onClick={() => setActiveColor(c)}
                  style={{ width: 22, height: 22, borderRadius: '50%', border: activeColor === c ? '2px solid #c9a96e' : '2px solid transparent',
                    background: c, cursor: 'pointer', padding: 0, flexShrink: 0,
                    boxShadow: c === '#ffffff' ? 'inset 0 0 0 1px rgba(255,255,255,0.2)' : 'none',
                  }}/>
              ))}
            </div>

            {/* Divider */}
            <div style={{ width: 1, height: 28, background: 'rgba(255,255,255,0.08)' }}/>

            {/* Brush size */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.4)' }}>Size</span>
              <input type="range" min={1} max={20} value={brushSize} onChange={e => setBrushSize(Number(e.target.value))}
                style={{ width: 72, accentColor: '#c9a96e' }}/>
              <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.5)', width: 18 }}>{brushSize}</span>
            </div>

            {/* Divider */}
            <div style={{ width: 1, height: 28, background: 'rgba(255,255,255,0.08)' }}/>

            {/* Undo / Redo / Clear */}
            <div style={{ display: 'flex', gap: 4 }}>
              <button onClick={undo} disabled={historyIndex <= 0} style={btnStyle('ghost', historyIndex <= 0)}>↩ Undo</button>
              <button onClick={redo} disabled={historyIndex >= history.length - 1} style={btnStyle('ghost', historyIndex >= history.length - 1)}>↪ Redo</button>
              <button onClick={clearDrawing} style={btnStyle('ghost')}>✕ Clear</button>
            </div>
          </div>

          {/* canvas */}
          <div style={{ background: '#1a2235', display: 'flex', justifyContent: 'center', alignItems: 'center', padding: 20 }}>
            <canvas ref={canvasRef} style={{ borderRadius: 8, boxShadow: '0 4px 24px rgba(0,0,0,0.4)' }}/>
          </div>

          {/* Attribution footer — CC BY 4.0 required */}
          <div style={{ background: '#060e1c', padding: '8px 14px', borderTop: '1px solid rgba(255,255,255,0.05)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.25)' }}>
              Base illustrations by{' '}
              <a href="https://smart.servier.com" target="_blank" rel="noopener noreferrer"
                style={{ color: 'rgba(201,169,110,0.5)', textDecoration: 'none' }}>
                Servier Medical Art
              </a>
              {' '}(smart.servier.com), CC BY 4.0
            </span>
            <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.2)' }}>
              Encounter {encounterId.slice(0, 8)}…
            </span>
          </div>
        </div>
      </div>
    </div>
  )
}

// ── Style helpers ─────────────────────────────────────────────────────────────
function btnStyle(variant: 'gold' | 'ghost', disabled = false): React.CSSProperties {
  const base: React.CSSProperties = {
    padding: '6px 14px', borderRadius: 6, border: 'none', cursor: disabled ? 'not-allowed' : 'pointer',
    fontSize: 12, fontWeight: 600, letterSpacing: '0.3px', opacity: disabled ? 0.4 : 1, transition: 'opacity 0.15s',
  }
  if (variant === 'gold') return { ...base, background: '#c9a96e', color: '#060e1c' }
  return { ...base, background: 'rgba(255,255,255,0.07)', color: 'rgba(255,255,255,0.6)' }
}
