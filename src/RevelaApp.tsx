import { useState, useEffect, useCallback } from 'react'

const SUPABASE_URL = 'https://mskormozwekezjmtcylv.supabase.co'
const ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1za29ybW96d2VrZXpqbXRjeWx2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY5NzMzMzAsImV4cCI6MjA5MjU0OTMzMH0.nO9Q31CZJWRSIieLdLYVMkdu5NDDWjf1z0TgwCjBpp0'
const BRIDGE = `${SUPABASE_URL}/functions/v1/cross-app-bridge`
const DEVICES_FN = `${SUPABASE_URL}/functions/v1/implantable-devices`

async function bridge(action: string, body: Record<string, unknown> = {}) {
  const res = await fetch(BRIDGE, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', apikey: ANON_KEY },
    body: JSON.stringify({ action, ...body }),
  })
  return res.json()
}

async function callDevices(action: string, token: string, body?: Record<string, unknown>) {
  const url = `${DEVICES_FN}?action=${action}`
  const res = await fetch(url, {
    method: body ? 'POST' : 'GET',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
    body: body ? JSON.stringify(body) : undefined,
  })
  return res.json()
}

// ── Styles ─────────────────────────────────────────────────────────────────
const s: Record<string, React.CSSProperties> = {
  shell:    { minHeight:'100vh', background:'#020a14', color:'#e8eaf0', fontFamily:"'DM Sans',sans-serif" },
  center:   { display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', minHeight:'100vh' },
  spin:     { width:32, height:32, border:'3px solid #3a4a6a', borderTopColor:'#c9a96e', borderRadius:'50%', animation:'spin 1s linear infinite' },
  errBox:   { maxWidth:480, margin:'80px auto', padding:32, background:'#0a1628', border:'1px solid rgba(255,107,107,0.3)' },
  hdr:      { background:'#060e1c', borderBottom:'1px solid rgba(201,169,110,0.15)', padding:'12px 24px', display:'flex', alignItems:'center', justifyContent:'space-between' },
  badge:    { background:'#c9a96e', color:'#020a14', padding:'4px 10px', fontSize:11, fontWeight:700, letterSpacing:'0.12em', fontFamily:'DM Mono,monospace' },
  pname:    { fontSize:20, fontWeight:600, fontFamily:'Rajdhani,sans-serif', letterSpacing:'0.02em' },
  meta:     { fontSize:12, color:'#8a9bc0', marginTop:2 },
  tabs:     { display:'flex', borderBottom:'1px solid rgba(201,169,110,0.1)', background:'#060e1c', overflowX:'auto' },
  tab:      { padding:'12px 22px', background:'transparent', border:'none', borderBottom:'2px solid transparent', color:'#3a4a6a', cursor:'pointer', fontSize:12, fontFamily:'DM Mono,monospace', letterSpacing:'0.06em', textTransform:'uppercase', whiteSpace:'nowrap', flexShrink:0 },
  tabOn:    { color:'#c9a96e', borderBottomColor:'#c9a96e' },
  body:     { padding:24 },
  grid:     { display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(280px,1fr))', gap:16 },
  card:     { background:'rgba(10,22,40,0.8)', border:'1px solid rgba(201,169,110,0.1)', padding:20 },
  cttl:     { fontSize:10, fontFamily:'DM Mono,monospace', color:'#c9a96e', letterSpacing:'0.12em', textTransform:'uppercase', marginBottom:10 },
  cbdy:     { fontSize:14, color:'#8a9bc0', lineHeight:1.6 },
  form:     { maxWidth:720, display:'flex', flexDirection:'column', gap:18 },
  row:      { display:'grid', gridTemplateColumns:'1fr 1fr', gap:16 },
  fg:       { display:'flex', flexDirection:'column', gap:6 },
  lbl:      { fontSize:10, fontFamily:'DM Mono,monospace', color:'#3a4a6a', letterSpacing:'0.1em', textTransform:'uppercase' },
  inp:      { background:'rgba(6,14,28,0.8)', border:'1px solid rgba(201,169,110,0.15)', color:'#e8eaf0', padding:'10px 14px', fontSize:14, fontFamily:"'DM Sans',sans-serif", outline:'none', width:'100%' },
  ta:       { background:'rgba(6,14,28,0.8)', border:'1px solid rgba(201,169,110,0.15)', color:'#e8eaf0', padding:'12px 14px', fontSize:14, fontFamily:"'DM Sans',sans-serif", outline:'none', resize:'vertical', lineHeight:1.6, width:'100%' },
  btn:      { background:'#c9a96e', color:'#020a14', border:'none', padding:'12px 28px', fontSize:12, fontFamily:'DM Mono,monospace', letterSpacing:'0.12em', textTransform:'uppercase', cursor:'pointer', fontWeight:500, alignSelf:'flex-start' },
  btnOff:   { background:'rgba(201,169,110,0.3)', color:'#020a14', border:'none', padding:'12px 28px', fontSize:12, fontFamily:'DM Mono,monospace', letterSpacing:'0.12em', textTransform:'uppercase', cursor:'not-allowed', fontWeight:500, alignSelf:'flex-start' },
  btnSm:    { background:'rgba(201,169,110,0.12)', color:'#c9a96e', border:'1px solid rgba(201,169,110,0.3)', padding:'8px 18px', fontSize:11, fontFamily:'DM Mono,monospace', letterSpacing:'0.1em', textTransform:'uppercase', cursor:'pointer' },
  btnDanger:{ background:'rgba(239,68,68,0.1)', color:'#ef4444', border:'1px solid rgba(239,68,68,0.3)', padding:'6px 14px', fontSize:11, fontFamily:'DM Mono,monospace', cursor:'pointer' },
  ok:       { background:'rgba(74,222,128,0.08)', border:'1px solid rgba(74,222,128,0.2)', color:'#4ade80', padding:'8px 16px', marginBottom:16, fontSize:13 },
  err:      { background:'rgba(255,107,107,0.08)', border:'1px solid rgba(255,107,107,0.2)', color:'#ff6b6b', padding:'8px 16px', marginBottom:16, fontSize:13 },
  closeBtn: { background:'rgba(255,107,107,0.1)', border:'1px solid rgba(255,107,107,0.3)', color:'#ff6b6b', padding:'6px 16px', cursor:'pointer', fontSize:12, fontFamily:'DM Mono,monospace', letterSpacing:'0.08em' },
  muted:    { color:'#3a4a6a', fontSize:13 },
  encTag:   { fontSize:11, background:'rgba(0,212,255,0.1)', color:'#00d4ff', padding:'3px 8px', fontFamily:'DM Mono,monospace' },
  // Device-specific
  deviceCard:  { background:'rgba(10,22,40,0.9)', border:'1px solid rgba(201,169,110,0.15)', padding:16, marginBottom:12 },
  deviceName:  { fontSize:15, fontWeight:600, fontFamily:'Rajdhani,sans-serif', color:'#c9a96e', marginBottom:4 },
  deviceMeta:  { fontSize:11, fontFamily:'DM Mono,monospace', color:'#3a4a6a', lineHeight:1.8 },
  activeBadge: { display:'inline-block', background:'rgba(74,222,128,0.1)', color:'#4ade80', border:'1px solid rgba(74,222,128,0.2)', padding:'2px 8px', fontSize:10, fontFamily:'DM Mono,monospace' },
  removedBadge:{ display:'inline-block', background:'rgba(148,163,184,0.08)', color:'#64748b', border:'1px solid rgba(148,163,184,0.15)', padding:'2px 8px', fontSize:10, fontFamily:'DM Mono,monospace' },
  udiBox:      { background:'rgba(0,212,255,0.04)', border:'1px solid rgba(0,212,255,0.12)', padding:14, marginBottom:20 },
  catBtn:      { padding:'6px 14px', fontSize:11, fontFamily:'DM Mono,monospace', cursor:'pointer', border:'1px solid rgba(201,169,110,0.15)', background:'transparent', color:'#3a4a6a', letterSpacing:'0.08em' },
  catBtnOn:    { background:'rgba(201,169,110,0.1)', color:'#c9a96e', borderColor:'rgba(201,169,110,0.35)' },
  catalogGrid: { display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(200px,1fr))', gap:8, maxHeight:180, overflowY:'auto', marginBottom:16 },
  catalogItem: { padding:'8px 12px', fontSize:12, cursor:'pointer', border:'1px solid rgba(255,255,255,0.05)', background:'rgba(10,22,40,0.5)', color:'#8a9bc0', textAlign:'left' as const },
  catalogItemOn:{ background:'rgba(201,169,110,0.08)', color:'#c9a96e', borderColor:'rgba(201,169,110,0.3)' },
  siteGrid:    { display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(150px,1fr))', gap:6, maxHeight:130, overflowY:'auto', marginBottom:16 },
  siteBtn:     { padding:'6px 10px', fontSize:11, cursor:'pointer', border:'1px solid rgba(255,255,255,0.05)', background:'rgba(10,22,40,0.5)', color:'#8a9bc0', textAlign:'left' as const },
  siteBtnOn:   { background:'rgba(201,169,110,0.08)', color:'#c9a96e', borderColor:'rgba(201,169,110,0.3)' },
}

// ── Device catalog ─────────────────────────────────────────────────────────
const DEVICE_CATALOG = [
  { snomed:'360046005', name:'Breast implant',              category:'plastic' },
  { snomed:'360052003', name:'Silicone breast implant',     category:'plastic' },
  { snomed:'360047001', name:'Saline breast implant',       category:'plastic' },
  { snomed:'257275005', name:'Tissue expander',             category:'plastic' },
  { snomed:'360199000', name:'Facial implant',              category:'plastic' },
  { snomed:'304120007', name:'Total hip prosthesis',        category:'ortho'   },
  { snomed:'304121006', name:'Total knee prosthesis',       category:'ortho'   },
  { snomed:'360041000', name:'Cardiac pacemaker',           category:'cardiac' },
  { snomed:'14106009',  name:'Cardiac defibrillator',       category:'cardiac' },
  { snomed:'272287005', name:'Coronary stent',              category:'cardiac' },
  { snomed:'53350007',  name:'Prosthetic heart valve',      category:'cardiac' },
  { snomed:'79733001',  name:'Intraocular lens',            category:'other'   },
  { snomed:'360048006', name:'Cochlear implant',            category:'other'   },
  { snomed:'303728001', name:'Spinal cord stimulator',      category:'other'   },
  { snomed:'77444004',  name:'Insulin pump',                category:'other'   },
]

const BODY_SITES = [
  { snomed:'80248007',  site:'Left breast'       },
  { snomed:'73056007',  site:'Right breast'      },
  { snomed:'361355005', site:'Bilateral breasts' },
  { snomed:'302551006', site:'Face'              },
  { snomed:'362875007', site:'Chest wall'        },
  { snomed:'816094009', site:'Abdomen'           },
  { snomed:'362905008', site:'Right hip'         },
  { snomed:'362906009', site:'Left hip'          },
  { snomed:'57999000',  site:'Right knee'        },
  { snomed:'8873007',   site:'Left knee'         },
  { snomed:'302512001', site:'Heart'             },
  { snomed:'816080008', site:'Spine'             },
]

// ── Implantable Devices Tab ────────────────────────────────────────────────
function DevicesTab({ patientId, token }: { patientId: number; token: string }) {
  const [subView,       setSubView]       = useState<'list'|'add'>('list')
  const [devices,       setDevices]       = useState<Record<string, unknown>[]>([])
  const [loading,       setLoading]       = useState(false)
  const [saving,        setSaving]        = useState(false)
  const [devError,      setDevError]      = useState('')
  const [devSuccess,    setDevSuccess]    = useState('')
  const [showAll,       setShowAll]       = useState(false)
  const [confirmRemove, setConfirmRemove] = useState<number|null>(null)
  const [removalDate,   setRemovalDate]   = useState(new Date().toISOString().split('T')[0])

  // Form state
  const [udiInput,        setUdiInput]        = useState('')
  const [udiDi,           setUdiDi]           = useState('')
  const [udiParsing,      setUdiParsing]      = useState(false)
  const [deviceName,      setDeviceName]      = useState('')
  const [snomedCode,      setSnomedCode]      = useState('')
  const [manufacturer,    setManufacturer]    = useState('')
  const [modelNumber,     setModelNumber]     = useState('')
  const [lotNumber,       setLotNumber]       = useState('')
  const [serialNumber,    setSerialNumber]    = useState('')
  const [implantDate,     setImplantDate]     = useState(new Date().toISOString().split('T')[0])
  const [expirationDate,  setExpirationDate]  = useState('')
  const [bodySite,        setBodySite]        = useState('')
  const [bodySiteSnomed,  setBodySiteSnomed]  = useState('')
  const [notes,           setNotes]           = useState('')
  const [catFilter,       setCatFilter]       = useState('plastic')

  const load = useCallback(async () => {
    setLoading(true)
    const d = await callDevices(`list&patient_id=${patientId}&active_only=${!showAll}`, token)
    setDevices(d.devices ?? [])
    setLoading(false)
  }, [patientId, token, showAll])

  useEffect(() => { load() }, [load])

  const handleUdiChange = async (val: string) => {
    setUdiInput(val)
    if (val.length < 8) { setUdiDi(''); return }
    setUdiParsing(true)
    const d = await callDevices('parse-udi', token, { udi: val })
    setUdiDi(d.udi_di || '')
    setUdiParsing(false)
  }

  const handleAdd = async () => {
    if (!deviceName) return
    setSaving(true); setDevError('')
    const d = await callDevices('add', token, {
      patient_id: patientId,
      udi: udiInput || null, udi_di: udiDi || null,
      device_name: deviceName, snomed_code: snomedCode || null,
      manufacturer: manufacturer || null, model_number: modelNumber || null,
      lot_number: lotNumber || null, serial_number: serialNumber || null,
      implant_date: implantDate || null, expiration_date: expirationDate || null,
      body_site: bodySite || null, body_site_snomed: bodySiteSnomed || null,
      notes: notes || null,
    })
    setSaving(false)
    if (d.success) {
      setDevSuccess(`${deviceName} recorded`)
      setSubView('list'); load()
      setUdiInput(''); setUdiDi(''); setDeviceName(''); setSnomedCode('')
      setManufacturer(''); setModelNumber(''); setLotNumber(''); setSerialNumber('')
      setImplantDate(new Date().toISOString().split('T')[0]); setExpirationDate('')
      setBodySite(''); setBodySiteSnomed(''); setNotes('')
    } else {
      setDevError(d.error || 'Failed to add device')
    }
  }

  const handleRemove = async (deviceId: number) => {
    await callDevices('remove', token, { device_id: deviceId, removal_date: removalDate })
    setConfirmRemove(null); load()
  }

  const filteredCatalog = DEVICE_CATALOG.filter(c => catFilter === 'all' || c.category === catFilter)
  const activeDevices   = devices.filter(d => d.is_active)
  const removedDevices  = devices.filter(d => !d.is_active)

  return (
    <div>
      {/* Sub-nav */}
      <div style={{ display:'flex', gap:8, marginBottom:20, borderBottom:'1px solid rgba(201,169,110,0.08)', paddingBottom:12 }}>
        {[['list','Device List'],['add','Add Device']].map(([k,l]) => (
          <button key={k} onClick={() => { setSubView(k as 'list'|'add'); setDevError('') }}
            style={{ ...s.catBtn, ...(subView===k ? s.catBtnOn : {}) }}>
            {l}
          </button>
        ))}
        <span style={{ marginLeft:'auto', fontSize:10, fontFamily:'DM Mono,monospace', color:'#2a3a5a', alignSelf:'center' }}>
          FDA UDI · SNOMED CT · §170.315(a)(14)
        </span>
      </div>

      {devError   && <div style={s.err}>{devError}</div>}
      {devSuccess && <div style={s.ok}>✓ {devSuccess}</div>}

      {/* ── Device List ── */}
      {subView === 'list' && (
        <div>
          <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:16 }}>
            <div style={{ display:'flex', alignItems:'center', gap:12 }}>
              <span style={{ fontSize:14, fontFamily:'Rajdhani,sans-serif', color:'#c9a96e', fontWeight:600 }}>
                Implantable Devices
              </span>
              <span style={{ ...s.encTag }}>{activeDevices.length} active</span>
              <label style={{ display:'flex', alignItems:'center', gap:6, fontSize:11, color:'#3a4a6a', cursor:'pointer' }}>
                <input type="checkbox" checked={showAll} onChange={e => setShowAll(e.target.checked)} />
                Show removed
              </label>
            </div>
            <button onClick={() => setSubView('add')} style={s.btnSm}>+ Add Device</button>
          </div>

          {loading ? (
            <div style={{ ...s.muted, textAlign:'center', padding:32 }}>Loading devices…</div>
          ) : !devices.length ? (
            <div style={{ textAlign:'center', padding:48 }}>
              <div style={{ fontSize:36, opacity:0.12, marginBottom:12 }}>🔧</div>
              <div style={s.muted}>No implantable devices on record.</div>
              <button onClick={() => setSubView('add')} style={{ ...s.btnSm, marginTop:16 }}>Record First Device</button>
            </div>
          ) : (
            <div>
              {activeDevices.map(d => (
                <div key={d.device_id as number} style={s.deviceCard}>
                  <div style={{ display:'flex', alignItems:'flex-start', justifyContent:'space-between', marginBottom:8 }}>
                    <div>
                      <div style={s.deviceName}>{d.device_name as string}</div>
                      {d.manufacturer && <div style={{ fontSize:12, color:'#8a9bc0' }}>{d.manufacturer as string}{d.model_number ? ` · ${d.model_number}` : ''}</div>}
                      {d.body_site && <div style={{ fontSize:11, color:'#3a4a6a', marginTop:2 }}>📍 {d.body_site as string}</div>}
                    </div>
                    <div style={{ display:'flex', flexDirection:'column', alignItems:'flex-end', gap:6 }}>
                      <span style={s.activeBadge}>Active</span>
                      {confirmRemove === d.device_id ? (
                        <div style={{ textAlign:'right' }}>
                          <div style={{ fontSize:11, color:'#ef4444', marginBottom:6 }}>Confirm removal?</div>
                          <input type="date" value={removalDate} onChange={e => setRemovalDate(e.target.value)}
                            style={{ ...s.inp, width:140, fontSize:11, padding:'4px 8px', marginBottom:6 }} />
                          <div style={{ display:'flex', gap:6 }}>
                            <button onClick={() => setConfirmRemove(null)} style={{ ...s.catBtn, fontSize:10 }}>Cancel</button>
                            <button onClick={() => handleRemove(d.device_id as number)} style={{ ...s.btnDanger, fontSize:10 }}>Confirm</button>
                          </div>
                        </div>
                      ) : (
                        <button onClick={() => setConfirmRemove(d.device_id as number)} style={s.btnDanger}>Remove</button>
                      )}
                    </div>
                  </div>
                  <div style={s.deviceMeta}>
                    {d.udi && <div>UDI: <span style={{ color:'#5a6a8a' }}>{d.udi as string}</span></div>}
                    {d.udi_di && <div>DI: <span style={{ color:'#5a6a8a' }}>{d.udi_di as string}</span></div>}
                    {d.lot_number && <div>Lot: <span style={{ color:'#5a6a8a' }}>{d.lot_number as string}</span></div>}
                    {d.serial_number && <div>S/N: <span style={{ color:'#5a6a8a' }}>{d.serial_number as string}</span></div>}
                    {d.implant_date && <div>Implanted: <span style={{ color:'#5a6a8a' }}>{new Date(d.implant_date as string).toLocaleDateString()}</span></div>}
                    {d.expiration_date && <div>Expires: <span style={{ color:'#5a6a8a' }}>{new Date(d.expiration_date as string).toLocaleDateString()}</span></div>}
                    {d.snomed_code && <div style={{ opacity:0.4 }}>SNOMED: {d.snomed_code as string}</div>}
                  </div>
                  {d.notes && <div style={{ fontSize:12, color:'#3a4a6a', marginTop:8, fontStyle:'italic' }}>{d.notes as string}</div>}
                </div>
              ))}

              {showAll && removedDevices.length > 0 && (
                <div>
                  <div style={{ fontSize:10, fontFamily:'DM Mono,monospace', color:'#2a3a5a', letterSpacing:'0.1em', marginTop:24, marginBottom:12 }}>
                    REMOVED DEVICES
                  </div>
                  {removedDevices.map(d => (
                    <div key={d.device_id as number} style={{ ...s.deviceCard, opacity:0.5 }}>
                      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                        <div>
                          <div style={{ ...s.deviceName, color:'#64748b' }}>{d.device_name as string}</div>
                          {d.removal_date && <div style={{ fontSize:11, color:'#3a4a6a' }}>Removed: {new Date(d.removal_date as string).toLocaleDateString()}</div>}
                        </div>
                        <span style={s.removedBadge}>Removed</span>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* ── Add Device ── */}
      {subView === 'add' && (
        <div style={{ maxWidth:720 }}>

          {/* UDI input */}
          <div style={s.udiBox}>
            <div style={{ fontSize:10, fontFamily:'DM Mono,monospace', color:'#00d4ff', letterSpacing:'0.12em', marginBottom:8 }}>
              📷 FDA UDI SCAN / ENTRY · GS1 · HIBC
            </div>
            <div style={{ position:'relative' }}>
              <input value={udiInput} onChange={e => handleUdiChange(e.target.value)}
                placeholder="Scan or type UDI barcode string…"
                style={{ ...s.inp, fontFamily:'DM Mono,monospace', fontSize:13 }} />
              {udiParsing && <span style={{ position:'absolute', right:12, top:12, fontSize:12, color:'#3a4a6a' }}>⟳</span>}
            </div>
            {udiDi && <div style={{ fontSize:11, color:'#4ade80', fontFamily:'DM Mono,monospace', marginTop:6 }}>✓ Device Identifier (DI): {udiDi}</div>}
            <div style={{ fontSize:10, color:'#2a3a5a', marginTop:4 }}>Auto-parses DI from UDI. Manual entry available below.</div>
          </div>

          {/* Catalog */}
          <div style={s.fg}>
            <div style={s.lbl}>Select Device Type</div>
            <div style={{ display:'flex', gap:6, flexWrap:'wrap', marginBottom:10 }}>
              {[['plastic','✨ Plastic Surgery'],['ortho','🦴 Orthopedic'],['cardiac','💓 Cardiac'],['other','Other'],['all','All']].map(([k,l]) => (
                <button key={k} onClick={() => setCatFilter(k)}
                  style={{ ...s.catBtn, ...(catFilter===k ? s.catBtnOn : {}) }}>{l}</button>
              ))}
            </div>
            <div style={s.catalogGrid}>
              {filteredCatalog.map(item => (
                <button key={item.snomed} onClick={() => { setDeviceName(item.name); setSnomedCode(item.snomed) }}
                  style={{ ...s.catalogItem, ...(snomedCode===item.snomed ? s.catalogItemOn : {}) }}>
                  {item.name}
                  <div style={{ fontSize:9, fontFamily:'DM Mono,monospace', opacity:0.35, marginTop:2 }}>{item.snomed}</div>
                </button>
              ))}
            </div>
          </div>

          {/* Device name */}
          <div style={s.fg}>
            <label style={s.lbl}>Device Name *</label>
            <input value={deviceName} onChange={e => setDeviceName(e.target.value)}
              placeholder="e.g. Silicone breast implant"
              style={s.inp} />
          </div>

          <div style={s.row}>
            <div style={s.fg}>
              <label style={s.lbl}>Manufacturer</label>
              <input value={manufacturer} onChange={e => setManufacturer(e.target.value)}
                placeholder="e.g. Allergan, Mentor"
                style={s.inp} />
            </div>
            <div style={s.fg}>
              <label style={s.lbl}>Model Number</label>
              <input value={modelNumber} onChange={e => setModelNumber(e.target.value)}
                placeholder="Catalog / model #"
                style={{ ...s.inp, fontFamily:'DM Mono,monospace' }} />
            </div>
            <div style={s.fg}>
              <label style={s.lbl}>Lot Number</label>
              <input value={lotNumber} onChange={e => setLotNumber(e.target.value)}
                style={{ ...s.inp, fontFamily:'DM Mono,monospace' }} />
            </div>
            <div style={s.fg}>
              <label style={s.lbl}>Serial Number</label>
              <input value={serialNumber} onChange={e => setSerialNumber(e.target.value)}
                style={{ ...s.inp, fontFamily:'DM Mono,monospace' }} />
            </div>
            <div style={s.fg}>
              <label style={s.lbl}>Implant Date</label>
              <input type="date" value={implantDate} onChange={e => setImplantDate(e.target.value)}
                style={{ ...s.inp, fontFamily:'DM Mono,monospace' }} />
            </div>
            <div style={s.fg}>
              <label style={s.lbl}>Expiration Date</label>
              <input type="date" value={expirationDate} onChange={e => setExpirationDate(e.target.value)}
                style={{ ...s.inp, fontFamily:'DM Mono,monospace' }} />
            </div>
          </div>

          {/* Body site */}
          <div style={s.fg}>
            <label style={s.lbl}>Body Site</label>
            <div style={s.siteGrid}>
              {BODY_SITES.map(site => (
                <button key={site.snomed} onClick={() => { setBodySite(site.site); setBodySiteSnomed(site.snomed) }}
                  style={{ ...s.siteBtn, ...(bodySite===site.site ? s.siteBtnOn : {}) }}>
                  {site.site}
                </button>
              ))}
            </div>
            {bodySiteSnomed && <div style={{ fontSize:10, fontFamily:'DM Mono,monospace', color:'#2a3a5a' }}>SNOMED: {bodySiteSnomed}</div>}
          </div>

          {/* DI manual if not parsed */}
          {!udiDi && (
            <div style={s.fg}>
              <label style={s.lbl}>Device Identifier (DI) — manual</label>
              <input value={udiDi} onChange={e => setUdiDi(e.target.value)}
                placeholder="14-digit GS1 DI"
                style={{ ...s.inp, fontFamily:'DM Mono,monospace' }} />
            </div>
          )}

          {/* Notes */}
          <div style={s.fg}>
            <label style={s.lbl}>Clinical Notes</label>
            <textarea value={notes} onChange={e => setNotes(e.target.value)} rows={3}
              placeholder="Fill volume, texture, profile, operative details…"
              style={s.ta} />
          </div>

          <div style={{ display:'flex', gap:10 }}>
            <button onClick={() => setSubView('list')}
              style={{ ...s.btnSm }}>
              Cancel
            </button>
            <button onClick={handleAdd} disabled={!deviceName || saving}
              style={!deviceName || saving ? s.btnOff : s.btn}>
              {saving ? 'Recording…' : 'Record Device →'}
            </button>
          </div>
        </div>
      )}
    </div>
  )
}

// ── Main App ───────────────────────────────────────────────────────────────
export default function RevelaApp() {
  const [token]    = useState(() => new URLSearchParams(location.search).get('token') ?? '')
  const [sess,  setSess]   = useState<Record<string, unknown> | null>(null)
  const [enc,   setEnc]    = useState<Record<string, unknown> | null>(null)
  const [intake,setIntake] = useState<Record<string, unknown> | null>(null)
  const [meds,  setMeds]   = useState<Record<string, unknown>[]>([])
  const [tab,   setTab]    = useState<'chart'|'prognote'|'opnote'|'postop'|'devices'>('chart')
  const [saving,setSaving] = useState(false)
  const [saved, setSaved]  = useState(false)
  const [error, setError]  = useState('')
  const [loading,setLoading]=useState(true)
  const [pn, setPn] = useState({ pain_level:0, progress_notes:'', plan:'' })
  const [on, setOn] = useState({ procedure_name:'', anesthesia:'General', duration_mins:60, complications:'None', notes:'' })
  const [po, setPo] = useState({ follow_up_date:'', wound_status:'', instructions:'' })

  useEffect(() => {
    if (!token) { setError('No session token. Access from PatientTracForge scheduling.'); setLoading(false); return }
    bridge('get_encounter', { token }).then(d => {
      if (d.error) { setError(d.error); setLoading(false); return }
      setSess(d.session); setEnc(d.encounter); setIntake(d.intake); setMeds(d.medications ?? [])
      if (d.encounter?.progress_notes) setPn((p:typeof pn) => ({...p, progress_notes: d.encounter.progress_notes, pain_level: d.encounter.pain_level ?? 0}))
      if (d.encounter?.procedure_name) setOn((o:typeof on) => ({...o, procedure_name: d.encounter.procedure_name, notes: d.encounter.notes ?? ''}))
      setLoading(false)
    }).catch((e: unknown) => { setError(String(e)); setLoading(false) })
  }, [token])

  async function save(type: string, data: Record<string, unknown>) {
    setSaving(true); setSaved(false); setError('')
    const r = await bridge('save_note', { token, note_type: type, note_data: data })
    setSaving(false)
    if (r.success) { setSaved(true); setTimeout(() => setSaved(false), 3000) }
    else setError(r.error ?? 'Save failed')
  }

  async function closeEnc() {
    if (!confirm('Close encounter? Billing AI will auto-run.')) return
    await bridge('callback', { token, event_type:'encounter_closed', payload:{closed_by:'revela'} })
    alert('Encounter closed. Return to scheduling.'); window.close()
  }

  if (loading) return <div style={s.shell}><div style={s.center}><div style={s.spin}/><div style={{...s.muted,marginTop:12}}>Loading encounter...</div></div></div>
  if (error && !sess) return <div style={s.shell}><div style={s.errBox}><div style={{fontSize:18,fontWeight:600,color:'#ff6b6b',marginBottom:8}}>⚠ Access Error</div><div style={s.cbdy}>{error}</div><div style={{...s.muted,marginTop:12,fontSize:12}}>Access this app from PatientTracForge → Check-In → Route to Revela.</div></div></div>

  const pat  = sess?.patient  as Record<string, unknown>
  const prov = sess?.provider as Record<string, unknown>
  const patientId = Number(sess?.patient_id || 0)

  return (
    <div style={s.shell}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}} * { box-sizing:border-box } select,input,textarea{color:#e8eaf0 !important}`}</style>

      <div style={s.hdr}>
        <div style={{display:'flex',alignItems:'center',gap:16}}>
          <div style={s.badge}>✦ REVELA</div>
          <div>
            <div style={s.pname}>{pat?.last_name as string}, {pat?.first_name as string}</div>
            <div style={s.meta}>DOB: {pat?.dob ? new Date(pat.dob as string).toLocaleDateString() : '—'} · {pat?.gender as string} · #{sess?.patient_id as number}</div>
          </div>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:10}}>
          <div style={{fontSize:11,color:'#8a9bc0',fontFamily:'DM Mono,monospace'}}>Dr. {prov?.last_name as string} · {prov?.specialty as string}</div>
          <div style={s.encTag}>Enc #{sess?.encounter_id as number}</div>
          <div style={{width:8,height:8,borderRadius:'50%',background:(enc?.status as string)==='open'?'#4ade80':'#f59e0b'}}/>
          <button onClick={closeEnc} style={s.closeBtn}>Close Encounter</button>
        </div>
      </div>

      <div style={s.tabs}>
        {([
          ['chart',   'Patient Chart'],
          ['prognote','Progress Note'],
          ['opnote',  'Operative Note'],
          ['postop',  'Post-Op Plan'],
          ['devices', '🔧 Implant Devices'],
        ] as const).map(([k,l]) => (
          <button key={k} onClick={() => setTab(k)} style={{...s.tab,...(tab===k?s.tabOn:{})}}>{l}</button>
        ))}
      </div>

      <div style={s.body}>
        {error && <div style={s.err}>{error}</div>}
        {saved && <div style={s.ok}>✓ Saved successfully — synced to PatientTracForge</div>}

        {tab==='chart' && (
          <div style={s.grid}>
            <div style={s.card}><div style={s.cttl}>Chief Complaint</div><div style={s.cbdy}>{enc?.chief_complaint as string || intake?.chief_complaint as string || '—'}</div></div>
            {intake && <div style={s.card}><div style={s.cttl}>AI Intake Summary</div><div style={s.cbdy}>{intake.ai_summary as string||'—'}</div>{intake.triage_level&&<div style={{marginTop:8,display:'inline-block',padding:'3px 8px',fontSize:11,background:'rgba(251,191,36,0.15)',color:'#fbbf24'}}>Triage: {intake.triage_level as string}</div>}</div>}
            <div style={s.card}><div style={s.cttl}>Allergies</div><div style={s.cbdy}>{pat?.allergies as string||'NKDA'}</div></div>
            <div style={s.card}><div style={s.cttl}>Current Medications ({meds.length})</div>{meds.length===0?<div style={s.muted}>No current medications</div>:meds.map((m,i)=><div key={i} style={{fontSize:13,color:'#8a9bc0',padding:'4px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}}><strong style={{color:'#e8eaf0'}}>{m.medication_name as string}</strong> {m.dosage as string} {m.route as string} — {m.frequency as string}</div>)}</div>
            <div style={s.card}><div style={s.cttl}>Contact</div><div style={s.cbdy}>{pat?.phone as string}<br/>{pat?.email as string}</div></div>
            <div style={s.card}><div style={s.cttl}>Provider / NPI</div><div style={s.cbdy}>Dr. {prov?.first_name as string} {prov?.last_name as string}, {prov?.credential as string}<br/><span style={{fontFamily:'DM Mono,monospace',fontSize:12}}>NPI {prov?.npi as string}</span></div></div>
          </div>
        )}

        {tab==='prognote' && (
          <div style={s.form}>
            <div style={s.fg}>
              <label style={s.lbl}>Pain Level (0–10)</label>
              <div style={{display:'flex',alignItems:'center',gap:12}}>
                <input type="range" min={0} max={10} value={pn.pain_level} onChange={e=>setPn(p=>({...p,pain_level:+e.target.value}))} style={{flex:1}}/>
                <span style={{background:'rgba(201,169,110,0.15)',color:'#c9a96e',padding:'4px 14px',fontFamily:'DM Mono,monospace',fontSize:14,minWidth:40,textAlign:'center'}}>{pn.pain_level}</span>
              </div>
            </div>
            <div style={s.fg}><label style={s.lbl}>Progress Notes *</label><textarea rows={8} value={pn.progress_notes} onChange={e=>setPn(p=>({...p,progress_notes:e.target.value}))} placeholder="Subjective findings, wound status, healing progress, objective assessment..." style={s.ta}/></div>
            <div style={s.fg}><label style={s.lbl}>Plan / Next Steps</label><textarea rows={4} value={pn.plan} onChange={e=>setPn(p=>({...p,plan:e.target.value}))} placeholder="Follow-up plan, prescription changes, referrals..." style={s.ta}/></div>
            <button disabled={saving||!pn.progress_notes} onClick={()=>save('surgical_prognote', pn as unknown as Record<string,unknown>)} style={saving||!pn.progress_notes?s.btnOff:s.btn}>{saving?'Saving...':'Save Progress Note → Scheduling'}</button>
          </div>
        )}

        {tab==='opnote' && (
          <div style={s.form}>
            <div style={s.row}>
              <div style={s.fg}><label style={s.lbl}>Procedure Name *</label><input value={on.procedure_name} onChange={e=>setOn(o=>({...o,procedure_name:e.target.value}))} placeholder="e.g. Augmentation mammaplasty" style={s.inp}/></div>
              <div style={s.fg}><label style={s.lbl}>Anesthesia</label><select value={on.anesthesia} onChange={e=>setOn(o=>({...o,anesthesia:e.target.value}))} style={s.inp}><option>General</option><option>Regional</option><option>Local + Sedation</option><option>Local</option></select></div>
            </div>
            <div style={s.row}>
              <div style={s.fg}><label style={s.lbl}>Duration (minutes)</label><input type="number" value={on.duration_mins} onChange={e=>setOn(o=>({...o,duration_mins:+e.target.value}))} style={s.inp}/></div>
              <div style={s.fg}><label style={s.lbl}>Complications</label><input value={on.complications} onChange={e=>setOn(o=>({...o,complications:e.target.value}))} placeholder="None / describe" style={s.inp}/></div>
            </div>
            <div style={s.fg}><label style={s.lbl}>Operative Notes *</label><textarea rows={10} value={on.notes} onChange={e=>setOn(o=>({...o,notes:e.target.value}))} placeholder="Detailed intraoperative findings, technique, closure method, implant specifics if applicable..." style={s.ta}/></div>
            <button disabled={saving||!on.procedure_name||!on.notes} onClick={()=>save('operative_notes', on as unknown as Record<string,unknown>)} style={saving||!on.procedure_name||!on.notes?s.btnOff:s.btn}>{saving?'Saving...':'Save Operative Note → Scheduling'}</button>
          </div>
        )}

        {tab==='postop' && (
          <div style={s.form}>
            <div style={s.row}>
              <div style={s.fg}><label style={s.lbl}>Follow-up Date</label><input type="date" value={po.follow_up_date} onChange={e=>setPo(p=>({...p,follow_up_date:e.target.value}))} style={s.inp}/></div>
              <div style={s.fg}><label style={s.lbl}>Wound / Healing Status</label><select value={po.wound_status} onChange={e=>setPo(p=>({...p,wound_status:e.target.value}))} style={s.inp}><option value="">Select...</option><option>Healing well — no concerns</option><option>Minor swelling — expected</option><option>Signs of infection — monitor</option><option>Requires intervention</option></select></div>
            </div>
            <div style={s.fg}><label style={s.lbl}>Patient Instructions *</label><textarea rows={8} value={po.instructions} onChange={e=>setPo(p=>({...p,instructions:e.target.value}))} placeholder="Activity restrictions, wound care, medications, warning signs, when to return to ER..." style={s.ta}/></div>
            <button disabled={saving||!po.instructions} onClick={()=>save('operative_notes', {...on, postop_plan:po} as unknown as Record<string,unknown>)} style={saving||!po.instructions?s.btnOff:s.btn}>{saving?'Saving...':'Save Post-Op Plan → Scheduling'}</button>
          </div>
        )}

        {tab==='devices' && patientId > 0 && (
          <DevicesTab patientId={patientId} token={token} />
        )}
      </div>
    </div>
  )
}
