import { useState, useEffect, useRef, useCallback } from 'react'
import { callMeds, DrugResult, Medication } from './MedicationsModule'

interface Props {
  patientId: number
  encounterId: number | undefined
  providerId: number
  orgId: string
  currentMeds: Medication[]
  onAdded: () => void
  onCancel: () => void
}

interface StrengthOption {
  rxcui: string
  name: string
  strength: string | null
  dose_form: string | null
  tty: string
}

const FREQUENCIES = [
  'Once daily (QD)', 'Twice daily (BID)', 'Three times daily (TID)',
  'Four times daily (QID)', 'Every morning (QAM)', 'Every night (QHS)',
  'Every 4 hours (Q4H)', 'Every 6 hours (Q6H)', 'Every 8 hours (Q8H)',
  'Every 12 hours (Q12H)', 'As needed (PRN)', 'Weekly', 'Biweekly', 'Monthly',
]

const ROUTES = [
  'Oral', 'Sublingual', 'Buccal', 'Topical', 'Transdermal', 'Inhaled',
  'Intranasal', 'Intravenous', 'Intramuscular', 'Subcutaneous', 'Rectal', 'Ophthalmic',
]

const DEA_COLORS: Record<string, string> = {
  'II': '#ff3b3b', 'III': '#e67e22', 'IV': '#f39c12', 'V': '#f1c40f',
}

export default function MedSearch({ patientId, encounterId, providerId, orgId, currentMeds, onAdded, onCancel }: Props) {
  const [query, setQuery] = useState('')
  const [results, setResults] = useState<DrugResult[]>([])
  const [suggestions, setSuggestions] = useState<string[]>([])
  const [searching, setSearching] = useState(false)
  const [selected, setSelected] = useState<DrugResult | null>(null)
  const [drugDetail, setDrugDetail] = useState<DrugResult | null>(null)
  const [strengths, setStrengths] = useState<StrengthOption[]>([])
  const [selectedStrength, setSelectedStrength] = useState<StrengthOption | null>(null)
  const [showDropdown, setShowDropdown] = useState(false)
  const [loadingDetail, setLoadingDetail] = useState(false)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [error, setError] = useState('')
  const [epcsConfirmed, setEpcsConfirmed] = useState(false)

  // Dosing fields
  const [frequency, setFrequency] = useState('')
  const [route, setRoute] = useState('')
  const [startDate, setStartDate] = useState(new Date().toISOString().split('T')[0])
  const [notes, setNotes] = useState('')

  const searchRef = useRef<HTMLDivElement>(null)
  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  // Debounced search
  useEffect(() => {
    if (query.length < 2) { setResults([]); setSuggestions([]); setShowDropdown(false); return }
    if (debounceRef.current) clearTimeout(debounceRef.current)
    debounceRef.current = setTimeout(async () => {
      setSearching(true)
      try {
        const d = await callMeds({ action: 'search', query })
        setResults(d.results || [])
        setSuggestions(d.suggestions || [])
        setShowDropdown(true)
      } finally {
        setSearching(false)
      }
    }, 350)
    return () => { if (debounceRef.current) clearTimeout(debounceRef.current) }
  }, [query])

  // Click outside to close dropdown
  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (searchRef.current && !searchRef.current.contains(e.target as Node)) setShowDropdown(false)
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [])

  const handleSelect = useCallback(async (drug: DrugResult) => {
    setSelected(drug)
    setShowDropdown(false)
    setQuery(drug.drug_name)
    setLoadingDetail(true)
    setEpcsConfirmed(false)
    try {
      // Get full drug detail
      const detail = await callMeds({ action: 'get_drug', rxcui: drug.rxcui })
      setDrugDetail(detail.drug || drug)
      // Auto-fill route from detail
      if (detail.drug?.route) setRoute(detail.drug.route)

      // Get available strengths for this ingredient
      const strengthData = await callMeds({ action: 'get_strengths', rxcui: drug.rxcui })
      const opts = (strengthData.products || []) as StrengthOption[]
      setStrengths(opts.filter((s: StrengthOption) => s.strength || s.dose_form))
      // Auto-select if only one
      if (opts.length === 1) setSelectedStrength(opts[0])
    } finally {
      setLoadingDetail(false)
    }
  }, [])

  const isControlled = drugDetail?.is_controlled || selected?.is_controlled
  const deaSchedule = drugDetail?.dea_schedule || selected?.dea_schedule
  const canSave = selected && (selectedStrength || drugDetail) && frequency && route &&
    (!isControlled || epcsConfirmed)

  // Check if already prescribed
  const alreadyPrescribed = currentMeds.some(m =>
    m.rxnorm_cui === selected?.rxcui ||
    m.drug_name?.toLowerCase().includes((selected?.base_name || '').toLowerCase())
  )

  const handleSave = async () => {
    if (!canSave || !selected) return
    setSaving(true)
    setError('')
    try {
      const targetDrug = selectedStrength
        ? { ...drugDetail, rxcui: selectedStrength.rxcui || selected.rxcui, drug_name: selectedStrength.name || selected.drug_name, strength: selectedStrength.strength, dose_form: selectedStrength.dose_form }
        : drugDetail || selected

      const d = await callMeds({
        action: 'save_medication',
        patient_id: patientId,
        encounter_id: encounterId,
        provider_id: providerId,
        org_id: orgId,
        medication: {
          rxcui: targetDrug.rxcui,
          drug_name: targetDrug.drug_name,
          dose_form: targetDrug.dose_form || null,
          strength: targetDrug.strength || null,
          dose: targetDrug.strength || null,
          route,
          frequency,
          drug_class: drugDetail?.drug_class || null,
          start_date: startDate,
          notes: notes || null,
        }
      })
      if (!d.success) throw new Error(d.error || 'Save failed')
      setSaved(true)
      setTimeout(() => onAdded(), 1000)
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Failed to save')
    } finally {
      setSaving(false)
    }
  }

  if (saved) return (
    <div className="text-center py-16 animate-fadeIn">
      <div style={{ fontSize: '3rem', marginBottom: 12 }}>✅</div>
      <div style={{ fontFamily: 'Rajdhani', fontSize: '1.25rem', color: '#27ae60', fontWeight: 700 }}>
        Medication saved
      </div>
      <div className="text-sm opacity-50 mt-1">Returning to medication list...</div>
    </div>
  )

  return (
    <div className="max-w-2xl mx-auto">
      <div className="flex items-center justify-between mb-5">
        <div style={{ fontFamily: 'Rajdhani', fontSize: '1rem', fontWeight: 700, color: '#c9a96e' }}>
          Add Medication
        </div>
        <button onClick={onCancel} className="text-xs opacity-40 hover:opacity-70">✕ Cancel</button>
      </div>

      {/* ── SEARCH BOX ─────────────────────────────────────── */}
      <div ref={searchRef} className="relative mb-5">
        <div className="relative">
          <input
            type="text"
            placeholder="Search by drug name, brand, or ingredient..."
            value={query}
            onChange={e => { setQuery(e.target.value); setSelected(null); setDrugDetail(null); setStrengths([]); setSelectedStrength(null) }}
            onFocus={() => results.length > 0 && setShowDropdown(true)}
            className="w-full px-4 py-3 pr-10 rounded text-sm"
            style={{
              background: 'rgba(0,212,255,0.05)',
              border: '1px solid rgba(0,212,255,0.25)',
              color: 'rgba(200,215,235,0.9)',
              outline: 'none',
              fontFamily: 'DM Sans',
            }}
          />
          {searching && (
            <div className="absolute right-3 top-3 animate-spin text-sm opacity-40">⟳</div>
          )}
          {!searching && query.length >= 2 && (
            <div className="absolute right-3 top-3 text-xs opacity-30">RxNorm</div>
          )}
        </div>

        {/* Dropdown results */}
        {showDropdown && (results.length > 0 || suggestions.length > 0) && (
          <div className="absolute z-50 w-full mt-1 rounded shadow-lg overflow-hidden"
            style={{ background: 'rgba(8,18,34,0.98)', border: '1px solid rgba(0,212,255,0.2)', maxHeight: 320, overflowY: 'auto' }}>

            {/* Spelling suggestions */}
            {suggestions.length > 0 && results.length === 0 && (
              <div className="px-3 py-2 text-xs opacity-50">
                Did you mean:&nbsp;
                {suggestions.map((s, i) => (
                  <button key={i} onClick={() => setQuery(s)}
                    className="underline mr-2 opacity-70 hover:opacity-100">{s}</button>
                ))}
              </div>
            )}

            {results.map((drug, i) => (
              <button key={`${drug.rxcui}-${i}`}
                onClick={() => handleSelect(drug)}
                className="w-full text-left px-4 py-2.5 transition-all hover:bg-white/5 border-b last:border-b-0"
                style={{ borderColor: 'rgba(255,255,255,0.04)' }}>
                <div className="flex items-center justify-between gap-2">
                  <div className="min-w-0">
                    <div className="text-sm truncate" style={{ color: 'rgba(200,215,235,0.9)' }}>
                      {drug.drug_name}
                    </div>
                    {drug.base_name && drug.base_name !== drug.drug_name && (
                      <div className="text-xs opacity-40">{drug.base_name}</div>
                    )}
                  </div>
                  <div className="flex items-center gap-1.5 flex-shrink-0">
                    {drug.is_controlled && (
                      <span className="text-xs px-1.5 py-0.5 rounded font-bold"
                        style={{ background: `${DEA_COLORS[drug.dea_schedule || 'IV']}20`, color: DEA_COLORS[drug.dea_schedule || 'IV'], border: `1px solid ${DEA_COLORS[drug.dea_schedule || 'IV']}40`, fontFamily: 'Rajdhani' }}>
                        C-{drug.dea_schedule}
                      </span>
                    )}
                    {drug.tty && (
                      <span className="text-xs opacity-25 font-mono">{drug.tty}</span>
                    )}
                  </div>
                </div>
              </button>
            ))}
          </div>
        )}
      </div>

      {/* ── SELECTED DRUG DETAIL ──────────────────────────── */}
      {selected && (
        <div className="animate-fadeIn">
          {loadingDetail ? (
            <div className="py-6 text-center hud-label animate-pulse">Loading drug information...</div>
          ) : (
            <>
              {/* Drug info card */}
              <div className="p-4 rounded mb-4"
                style={{ background: 'rgba(10,22,40,0.8)', border: '1px solid rgba(0,212,255,0.15)' }}>
                <div className="flex items-start justify-between gap-3">
                  <div>
                    <div className="font-bold mb-1"
                      style={{ fontFamily: 'Rajdhani', fontSize: '1.1rem', color: '#c9a96e' }}>
                      {drugDetail?.base_name || selected.base_name || selected.drug_name}
                    </div>
                    {drugDetail?.brand_names && drugDetail.brand_names.length > 0 && (
                      <div className="text-xs opacity-40 mb-1">
                        Brand: {drugDetail.brand_names.join(', ')}
                      </div>
                    )}
                    {drugDetail?.drug_class && drugDetail.drug_class.length > 0 && (
                      <div className="text-xs opacity-50">{drugDetail.drug_class[0]}</div>
                    )}
                  </div>
                  <div className="text-right">
                    <div className="text-xs font-mono opacity-30 mb-1">RxCUI {selected.rxcui}</div>
                    {isControlled && (
                      <div className="px-2 py-1 rounded text-xs font-bold"
                        style={{
                          background: `${DEA_COLORS[deaSchedule || 'IV']}18`,
                          border: `1px solid ${DEA_COLORS[deaSchedule || 'IV']}50`,
                          color: DEA_COLORS[deaSchedule || 'IV'],
                          fontFamily: 'Rajdhani',
                        }}>
                        ⚠ Schedule {deaSchedule} Controlled Substance
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* Already prescribed warning */}
              {alreadyPrescribed && (
                <div className="p-3 rounded mb-4 text-sm"
                  style={{ background: 'rgba(231,76,60,0.1)', border: '1px solid rgba(231,76,60,0.3)', color: '#e74c3c' }}>
                  ⚠ This patient is already prescribed a medication with a similar ingredient. Review for duplicate therapy.
                </div>
              )}

              {/* ── STRENGTH SELECTOR ─────────────────────── */}
              {strengths.length > 0 && (
                <div className="mb-4">
                  <div className="hud-label text-xs mb-2">Select Strength / Formulation</div>
                  <div className="grid grid-cols-2 gap-2 max-h-48 overflow-y-auto pr-1">
                    {strengths.map((s, i) => {
                      const isSelected = selectedStrength?.rxcui === s.rxcui && selectedStrength?.name === s.name
                      return (
                        <button key={`${s.rxcui}-${i}`}
                          onClick={() => setSelectedStrength(s)}
                          className="p-2.5 rounded text-left text-xs transition-all"
                          style={{
                            background: isSelected ? 'rgba(201,169,110,0.15)' : 'rgba(255,255,255,0.03)',
                            border: isSelected ? '1px solid rgba(201,169,110,0.5)' : '1px solid rgba(255,255,255,0.07)',
                            color: isSelected ? '#c9a96e' : 'rgba(200,215,235,0.6)',
                          }}>
                          <div className="font-semibold">{s.strength || '—'}</div>
                          <div className="opacity-60">{s.dose_form || s.name}</div>
                        </button>
                      )
                    })}
                  </div>
                </div>
              )}

              {/* ── DOSING FIELDS ─────────────────────────── */}
              <div className="grid grid-cols-2 gap-3 mb-4">
                {/* Frequency */}
                <div>
                  <div className="hud-label text-xs mb-1.5">Frequency <span className="text-red-400">*</span></div>
                  <select value={frequency} onChange={e => setFrequency(e.target.value)}
                    className="w-full px-3 py-2 rounded text-xs"
                    style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.85)', outline: 'none' }}>
                    <option value="">Select frequency...</option>
                    {FREQUENCIES.map(f => <option key={f} value={f}>{f}</option>)}
                  </select>
                </div>

                {/* Route */}
                <div>
                  <div className="hud-label text-xs mb-1.5">Route <span className="text-red-400">*</span></div>
                  <select value={route} onChange={e => setRoute(e.target.value)}
                    className="w-full px-3 py-2 rounded text-xs"
                    style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.85)', outline: 'none' }}>
                    <option value="">Select route...</option>
                    {ROUTES.map(r => <option key={r} value={r}>{r}</option>)}
                  </select>
                </div>

                {/* Start date */}
                <div>
                  <div className="hud-label text-xs mb-1.5">Start Date</div>
                  <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)}
                    className="w-full px-3 py-2 rounded text-xs"
                    style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.85)', outline: 'none' }} />
                </div>

                {/* Notes */}
                <div>
                  <div className="hud-label text-xs mb-1.5">Notes</div>
                  <input type="text" value={notes} onChange={e => setNotes(e.target.value)}
                    placeholder="e.g. titrate up, take with food..."
                    className="w-full px-3 py-2 rounded text-xs"
                    style={{ background: 'rgba(0,212,255,0.04)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.85)', outline: 'none' }} />
                </div>
              </div>

              {/* ── EPCS GATE for controlled substances ───── */}
              {isControlled && (
                <div className="p-4 rounded mb-4"
                  style={{ background: 'rgba(255,59,59,0.07)', border: '1px solid rgba(255,59,59,0.35)' }}>
                  <div className="flex items-start gap-3 mb-3">
                    <span style={{ fontSize: '1.3rem' }}>🔐</span>
                    <div>
                      <div className="text-sm font-bold mb-1"
                        style={{ fontFamily: 'Rajdhani', color: '#ff3b3b', letterSpacing: '0.04em' }}>
                        Controlled Substance — DEA Schedule {deaSchedule}
                      </div>
                      <div className="text-xs leading-relaxed opacity-60">
                        DEA 21 CFR Part 1311 requires prescriber identity verification before ordering a Schedule {deaSchedule} controlled substance. Full EPCS transmission requires DrFirst or Surescripts integration.
                      </div>
                    </div>
                  </div>
                  <label className="flex items-start gap-2 cursor-pointer">
                    <input type="checkbox" checked={epcsConfirmed} onChange={e => setEpcsConfirmed(e.target.checked)}
                      className="mt-0.5 flex-shrink-0" />
                    <span className="text-xs leading-relaxed" style={{ color: 'rgba(200,215,235,0.75)' }}>
                      I confirm my identity as the prescribing provider and that this Schedule {deaSchedule} substance is medically necessary for this patient. This action will be logged to the EPCS audit trail.
                    </span>
                  </label>
                </div>
              )}

              {/* Error */}
              {error && (
                <div className="p-3 rounded mb-3 text-sm"
                  style={{ background: 'rgba(255,59,59,0.1)', color: '#ff3b3b', border: '1px solid rgba(255,59,59,0.3)' }}>
                  {error}
                </div>
              )}

              {/* Save button */}
              <div className="flex gap-3">
                <button onClick={onCancel}
                  className="px-5 py-3 rounded text-sm"
                  style={{ background: 'rgba(255,255,255,0.03)', border: '1px solid rgba(255,255,255,0.1)', color: 'rgba(200,215,235,0.4)', fontFamily: 'Rajdhani' }}>
                  Cancel
                </button>
                <button onClick={handleSave} disabled={!canSave || saving}
                  className="flex-1 py-3 rounded text-sm font-bold transition-all disabled:opacity-30"
                  style={{
                    background: canSave ? 'rgba(201,169,110,0.15)' : 'rgba(255,255,255,0.03)',
                    border: canSave ? '1px solid rgba(201,169,110,0.5)' : '1px solid rgba(255,255,255,0.08)',
                    color: canSave ? '#c9a96e' : 'rgba(200,215,235,0.3)',
                    fontFamily: 'Rajdhani', letterSpacing: '0.08em',
                  }}>
                  {saving ? 'Saving...' : `Add ${drugDetail?.base_name || selected.drug_name} →`}
                </button>
              </div>
            </>
          )}
        </div>
      )}
    </div>
  )
}
