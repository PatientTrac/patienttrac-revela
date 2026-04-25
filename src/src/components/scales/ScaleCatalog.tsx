import { useState, useEffect } from 'react'
import { callScales, ScaleMeta } from './RatingScalesModule'

const CATEGORIES = [
  { id: 'all',           label: 'All Scales' },
  { id: 'depression',    label: 'Depression' },
  { id: 'anxiety',       label: 'Anxiety' },
  { id: 'suicide_risk',  label: 'Suicide Risk' },
  { id: 'trauma',        label: 'Trauma / PTSD' },
  { id: 'bipolar',       label: 'Bipolar' },
  { id: 'psychosis',     label: 'Psychosis' },
  { id: 'adhd',          label: 'ADHD' },
  { id: 'substance_use', label: 'Substance Use' },
  { id: 'cognitive',     label: 'Cognitive' },
  { id: 'global',        label: 'Global / CGI' },
]

const CATEGORY_COLORS: Record<string, string> = {
  depression:    '#c9a96e',
  anxiety:       '#00d4ff',
  suicide_risk:  '#ff3b3b',
  trauma:        '#9b59b6',
  bipolar:       '#e67e22',
  psychosis:     '#e74c3c',
  adhd:          '#2ecc71',
  substance_use: '#1abc9c',
  cognitive:     '#3498db',
  global:        '#95a5a6',
}

interface Props {
  orgId: string
  onSelect: (scale: ScaleMeta) => void
}

export default function ScaleCatalog({ orgId, onSelect }: Props) {
  const [scales, setScales] = useState<ScaleMeta[]>([])
  const [loading, setLoading] = useState(true)
  const [category, setCategory] = useState('all')
  const [search, setSearch] = useState('')

  useEffect(() => {
    setLoading(true)
    callScales({ action: 'get_catalog', org_id: orgId })
      .then(d => setScales(d.scales || []))
      .finally(() => setLoading(false))
  }, [orgId])

  const filtered = scales.filter(s => {
    const matchCat = category === 'all' || s.category === category
    const matchSearch = !search ||
      s.scale_code.toLowerCase().includes(search.toLowerCase()) ||
      s.scale_name.toLowerCase().includes(search.toLowerCase()) ||
      s.scale_full_name?.toLowerCase().includes(search.toLowerCase()) ||
      s.description?.toLowerCase().includes(search.toLowerCase())
    return matchCat && matchSearch
  })

  return (
    <div>
      {/* Search */}
      <input
        type="text"
        placeholder="Search scales..."
        value={search}
        onChange={e => setSearch(e.target.value)}
        className="w-full mb-4 px-3 py-2 rounded text-sm"
        style={{ background: 'rgba(0,212,255,0.05)', border: '1px solid rgba(0,212,255,0.2)', color: 'rgba(200,215,235,0.9)', outline: 'none' }}
      />

      {/* Category pills */}
      <div className="flex flex-wrap gap-2 mb-5">
        {CATEGORIES.map(c => (
          <button key={c.id} onClick={() => setCategory(c.id)}
            className="px-3 py-1 rounded text-xs transition-all"
            style={{
              fontFamily: 'Rajdhani', fontWeight: 600, letterSpacing: '0.05em',
              background: category === c.id ? 'rgba(0,212,255,0.15)' : 'rgba(0,212,255,0.04)',
              border: category === c.id ? '1px solid rgba(0,212,255,0.5)' : '1px solid rgba(0,212,255,0.12)',
              color: category === c.id ? '#00d4ff' : 'rgba(200,215,235,0.6)',
            }}>
            {c.label}
          </button>
        ))}
      </div>

      {loading ? (
        <div className="text-center py-10 hud-label animate-pulse">Loading scales...</div>
      ) : (
        <div className="grid grid-cols-1 gap-3">
          {filtered.map(scale => (
            <div key={scale.scale_code}
              className="p-4 rounded cursor-pointer transition-all hover:border-opacity-60"
              style={{ background: 'rgba(10,22,40,0.7)', border: '1px solid rgba(0,212,255,0.12)' }}
              onClick={() => !scale.license_required && onSelect(scale)}>
              <div className="flex items-start justify-between gap-4">
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 mb-1 flex-wrap">
                    {/* Category badge */}
                    <span className="px-2 py-0.5 rounded text-xs font-bold"
                      style={{
                        background: `${CATEGORY_COLORS[scale.category] || '#666'}22`,
                        color: CATEGORY_COLORS[scale.category] || '#999',
                        border: `1px solid ${CATEGORY_COLORS[scale.category] || '#666'}44`,
                        fontFamily: 'Rajdhani',
                      }}>
                      {scale.category.replace('_', ' ').toUpperCase()}
                    </span>
                    {/* Scale name */}
                    <span className="font-bold" style={{ fontFamily: 'Rajdhani', fontSize: '1rem', color: '#c9a96e' }}>
                      {scale.scale_code}
                    </span>
                    <span className="text-sm opacity-60">{scale.scale_full_name || scale.scale_name}</span>
                    {/* Licensed badge */}
                    {scale.license_required && (
                      <span className="px-2 py-0.5 rounded text-xs"
                        style={{ background: 'rgba(255,59,59,0.1)', border: '1px solid rgba(255,59,59,0.3)', color: '#ff3b3b', fontFamily: 'Rajdhani' }}>
                        LICENSE REQUIRED
                      </span>
                    )}
                  </div>
                  <p className="text-xs mb-2 leading-relaxed" style={{ color: 'rgba(200,215,235,0.55)' }}>
                    {scale.description}
                  </p>
                  {scale.license_required && scale.license_note && (
                    <p className="text-xs" style={{ color: 'rgba(255,59,59,0.7)' }}>⚠️ {scale.license_note}</p>
                  )}
                </div>

                {/* Meta */}
                <div className="flex-shrink-0 text-right space-y-1">
                  <div className="text-xs" style={{ color: 'rgba(200,215,235,0.4)' }}>
                    {scale.total_items} items · {scale.estimated_minutes} min
                  </div>
                  <div className="text-xs" style={{ color: 'rgba(200,215,235,0.4)' }}>
                    Score 0–{scale.max_score}
                  </div>
                  {scale.cpt_code && (
                    <div className="text-xs px-2 py-0.5 rounded"
                      style={{ background: 'rgba(0,212,255,0.06)', color: 'rgba(0,212,255,0.6)', border: '1px solid rgba(0,212,255,0.1)' }}>
                      CPT {scale.cpt_code}
                    </div>
                  )}
                  {!scale.license_required && (
                    <button
                      onClick={e => { e.stopPropagation(); onSelect(scale) }}
                      className="text-xs px-3 py-1 rounded transition-all mt-1"
                      style={{ background: 'rgba(201,169,110,0.12)', border: '1px solid rgba(201,169,110,0.4)', color: '#c9a96e', fontFamily: 'Rajdhani', fontWeight: 700 }}>
                      Administer →
                    </button>
                  )}
                </div>
              </div>
            </div>
          ))}

          {filtered.length === 0 && (
            <div className="text-center py-8 opacity-40 text-sm">No scales match your search.</div>
          )}
        </div>
      )}
    </div>
  )
}
