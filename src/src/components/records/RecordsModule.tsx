import { useState } from 'react'
import RequestsList from './RequestsList'
import NewRequestForm from './NewRequestForm'
import ReportBuilder from './ReportBuilder'
import ThirdPartyPortal from './ThirdPartyPortal'
import type { RecordsRequest, GeneratedReport } from './types'

interface Props {
  patientId: string
  encounterId?: string
  providerId: string
  orgId: string
  appSource: 'mind' | 'revela'
  patientName: string
  providerName: string
  providerCreds: string
  providerNpi?: string
}

type Tab = 'requests' | 'new_request' | 'build_report' | 'third_party'

export default function RecordsModule(props: Props) {
  const [activeTab, setActiveTab] = useState<Tab>('requests')
  const [selectedRequest, setSelectedRequest] = useState<RecordsRequest | null>(null)
  const [selectedReport, setSelectedReport] = useState<GeneratedReport | null>(null)

  const tabs: { id: Tab; label: string; icon: string }[] = [
    { id: 'requests', label: 'Requests', icon: '📋' },
    { id: 'new_request', label: 'New Request', icon: '➕' },
    { id: 'build_report', label: 'Build Report', icon: '📄' },
    { id: 'third_party', label: 'Third-Party Access', icon: '🔗' },
  ]

  return (
    <div style={{ fontFamily: 'DM Sans, sans-serif' }}>
      {/* Module Header */}
      <div className="flex items-center justify-between px-6 py-4 border-b"
        style={{ borderColor: 'rgba(0,212,255,0.15)', background: 'rgba(6,14,28,0.6)' }}>
        <div className="flex items-center gap-3">
          <span className="text-2xl">🗂️</span>
          <div>
            <div className="hud-title text-lg" style={{ color: '#c9a96e' }}>
              Records & Reports
            </div>
            <div className="hud-label">{props.patientName}</div>
          </div>
        </div>
        <div className="hud-label px-3 py-1 rounded"
          style={{ background: 'rgba(0,212,255,0.08)', border: '1px solid rgba(0,212,255,0.2)', color: '#00d4ff' }}>
          {props.appSource === 'mind' ? 'PatientTrac Mind' : 'Revela'}
        </div>
      </div>

      {/* Tab nav */}
      <div className="flex border-b px-4"
        style={{ borderColor: 'rgba(0,212,255,0.12)', background: 'rgba(10,22,40,0.7)' }}>
        {tabs.map((tab) => {
          const active = activeTab === tab.id
          return (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={`flex items-center gap-2 px-4 py-3 text-xs transition-all ${active ? 'tab-active' : 'tab-inactive'}`}
              style={{ fontFamily: 'Rajdhani', fontWeight: 600, letterSpacing: '0.08em', textTransform: 'uppercase' }}
            >
              <span>{tab.icon}</span>
              {tab.label}
            </button>
          )
        })}
      </div>

      {/* Tab content */}
      <div className="p-5 animate-fadeIn">
        {activeTab === 'requests' && (
          <RequestsList
            {...props}
            onSelectRequest={(r) => { setSelectedRequest(r); setActiveTab('build_report') }}
          />
        )}
        {activeTab === 'new_request' && (
          <NewRequestForm
            {...props}
            onCreated={() => setActiveTab('requests')}
          />
        )}
        {activeTab === 'build_report' && (
          <ReportBuilder
            {...props}
            initialRequest={selectedRequest}
            onReportGenerated={(r) => setSelectedReport(r)}
          />
        )}
        {activeTab === 'third_party' && (
          <ThirdPartyPortal
            {...props}
            selectedReport={selectedReport}
          />
        )}
      </div>
    </div>
  )
}
