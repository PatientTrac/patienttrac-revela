import React from 'react'
import ReactDOM from 'react-dom/client'
import { BrowserRouter } from 'react-router-dom'
import { QueryClient, QueryClientProvider } from '@tanstack/react-query'
import App from './App'
import './styles/index.css'
import { supabase } from './lib/supabase'
import { initAudit } from './lib/audit'
import { useAppStore } from './lib/store'
import { registerAppStore } from '../../packages/ui/src/store'

// ── Shared package store registration ────────────────────────
// Must run before any @patienttrac/ui component is rendered
registerAppStore(useAppStore)

// ── Session → store hydration ─────────────────────────────────
function hydrateStore(session: { user: { id: string; email?: string; user_metadata?: Record<string, unknown> }; access_token: string } | null) {
  if (!session) return
  const meta = session.user.user_metadata || {}
  useAppStore.getState().setSession({
    org_id:       (meta.org_id as string)      || '00000000-0000-0000-0000-000000000001',
    provider_id:  (meta.provider_id as string) || session.user.id,
    access_token: session.access_token,
  })
}

const queryClient = new QueryClient({
  defaultOptions: { queries: { staleTime: 1000 * 60 * 5, retry: 1 } },
})

const queryClient = new QueryClient({
  defaultOptions: { queries: { staleTime: 1000 * 60 * 5, retry: 1 } },
})

// ── §170.315(d)(2) Audit Initialization ──────────────────────
// PatientTrac Revela — plastic surgery clinical module
async function setupAudit() {
  const { data: { session } } = await supabase.auth.getSession()
  hydrateStore(session)
  if (session?.user) {
    const meta = session.user.user_metadata || {}
    initAudit({
      org_id:     (meta.org_id as string)   || '00000000-0000-0000-0000-000000000001',
      user_id:    session.user.id,
      user_email: session.user.email        || '',
      user_role:  (meta.role as string)     || 'provider',
      session_id: session.access_token.slice(-16),
      app_source: 'revela',
    })
  } else {
    // Cross-app bridge token from PatientTracForge
    const params    = new URLSearchParams(window.location.search)
    const orgId     = params.get('org_id')      || '00000000-0000-0000-0000-000000000001'
    const providerId = params.get('provider_id') || ''
    useAppStore.getState().setSession({ org_id: orgId, provider_id: providerId })
    initAudit({ org_id: orgId, user_id: providerId, user_role: 'provider', app_source: 'revela' })
  }
}

setupAudit()

supabase.auth.onAuthStateChange((_event, session) => {
  hydrateStore(session)
  if (session?.user) {
    const meta = session.user.user_metadata || {}
    initAudit({
      org_id:     (meta.org_id as string)   || '00000000-0000-0000-0000-000000000001',
      user_id:    session.user.id,
      user_email: session.user.email        || '',
      user_role:  (meta.role as string)     || 'provider',
      session_id: session.access_token.slice(-16),
      app_source: 'revela',
    })
  }
})

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <QueryClientProvider client={queryClient}>
      <BrowserRouter>
        <App />
      </BrowserRouter>
    </QueryClientProvider>
  </React.StrictMode>
)
