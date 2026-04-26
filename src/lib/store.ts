// patienttrac-revela/src/lib/store.ts
// Minimal zustand store satisfying the @patienttrac/ui AppStore interface.
// Revela's clinical state lives in VB6 hooks (useVB6Forms.ts) — this store
// exists solely to provide session/context to shared package components
// (MedicationsModule, AllergyModule, DiagnosisModule, Compliance, etc.)

import { create } from 'zustand'
import type { AppSession, AppContext } from '../../../packages/ui/src/types'

interface RevelaStore {
  session: AppSession | null
  context: AppContext | null

  // Setters — called from App.tsx or auth listeners
  setSession: (s: AppSession) => void
  setContext:  (c: AppContext)  => void
  clearSession: () => void
}

export const useAppStore = create<RevelaStore>((set) => ({
  session: null,
  context: null,
  setSession:   (s) => set({ session: s }),
  setContext:   (c) => set({ context: c }),
  clearSession: ()  => set({ session: null, context: null }),
}))
