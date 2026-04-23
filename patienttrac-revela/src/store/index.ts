import { create } from 'zustand'

export interface Flag {
  id: string
  severity: 'critical' | 'warning' | 'info'
  category: string
  title: string
  description: string
  recommendedAction: string
  dismissed: boolean
}

interface RevelaStore {
  currentUser: { name: string; role: string } | null
  setCurrentUser: (user: { name: string; role: string } | null) => void

  selectedPatientId: string | null
  setSelectedPatient: (id: string | null) => void

  flags: Flag[]
  dismissFlag: (flagId: string) => void
  addFlags: (flags: Flag[]) => void

  sidebarOpen: boolean
  toggleSidebar: () => void

  activeModule: string
  setActiveModule: (module: string) => void
}

export const useRevelaStore = create<RevelaStore>((set) => ({
  currentUser: { name: 'Dr. Williams', role: 'surgeon' },
  setCurrentUser: (user) => set({ currentUser: user }),

  selectedPatientId: null,
  setSelectedPatient: (id) => set({ selectedPatientId: id }),

  flags: [],
  dismissFlag: (flagId) =>
    set((state) => ({
      flags: state.flags.map((f) =>
        f.id === flagId ? { ...f, dismissed: true } : f
      ),
    })),
  addFlags: (newFlags) =>
    set((state) => ({ flags: [...state.flags, ...newFlags] })),

  sidebarOpen: true,
  toggleSidebar: () =>
    set((state) => ({ sidebarOpen: !state.sidebarOpen })),

  activeModule: 'dashboard',
  setActiveModule: (module) => set({ activeModule: module }),
}))
