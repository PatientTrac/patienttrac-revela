import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom'
import { QueryClient, QueryClientProvider } from '@tanstack/react-query'
import Layout from './components/layout/Layout'
import Dashboard from './components/modules/Dashboard'
import IntakeModule from './components/modules/intake/IntakeModule'
import ORScheduling from './components/modules/or-scheduling/ORScheduling'
import Proposals from './components/modules/proposals/Proposals'
import Inventory from './components/modules/inventory/Inventory'
import Accounting from './components/modules/accounting/Accounting'
import Staff from './components/modules/staff/Staff'

const queryClient = new QueryClient({
  defaultOptions: {
    queries: { staleTime: 1000 * 60 * 5, retry: 1 },
  },
})

export default function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <BrowserRouter>
        <Routes>
          <Route path="/" element={<Layout />}>
            <Route index element={<Navigate to="/dashboard" replace />} />
            <Route path="dashboard" element={<Dashboard />} />
            <Route path="intake" element={<IntakeModule />} />
            <Route path="intake/:patientId" element={<IntakeModule />} />
            <Route path="or-scheduling" element={<ORScheduling />} />
            <Route path="proposals" element={<Proposals />} />
            <Route path="inventory" element={<Inventory />} />
            <Route path="accounting" element={<Accounting />} />
            <Route path="staff" element={<Staff />} />
          </Route>
        </Routes>
      </BrowserRouter>
    </QueryClientProvider>
  )
}
