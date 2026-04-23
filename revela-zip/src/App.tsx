import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom'
import IntakeModule from '@/components/modules/intake/IntakeModule'
import SurgeonIPad from '@/components/modules/surgeon/SurgeonIPad'

export default function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<Navigate to="/intake" replace />} />
        <Route path="/intake" element={<IntakeModule />} />
        <Route path="/surgeon" element={<SurgeonIPad />} />
      </Routes>
    </BrowserRouter>
  )
}
