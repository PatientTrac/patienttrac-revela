# PatientTrac Revela — Surgical Drawing Tool

## Files
- `src/components/SurgicalDrawingTool.tsx` → drop into `patienttrac-revela/src/components/`
- `024_surgical_photos_drawings.sql` → already applied to Supabase (migration 024)

## Setup
1. Install Fabric.js: `npm install fabric`
2. Create Supabase Storage buckets:
   - `patient-photos` (private)
   - `revela-templates` (public)
3. Download SVG templates from smart.servier.com and upload to `revela-templates` bucket
4. Import component: `import SurgicalDrawingTool from '../components/SurgicalDrawingTool'`

## Usage
```tsx
<SurgicalDrawingTool
  encounterId={encounter.id}
  patientId={patient.id}
  procedureType="breast"
  onSave={(drawingId) => console.log('Saved:', drawingId)}
/>
```

## Template keys
breast | body_female | face_female | pectoral | body_male | face_male

## Attribution (required by CC BY 4.0)
"Illustrations by Servier Medical Art (smart.servier.com), CC BY 4.0"
Already included in component footer.
