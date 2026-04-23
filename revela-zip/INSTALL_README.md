# PatientTrac Revela — File Installation Guide

## What's in this zip

```
src/
  lib/
    riskEngine.ts        ← AI risk intelligence engine (all 12 risk categories)
    supabase.ts          ← Supabase client + typed database helpers
  components/
    modules/
      intake/
        IntakeModule.tsx ← Full AI intake form with live risk assessment
      surgeon/
        SurgeonIPad.tsx  ← iPad surgeon interface (drawing, voice, camera)
  App.tsx                ← Main React router (connects all routes)
index.html               ← Futuristic HUD homepage (replaces existing)
.env.local               ← Environment variables (Supabase keys pre-filled)
```

## How to install

1. Extract this zip to your Desktop
2. Copy ALL files into: `C:\Users\wayne\patienttrac-revela\`
   - When asked "Replace existing files?" — click YES for index.html and App.tsx
   - All src/ files are NEW — just copy them in

3. Open Git Bash and run:
```bash
cd /c/Users/wayne/patienttrac-revela
npm install @supabase/supabase-js react-router-dom
git add .
git commit -m "feat: AI intake module + risk engine + surgeon iPad + Supabase client"
git push https://YOUR_TOKEN@github.com/PatientTrac/patienttrac-revela.git master
```

## Routes after deploy

- `/`          → Redirects to intake
- `/intake`    → AI Patient Intake with live risk assessment
- `/surgeon`   → iPad surgeon interface

## Supabase connection

Already configured. Project: mskormozwekezjmtcylv (us-east-1)
68 tables live across cr and terms schemas.
