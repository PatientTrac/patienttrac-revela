# PatientTrac Revela — 4 New Features

## Files to deploy → patienttrac-revela/src/components/

1. RevelaLogin.tsx         — Email + password + Google Authenticator TOTP only
2. NewPatientRevela.tsx    — Create new patients directly from Revela
3. RevelaAIAssistant.tsx   — HUD AI Avatar demo + Interactive AI assistant

## How to integrate in App.tsx

```tsx
import RevelaLogin from './components/RevelaLogin'
import NewPatientRevela from './components/NewPatientRevela'
import RevelaAIAssistant from './components/RevelaAIAssistant'

// In your App component:
// 1. If no token and not authenticated → show RevelaLogin
// 2. Add + New Patient button → opens NewPatientRevela modal
// 3. Always render RevelaAIAssistant (fixed position, bottom-right)

<RevelaAIAssistant
  orgId={orgId}
  providerId={userId}
  providerEmail={userEmail}
/>

<NewPatientRevela
  orgId={orgId}
  providerId={userId}
  onPatientCreated={(pt) => console.log('Created:', pt)}
  onClose={() => setShowNewPatient(false)}
/>
```

## Features

### 1. RevelaLogin
- Email + password + Google Authenticator TOTP
- First login: QR code appears automatically → scan once → done
- Every login after: email + password + 6-digit code
- NO Google OAuth, NO SSO

### 2. NewPatientRevela
- Full patient registration form inside Revela
- Demographics, contact, emergency contact, clinical notes, insurance
- Writes directly to cr.patient in Supabase
- Tags patient as created_source = 'revela'

### 3. RevelaAIAssistant (HUD Avatar + Demo)

#### Demo Mode
- Click "▶ Start Demo" button (bottom-right)
- 10-step guided tour with voice narration
- Highlights each EMR feature with description
- Progress bar, step dots, skip/next controls
- Text-to-speech using browser Web Speech API

#### AI Assistant
- Click the diamond avatar button (bottom-right)
- Ask questions by TEXT or VOICE (🎤 button)
- Select feedback type: question / suggestion / feature_request / bug
- All interactions saved to cr.revela_ai_feedback in Supabase
- Suggestions automatically emailed to wayne@patienttracforge.com

## Database
Migration 025 already applied:
- Table: cr.revela_ai_feedback
- Fields: org_id, provider_id, provider_email, feedback_type, 
          input_method, content, ai_response, status, email_sent
- RLS: org isolation enforced

## Review Dashboard
Query cr.revela_ai_feedback in Supabase dashboard:
SELECT * FROM cr.revela_ai_feedback ORDER BY created_at DESC;
