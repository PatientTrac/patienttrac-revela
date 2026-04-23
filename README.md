# PatientTrac Revela
### Plastic & Reconstructive Surgery EMR Module

> *Where precision meets possibility.*

Revela is the AI-powered clinical intelligence platform built exclusively for plastic and reconstructive surgery practices — combining AI-guided patient intake, smart documentation, OR scheduling, surgical proposal generation, inventory management, and full practice accounting in one elegant system.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Frontend | React 18, TypeScript, Vite |
| Styling | Tailwind CSS + Revela CSS Design Tokens |
| AI Engine | Anthropic Claude API (claude-sonnet-4-20250514) |
| Backend/API | Netlify Serverless Functions |
| Database | PostgreSQL (Supabase) with Row-Level Security |
| Auth | JWT + TOTP MFA, httpOnly cookies |
| Hosting | Netlify (CI/CD via GitHub Actions) |
| Testing | Vitest, Playwright, React Testing Library |

---

## Branch Structure

```
main          →  Production (auto-deploys to Netlify)
develop       →  Staging / integration branch
feature/*     →  Individual feature branches (linked to Jira stories)
hotfix/*      →  Emergency production fixes
```

**Branch naming convention:** `feature/SCRUM-{number}-short-description`
Example: `feature/SCRUM-24-ai-patient-intake`

---

## Getting Started

### Prerequisites
- Node.js 20+
- npm 10+
- PostgreSQL 15+ (or Supabase account)

### Setup

```bash
# Clone the repository
git clone https://github.com/YOUR_ORG/patienttrac-revela.git
cd patienttrac-revela

# Install dependencies
npm install

# Copy environment variables
cp .env.example .env.local

# Fill in your .env.local values (see Environment Variables section)

# Start development server
npm run dev
```

Open [http://localhost:5173](http://localhost:5173) to view the app.

---

## Environment Variables

```env
# Anthropic AI
ANTHROPIC_API_KEY=sk-ant-...

# Database (Supabase)
SUPABASE_URL=https://your-project.supabase.co
SUPABASE_ANON_KEY=eyJ...
SUPABASE_SERVICE_ROLE_KEY=eyJ...

# Authentication
JWT_SECRET=your-jwt-secret-min-64-chars
JWT_REFRESH_SECRET=your-refresh-secret-min-64-chars
SESSION_TIMEOUT_MINUTES=15

# App
VITE_APP_NAME=PatientTrac Revela
VITE_APP_ENV=development
VITE_API_BASE_URL=http://localhost:8888/.netlify/functions

# Email (for MFA, notifications)
SMTP_HOST=
SMTP_PORT=587
SMTP_USER=
SMTP_PASS=
```

> ⚠️ **Never commit `.env.local` or any file containing real credentials to version control.**

---

## Project Structure

```
patienttrac-revela/
├── src/
│   ├── components/
│   │   ├── ui/              # Base UI components (Button, Card, Badge, Input)
│   │   ├── layout/          # AppShell, Sidebar, Header, NavBar
│   │   ├── dashboard/       # DashboardPage, KPICards, ORStatusGrid, FlagPanel
│   │   ├── patients/        # PatientList, PatientRecord, PatientQueue
│   │   ├── intake/          # AIInterview, IntakeSection, ClinicalFlag
│   │   ├── or/              # ORCalendar, BookingModal, StaffAssignment
│   │   ├── proposals/       # ProposalBuilder, CostingEngine, PDFExport
│   │   ├── inventory/       # InventoryDashboard, StockAlert, POGenerator
│   │   ├── staff/           # StaffList, ScheduleBuilder, Credentials
│   │   └── accounting/      # FinancialDashboard, PLReport, ARaging
│   ├── pages/
│   │   ├── DashboardPage.tsx
│   │   ├── PatientsPage.tsx
│   │   ├── IntakePage.tsx
│   │   ├── ORSchedulePage.tsx
│   │   ├── ProposalsPage.tsx
│   │   ├── InventoryPage.tsx
│   │   ├── AccountingPage.tsx
│   │   ├── StaffPage.tsx
│   │   └── LoginPage.tsx
│   ├── api/
│   │   ├── anthropic.ts     # Claude API wrapper
│   │   ├── patients.ts      # Patient record API calls
│   │   ├── or.ts            # OR scheduling API
│   │   ├── proposals.ts     # Proposal generation API
│   │   └── auth.ts          # Auth API calls
│   ├── hooks/
│   │   ├── useAI.ts         # AI query hook with loading/error states
│   │   ├── usePatient.ts    # Patient data hook
│   │   ├── useAuth.ts       # Auth state hook
│   │   └── useOR.ts         # OR scheduling hook
│   ├── styles/
│   │   ├── tokens.css       # Revela design tokens (colors, typography)
│   │   └── globals.css      # Global reset + base styles
│   ├── types/
│   │   ├── patient.ts       # Patient type definitions
│   │   ├── or.ts            # OR / procedure types
│   │   └── auth.ts          # Auth / user role types
│   ├── utils/
│   │   ├── crypto.ts        # Field-level encryption helpers
│   │   ├── audit.ts         # Audit log writer
│   │   └── format.ts        # Date, currency, MRN formatters
│   └── main.tsx
├── netlify/
│   └── functions/
│       ├── ai-intake.ts     # AI interview + flag generation
│       ├── ai-proposal.ts   # Surgical proposal generation
│       ├── ai-summary.ts    # Physician briefing summary
│       └── auth.ts          # Auth serverless handler
├── public/
│   └── revela-logo.svg
├── docs/
│   ├── architecture.md      # System architecture overview
│   ├── hipaa-controls.md    # HIPAA compliance controls map
│   ├── api-reference.md     # Internal API documentation
│   └── sprint-planning.md   # Sprint history and velocity
├── tests/
│   ├── unit/                # Vitest unit tests
│   ├── e2e/                 # Playwright end-to-end tests
│   └── components/          # React Testing Library tests
├── .github/
│   └── workflows/
│       ├── ci.yml           # Run tests on PR
│       └── deploy.yml       # Deploy to Netlify on main push
├── .env.example
├── .eslintrc.json
├── .prettierrc
├── netlify.toml
├── vite.config.ts
├── tailwind.config.ts
├── tsconfig.json
└── package.json
```

---

## Revela Design Tokens

```css
/* Core brand colors — imported globally from src/styles/tokens.css */
--revela-navy:         #0a1628;   /* Primary brand background */
--revela-surgical-blue: #1a3a5c; /* Secondary / depth */
--revela-gold:         #c9a96e;  /* Accent / luxury / CTAs */
--revela-ivory:        #f0ece4;  /* Light text / warm white bg */
--revela-steel-teal:   #4a7fa8;  /* AI indicators / data UI */

/* Semantic */
--revela-flag-red:     #e11d48;  /* Critical clinical alert */
--revela-flag-amber:   #d97706;  /* Warning / review needed */
--revela-flag-blue:    #3b82f6;  /* Informational / planning note */
--revela-flag-green:   #16a34a;  /* AI pre-filled / confirmed */
```

---

## Sprint 1 — Stories (Days 1–30)

| Jira | Story | Priority |
|---|---|---|
| SCRUM-21 | React + Vite setup with Revela design tokens | High |
| SCRUM-22 | Physician auth, MFA & RBAC | High |
| SCRUM-23 | Main dashboard — patient queue, OR status, flags | High |
| SCRUM-24 | AI patient intake interview — 6-section engine | High |
| SCRUM-25 | Patient record data model & MRN search | High |
| SCRUM-26 | OR scheduling calendar — booking & conflict detection | High |
| SCRUM-27 | AI surgical proposal generator & PDF export | Medium |
| SCRUM-28 | HIPAA backend — DB, encryption, audit logging | High |
| SCRUM-29 | Surgical inventory tracking & AI purchase orders | Medium |
| SCRUM-30 | Testing framework — Vitest, Playwright, CI pipeline | Medium |

---

## HIPAA Compliance

Revela is architected for HIPAA compliance from day one:

- All PHI encrypted at rest (AES-256) and in transit (TLS 1.3)
- Audit log on every PHI read/write event (user, action, timestamp, IP)
- Session auto-timeout at 15 minutes idle
- MFA required for all user accounts
- No PHI in URL parameters or browser localStorage
- API rate limiting and IP allowlisting available
- BAA-ready architecture documentation in `/docs/hipaa-controls.md`

---

## MVP Target: 90 Days

| Phase | Days | Focus |
|---|---|---|
| Sprint 1 | 1–30 | Foundation, auth, dashboard, AI intake, patient records, OR scheduling |
| Sprint 2 | 31–60 | Proposals, accounting, inventory, staff management, compliance |
| Sprint 3 | 61–90 | QA, E2E testing, performance, beta practice onboarding, launch |

---

## Contributing

1. Create a branch from `develop`: `git checkout -b feature/SCRUM-{number}-description`
2. Write tests alongside your code
3. Run `npm run lint && npm run test` before pushing
4. Open a PR into `develop` — CI must pass before merge
5. Squash merge to keep history clean

---

## License

Proprietary — PatientTrac / AegisIQ. All rights reserved.
