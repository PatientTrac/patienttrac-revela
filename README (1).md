# VB6 Migration Toolkit — PatientTrac Revela

Copy these tools into your PatientTrac Revela project root.

## Quick Start

### Step 1 — Parse a single VB6 form
```
node parse-vb6.js C:\path\to\YourForm.frm --ts
```
Outputs:
- `YourForm.parsed.json` — all controls, variables, validation rules, business rules
- `YourForm.generated.ts` — TypeScript skeleton ready to fill in

### Step 2 — Batch parse all forms
1. Create a folder called `vb6-forms/`
2. Copy all your .frm files into it
3. Run:
```
node batch-parse.js ./vb6-forms/
```
Outputs:
- `migration-manifest.json` — full inventory of all forms
- One `.parsed.json` per form
- Suggested Jira story for each form mapped to correct Epic

### Step 3 — Upload .frm files here to Claude
Paste the contents of any `.frm` file directly into chat.
Claude will read the VB6 code and produce:
- TypeScript types and interfaces
- Validation functions
- Business rule hooks
- React component structure

## What gets extracted
| VB6 element         | Converted to                        |
|---------------------|-------------------------------------|
| Form fields         | TypeScript interface                 |
| Validation subs     | validate() function                 |
| Calculate subs      | Pure calculation functions          |
| Constants           | TypeScript const exports            |
| SQL queries         | API service stubs                   |
| Event handlers      | React event handler stubs           |
| ComboBox options    | TypeScript union types / enums      |

## File locations in Revela project
| Output              | Location                            |
|---------------------|-------------------------------------|
| Domain types        | src/types/index.ts                  |
| Business rules      | src/lib/rules/{formName}.ts         |
| Validation hooks    | src/hooks/use{FormName}Validation.ts|
| API stubs           | src/services/{formName}Service.ts   |

## Jira Epic mapping (auto-assigned by batch parser)
| VB6 form keyword    | Revela module     | Jira Epic  |
|---------------------|-------------------|------------|
| patient/intake      | AI Intake         | SCRUM-6    |
| OR/schedule/surgery | OR Scheduling     | SCRUM-7    |
| proposal/cost/fee   | Proposals         | SCRUM-8    |
| inventory/supply    | Inventory         | SCRUM-9    |
| account/billing     | Accounting        | SCRUM-10   |
| staff/credential    | Staff             | SCRUM-11   |
