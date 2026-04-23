# PatientTrac Revela — HIPAA Compliance Controls

**Last updated:** April 23, 2026  
**Status:** Architecture review — pre-BAA

---

## Technical Safeguards (45 CFR § 164.312)

### Access Controls
| Control | Implementation | Status |
|---|---|---|
| Unique user identification | Each user has UUID + email login | ✅ Planned |
| Emergency access procedure | Break-glass admin role documented | ✅ Planned |
| Automatic logoff | 15-minute idle session timeout | ✅ Planned |
| Encryption of ePHI | AES-256 at rest, TLS 1.3 in transit | ✅ Planned |

### Audit Controls
| Control | Implementation | Status |
|---|---|---|
| Audit log — read events | Every patient record read logged | ✅ Planned |
| Audit log — write events | All create/update/delete on PHI logged | ✅ Planned |
| Audit log — auth events | Login, logout, failed login logged | ✅ Planned |
| Log retention | 6 years minimum (HIPAA requirement) | ✅ Planned |
| Log integrity | Append-only table, no delete permission | ✅ Planned |

### Integrity Controls
| Control | Implementation | Status |
|---|---|---|
| ePHI integrity | Row-level checksums on patient records | ⬜ Planned |
| Transmission integrity | TLS 1.3, HTTPS enforced via HSTS | ✅ Planned |

### Transmission Security
| Control | Implementation | Status |
|---|---|---|
| TLS 1.3 | Enforced on all endpoints via Netlify | ✅ Planned |
| HSTS | max-age=63072000, includeSubDomains | ✅ Planned |
| No PHI in URLs | All PHI passed in request body only | ✅ Planned |
| No PHI in localStorage | httpOnly cookies only for auth tokens | ✅ Planned |

---

## Administrative Safeguards (45 CFR § 164.308)

| Requirement | Action | Owner |
|---|---|---|
| Security Officer designation | Appoint HIPAA Security Officer | Legal/HR |
| Risk Analysis | Conduct formal risk assessment | Security Officer |
| Workforce training | Annual HIPAA training for all staff | HR |
| BAA with Anthropic | Execute BAA before PHI sent to AI | Legal |
| BAA with Supabase/Netlify | Execute BAA with each vendor | Legal |
| Incident response plan | Document breach notification procedure | Security Officer |
| Contingency plan | Disaster recovery and backup plan | Engineering |

---

## Physical Safeguards (45 CFR § 164.310)

Revela is a cloud-native SaaS application. Physical safeguards are inherited from:
- **Netlify** (hosting) — SOC 2 Type II certified
- **Supabase** (database) — hosted on AWS, SOC 2 certified
- **Anthropic** (AI) — BAA required before PHI transmission

---

## AI & PHI Policy

Revela uses the Anthropic Claude API for AI-powered intake, proposals, and summaries.

**Before go-live, the following must be completed:**
1. Execute a Business Associate Agreement (BAA) with Anthropic
2. Confirm Anthropic's HIPAA-eligible API tier is activated
3. Implement field-level consent flags — PHI may only be sent to AI after patient consent is recorded
4. Implement PHI de-identification option for AI calls where full PHI is not required
5. Audit log every API call that contains PHI

**PHI fields that may be sent to AI (with consent):**
- Patient age range (not exact DOB)
- Diagnosis category
- Procedure type
- Surgical history summary
- Medication list (de-identified where possible)

**PHI fields that must NEVER be sent to AI:**
- Full name
- Date of birth (exact)
- MRN
- Social Security Number
- Insurance member ID
- Address, phone, email

---

## Breach Notification Requirements

Under HIPAA Breach Notification Rule (45 CFR §§ 164.400–414):

- **Individual notification:** Within 60 days of breach discovery
- **HHS notification:** Within 60 days (or annual report if <500 individuals)
- **Media notification:** If breach affects >500 residents of a state

**Breach response contacts:**
- Security Officer: [TBD]
- Legal counsel: [TBD]
- HHS Office for Civil Rights: https://www.hhs.gov/ocr/

---

## Pre-Launch HIPAA Checklist

- [ ] BAA executed with Anthropic
- [ ] BAA executed with Supabase
- [ ] BAA executed with Netlify
- [ ] Risk analysis completed and documented
- [ ] HIPAA Security Officer appointed
- [ ] All staff completed HIPAA training
- [ ] Penetration test completed by qualified third party
- [ ] Audit logging tested end-to-end
- [ ] Session timeout tested (15 minutes)
- [ ] Encryption verified at rest and in transit
- [ ] Incident response plan documented
- [ ] PHI consent workflow implemented and tested
- [ ] AI PHI policy implemented and enforced in code
