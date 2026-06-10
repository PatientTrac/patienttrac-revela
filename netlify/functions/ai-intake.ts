// ============================================================
// PatientTrac Revela — AI Intake Clinical Flag Generator
// Netlify Function: /.netlify/functions/ai-intake
// ============================================================

import type { Handler, HandlerEvent } from '@netlify/functions';

const ANTHROPIC_API = 'https://api.anthropic.com/v1/messages';
const MODEL = 'claude-fable-5';

const SYSTEM_PROMPT = `You are the Revela AI clinical assistant for PatientTrac — a HIPAA-compliant
EMR platform for plastic and reconstructive surgery. Your role is to:

1. Analyze patient intake data for each section of the interview
2. Generate clinical flags (red/amber/blue) for issues requiring physician attention
3. Produce a concise physician briefing summary at completion

CLINICAL FLAG RULES:
- RED: Immediate safety concern — do not schedule OR without physician clearance
  Examples: BP >140/90, unresolved cardiac issues, active anticoagulation without bridge plan
- AMBER: Requires review or action before surgery
  Examples: Medication interactions, prior abdominal surgery affecting donor site, CME due
- BLUE: Planning note for surgical team
  Examples: Preferred reconstruction type, insurance coverage gaps, contralateral symmetry interest

OUTPUT FORMAT: Always respond with valid JSON only. No markdown, no preamble.
Schema: { flags: ClinicalFlag[], briefing?: string }

ClinicalFlag schema:
{
  severity: "red" | "amber" | "blue",
  category: "vitals" | "medication" | "surgical" | "anesthesia" | "psychological" | "planning",
  title: string,          // max 60 chars
  detail: string,         // max 200 chars, clinical detail
  suggestedAction: string // max 80 chars, what to do
}`;

interface IntakeRequest {
  section: 1 | 2 | 3 | 4 | 5 | 6;
  sectionData: Record<string, unknown>;
  patientContext?: Record<string, unknown>;
  generateBriefing?: boolean;
  encounterId?: string;
  orgId?: string;
  providerId?: string;
}

// ── PHI scrubber ─────────────────────────────────────────────────────────────
const PHI_KEYS = ['patient_name','first_name','last_name','name','dob','date_of_birth',
  'ssn','social_security','mrn','medical_record','address','street','city','state','zip',
  'phone','mobile','email','insurance_id','member_id','group_number','patient_id']

function scrubPHI(data: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  for (const [k, v] of Object.entries(data)) {
    if (PHI_KEYS.some(f => k.toLowerCase().includes(f))) {
      out[k] = '[PHI-SCRUBBED]'
    } else if (v && typeof v === 'object' && !Array.isArray(v)) {
      out[k] = scrubPHI(v as Record<string, unknown>)
    } else {
      out[k] = v
    }
  }
  return out
}

// ── Audit log ─────────────────────────────────────────────────────────────────
async function logAudit(opts: {
  functionName: string; action: string; encounterId?: string;
  orgId?: string; providerId?: string; latencyMs: number; phiScrubbed: boolean;
}) {
  const supabaseUrl = process.env.SUPABASE_URL
  const svcKey = process.env.SUPABASE_SERVICE_ROLE_KEY
  if (!supabaseUrl || !svcKey) return
  try {
    await fetch(`${supabaseUrl}/rest/v1/ai_audit_log`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': svcKey,
        'Authorization': `Bearer ${svcKey}`,
        'Content-Profile': 'cr',
        'Prefer': 'return=minimal',
      },
      body: JSON.stringify({
        function_name: opts.functionName,
        model_used: MODEL,
        action: opts.action,
        encounter_id: opts.encounterId ? Number(opts.encounterId) : null,
        org_id: opts.orgId ?? null,
        phi_scrubbed: opts.phiScrubbed,
        latency_ms: opts.latencyMs,
        specialty: 'plastic_surgery',
      }),
    })
  } catch { /* non-blocking */ }
}

export const handler: Handler = async (event: HandlerEvent) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return { statusCode: 500, body: JSON.stringify({ error: 'AI service not configured' }) };
  }

  let body: IntakeRequest;
  try {
    body = JSON.parse(event.body ?? '{}');
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid request body' }) };
  }

  // Scrub PHI before sending to Claude
  const cleanSection = scrubPHI(body.sectionData ?? {})
  const cleanContext = body.patientContext ? scrubPHI(body.patientContext) : {}

  const sectionLabels: Record<number, string> = {
    1: 'Oncology and mastectomy history',
    2: 'Medical history and comorbidities',
    3: 'Social history and lifestyle',
    4: 'Psychological assessment and patient goals',
    5: 'Procedure preference and surgical planning',
    6: 'Complete intake summary — generate full physician briefing',
  };

  const userMessage = body.generateBriefing
    ? `Generate a complete physician briefing paragraph based on all intake sections.
       Patient context (PHI scrubbed): ${JSON.stringify(cleanContext)}
       Final section data (PHI scrubbed): ${JSON.stringify(cleanSection)}`
    : `Analyze section ${body.section} (${sectionLabels[body.section]}) of the patient intake.
       Section data (PHI scrubbed): ${JSON.stringify(cleanSection)}
       Context from prior sections (PHI scrubbed): ${JSON.stringify(cleanContext)}
       Generate clinical flags for any issues. Return empty flags array if none.`;

  const t0 = Date.now();

  try {
    const response = await fetch(ANTHROPIC_API, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: 1200,
        system: SYSTEM_PROMPT,
        messages: [{ role: 'user', content: userMessage }],
      }),
    });

    if (!response.ok) throw new Error(`Anthropic API error: ${response.status}`);

    const data = await response.json();
    const text = data.content?.[0]?.text ?? '{}';

    let parsed: unknown;
    try { parsed = JSON.parse(text) } catch { parsed = { flags: [], briefing: text } }

    void logAudit({
      functionName: 'ai-intake', action: `section_${body.section}`, phiScrubbed: true,
      encounterId: body.encounterId, orgId: body.orgId, providerId: body.providerId,
      latencyMs: Date.now() - t0,
    });

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(parsed),
    };
  } catch (error) {
    console.error('AI intake error:', error);
    void logAudit({ functionName: 'ai-intake', action: 'error', phiScrubbed: true, latencyMs: Date.now() - t0 });
    return { statusCode: 500, body: JSON.stringify({ error: 'AI service error', flags: [] }) };
  }
};
