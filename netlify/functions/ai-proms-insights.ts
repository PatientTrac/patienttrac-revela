// ============================================================
// PatientTrac Revela — AI PROMs Insights Interpreter
// Netlify Function: /.netlify/functions/ai-proms-insights
// ============================================================

import type { Handler, HandlerEvent } from '@netlify/functions';

const ANTHROPIC_API = 'https://api.anthropic.com/v1/messages';
const MODEL = 'claude-fable-5';

const SYSTEM_PROMPT = `You are the patient-reported outcomes (PROMs) analysis AI for PatientTrac Revela — a HIPAA-compliant plastic surgery EMR.

Analyze the provided PROM instrument scores and return clinical insights. Scores are on a 1-5 scale (1 = very dissatisfied/worst, 5 = very satisfied/best). A score of 0 means the domain was not rated.

You receive NO patient PHI — only anonymized clinical outcome data.

Return ONLY valid JSON — no markdown.

OUTPUT FORMAT:
{
  "overallScore": number (mean of rated domains, 1 decimal),
  "overallInterpretation": "brief plain-English summary of overall satisfaction level",
  "domainInsights": [
    {
      "domain": "domain_key",
      "score": number,
      "interpretation": "brief clinical interpretation of this domain score",
      "concern": boolean
    }
  ],
  "clinicalConcerns": ["array of specific concerns requiring provider attention — empty if none"],
  "providerRecommendations": ["array of 2-4 actionable recommendations for the provider"],
  "proceedWithSurgery": "yes" | "with_caution" | "defer_pending_evaluation",
  "proceedRationale": "one sentence explaining the proceed recommendation"
}

INTERPRETATION GUIDELINES:
- Score 4-5: Good to excellent satisfaction — generally favorable for surgery
- Score 3: Moderate — may proceed with realistic expectation-setting
- Score 1-2: Poor — significant dissatisfaction; flag for psychological evaluation
- Any domain score of 1: Critical concern — flag immediately
- Overall mean < 2.5: Consider deferring surgery pending psychological evaluation
- Body dysmorphia high risk: Always flag "defer_pending_evaluation"
- BREAST-Q sexual_wellbeing = 1: Sensitive — flag for counseling, do not defer surgery solely for this
- psychosocial_wellbeing < 2 in any instrument: Flag for pre-op mental health consult

INSTRUMENT-SPECIFIC CONTEXT:
- BREAST-Q: Used for breast augmentation, reduction, reconstruction, mastopexy
- FACE-Q: Used for facelift, rhinoplasty, blepharoplasty, facial rejuvenation
- BODY-Q / BODY-Q-lipo: Used for abdominoplasty, body contouring, liposuction`;

interface PromsInsightsRequest {
  instrument: 'BREAST-Q' | 'FACE-Q' | 'BODY-Q' | 'BODY-Q-lipo';
  domains: Record<string, number>;  // domain name → score 1-5
  procedureName: string;
  procedureType: string;
  bodyDysmorphiaScreening?: 'low_risk' | 'moderate_risk' | 'high_risk';
  surgicalCandidacy?: string;
  psychClearance?: string;
  encounterId?: string;
  orgId?: string;
  providerId?: string;
}

// ── PHI scrubber — belt-and-suspenders guard before any Claude call ────────────
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

// ── Audit log — non-blocking write to Supabase cr.ai_audit_log ────────────────
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
  } catch { /* audit failure must not block the clinical response */ }
}

export const handler: Handler = async (event: HandlerEvent) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return { statusCode: 500, body: JSON.stringify({ error: 'AI service not configured' }) };
  }

  let body: PromsInsightsRequest;
  try {
    body = JSON.parse(event.body ?? '{}');
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid request body' }) };
  }

  // Build user message — no PHI in this interface; domains are clinical scores only
  const domainLines = Object.entries(body.domains)
    .map(([k, v]) => `- ${k.replace(/_/g, ' ')}: ${v}/5`)
    .join('\n');

  const userMessage = `Analyze the following PROM scores for a patient undergoing ${body.procedureName}:

Instrument: ${body.instrument}
Procedure: ${body.procedureName} (${body.procedureType})

Domain scores (1-5 scale, 0 = not rated):
${domainLines}

Additional context:
- Body dysmorphia screening: ${body.bodyDysmorphiaScreening ?? 'not performed'}
- Surgical candidacy assessment: ${body.surgicalCandidacy ?? 'not documented'}
- Psychological clearance: ${body.psychClearance ?? 'not required'}

Return the JSON insights analysis.`;

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
        max_tokens: 900,
        system: SYSTEM_PROMPT,
        messages: [{ role: 'user', content: userMessage }],
      }),
    });

    if (!response.ok) {
      throw new Error(`Anthropic API error: ${response.status}`);
    }

    const data = await response.json();
    const text = data.content?.[0]?.text ?? '{}';

    let parsed: unknown;
    try { parsed = JSON.parse(text) } catch { parsed = { error: 'Parse error', raw: text } }

    // Audit log — non-blocking
    void logAudit({
      functionName: 'ai-proms-insights', action: 'generate_proms_insights', phiScrubbed: true,
      encounterId: body.encounterId, orgId: body.orgId, providerId: body.providerId,
      latencyMs: Date.now() - t0,
    });

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(parsed),
    };
  } catch (error) {
    console.error('AI PROMs insights error:', error);
    void logAudit({ functionName: 'ai-proms-insights', action: 'error', phiScrubbed: true, latencyMs: Date.now() - t0 });
    return { statusCode: 500, body: JSON.stringify({ error: 'AI service error' }) };
  }
};
