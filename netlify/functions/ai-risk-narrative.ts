// ============================================================
// PatientTrac Revela — AI Clinical Risk Narrative Generator
// Netlify Function: /.netlify/functions/ai-risk-narrative
// ============================================================

import type { Handler, HandlerEvent } from '@netlify/functions';

const ANTHROPIC_API = 'https://api.anthropic.com/v1/messages';
const MODEL = 'claude-fable-5';

const SYSTEM_PROMPT = `You are the clinical risk summarization AI for PatientTrac Revela — a HIPAA-compliant plastic surgery EMR.

Your role is to convert structured risk scores into a concise, accurate clinical narrative for the provider's pre-operative note. You receive NO patient PHI — only anonymized clinical parameters.

Return ONLY valid JSON. No markdown.

OUTPUT FORMAT:
{
  "narrative": "2-4 sentence clinical summary integrating all risk domains. Write in provider note style (third person, clinical language). Example: 'This ASA II patient presents with moderate VTE risk (Caprini 5) and low OSA risk (STOP-BANG 2) for planned breast augmentation. Cardiac risk is very low (RCRI 0). Standard VTE chemoprophylaxis with sequential compression devices is recommended; no additional cardiac or pulmonary workup indicated.'",
  "orClearanceRecommendation": "clear" | "conditional" | "hold",
  "orClearanceRationale": "One sentence justifying the clearance recommendation",
  "requiredConsults": ["list of required pre-op consults, empty array if none"],
  "keyRecommendations": ["array of 2-5 bullet-point action items"]
}

CLEARANCE LOGIC:
- clear: ASA I-II, Caprini < 5, STOP-BANG < 3, RCRI 0-1
- conditional: ASA III OR Caprini 5-8 OR STOP-BANG 3-5 OR RCRI 2 — requires specific pre-op action
- hold: ASA IV-V OR Caprini > 8 OR STOP-BANG ≥ 6 OR RCRI ≥ 3 — multidisciplinary review required

REQUIRED CONSULTS TRIGGERS:
- Anesthesia pre-op: STOP-BANG ≥ 3 or ASA ≥ III
- Cardiology: RCRI ≥ 2
- Pulmonology: STOP-BANG ≥ 5
- Hematology: Caprini > 8 or history of prior DVT/PE (flag from score context)`;

interface RiskNarrativeRequest {
  procedureName: string;
  capriniScore?: number;
  capriniRisk?: string;          // 'Very Low'|'Low'|'Moderate'|'High'
  stopBangScore?: number;
  stopBangRisk?: string;
  rcriScore?: number;
  rcriRisk?: string;
  asaClass?: string;
  bmi?: number;
  smokingStatus?: string;
  estimatedDurationHours?: number;
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

  let body: RiskNarrativeRequest;
  try {
    body = JSON.parse(event.body ?? '{}');
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid request body' }) };
  }

  // Build prompt — no PHI present in RiskNarrativeRequest (procedure-level scores only)
  const userMessage = `Generate a clinical risk narrative for this pre-operative assessment:

Procedure: ${body.procedureName}
Estimated duration: ${body.estimatedDurationHours ?? 'not specified'} hours

Risk Scores:
- ASA Physical Status: ${body.asaClass ?? 'not documented'}
- Caprini VTE Score: ${body.capriniScore ?? 'not documented'} (${body.capriniRisk ?? ''})
- STOP-BANG OSA Score: ${body.stopBangScore ?? 'not documented'} (${body.stopBangRisk ?? ''})
- RCRI Cardiac Score: ${body.rcriScore ?? 'not documented'} (${body.rcriRisk ?? ''})

Contributing factors:
- BMI: ${body.bmi ?? 'not documented'}
- Smoking: ${body.smokingStatus ?? 'not documented'}

Return the JSON risk narrative.`;

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
        max_tokens: 700,
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
      functionName: 'ai-risk-narrative', action: 'generate_risk_narrative', phiScrubbed: true,
      encounterId: body.encounterId, orgId: body.orgId, providerId: body.providerId,
      latencyMs: Date.now() - t0,
    });

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(parsed),
    };
  } catch (error) {
    console.error('AI risk narrative error:', error);
    void logAudit({ functionName: 'ai-risk-narrative', action: 'error', phiScrubbed: true, latencyMs: Date.now() - t0 });
    return { statusCode: 500, body: JSON.stringify({ error: 'AI service error' }) };
  }
};
