// ============================================================
// PatientTrac Revela — AI Operative Note Draft Generator
// Netlify Function: /.netlify/functions/ai-op-note-draft
// ============================================================

import type { Handler, HandlerEvent } from '@netlify/functions';

const ANTHROPIC_API = 'https://api.anthropic.com/v1/messages';
const MODEL = 'claude-fable-5';

const SYSTEM_PROMPT = `You are the operative note drafting AI for PatientTrac Revela — a HIPAA-compliant EMR for plastic and reconstructive surgery.

Generate a complete, dictation-quality operative note narrative from the provided structured data. You receive NO patient PHI (no name, DOB, MRN, SSN, address, insurance). The note will be attributed to the surgeon separately by the EMR.

Return ONLY valid JSON — no markdown.

OUTPUT FORMAT:
{
  "preoperativeDiagnosis": "Clear pre-operative diagnosis statement",
  "postoperativeDiagnosis": "Post-operative diagnosis (same as pre-op unless changed intraoperatively)",
  "procedurePerformed": "Full procedure name with laterality if applicable",
  "operativeNarrative": "3-6 paragraph dictation-quality narrative covering: patient positioning and prep, anesthesia induction, surgical approach and technique, key intraoperative steps, implant/tissue details if applicable, closure technique, sponge/instrument counts, drain placement, dressing. Write in standard operative note style using past tense. Do not include patient name or identifiers.",
  "findings": "1-2 sentences describing intraoperative findings",
  "complications": "Complications noted, or 'The procedure was completed without complication.'",
  "ebl": "Estimated blood loss in mL",
  "disposition": "Patient disposition statement"
}

STYLE GUIDELINES:
- Use standard plastic surgery dictation language
- Past tense throughout
- Mention universal protocol compliance if provided
- Mention antibiotic prophylaxis timing if SSI data provided
- Be specific about technique, tissue planes, and closure layers
- Wound class should be stated as: "The wound was classified as {class}."`;

interface OpNoteDraftRequest {
  procedureName: string;
  procedureType: string;
  laterality?: string;
  anesthesiaType?: string;
  anesthesiaProvider?: string;
  surgeonPrimary?: string;
  positioning?: string;
  prepSolution?: string;
  estimatedBloodLossML?: number;
  crystalloidGivenML?: number;
  drainsPlaced?: boolean;
  drainDetails?: string;
  implantDetails?: string;
  tissueRemovedDetails?: string;
  fatHarvestedDetails?: string;
  complications?: string;
  condition?: string;
  disposition?: string;
  asaClass?: string;
  capriniScore?: number;
  stopBangScore?: number;
  rcriScore?: number;
  woundClass?: string;
  universalProtocolComplete?: boolean;
  ssiAntibiotic?: string;
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

  let body: OpNoteDraftRequest;
  try {
    body = JSON.parse(event.body ?? '{}');
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid request body' }) };
  }

  // Scrub entire body before any use — surgeon name is preserved (provider, not patient PHI)
  const scrubbed = scrubPHI(body as unknown as Record<string, unknown>) as unknown as OpNoteDraftRequest;

  const userMessage = `Generate an operative note for the following case:

Procedure: ${scrubbed.procedureName} (${scrubbed.laterality ?? 'not specified'})
Anesthesia: ${scrubbed.anesthesiaType ?? 'not documented'} by ${scrubbed.anesthesiaProvider ?? 'anesthesia team'}
Surgeon: ${scrubbed.surgeonPrimary ?? '[PROVIDER]'}
Positioning: ${scrubbed.positioning ?? 'not documented'}
Prep solution: ${scrubbed.prepSolution ?? 'not documented'}

Intraoperative data:
- EBL: ${scrubbed.estimatedBloodLossML ?? 'minimal'} mL
- IV crystalloid: ${scrubbed.crystalloidGivenML ?? 'not recorded'} mL
- Drains placed: ${scrubbed.drainsPlaced ? 'Yes — ' + (scrubbed.drainDetails ?? '') : 'None'}
- Implant details: ${scrubbed.implantDetails ?? 'Not applicable'}
- Tissue removed: ${scrubbed.tissueRemovedDetails ?? 'Not applicable'}
- Fat harvested: ${scrubbed.fatHarvestedDetails ?? 'Not applicable'}

Safety:
- ASA Class: ${scrubbed.asaClass ?? 'not documented'}
- Universal protocol: ${scrubbed.universalProtocolComplete ? 'Complete' : 'Partial/not documented'}
- SSI antibiotic: ${scrubbed.ssiAntibiotic ?? 'not documented'}
- Wound classification: ${scrubbed.woundClass ?? 'not documented'}

Patient condition at end: ${scrubbed.condition ?? 'stable'}
Disposition: ${scrubbed.disposition ?? 'PACU'}
Complications: ${scrubbed.complications || 'None'}

Generate the complete operative note narrative.`;

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
        max_tokens: 1500,
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
      functionName: 'ai-op-note-draft', action: 'generate_op_note_draft', phiScrubbed: true,
      encounterId: body.encounterId, orgId: body.orgId, providerId: body.providerId,
      latencyMs: Date.now() - t0,
    });

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(parsed),
    };
  } catch (error) {
    console.error('AI op note draft error:', error);
    void logAudit({ functionName: 'ai-op-note-draft', action: 'error', phiScrubbed: true, latencyMs: Date.now() - t0 });
    return { statusCode: 500, body: JSON.stringify({ error: 'AI service error' }) };
  }
};
