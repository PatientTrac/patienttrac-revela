// ============================================================
// PatientTrac Revela — AI Intake Serverless Function
// Netlify Function: /.netlify/functions/ai-intake
// ============================================================

import type { Handler, HandlerEvent } from '@netlify/functions';

const ANTHROPIC_API = 'https://api.anthropic.com/v1/messages';

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

  const sectionLabels: Record<number, string> = {
    1: 'Oncology and mastectomy history',
    2: 'Medical history and comorbidities',
    3: 'Social history and lifestyle',
    4: 'Psychological assessment and patient goals',
    5: 'Procedure preference and surgical planning',
    6: 'Complete intake summary — generate full physician briefing',
  };

  const userMessage = body.generateBriefing
    ? `Generate a complete physician briefing paragraph for this patient based on all intake sections. 
       Patient context: ${JSON.stringify(body.patientContext ?? {})}
       Final section data: ${JSON.stringify(body.sectionData)}`
    : `Analyze section ${body.section} (${sectionLabels[body.section]}) of the patient intake interview.
       Section data: ${JSON.stringify(body.sectionData)}
       Patient context from prior sections: ${JSON.stringify(body.patientContext ?? {})}
       Generate clinical flags for any issues found. Return empty flags array if none.`;

  try {
    const response = await fetch(ANTHROPIC_API, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1000,
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
    try {
      parsed = JSON.parse(text);
    } catch {
      parsed = { flags: [], briefing: text };
    }

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(parsed),
    };
  } catch (error) {
    console.error('AI intake error:', error);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: 'AI service error', flags: [] }),
    };
  }
};
