// ============================================================
// PatientTrac Revela — AI Surgical Proposal Generator
// Netlify Function: /.netlify/functions/ai-proposal
// ============================================================

import type { Handler, HandlerEvent } from '@netlify/functions';

const ANTHROPIC_API = 'https://api.anthropic.com/v1/messages';

const SYSTEM_PROMPT = `You are the Revela billing and proposal AI for PatientTrac — a HIPAA-compliant 
EMR for plastic and reconstructive surgery. Your role is to generate accurate, itemized surgical 
cost proposals and patient-facing financing summaries.

FEE GUIDELINES (adjust based on provided overrides):
- Surgeon fee: $4,000–$12,000 depending on procedure complexity
- Anesthesia: $600/hour (estimate from procedure duration)
- OR facility fee: $3,000–$6,000 depending on duration and suite tier
- Implants/materials: actual cost from inventory if available, else estimate
- Post-op visits (2 standard): $350–$500 total
- Garments/supplies: $150–$250

FINANCING OPTIONS:
1. Pay in full — apply 5% courtesy discount
2. CareCredit 24-month 0% — no discount, calculate $X/month
3. In-house plan — 20% down, balance over 10 months, 0% interest

OUTPUT FORMAT: Valid JSON only. No markdown.
Schema: { lineItems: CostLineItem[], totalEstimate: number, financingOptions: FinancingOption[], 
          patientSummary: string }

CostLineItem: { label, amount, category, editable: true }
FinancingOption: { type, label, totalAmount, downPayment?, monthlyPayment?, termMonths?, discountApplied? }
patientSummary: 2-sentence plain English summary for patient-facing proposal cover page`;

interface ProposalRequest {
  procedureType: string;
  procedureName: string;
  estimatedDurationHours: number;
  patientName: string;
  surgeonName: string;
  feeOverrides?: Record<string, number>;
  includeContralatéral?: boolean;
  includeNAR?: boolean;
}

export const handler: Handler = async (event: HandlerEvent) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return { statusCode: 500, body: JSON.stringify({ error: 'AI service not configured' }) };
  }

  let body: ProposalRequest;
  try {
    body = JSON.parse(event.body ?? '{}');
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid request body' }) };
  }

  const userMessage = `Generate a complete surgical cost proposal for the following:

Procedure: ${body.procedureName} (${body.procedureType})
Estimated duration: ${body.estimatedDurationHours} hours
Surgeon: ${body.surgeonName}
Patient: ${body.patientName}
${body.feeOverrides ? `Fee overrides: ${JSON.stringify(body.feeOverrides)}` : ''}
${body.includeContralatéral ? 'Include: Contralateral symmetry procedure (augmentation)' : ''}
${body.includeNAR ? 'Include: Staged nipple-areola reconstruction (separate line item)' : ''}

Generate all line items, total estimate, three financing options, and patient summary.`;

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
      parsed = { error: 'Parse error', raw: text };
    }

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(parsed),
    };
  } catch (error) {
    console.error('AI proposal error:', error);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: 'AI service error' }),
    };
  }
};
