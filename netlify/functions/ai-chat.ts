// ============================================================
// PatientTrac Revela — AI Assistant Chat Proxy
// Netlify Function: /.netlify/functions/ai-chat
//
// SECURITY: All Claude API calls must go through this server-side
// function. The ANTHROPIC_API_KEY must NEVER be in browser code.
// ============================================================

import type { Handler, HandlerEvent } from '@netlify/functions';

const ANTHROPIC_API = 'https://api.anthropic.com/v1/messages';
const MODEL = 'claude-fable-5';

const SYSTEM_PROMPT = `You are the AI assistant for PatientTrac Revela — a HIPAA-compliant EMR for
plastic and reconstructive surgery. You help physicians navigate the EMR, answer clinical
documentation questions, and route feedback to the PatientTrac team.

IMPORTANT:
- Never ask for, repeat, or store patient PHI (names, DOB, SSN, MRN, addresses, insurance)
- If a user attempts to send PHI, acknowledge the question clinically without repeating the PHI
- Route feature requests and bug reports to: wayne@patienttracforge.com
- Modules available: Initial Exam, Breast, Body Contour, Liposuction, Face/Neck, Operative Note,
  Surgical Drawing Tool, Chart, Progress Notes, Post-Op Plan, New Patient Registration
- This system uses Supabase PostgreSQL (cr schema) and Netlify Functions
- AI clinical features use claude-fable-5 for reasoning tasks

Keep responses concise and clinically useful. For technical questions about the EMR, be precise.`;

interface ChatRequest {
  messages: { role: 'user' | 'assistant'; content: string }[];
  orgId?: string;
  providerId?: string;
  encounterId?: string;
}

// ── PHI guard — remove identifiers from message content before logging ────────
const PHI_PATTERN = /\b(\d{3}-\d{2}-\d{4}|\d{9}|[A-Z]{2}\d{6,})\b/g

function redactForLog(text: string): string {
  return text.replace(PHI_PATTERN, '[REDACTED]')
}

// ── Audit log ─────────────────────────────────────────────────────────────────
async function logAudit(opts: {
  action: string; encounterId?: string; orgId?: string;
  providerId?: string; latencyMs: number;
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
        function_name: 'ai-chat',
        model_used: MODEL,
        action: opts.action,
        encounter_id: opts.encounterId ? Number(opts.encounterId) : null,
        org_id: opts.orgId ?? null,
        phi_scrubbed: true,
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

  let body: ChatRequest;
  try {
    body = JSON.parse(event.body ?? '{}');
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid request body' }) };
  }

  if (!body.messages || !Array.isArray(body.messages) || body.messages.length === 0) {
    return { statusCode: 400, body: JSON.stringify({ error: 'messages array required' }) };
  }

  // Cap message history to last 20 exchanges to limit token use
  const recentMessages = body.messages.slice(-20)

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
        max_tokens: 1024,
        system: SYSTEM_PROMPT,
        messages: recentMessages,
      }),
    });

    if (!response.ok) throw new Error(`Anthropic API error: ${response.status}`);

    const data = await response.json();
    const replyText: string = data.content?.[0]?.text ?? 'I encountered an error. Please try again.';

    void logAudit({
      action: 'chat_message',
      encounterId: body.encounterId,
      orgId: body.orgId,
      providerId: body.providerId,
      latencyMs: Date.now() - t0,
    });

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ reply: replyText }),
    };
  } catch (error) {
    console.error('AI chat error:', error);
    void logAudit({ action: 'error', latencyMs: Date.now() - t0 });
    return { statusCode: 500, body: JSON.stringify({ error: 'AI service error' }) };
  }
};
