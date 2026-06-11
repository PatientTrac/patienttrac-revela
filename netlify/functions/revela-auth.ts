// ============================================================
// PatientTrac Revela — Cross-App JWT Exchange
// Netlify Function: /.netlify/functions/revela-auth
//
// Validates a PatientTracForge cross-app token server-side,
// mints a Supabase magic-link token for the provider, ensures
// saas.org_members has the provider's auth user linked, and
// returns a token_hash the browser exchanges for a real session.
// No PHI is logged. All Supabase reads use service role key.
// ============================================================

import type { Handler, HandlerEvent } from '@netlify/functions';

const SUPABASE_URL = process.env.SUPABASE_URL;
const SVC_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;

const svcHeaders = {
  'Content-Type': 'application/json',
  'apikey': SVC_KEY ?? '',
  'Authorization': `Bearer ${SVC_KEY ?? ''}`,
};

async function callRpc(fn: string, params: Record<string, unknown>) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/rpc/${fn}`, {
    method: 'POST',
    headers: svcHeaders,
    body: JSON.stringify(params),
  });
  if (!r.ok) throw new Error(`RPC ${fn} failed: ${r.status} ${await r.text()}`);
  return r.json();
}

async function queryTable(schema: string, table: string, qs: string) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${table}?${qs}`, {
    headers: { ...svcHeaders, 'Accept-Profile': schema },
  });
  if (!r.ok) throw new Error(`Query ${schema}.${table} failed: ${r.status}`);
  return r.json();
}

async function upsertTable(schema: string, table: string, row: Record<string, unknown>) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${table}`, {
    method: 'POST',
    headers: {
      ...svcHeaders,
      'Content-Profile': schema,
      'Prefer': 'resolution=ignore-duplicates',
    },
    body: JSON.stringify(row),
  });
  // 201 Created or 200 OK — both fine; ignore body
  if (r.status >= 400) throw new Error(`Upsert ${schema}.${table} failed: ${r.status}`);
}

export const handler: Handler = async (event: HandlerEvent) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  if (!SUPABASE_URL || !SVC_KEY) {
    return { statusCode: 500, body: JSON.stringify({ error: 'Service not configured' }) };
  }

  let token: string;
  try {
    const body = JSON.parse(event.body ?? '{}');
    token = body.token;
    if (!token || typeof token !== 'string') throw new Error('Missing token');
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid request body' }) };
  }

  try {
    // ── Step 1: Validate cross-app token (server-side, correct param: p_token) ──
    const tokenData = await callRpc('validate_cross_app_token', { p_token: token });

    if (!tokenData?.valid) {
      return {
        statusCode: 401,
        body: JSON.stringify({ error: tokenData?.error ?? 'Invalid or expired token' }),
      };
    }

    const { provider_id, org_id, patient, encounter } = tokenData;

    // ── Step 2: Look up provider email from cr.providers ──
    const [provider] = await queryTable(
      'cr',
      'providers',
      `provider_id=eq.${provider_id}&select=email,first_name,last_name,credential&limit=1`,
    );

    if (!provider?.email) {
      return { statusCode: 422, body: JSON.stringify({ error: 'Provider record has no email address' }) };
    }

    // ── Step 3: Mint a magic-link token (creates auth user if not yet exists) ──
    const linkRes = await fetch(`${SUPABASE_URL}/auth/v1/admin/generate_link`, {
      method: 'POST',
      headers: svcHeaders,
      body: JSON.stringify({
        type: 'magiclink',
        email: provider.email,
        data: { org_id, provider_id },
      }),
    });

    if (!linkRes.ok) {
      const errText = await linkRes.text();
      throw new Error(`Auth link generation failed: ${errText}`);
    }

    const linkData = await linkRes.json();
    const hashedToken: string = linkData?.properties?.hashed_token ?? linkData?.hashed_token;
    const authUserId: string | undefined = linkData?.user?.id;

    if (!hashedToken) throw new Error('No hashed_token returned from auth API');

    // ── Step 4: Ensure saas.org_members row exists for this user + org ──
    if (authUserId) {
      await upsertTable('saas', 'org_members', {
        user_id: authUserId,
        org_id,
        email: provider.email,
        first_name: provider.first_name,
        last_name: provider.last_name,
        role: 'provider',
        is_active: true,
      });
    }

    // ── Step 5: Return token_hash + pre-loaded patient context ──
    const providerName = [
      provider.first_name,
      provider.last_name,
      provider.credential ? `, ${provider.credential}` : '',
    ].join(' ').trim();

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        token_hash: hashedToken,
        patient_context: {
          patient_id:        patient.patient_id,
          encounter_id:      encounter.encounter_id,
          patient_name:      `${patient.first_name} ${patient.last_name}`.trim(),
          patient_sex:       patient.gender,
          patient_dob:       patient.dob,
          patient_photo_url: patient.photo_url ?? null,
          chief_complaint:   encounter.chief_complaint ?? null,
          provider_id,
          provider_name:     providerName,
          org_id,
        },
      }),
    };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Authentication failed';
    return { statusCode: 500, body: JSON.stringify({ error: msg }) };
  }
};
