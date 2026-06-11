// ============================================================
// PatientTrac Revela — EHI Export
// Netlify Function: /.netlify/functions/ehi-export
//
// Regulatory: ONC 21st Century Cures Act §170.315(b)(10)
// Electronic Health Information (EHI) Export
//
// Exports all structured clinical data for a patient/encounter
// so that patients and providers can exercise their right to
// access, download, and transmit their health information.
//
// NOTE: This function intentionally returns PHI — that is the
// legal purpose of EHI export. The caller must be authenticated
// via a valid cross-app bridge token validated against Supabase.
//
// Expected latency: 5–10 seconds for large records due to
// parallel Supabase fetches. Netlify background function timeout
// is 26 seconds; this function should comfortably fit within it.
// ============================================================

import type { Handler, HandlerEvent } from '@netlify/functions';

interface EHIExportRequest {
  patientId: number;
  encounterId: number;
  orgId: string;
  requestedBy: 'patient' | 'provider' | 'care_team';
  format?: 'json' | 'fhir_r4';  // default: json (fhir_r4 reserved for future implementation)
  token: string;  // cross-app bridge token for auth validation
}

// ── Audit log — non-blocking write to Supabase cr.ai_audit_log ────────────────
async function logAudit(opts: {
  functionName: string;
  action: string;
  encounterId?: number;
  orgId?: string;
  latencyMs: number;
  phiScrubbed: boolean;
  exportRequestedBy?: string;
}) {
  const supabaseUrl = process.env.SUPABASE_URL;
  const svcKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!supabaseUrl || !svcKey) return;
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
        action: opts.action,
        encounter_id: opts.encounterId ?? null,
        org_id: opts.orgId ?? null,
        phi_scrubbed: opts.phiScrubbed,
        latency_ms: opts.latencyMs,
        specialty: 'plastic_surgery',
        export_requested_by: opts.exportRequestedBy ?? null,
      }),
    });
  } catch { /* audit failure must not block the clinical response */ }
}

// ── Helper: fetch a single Supabase table with service role ──────────────────
async function fetchTable(
  supabaseUrl: string,
  svcKey: string,
  path: string,
  tableName: string,
): Promise<unknown> {
  try {
    const res = await fetch(`${supabaseUrl}/rest/v1/${path}`, {
      method: 'GET',
      headers: {
        'apikey': svcKey,
        'Authorization': `Bearer ${svcKey}`,
        'Content-Profile': 'cr',
        'Accept': 'application/json',
      },
    });
    if (!res.ok) {
      const errText = await res.text().catch(() => res.statusText);
      console.error(`EHI fetch error [${tableName}]: ${res.status} ${errText}`);
      return { error: 'fetch failed', table: tableName, status: res.status };
    }
    return await res.json();
  } catch (err) {
    console.error(`EHI fetch exception [${tableName}]:`, err);
    return { error: 'fetch failed', table: tableName };
  }
}

export const handler: Handler = async (event: HandlerEvent) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  const supabaseUrl = process.env.SUPABASE_URL;
  const svcKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!supabaseUrl || !svcKey) {
    return { statusCode: 500, body: JSON.stringify({ error: 'EHI export service not configured' }) };
  }

  let body: EHIExportRequest;
  try {
    body = JSON.parse(event.body ?? '{}');
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid request body' }) };
  }

  // Validate required fields
  if (!body.patientId || !body.encounterId || !body.orgId || !body.token) {
    return {
      statusCode: 400,
      body: JSON.stringify({ error: 'Missing required fields: patientId, encounterId, orgId, token' }),
    };
  }

  // ── Step 1: Validate cross-app bridge token ─────────────────────────────────
  try {
    const validRes = await fetch(`${supabaseUrl}/rest/v1/rpc/validate_cross_app_token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': svcKey,
        'Authorization': `Bearer ${svcKey}`,
        'Content-Profile': 'cr',
      },
      body: JSON.stringify({ token_string: body.token }),
    });
    const valid = await validRes.json();
    if (!valid) {
      return { statusCode: 401, body: JSON.stringify({ error: 'Invalid session token' }) };
    }
  } catch (err) {
    console.error('Token validation error:', err);
    return { statusCode: 401, body: JSON.stringify({ error: 'Token validation failed' }) };
  }

  const t0 = Date.now();
  const { patientId, encounterId } = body;

  // ── Step 2: Fetch all EHI data in parallel ──────────────────────────────────
  // Individual fetch failures are captured per-table rather than failing the whole export.
  const [
    patientResult,
    encounterResult,
    consultationResult,
    bodyPartResult,
    operativeNoteResult,
    medicationsResult,
    allergiesResult,
    diagnosesResult,
    surgHistResult,
    vitalsResult,
  ] = await Promise.allSettled([
    fetchTable(supabaseUrl, svcKey, `patient?patient_id=eq.${patientId}&select=*`, 'patient'),
    fetchTable(supabaseUrl, svcKey, `encounter?encounter_id=eq.${encounterId}&select=*`, 'encounter'),
    fetchTable(supabaseUrl, svcKey, `surgical_consultation?encounter_id=eq.${encounterId}&select=*`, 'surgical_consultation'),
    fetchTable(supabaseUrl, svcKey, `body_part_documentation?encounter_id=eq.${encounterId}&select=*`, 'body_part_documentation'),
    fetchTable(supabaseUrl, svcKey, `operative_note_detailed?encounter_id=eq.${encounterId}&select=*`, 'operative_note_detailed'),
    fetchTable(supabaseUrl, svcKey, `patient_medications?patient_id=eq.${patientId}&select=*`, 'patient_medications'),
    fetchTable(supabaseUrl, svcKey, `patient_allergies?patient_id=eq.${patientId}&select=*`, 'patient_allergies'),
    fetchTable(supabaseUrl, svcKey, `patient_diagnosis?patient_id=eq.${patientId}&select=*`, 'patient_diagnosis'),
    fetchTable(supabaseUrl, svcKey, `patient_surghist?patient_id=eq.${patientId}&select=*`, 'patient_surghist'),
    fetchTable(supabaseUrl, svcKey, `patient_vitals?patient_id=eq.${patientId}&select=*`, 'patient_vitals'),
  ]);

  // Unwrap settled results — rejected promises produce a structured error object
  const unwrap = (result: PromiseSettledResult<unknown>, tableName: string): unknown =>
    result.status === 'fulfilled'
      ? result.value
      : { error: 'fetch failed', table: tableName };

  const patientData        = unwrap(patientResult,        'patient');
  const encounterData      = unwrap(encounterResult,      'encounter');
  const consultationData   = unwrap(consultationResult,   'surgical_consultation');
  const bodyPartData       = unwrap(bodyPartResult,       'body_part_documentation');
  const operativeNoteData  = unwrap(operativeNoteResult,  'operative_note_detailed');
  const medicationsData    = unwrap(medicationsResult,    'patient_medications');
  const allergiesData      = unwrap(allergiesResult,      'patient_allergies');
  const diagnosesData      = unwrap(diagnosesResult,      'patient_diagnosis');
  const surgHistData       = unwrap(surgHistResult,       'patient_surghist');
  const vitalsData         = unwrap(vitalsResult,         'patient_vitals');

  // ── Step 3: Write audit log (non-blocking) ──────────────────────────────────
  void logAudit({
    functionName: 'ehi-export',
    action: 'ehi_export',
    encounterId: Number(body.encounterId),
    orgId: body.orgId,
    latencyMs: Date.now() - t0,
    phiScrubbed: false,  // This endpoint intentionally exports PHI — that is its regulatory purpose
    exportRequestedBy: body.requestedBy,
  });

  // ── Step 4: Build and return the EHI export bundle ─────────────────────────
  const bundle = {
    exportVersion: '1.0',
    standard: 'ONC_21st_Century_Cures_EHI',
    regulatory: 'ONC §170.315(b)(10) — Electronic Health Information Export',
    exportTimestamp: new Date().toISOString(),
    patientId: body.patientId,
    encounterId: body.encounterId,
    orgId: body.orgId,
    requestedBy: body.requestedBy,
    data: {
      patient: patientData,
      encounter: encounterData,
      surgicalConsultation: consultationData,
      bodyPartDocumentation: bodyPartData,
      operativeNote: operativeNoteData,
      medications: medicationsData,
      allergies: allergiesData,
      diagnoses: diagnosesData,
      surgicalHistory: surgHistData,
      vitals: vitalsData,
    },
  };

  return {
    statusCode: 200,
    headers: {
      'Content-Type': 'application/json',
      'Content-Disposition': `attachment; filename="ehi-export-encounter-${body.encounterId}.json"`,
      'X-EHI-Export-Version': '1.0',
      'X-Regulatory': 'ONC-21CCA-170.315(b)(10)',
    },
    body: JSON.stringify(bundle),
  };
};
