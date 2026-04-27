import { serve } from 'https://deno.land/std@0.168.0/http/server.ts'
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders })

  try {
    const supabase = createClient(
      Deno.env.get('SUPABASE_URL')!,
      Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
    )

    const { member_id, invited_by_name } = await req.json()
    if (!member_id) throw new Error('member_id required')

    // Fetch member record
    const { data: member, error: memberErr } = await supabase
      .from('org_members')
      .select('id, email, first_name, last_name, role, invite_token, org_id')
      .eq('id', member_id)
      .single()

    if (memberErr || !member) throw new Error('Member not found')
    if (!member.invite_token) throw new Error('No invite token on record')

    // Fetch org name
    const { data: org } = await supabase
      .from('organizations')
      .select('org_name')
      .eq('org_id', member.org_id)
      .single()

    const orgName = org?.org_name ?? 'PatientTracForge'
    const inviteUrl = `https://patienttracforge.com/admin/accept-invite?token=${member.invite_token}&email=${encodeURIComponent(member.email)}`
    const resendKey = Deno.env.get('RESEND_API_KEY')!
    const fromEmail = Deno.env.get('FROM_EMAIL') ?? 'noreply@patienttracforge.com'

    const roleLabel: Record<string, string> = {
      provider: 'Provider',
      admin: 'Administrator',
      billing: 'Billing Specialist',
      staff: 'Staff',
      readonly: 'Read-Only User',
      super_admin: 'Super Administrator',
    }

    const emailHtml = buildInviteEmail({
      firstName: member.first_name ?? 'Doctor',
      lastName: member.last_name ?? '',
      role: roleLabel[member.role] ?? member.role,
      orgName,
      inviteUrl,
      invitedByName: invited_by_name ?? 'The PatientTracForge Team',
    })

    // Send via Resend
    const resendRes = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${resendKey}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        from: `PatientTracForge <${fromEmail}>`,
        to: [member.email],
        subject: `You're invited to PatientTracForge — ${orgName}`,
        html: emailHtml,
      }),
    })

    if (!resendRes.ok) {
      const err = await resendRes.text()
      throw new Error(`Resend error: ${err}`)
    }

    // Update invite_sent_at
    await supabase
      .from('org_members')
      .update({ invite_sent_at: new Date().toISOString() })
      .eq('id', member_id)

    return new Response(
      JSON.stringify({ success: true, email: member.email }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )

  } catch (err) {
    return new Response(
      JSON.stringify({ error: err.message }),
      { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )
  }
})

// ─── Email Template ───────────────────────────────────────────────────────────

function buildInviteEmail({ firstName, lastName, role, orgName, inviteUrl, invitedByName }: {
  firstName: string
  lastName: string
  role: string
  orgName: string
  inviteUrl: string
  invitedByName: string
}) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>You're invited to PatientTracForge</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { background:#f0f2f5; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
  .wrapper { max-width:600px; margin:40px auto; background:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 4px 24px rgba(0,0,0,0.08); }
  .header { background:#060e1c; padding:36px 40px; text-align:center; }
  .header-logo { display:inline-flex; align-items:center; gap:10px; }
  .logo-diamond { width:32px; height:32px; }
  .logo-text { color:#c9a96e; font-size:22px; font-weight:600; letter-spacing:0.5px; }
  .header-tagline { color:rgba(201,169,110,0.6); font-size:12px; margin-top:6px; letter-spacing:1.5px; text-transform:uppercase; }
  .body { padding:44px 40px; }
  .greeting { font-size:24px; font-weight:600; color:#060e1c; margin-bottom:8px; }
  .subtext { font-size:15px; color:#64748b; line-height:1.6; margin-bottom:28px; }
  .role-badge { display:inline-block; background:#f0f4ff; color:#3b5bdb; font-size:12px; font-weight:600; padding:4px 12px; border-radius:20px; letter-spacing:0.5px; text-transform:uppercase; margin-bottom:28px; }
  .steps { background:#f8fafc; border-radius:10px; padding:24px 28px; margin-bottom:32px; }
  .steps-title { font-size:13px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:1px; margin-bottom:16px; }
  .step { display:flex; align-items:flex-start; gap:14px; margin-bottom:14px; }
  .step:last-child { margin-bottom:0; }
  .step-num { width:24px; height:24px; background:#060e1c; color:#c9a96e; border-radius:50%; font-size:12px; font-weight:600; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
  .step-text { font-size:14px; color:#334155; line-height:1.5; padding-top:3px; }
  .step-text strong { color:#060e1c; }
  .cta-wrap { text-align:center; margin-bottom:32px; }
  .cta-btn { display:inline-block; background:#c9a96e; color:#060e1c; font-size:15px; font-weight:700; padding:16px 40px; border-radius:8px; text-decoration:none; letter-spacing:0.3px; }
  .cta-sub { font-size:12px; color:#94a3b8; margin-top:10px; }
  .divider { border:none; border-top:1px solid #e2e8f0; margin:28px 0; }
  .mfa-note { background:#fffbeb; border:1px solid #fcd34d; border-radius:8px; padding:16px 20px; margin-bottom:28px; }
  .mfa-note-title { font-size:13px; font-weight:600; color:#92400e; margin-bottom:6px; }
  .mfa-note-text { font-size:13px; color:#78350f; line-height:1.6; }
  .link-fallback { font-size:12px; color:#94a3b8; line-height:1.6; word-break:break-all; }
  .link-fallback a { color:#3b5bdb; }
  .footer { background:#f8fafc; padding:24px 40px; text-align:center; border-top:1px solid #e2e8f0; }
  .footer-text { font-size:12px; color:#94a3b8; line-height:1.7; }
  .footer-links a { color:#64748b; text-decoration:none; margin:0 8px; }
</style>
</head>
<body>
<div class="wrapper">

  <!-- Header -->
  <div class="header">
    <div class="header-logo">
      <svg class="logo-diamond" viewBox="0 0 32 32" fill="none">
        <polygon points="16,2 30,16 16,30 2,16" fill="none" stroke="#c9a96e" stroke-width="1.5"/>
        <polygon points="16,7 25,16 16,25 7,16" fill="#c9a96e" opacity="0.3"/>
        <polygon points="16,11 21,16 16,21 11,16" fill="#c9a96e"/>
      </svg>
      <span class="logo-text">PatientTracForge</span>
    </div>
    <div class="header-tagline">AI-Powered Practice Management · HIPAA Compliant</div>
  </div>

  <!-- Body -->
  <div class="body">
    <div class="greeting">Welcome, Dr. ${lastName}.</div>
    <p class="subtext">
      ${invitedByName} has invited you to join <strong>${orgName}</strong> on PatientTracForge —
      the AI-powered EMR platform for scheduling, clinical documentation, and billing.
    </p>

    <div class="role-badge">${role}</div>

    <!-- Steps -->
    <div class="steps">
      <div class="steps-title">Getting started — 3 steps</div>
      <div class="step">
        <div class="step-num">1</div>
        <div class="step-text">Click the button below to <strong>accept your invitation</strong> and set your password.</div>
      </div>
      <div class="step">
        <div class="step-num">2</div>
        <div class="step-text">Install <strong>Google Authenticator</strong> on your phone (App Store or Google Play).</div>
      </div>
      <div class="step">
        <div class="step-num">3</div>
        <div class="step-text">Scan the QR code shown after login to <strong>activate your MFA</strong>. Required for all provider accounts.</div>
      </div>
    </div>

    <!-- CTA -->
    <div class="cta-wrap">
      <a href="${inviteUrl}" class="cta-btn">Accept Invitation →</a>
      <div class="cta-sub">This link expires in 72 hours</div>
    </div>

    <!-- MFA note -->
    <div class="mfa-note">
      <div class="mfa-note-title">⚠ MFA is required for your account</div>
      <div class="mfa-note-text">
        As a provider, your account requires two-factor authentication via Google Authenticator.
        You will be prompted to set this up on your first login — please have your phone ready.
        Each person sets up MFA on their own device.
      </div>
    </div>

    <hr class="divider"/>

    <p class="link-fallback">
      If the button above doesn't work, copy and paste this link into your browser:<br/>
      <a href="${inviteUrl}">${inviteUrl}</a>
    </p>
  </div>

  <!-- Footer -->
  <div class="footer">
    <div class="footer-text">
      This invitation was sent to ${firstName} ${lastName} by ${invitedByName}.<br/>
      PatientTracForge · HIPAA Compliant · 256-bit TLS · Google Authenticator MFA<br/><br/>
      <div class="footer-links">
        <a href="https://patienttracforge.com/privacy">Privacy Policy</a>
        <a href="https://patienttracforge.com/terms">Terms of Service</a>
        <a href="https://patienttracforge.com/hipaa">HIPAA Notice</a>
        <a href="mailto:support@patienttracforge.com">Support</a>
      </div>
    </div>
  </div>

</div>
</body>
</html>`
}
