import type { Handler, HandlerEvent } from '@netlify/functions'

export const handler: Handler = async (event: HandlerEvent) => {
  if (event.httpMethod !== 'POST') return { statusCode: 405, body: 'Method Not Allowed' }

  const apiKey = process.env.ANTHROPIC_API_KEY
  if (!apiKey) return { statusCode: 500, body: JSON.stringify({ error: 'API key not configured' }) }

  try {
    const { procedure, patientName, lineItems, total } = JSON.parse(event.body || '{}')

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1200,
        messages: [{
          role: 'user',
          content: `Generate a professional patient-facing surgical proposal letter for PatientTrac Revela.
Patient: ${patientName}
Procedure: ${procedure}
Line items: ${JSON.stringify(lineItems)}
Total: $${total}
Include: warm professional tone, procedure overview, cost breakdown narrative, recovery expectations, next steps.
Return JSON: { "subject": "string", "body": "string", "callToAction": "string" }`,
        }],
      }),
    })

    const data = await response.json()
    const text = data.content?.[0]?.text || '{}'
    const proposal = JSON.parse(text.replace(/```json|```/g, '').trim())

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(proposal),
    }
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: String(err) }) }
  }
}
