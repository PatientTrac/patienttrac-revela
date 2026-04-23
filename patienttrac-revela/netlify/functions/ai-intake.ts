import type { Handler, HandlerEvent } from '@netlify/functions'

const ANTHROPIC_API = 'https://api.anthropic.com/v1/messages'

export const handler: Handler = async (event: HandlerEvent) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' }
  }

  const apiKey = process.env.ANTHROPIC_API_KEY
  if (!apiKey) {
    return { statusCode: 500, body: JSON.stringify({ error: 'API key not configured' }) }
  }

  try {
    const { patientContext, section, answers } = JSON.parse(event.body || '{}')

    const systemPrompt = `You are a clinical AI assistant for PatientTrac Revela, a plastic and reconstructive surgery EMR. 
Your role is to:
1. Analyze patient intake responses and identify clinical risk flags
2. Pre-fill intake forms from provided EMR chart data
3. Generate structured SOAP notes from physician dictation
4. Create AI summaries for physician pre-consultation review

Always respond in valid JSON. Flag severity levels: "critical" (requires immediate action), "warning" (requires review), "info" (informational).
Focus on plastic surgery-specific risks: prior radiation, anticoagulants, cardiovascular status, BMI, smoking, prior abdominal surgery for flap donors, implant history.
Never provide direct medical advice — only flag for physician review.`

    const userPrompt = `Patient context: ${JSON.stringify(patientContext)}
Current intake section: ${section}
Current answers: ${JSON.stringify(answers)}

Analyze this intake data and return JSON with:
{
  "flags": [{ "severity": "critical|warning|info", "category": "string", "title": "string", "description": "string", "recommendedAction": "string" }],
  "aiSummary": "string — 2-3 sentence clinical summary for physician",
  "preFillSuggestions": { "fieldId": "suggested value" },
  "proceedRecommendation": "clear|review|hold"
}`

    const response = await fetch(ANTHROPIC_API, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1500,
        system: systemPrompt,
        messages: [{ role: 'user', content: userPrompt }],
      }),
    })

    const data = await response.json()
    const content = data.content?.[0]?.text || '{}'

    let parsed
    try {
      parsed = JSON.parse(content)
    } catch {
      parsed = { aiSummary: content, flags: [], preFillSuggestions: {}, proceedRecommendation: 'review' }
    }

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(parsed),
    }
  } catch (err) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: 'AI intake processing failed', details: String(err) }),
    }
  }
}
