import type { Handler, HandlerEvent } from '@netlify/functions'

export const handler: Handler = async (event: HandlerEvent) => {
  if (event.httpMethod !== 'POST') return { statusCode: 405, body: 'Method Not Allowed' }

  const apiKey = process.env.ANTHROPIC_API_KEY
  if (!apiKey) return { statusCode: 500, body: JSON.stringify({ error: 'API key not configured' }) }

  try {
    const { patientData } = JSON.parse(event.body || '{}')

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1000,
        messages: [{
          role: 'user',
          content: `Review this plastic surgery patient record and generate clinical flags. Return only valid JSON array of flags.
Patient: ${JSON.stringify(patientData)}
Format: [{"severity":"critical|warning|info","category":"string","title":"string","description":"string","recommendedAction":"string"}]`,
        }],
      }),
    })

    const data = await response.json()
    const text = data.content?.[0]?.text || '[]'
    const flags = JSON.parse(text.replace(/```json|```/g, '').trim())

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ flags }),
    }
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: String(err) }) }
  }
}
