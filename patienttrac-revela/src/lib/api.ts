const BASE = import.meta.env.DEV
  ? 'http://localhost:8888/.netlify/functions'
  : '/.netlify/functions'

async function post<T>(endpoint: string, body: unknown): Promise<T> {
  const res = await fetch(`${BASE}/${endpoint}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  if (!res.ok) throw new Error(`API error ${res.status}: ${await res.text()}`)
  return res.json()
}

export const api = {
  intake: {
    analyze: (patientContext: unknown, section: string, answers: unknown) =>
      post('ai-intake', { patientContext, section, answers }),
  },
  flags: {
    generate: (patientData: unknown) =>
      post('ai-flags', { patientData }),
  },
  proposals: {
    generate: (data: {
      procedure: string
      patientName: string
      lineItems: unknown[]
      total: number
    }) => post('generate-proposal', data),
  },
}
