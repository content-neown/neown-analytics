module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { summary, sheetName, type } = req.body || {};
  const isRevenue = type === 'revenue';

  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY is not set in Vercel Environment Variables.' });
  }

  const systemPrompt = isRevenue
    ? `You are a senior revenue analyst for neOwn — India's largest children's book rental subscription service. Analyse revenue achievement data across Sales, Renewals (Manual + Automation), and Self-Serve channels. Use Indian number formatting (L, Cr) when appropriate. Be concise and direct.

Return ONLY a valid JSON object — no markdown, no backticks, no extra text:
{
  "summary": "2-3 sentence overview of achievement vs target",
  "highlights": ["3 positive observations"],
  "concerns": ["2-3 areas needing attention"],
  "recommendations": ["3-4 specific actionable items"],
  "aiCharts": [
    {"title": "chart idea", "insight": "what this metric would reveal"}
  ]
}`
    : `You are a senior performance marketing analyst for neOwn — India's largest children's book rental subscription service. Analyse Facebook and Google ad data. Use Indian number formatting (L, Cr) when appropriate. Be concise and direct.

Return ONLY a valid JSON object — no markdown, no backticks, no extra text:
{
  "summary": "2-3 sentence overview",
  "highlights": ["3 positive observations"],
  "concerns": ["2-3 areas needing attention"],
  "recommendations": ["3-4 specific actionable items"],
  "aiCharts": [
    {"title": "chart idea", "insight": "what this metric would reveal"}
  ]
}`;

  try {
    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1200,
        system: systemPrompt,
        messages: [{ role: 'user', content: `Analyse this data for ${sheetName}:\n\n${summary}` }],
      }),
    });

    const data  = await resp.json();
    if (!resp.ok) throw new Error(data?.error?.message || 'Claude API error');
    const text  = data.content?.[0]?.text || '{}';
    const clean = text.replace(/```json|```/g, '').trim();
    const ins   = JSON.parse(clean);
    res.json(ins);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
