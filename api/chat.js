module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { messages, context } = req.body || {};

  if (!process.env.ANTHROPIC_API_KEY) {
    return res.json({ reply: 'ANTHROPIC_API_KEY is not configured in Vercel.' });
  }

  try {
    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 500,
        system: `You are a performance marketing analyst for neOwn — India's largest children's book rental subscription service. You help interpret Facebook and Google ad data. Be concise (3-5 sentences max). Use Indian formatting (₹, L, Cr). Reference specific numbers when you can.

Current data:
${context || 'No data loaded yet.'}`,
        messages: (messages || [])
          .filter(m => m.role === 'user' || m.role === 'assistant')
          .slice(-12),
      }),
    });

    const data  = await resp.json();
    if (!resp.ok) throw new Error(data?.error?.message || 'Claude error');
    const reply = data.content?.[0]?.text || 'No response.';
    res.json({ reply });
  } catch (err) {
    res.json({ reply: 'Error: ' + err.message });
  }
};
