const SHEET_ID = process.env.SHEET_ID || '13q9R1M2RAxSQ5rsujRLDOtO-ip-eH6eH8Or-_9ZjiTQ';
const API_KEY  = process.env.GOOGLE_SHEETS_API_KEY;
const BASE     = 'https://sheets.googleapis.com/v4/spreadsheets';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  if (!API_KEY) {
    return res.status(500).json({
      error: 'GOOGLE_SHEETS_API_KEY is not set. Add it in Vercel → Settings → Environment Variables.'
    });
  }

  const { action, sheet } = req.query;

  try {
    if (action === 'list') {
      const url  = `${BASE}/${SHEET_ID}?key=${API_KEY}&fields=sheets.properties.title`;
      const resp = await fetch(url);
      const data = await resp.json();
      if (data.error) throw new Error(data.error.message || 'Sheets API error');
      const sheets = (data.sheets || []).map(s => s.properties.title);
      return res.json({ sheets });
    }

    if (sheet) {
      const range = encodeURIComponent(sheet);
      const url   = `${BASE}/${SHEET_ID}/values/${range}?key=${API_KEY}`;
      const resp  = await fetch(url);
      const data  = await resp.json();
      if (data.error) throw new Error(data.error.message || 'Sheets API error');
      return res.json({ values: data.values || [] });
    }

    res.status(400).json({ error: 'Missing ?action=list or ?sheet=SheetName' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
