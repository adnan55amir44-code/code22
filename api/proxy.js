/**
 * Vercel Serverless Function
 * Purpose: same-origin proxy from Vercel frontend to Google Apps Script Web App.
 *
 * Set the Vercel environment variable:
 * GAS_EXEC_URL = https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
 */

export default async function handler(req, res) {
  const target = process.env.GAS_EXEC_URL;
  if (!target) {
    return res.status(500).json({
      success: false,
      message: 'GAS_EXEC_URL is not configured on Vercel.'
    });
  }

  if (!['GET', 'POST', 'OPTIONS'].includes(req.method)) {
    return res.status(405).json({ success: false, message: 'Method not allowed.' });
  }

  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(204).end();
  }

  try {
    const init = {
      method: req.method,
      redirect: 'follow',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      }
    };

    if (req.method === 'POST') {
      let payload = req.body;
      if (typeof payload !== 'string') payload = JSON.stringify(payload ?? {});
      init.body = payload;
    }

    const upstream = await fetch(target, init);
    const text = await upstream.text();

    res.status(upstream.status || 200);
    res.setHeader('Content-Type', upstream.headers.get('content-type') || 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');
    res.send(text);
  } catch (error) {
    console.error('GAS proxy error:', error);
    res.status(502).json({
      success: false,
      message: 'Unable to reach the Google Apps Script backend.'
    });
  }
}