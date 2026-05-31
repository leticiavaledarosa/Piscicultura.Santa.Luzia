// api/foto.js — Vercel Function para receber foto e enviar ao Google Drive via Apps Script
export default async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { base64, nome, cage_id, data, chave } = req.body;

    if (chave !== 'AQUALUZIA2026') {
      return res.status(403).json({ error: 'Chave inválida' });
    }

    const WEBHOOK = 'https://script.google.com/macros/s/AKfycbzlAyW969zFIUaUHydehZVmr0yXEM6REAtnte9VBY1NcfokhHyG_6gN8BY805TgHWiV/exec';

    const payload = {
      chave: 'AQUALUZIA2026',
      tipo: 'foto',
      data: {
        base64,
        nome,
        cage_id,
        data,
        pasta: '1dCbTHQVZEL-ncLN5461kbWPJ8TP5AErq'
      }
    };

    const response = await fetch(WEBHOOK, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    return res.status(200).json({ ok: true, msg: 'Foto enviada!' });

  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
