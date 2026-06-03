export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { sintomas, foto } = req.body;

  if (!sintomas) {
    return res.status(400).json({ error: 'Sintomas são obrigatórios' });
  }

  try {
    const messages = [];
    
    if (foto) {
      messages.push({
        role: 'user',
        content: [
          {
            type: 'image',
            source: { type: 'base64', media_type: 'image/jpeg', data: foto }
          },
          { type: 'text', text: sintomas }
        ]
      });
    } else {
      messages.push({ role: 'user', content: sintomas });
    }

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1000,
        messages
      })
    });

    const data = await response.json();
    const diagnostico = data.content?.[0]?.text || 'Não foi possível analisar.';

    res.status(200).json({ diagnostico });
  } catch (error) {
    console.error('Erro diagnóstico:', error);
    res.status(500).json({ error: 'Erro ao processar diagnóstico: ' + error.message });
  }
}
