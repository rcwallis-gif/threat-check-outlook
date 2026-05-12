const express = require('express');
const { OpenAI } = require('openai');
const cors = require('cors');
const path = require('path');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

// Serve the Outlook task pane files
app.use(express.static(__dirname));   // serves taskpane.html, taskpane.js, etc.

const openai = new OpenAI({
  apiKey: process.env.XAI_API_KEY,
  baseURL: 'https://api.x.ai/v1'
});

const SYSTEM_PROMPT = `You are a senior cybersecurity analyst. Analyze the FULL RAW EML of an email and return ONLY a concise report.
Format exactly like this:

THREAT LEVEL: Safe / Suspicious / Phishing / Spam

REASONS:
• Bullet point 1
• Bullet point 2
...

RECOMMENDATION: [one sentence]

Be extremely precise about:
- Sender spoofing / DMARC
- Suspicious links / domains
- Urgency / threats
- Attachments
- Header anomalies`;

app.post('/analyze', async (req, res) => {
  try {
    const { emlBase64, subject } = req.body;
    const emlText = Buffer.from(emlBase64, 'base64').toString('utf-8');

    const completion = await openai.chat.completions.create({
      model: "grok-4",
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: `Email subject: ${subject}\n\nFull raw EML:\n\n${emlText}` }
      ],
      temperature: 0.2,
      max_tokens: 800
    });

    const reportText = completion.choices[0].message.content;
    const isThreat = /phishing|spam|suspicious/i.test(reportText);

    res.json({
      report: reportText,
      isThreat,
      confidence: "High confidence"
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Grok analysis failed" });
  }
});

// Serve taskpane.html when someone goes to the root URL
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'taskpane.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Threat Check backend + frontend running on port ${PORT}`));