// server.js - Express 後端（供 Render 部署）
// 請確保在 Render 的 Environment variables 裡設定 GEMINI_API_KEY（或 ALLOWED_ORIGINS）
const express = require('express');
const axios = require('axios');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');

const app = express();

// ALLOWED_ORIGINS 解析：支援空值（表示允許所有），或以逗號分隔的清單；會 trim 並過濾空字串
let allowed = ['*'];
if (typeof process.env.ALLOWED_ORIGINS !== 'undefined') {
  const raw = process.env.ALLOWED_ORIGINS;
  if (raw === '' || raw.trim() === '*') {
    allowed = ['*'];
  } else {
    allowed = raw.split(',').map(s => s.trim()).filter(Boolean);
    // 如果結果為空，視為允許所有（避免把空字串當作唯一項目造成拒絕）
    if (allowed.length === 0) allowed = ['*'];
  }
}
console.log('[startup] ALLOWED_ORIGINS =', allowed);

app.use(cors({
  origin: (origin, callback) => {
    // 若無 origin，通常為 server-to-server 或同源直接呼叫（例如 curl），允許
    if (!origin) return callback(null, true);

    // 若設定為 '*'，則允許所有 origin
    if (allowed.includes('*')) return callback(null, true);

    // 若 origin 明確包含於 allowed 清單，允許
    if (allowed.includes(origin)) return callback(null, true);

    // 否則拒絕（以 Error 回傳，會走到後端的錯誤處理）
    const err = new Error('Origin not allowed by CORS');
    err.status = 403;
    console.warn('[cors] Rejected origin:', origin);
    return callback(err);
  }
}));

app.use(bodyParser.json({ limit: '200kb' })); // 限制 payload 大小
app.use(express.static(path.join(__dirname))); // serve index_AI.html, railwayData.js 等靜態檔案

const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

app.post('/api/gemini', async (req, res) => {
  const { content, rounds } = req.body;
  if (!content || !Array.isArray(rounds) || rounds.length === 0) {
    return res.status(400).json({ error: "content 與 rounds 欄位不可為空且 rounds 必須為陣列且至少有一筆" });
  }

  if (!GEMINI_API_KEY) {
    return res.status(500).json({ error: "Server 未設定 GEMINI_API_KEY（請在環境變數設定）" });
  }

  if (String(content).length > 20000) return res.status(400).json({ error: 'content 太長' });

  const roundsText = rounds.map((r, idx) => `第${idx+1}次回復內容:\n${r.handling}\n第${idx+1}次審查意見:\n${r.review}\n`).join('\n');
  const prompt = `
請根據「原始開立的項目內容」及各次「鐵路機構回復內容」和「審查意見內容」，綜合判斷回復是否符合改善方向。
請只回覆 JSON 格式，不要其他說明文字。
{
  "fulfill": "是",
  "reason": "..."
}
原始開立項目內容如下：
${content}

各次回復與審查內容如下：
${roundsText}
`;

  console.log(`[api/gemini] prompt length: ${prompt.length}, rounds: ${rounds.length}`);

  try {
    const apiRes = await axios.post(GEMINI_URL, {
      contents: [{ role: "user", parts: [{ text: prompt }] }]
    }, { headers: { 'Content-Type': 'application/json' }, timeout: 180000 });

    let aiReply = '';
    if (apiRes.data && Array.isArray(apiRes.data.candidates) && apiRes.data.candidates[0]?.content?.parts[0]?.text) {
      aiReply = apiRes.data.candidates[0].content.parts[0].text;
    } else {
      console.error('Unexpected google api response shape:', apiRes.data);
      return res.status(502).json({ error: 'API 格式異常', raw: apiRes.data });
    }

    let result = {};
    try {
      const matched = aiReply.match(/\{[\s\S]*\}/);
      if (matched) result = JSON.parse(matched[0]);
      else result = { error: 'AI 回覆非 JSON 格式', raw: aiReply };
    } catch (e) {
      result = { error: '解析 AI 回覆 JSON 失敗', raw: aiReply };
    }

    res.setHeader('Content-Type', 'application/json');
    return res.json(result);
  } catch (err) {
    console.error('Gemini Error:', err.response?.status || err.message, err.response?.data || '');
    if (err.response) return res.status(502).json({ error: `Google API ${err.response.status}`, raw: err.response.data });
    return res.status(500).json({ error: err.message });
  }
});

// 全域錯誤處理（確保不會回傳 HTML）
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err && err.stack ? err.stack : err);
  const status = err && err.status ? err.status : 500;
  res.status(status).json({ error: err.message || 'Internal Server Error' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));