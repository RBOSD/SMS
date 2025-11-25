// server.js - Express 後端
// 請確保在 Render 的 Environment variables 裡設定 GEMINI_API_KEY 與 DATABASE_URL
require('dotenv').config(); // 載入本地 .env (Render 上會自動忽略)
const express = require('express');
const axios = require('axios');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const { Pool } = require('pg'); // 引入 PostgreSQL 套件

const app = express();

// 1. 設定資料庫連線池
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: {
    rejectUnauthorized: false // Render 的 PostgreSQL 需要 SSL 連線
  }
});

// 測試資料庫連線
pool.connect((err, client, release) => {
  if (err) {
    return console.error('[Database] Connection error:', err.stack);
  }
  console.log('[Database] Connected successfully!');
  release();
});

// ... (原本的 CORS 設定與 app.use 保持不變) ...
let allowed = ['*'];
if (typeof process.env.ALLOWED_ORIGINS !== 'undefined') {
  // ... (保留原本的邏輯)
  const raw = process.env.ALLOWED_ORIGINS;
  if (raw === '' || raw.trim() === '*') {
      allowed = ['*'];
  } else {
      allowed = raw.split(',').map(s => s.trim()).filter(Boolean);
      if (allowed.length === 0) allowed = ['*'];
  }
}

app.use(cors({
  origin: (origin, callback) => {
      if (!origin) return callback(null, true);
      if (allowed.includes('*')) return callback(null, true);
      if (allowed.includes(origin)) return callback(null, true);
      const err = new Error('Origin not allowed by CORS');
      err.status = 403;
      return callback(err);
  }
}));

app.use(bodyParser.json({ limit: '200kb' }));
app.use(express.static(path.join(__dirname)));

// 2. 新增一個簡單的 API 來初始化資料庫表格 (一次性執行)
// 警告：這只是為了方便建立表格，正式上線後建議移除或加強權限
app.get('/api/init-db', async (req, res) => {
  try {
    const client = await pool.connect();
    
    // 建立使用者表 (Users)
    await client.query(`
      CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        username VARCHAR(50) UNIQUE NOT NULL,
        password_hash VARCHAR(255) NOT NULL,
        role VARCHAR(20) DEFAULT 'user', -- 'admin' or 'user'
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `);

    // 建立事項表 (Issues) - 這裡先用通用欄位，之後可依你的 Excel 匯入欄位調整
    await client.query(`
      CREATE TABLE IF NOT EXISTS issues (
        id SERIAL PRIMARY KEY,
        title VARCHAR(255),
        content TEXT,
        status VARCHAR(50),
        created_by VARCHAR(50),
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `);

    client.release();
    res.send('資料庫表格初始化成功！(Users & Issues)');
  } catch (err) {
    console.error(err);
    res.status(500).send('初始化失敗: ' + err.message);
  }
});

// ... (原本的 GEMINI API 程式碼保持不變) ...
const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

app.post('/api/gemini', async (req, res) => {
  // ... (保留原本內容) ...
  // 注意：如果要將 GEMINI 的結果存入資料庫，未來可以在這裡加入 pool.query(...)
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

// 全域錯誤處理
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err && err.stack ? err.stack : err);
  const status = err && err.status ? err.status : 500;
  res.status(status).json({ error: err.message || 'Internal Server Error' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));