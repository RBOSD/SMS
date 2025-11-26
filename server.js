require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const { Pool } = require('pg');
const session = require('express-session');
const pgSession = require('connect-pg-simple')(session);
const bcrypt = require('bcrypt');
const axios = require('axios');

// [安全性修正] 檢查必要環境變數
if (!process.env.SESSION_SECRET) {
    console.error("❌ Critical Error: SESSION_SECRET is missing in environment variables.");
    process.exit(1);
}

const app = express();

const pool = new Pool({ 
    connectionString: process.env.DATABASE_URL, 
    ssl: { rejectUnauthorized: false } 
});

app.use(session({
  store: new pgSession({ pool: pool, tableName: 'session', createTableIfMissing: true }),
  secret: process.env.SESSION_SECRET, // [安全性修正] 移除預設值 'secret'
  resave: false, 
  saveUninitialized: false,
  cookie: { 
      maxAge: 30 * 24 * 60 * 60 * 1000, 
      httpOnly: true,
      // secure: true // [建議] 如果您的網站有 HTTPS (Render有)，建議加上這行，讓 Cookie 只能透過加密傳輸
  }
}));

app.use(cors({ credentials: true, origin: true }));

// [安全性修正] 限制請求大小，避免 DoS 攻擊 (視需求調整)
app.use(bodyParser.json({ limit: '10mb' })); 

// [架構修正] 設定靜態檔案目錄
app.use(express.static(path.join(__dirname, 'public')));

// ... (其餘路由程式碼保持不變)