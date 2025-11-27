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
const rateLimit = require('express-rate-limit');

console.log("🚀 Server starting (Username Login Mode)...");

// 環境變數檢查
if (!process.env.SESSION_SECRET || !process.env.DATABASE_URL) {
    console.error("❌ Critical Error: SESSION_SECRET or DATABASE_URL is missing.");
    process.exit(1);
}

const app = express();

// [關鍵設定] 解決 Render 上的 Rate Limit 錯誤
app.set('trust proxy', 1);

const pool = new Pool({ 
    connectionString: process.env.DATABASE_URL, 
    ssl: { rejectUnauthorized: false } 
});

pool.on('error', (err) => {
    console.error('❌ Database Pool Error:', err);
    process.exit(-1);
});

// Middleware
app.use(session({
  store: new pgSession({ pool: pool, tableName: 'session', createTableIfMissing: true }),
  secret: process.env.SESSION_SECRET,
  resave: false, saveUninitialized: false,
  cookie: { maxAge: 30 * 24 * 60 * 60 * 1000, httpOnly: true, secure: true }
}));

app.use(cors({ credentials: true, origin: true }));
app.use(bodyParser.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// [保留] 暴力破解防護 (針對 username)
const loginLimiter = rateLimit({
	windowMs: 15 * 60 * 1000, // 15 分鐘
	limit: 10, // 允許錯誤 10 次
	message: { error: '嘗試次數過多，請 15 分鐘後再試' },
	standardHeaders: true, 
	legacyHeaders: false, 
});

// 權限檢查
function checkAuth(req, res, next) { if (req.session.userId) next(); else res.status(401).json({ error: '請先登入' }); }
function checkAdmin(req, res, next) { if (req.session.role === 'admin') next(); else res.status(403).json({ error: '權限不足' }); }
function checkManager(req, res, next) { if (['admin', 'manager'].includes(req.session.role)) next(); else res.status(403).json({ error: '權限不足' }); }
function checkEditor(req, res, next) { if (['admin', 'manager', 'editor'].includes(req.session.role)) next(); else res.status(403).json({ error: '權限不足' }); }

// DB 初始化 (維持原樣，確保 username 存在)
async function initDB() {
    let client;
    try {
        client = await pool.connect();
        await client.query(`
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY, 
                username VARCHAR(100) UNIQUE NOT NULL, 
                name VARCHAR(50), 
                password_hash VARCHAR(255) NOT NULL, 
                role VARCHAR(20) DEFAULT 'viewer', 
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        `);
        // 嘗試補 email 欄位但不強制使用
        try { await client.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS email VARCHAR(255) UNIQUE;`); } catch(e){}
        
        await client.query(`CREATE TABLE IF NOT EXISTS login_logs (id SERIAL PRIMARY KEY, user_id INTEGER, ip_address VARCHAR(50), login_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP);`);
        await client.query(`CREATE TABLE IF NOT EXISTS issues (id SERIAL PRIMARY KEY, title VARCHAR(255), content TEXT, status VARCHAR(50), year VARCHAR(20), unit VARCHAR(50), category VARCHAR(50), inspection_category VARCHAR(50), division VARCHAR(50), handling TEXT, review TEXT, raw_data JSONB, created_by VARCHAR(50), created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);`);
        console.log("✅ Tables ready.");
    } catch (err) { console.error("InitDB Error:", err); } 
    finally { if (client) client.release(); }
}
initDB();

// --- API Routes ---

// [還原] 傳統 Username 登入
app.post('/api/auth/login', loginLimiter, async (req, res) => {
    const { username, password } = req.body; // 只收 username
    
    if (!username) return res.status(400).json({ error: '請輸入使用者名稱' });
    
    try {
        // 直接查 username，完全不經過 email 邏輯
        const r = await pool.query('SELECT * FROM users WHERE username = $1', [username]);
        if (r.rows.length === 0) return res.status(401).json({ error: '帳號不存在' });
        
        const user = r.rows[0];
        const match = await bcrypt.compare(password, user.password_hash);
        
        if (match) {
            req.session.userId = user.id; 
            req.session.username = user.username;
            req.session.name = user.name; 
            req.session.role = user.role;
            
            const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
            pool.query('INSERT INTO login_logs (user_id, ip_address) VALUES ($1, $2)', [user.id, ip]).catch(()=>{});
            
            res.json({ success: true });
        } else {
            res.status(401).json({ error: '密碼錯誤' });
        }
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/auth/logout', (req, res) => { req.session.destroy(); res.json({ success: true }); });
app.get('/api/auth/me', (req, res) => {
    if (req.session.userId) res.json({ isLogin: true, ...req.session });
    else res.json({ isLogin: false });
});

// 使用者管理 (保留中文職稱顯示)
const ROLE_MAP = { 'admin': '系統管理員', 'manager': '資料管理者', 'editor': '審查人員', 'viewer': '檢視人員' };
app.get('/api/users', checkAuth, checkAdmin, async (req, res) => {
    const r = await pool.query('SELECT * FROM users ORDER BY id ASC');
    const data = r.rows.map(u => ({ ...u, role_display: ROLE_MAP[u.role] || u.role }));
    res.json(data);
});

app.post('/api/users', checkAuth, checkAdmin, async (req, res) => {
    // 恢復為 username 必填
    const { username, name, password, role, email } = req.body; 
    try {
        const hash = await bcrypt.hash(password, 10);
        // email 為選填
        await pool.query(
            `INSERT INTO users (username, name, password_hash, role, email) VALUES ($1, $2, $3, $4, $5)`, 
            [username, name, hash, role, email || null]
        );
        res.json({ success: true });
    } catch (e) { 
        if(e.code === '23505') return res.status(400).json({error:'帳號名稱已存在'}); 
        res.status(400).json({ error: '建立失敗' }); 
    }
});

app.put('/api/users/:id', checkAuth, checkAdmin, async (req, res) => {
    const { name, role, password } = req.body;
    let sql = 'UPDATE users SET name=$1, role=$2', params = [name, role];
    if (password) { sql += ', password_hash=$3'; params.push(await bcrypt.hash(password, 10)); }
    sql += ` WHERE id=$${params.length+1}`; params.push(req.params.id);
    await pool.query(sql, params);
    res.json({ success: true });
});
app.delete('/api/users/:id', checkAuth, checkAdmin, async (req, res) => {
    if(parseInt(req.params.id) === req.session.userId) return res.status(400).json({ error: '不能刪除自己' });
    await pool.query('DELETE FROM users WHERE id=$1', [req.params.id]);
    res.json({ success: true });
});

// 事項管理 (保留批次刪除與新權限)
app.get('/api/issues', checkAuth, async (req, res) => {
    try {
        const r = await pool.query('SELECT * FROM issues ORDER BY created_at DESC');
        const data = r.rows.map(row => ({ ...(row.raw_data || {}), ...row, id: String(row.id) }));
        res.json(data);
    } catch (e) { res.status(500).json({ error: e.message }); }
});
// [新增] 批次刪除
app.post('/api/issues/batch-delete', checkAuth, checkManager, async (req, res) => {
    const { ids } = req.body; 
    if (!ids || !Array.isArray(ids) || ids.length === 0) return res.status(400).json({ error: '未選擇' });
    try { await pool.query('DELETE FROM issues WHERE id = ANY($1::int[])', [ids]); res.json({ success: true }); } 
    catch (e) { res.status(500).json({ error: '刪除失敗' }); }
});
app.delete('/api/issues/:id', checkAuth, checkManager, async (req, res) => { await pool.query('DELETE FROM issues WHERE id=$1', [req.params.id]); res.json({ success: true }); });
app.put('/api/issues/:id', checkAuth, checkEditor, async (req, res) => {
    const { status, round, handling, review } = req.body; const id = req.params.id;
    const r = await pool.query('SELECT * FROM issues WHERE id=$1', [id]); if (r.rows.length === 0) return res.status(404).json({ error: '找不到' });
    let raw = r.rows[0].raw_data || {}; raw.status = status; const suffix = parseInt(round) === 1 ? '' : round;
    raw['handling'+suffix] = handling; raw['review'+suffix] = review;
    let sql = 'UPDATE issues SET status=$1, raw_data=$2'; let params = [status, JSON.stringify(raw)];
    if(parseInt(round) === 1) { sql += ', handling=$3, review=$4'; params.push(handling, review); }
    sql += ` WHERE id=$${params.length+1}`; params.push(id); await pool.query(sql, params); res.json({ success: true });
});
app.post('/api/issues/import', checkAuth, checkManager, async (req, res) => {
    const { data, round, reviewDate, actualReplyDate } = req.body; const targetRound = parseInt(round || 1); const suffix = targetRound === 1 ? '' : targetRound;
    const client = await pool.connect(); try { await client.query('BEGIN'); for (const item of data) {
            const check = await client.query('SELECT id, raw_data FROM issues WHERE title = $1', [item.number]);
            const newHandling = item.handling || ''; const newReview = item.review || ''; const newStatus = item.status || ''; const newContent = item.content || '';
            if (check.rows.length > 0) {
                const existing = check.rows[0]; let raw = existing.raw_data || {}; raw['handling'+suffix] = newHandling; raw['review'+suffix] = newReview; raw['round'+targetRound+'Date'] = reviewDate; raw['round'+targetRound+'ActualDate'] = actualReplyDate; raw.status = newStatus;
                let sql = 'UPDATE issues SET status=$1, raw_data=$2, created_at=CURRENT_TIMESTAMP'; let params = [newStatus, JSON.stringify(raw)];
                if (targetRound === 1) { sql += ', content=$3, year=$4, unit=$5, handling=$6, review=$7'; params.push(newContent, item.year, item.unit, newHandling, newReview); }
                sql += ` WHERE id=$${params.length+1}`; params.push(existing.id); await client.query(sql, params);
            } else {
                let raw = { ...item }; raw['handling'+suffix] = newHandling; raw['review'+suffix] = newReview; raw['round'+targetRound+'Date'] = reviewDate;
                await client.query('INSERT INTO issues (title, content, status, year, unit, handling, review, raw_data) VALUES ($1, $2, $3, $4, $5, $6, $7, $8)', [item.number, newContent, newStatus, item.year, item.unit, (targetRound===1?newHandling:''), (targetRound===1?newReview:''), JSON.stringify(raw)]);
            }
        } await client.query('COMMIT'); res.json({ success: true });
    } catch (e) { await client.query('ROLLBACK'); res.status(500).json({ error: e.message }); } finally { client.release(); }
});

const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
app.post('/api/gemini', checkAuth, checkEditor, async (req, res) => {
    const { content, rounds } = req.body; if (!GEMINI_API_KEY) return res.status(500).json({ error: "No API Key" });
    const prompt = `Role: 監理機關... Finding: ${content}, Action: ${JSON.stringify(rounds)} ... Output JSON.`;
    try { const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;
        const r = await axios.post(url, { contents: [{ parts: [{ text: prompt }] }] }); let txt = r.data.candidates[0].content.parts[0].text;
        const jsonMatch = txt.match(/{[\s\S]*}/); if (jsonMatch) { try { res.json(JSON.parse(jsonMatch[0])); } catch (e) { res.json({ fulfill: "失敗", reason: "JSON Err" }); } } else { res.json({ fulfill: "失敗", reason: "No JSON" }); }
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Server running on ${PORT}`));