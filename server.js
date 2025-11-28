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

console.log("🚀 Server starting...");

if (!process.env.SESSION_SECRET || !process.env.DATABASE_URL) {
    console.error("❌ Critical Error: SESSION_SECRET or DATABASE_URL is missing.");
    process.exit(1);
}

const app = express();
app.set('trust proxy', 1);

const pool = new Pool({ 
    connectionString: process.env.DATABASE_URL, 
    ssl: { rejectUnauthorized: false } 
});

pool.on('error', (err) => {
    console.error('❌ Database Pool Error:', err);
    process.exit(-1);
});

app.use(session({
  store: new pgSession({ pool: pool, tableName: 'session', createTableIfMissing: true }),
  secret: process.env.SESSION_SECRET,
  resave: false, 
  saveUninitialized: false,
  cookie: { maxAge: 30 * 24 * 60 * 60 * 1000, httpOnly: true, secure: true }
}));

app.use(cors({ credentials: true, origin: true }));
app.use(bodyParser.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const loginLimiter = rateLimit({
	windowMs: 15 * 60 * 1000,
	limit: 10,
	message: { error: '嘗試登入次數過多，請 15 分鐘後再試' },
	standardHeaders: true, 
	legacyHeaders: false, 
});

function checkAuth(req, res, next) { if (req.session.userId) next(); else res.status(401).json({ error: '請先登入' }); }
function checkAdmin(req, res, next) { if (req.session.role === 'admin') next(); else res.status(403).json({ error: '權限不足' }); }
function checkManager(req, res, next) { if (['admin', 'manager'].includes(req.session.role)) next(); else res.status(403).json({ error: '權限不足' }); }
function checkEditor(req, res, next) { if (['admin', 'manager', 'editor'].includes(req.session.role)) next(); else res.status(403).json({ error: '權限不足' }); }

// Helper: 寫入操作紀錄
async function logAction(userId, action, details, ip) {
    try {
        await pool.query(
            'INSERT INTO action_logs (user_id, action, details, ip_address) VALUES ($1, $2, $3, $4)',
            [userId, action, details, ip]
        );
    } catch (e) {
        console.error("Log Action Error:", e.message);
    }
}

async function initDB() {
    let client;
    try {
        client = await pool.connect();
        console.log("✅ Database connected.");

        await client.query(`
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY, 
                username VARCHAR(100) UNIQUE NOT NULL, 
                name VARCHAR(50), 
                password_hash VARCHAR(255) NOT NULL, 
                role VARCHAR(20) DEFAULT 'viewer', 
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                email VARCHAR(255) UNIQUE
            );
        `);
        
        await client.query(`CREATE TABLE IF NOT EXISTS login_logs (id SERIAL PRIMARY KEY, user_id INTEGER, ip_address VARCHAR(50), login_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP);`);
        await client.query(`CREATE INDEX IF NOT EXISTS idx_login_logs_time ON login_logs(login_time DESC);`);

        // [新增] 詳細操作紀錄表
        await client.query(`
            CREATE TABLE IF NOT EXISTS action_logs (
                id SERIAL PRIMARY KEY,
                user_id INTEGER,
                action VARCHAR(50),
                details TEXT,
                ip_address VARCHAR(50),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        `);
        await client.query(`CREATE INDEX IF NOT EXISTS idx_action_logs_time ON action_logs(created_at DESC);`);

        await client.query(`CREATE TABLE IF NOT EXISTS issues (id SERIAL PRIMARY KEY, title VARCHAR(255), content TEXT, status VARCHAR(50), year VARCHAR(20), unit VARCHAR(50), category VARCHAR(50), inspection_category VARCHAR(50), division VARCHAR(50), handling TEXT, review TEXT, raw_data JSONB, created_by VARCHAR(50), created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);`);
        
        console.log("✅ Tables ready.");
    } catch (err) { console.error("InitDB Error:", err); } 
    finally { if (client) client.release(); }
}
initDB();

// --- API Routes ---

app.post('/api/auth/login', loginLimiter, async (req, res) => {
    const { username, password } = req.body;
    if (!username) return res.status(400).json({ error: '請輸入帳號' });

    try {
        const r = await pool.query('SELECT * FROM users WHERE username = $1', [username]);
        if (r.rows.length === 0) return res.status(401).json({ error: '帳號或密碼錯誤' });
        
        const user = r.rows[0];
        if (await bcrypt.compare(password, user.password_hash)) {
            req.session.userId = user.id; 
            req.session.username = user.username;
            req.session.name = user.name; 
            req.session.role = user.role;
            
            let ip = req.ip || req.headers['x-forwarded-for'] || req.connection.remoteAddress || 'unknown';
            if (ip === '::1') ip = '127.0.0.1';
            if (ip.includes(',')) ip = ip.split(',')[0].trim();

            pool.query('INSERT INTO login_logs (user_id, ip_address) VALUES ($1, $2)', [user.id, ip])
                .catch(err => console.error("Login Log Error:", err.message));
            
            res.json({ success: true });
        } else {
            res.status(401).json({ error: '帳號或密碼錯誤' });
        }
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/auth/logout', (req, res) => { req.session.destroy(); res.json({ success: true }); });
app.get('/api/auth/me', (req, res) => {
    if (req.session.userId) res.json({ isLogin: true, ...req.session });
    else res.json({ isLogin: false });
});

const ROLE_MAP = { 'admin': '系統管理員', 'manager': '資料管理者', 'editor': '審查人員', 'viewer': '檢視人員' };
app.get('/api/users', checkAuth, checkAdmin, async (req, res) => {
    const r = await pool.query('SELECT id, username, email, name, role, created_at FROM users ORDER BY id ASC');
    const data = r.rows.map(u => ({ ...u, role_display: ROLE_MAP[u.role] || u.role }));
    res.json(data);
});

app.post('/api/users', checkAuth, checkAdmin, async (req, res) => {
    const { username, name, password, role } = req.body;
    try {
        const hash = await bcrypt.hash(password, 10);
        await pool.query(`INSERT INTO users (username, name, password_hash, role) VALUES ($1, $2, $3, $4)`, [username, name, hash, role]);
        // Log Action
        logAction(req.session.userId, 'CREATE_USER', `建立使用者: ${username} (${role})`, req.ip);
        res.json({ success: true });
    } catch (e) { 
        if(e.code === '23505') return res.status(400).json({error:'帳號已存在'}); 
        res.status(400).json({ error: '建立失敗' }); 
    }
});

app.put('/api/users/:id', checkAuth, checkAdmin, async (req, res) => {
    const { name, role, password } = req.body;
    let sql = 'UPDATE users SET name=$1, role=$2', params = [name, role];
    if (password) { sql += ', password_hash=$3'; params.push(await bcrypt.hash(password, 10)); }
    sql += ` WHERE id=$${params.length+1}`; params.push(req.params.id);
    await pool.query(sql, params);
    // Log Action
    logAction(req.session.userId, 'UPDATE_USER', `更新使用者 ID: ${req.params.id}`, req.ip);
    res.json({ success: true });
});

app.delete('/api/users/:id', checkAuth, checkAdmin, async (req, res) => {
    if(parseInt(req.params.id) === req.session.userId) return res.status(400).json({ error: '不能刪除自己' });
    await pool.query('DELETE FROM users WHERE id=$1', [req.params.id]);
    // Log Action
    logAction(req.session.userId, 'DELETE_USER', `刪除使用者 ID: ${req.params.id}`, req.ip);
    res.json({ success: true });
});

app.get('/api/admin/logs', checkAuth, checkAdmin, async (req, res) => {
    try {
        const r = await pool.query(`
            SELECT l.login_time, l.ip_address, u.username, u.name 
            FROM login_logs l
            LEFT JOIN users u ON l.user_id = u.id
            ORDER BY l.login_time DESC 
            LIMIT 100
        `);
        res.json(r.rows);
    } catch (e) { res.status(500).json({ error: '無法讀取紀錄' }); }
});

// [新增] 讀取操作紀錄 API
app.get('/api/admin/action_logs', checkAuth, checkAdmin, async (req, res) => {
    try {
        const r = await pool.query(`
            SELECT a.created_at, a.action, a.details, a.ip_address, u.username, u.name 
            FROM action_logs a
            LEFT JOIN users u ON a.user_id = u.id
            ORDER BY a.created_at DESC 
            LIMIT 100
        `);
        res.json(r.rows);
    } catch (e) { res.status(500).json({ error: '無法讀取操作紀錄' }); }
});


app.get('/api/issues', checkAuth, async (req, res) => {
    try {
        const r = await pool.query('SELECT * FROM issues ORDER BY created_at DESC');
        const data = r.rows.map(row => ({ ...(row.raw_data || {}), ...row, id: String(row.id) }));
        res.json(data);
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/issues/batch-delete', checkAuth, checkManager, async (req, res) => {
    const { ids } = req.body; 
    if (!ids || !Array.isArray(ids) || ids.length === 0) return res.status(400).json({ error: '未選擇項目' });
    try { 
        await pool.query('DELETE FROM issues WHERE id = ANY($1::int[])', [ids]); 
        // Log Action
        logAction(req.session.userId, 'BATCH_DELETE', `批次刪除 ${ids.length} 筆資料`, req.ip);
        res.json({ success: true, message: `已刪除 ${ids.length} 筆資料` }); 
    } catch (e) { res.status(500).json({ error: '批次刪除失敗' }); }
});

app.delete('/api/issues/:id', checkAuth, checkManager, async (req, res) => { 
    await pool.query('DELETE FROM issues WHERE id=$1', [req.params.id]); 
    // Log Action
    logAction(req.session.userId, 'DELETE_ISSUE', `刪除事項 ID: ${req.params.id}`, req.ip);
    res.json({ success: true }); 
});

app.put('/api/issues/:id', checkAuth, checkEditor, async (req, res) => {
    const { status, round, handling, review } = req.body; 
    const id = req.params.id;
    const r = await pool.query('SELECT * FROM issues WHERE id=$1', [id]); 
    if (r.rows.length === 0) return res.status(404).json({ error: '找不到事項' });

    let raw = r.rows[0].raw_data || {}; 
    raw.status = status; 
    const suffix = parseInt(round) === 1 ? '' : round;
    raw['handling'+suffix] = handling; 
    raw['review'+suffix] = review;
    
    let sql = 'UPDATE issues SET status=$1, raw_data=$2'; 
    let params = [status, JSON.stringify(raw)];
    
    if(parseInt(round) === 1) { 
        sql += ', handling=$3, review=$4'; 
        params.push(handling, review); 
    }
    sql += ` WHERE id=$${params.length+1}`; 
    params.push(id); 
    
    await pool.query(sql, params); 
    // Log Action
    logAction(req.session.userId, 'UPDATE_ISSUE', `更新事項 ID: ${id}, 回合: ${round}, 狀態: ${status}`, req.ip);
    res.json({ success: true });
});

app.post('/api/issues/import', checkAuth, checkManager, async (req, res) => {
    const { data, round, reviewDate, actualReplyDate } = req.body; 
    const targetRound = parseInt(round || 1); 
    const suffix = targetRound === 1 ? '' : targetRound;
    const client = await pool.connect(); 
    try { 
        await client.query('BEGIN'); 
        let countNew = 0, countUpdate = 0;
        for (const item of data) {
            const check = await client.query('SELECT id, raw_data FROM issues WHERE title = $1', [item.number]);
            const newHandling = item.handling || ''; 
            const newReview = item.review || ''; 
            const newStatus = item.status || ''; 
            const newContent = item.content || '';
            if (check.rows.length > 0) {
                countUpdate++;
                const existing = check.rows[0]; 
                let raw = existing.raw_data || {}; 
                raw['handling'+suffix] = newHandling; 
                raw['review'+suffix] = newReview; 
                raw['round'+targetRound+'Date'] = reviewDate; 
                raw['round'+targetRound+'ActualDate'] = actualReplyDate; 
                raw.status = newStatus;
                let sql = 'UPDATE issues SET status=$1, raw_data=$2, created_at=CURRENT_TIMESTAMP'; 
                let params = [newStatus, JSON.stringify(raw)];
                if (targetRound === 1) { 
                    sql += ', content=$3, year=$4, unit=$5, handling=$6, review=$7'; 
                    params.push(newContent, item.year, item.unit, newHandling, newReview); 
                }
                sql += ` WHERE id=$${params.length+1}`; 
                params.push(existing.id); 
                await client.query(sql, params);
            } else {
                countNew++;
                let raw = { ...item }; 
                raw['handling'+suffix] = newHandling; 
                raw['review'+suffix] = newReview; 
                raw['round'+targetRound+'Date'] = reviewDate;
                await client.query('INSERT INTO issues (title, content, status, year, unit, handling, review, raw_data) VALUES ($1, $2, $3, $4, $5, $6, $7, $8)', 
                    [item.number, newContent, newStatus, item.year, item.unit, (targetRound===1?newHandling:''), (targetRound===1?newReview:''), JSON.stringify(raw)]);
            }
        } 
        await client.query('COMMIT'); 
        // Log Action
        logAction(req.session.userId, 'IMPORT', `匯入資料: 新增 ${countNew}, 更新 ${countUpdate}, 回合: ${targetRound}`, req.ip);
        res.json({ success: true });
    } catch (e) { 
        await client.query('ROLLBACK'); 
        res.status(500).json({ error: e.message }); 
    } finally { client.release(); }
});

const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
app.post('/api/gemini', checkAuth, checkEditor, async (req, res) => {
    const { content, rounds } = req.body; 
    if (!GEMINI_API_KEY) return res.status(500).json({ error: "No API Key configured" });

    const prompt = `
    Role: 監理機關審查人員.
    Task: 針對「開立事項 (Finding)」審查營運機構回報的「辦理情形 (Action)」。
    開立事項: ${content}
    辦理情形: ${JSON.stringify(rounds)}
    請判斷辦理情形是否足以解除列管。
    語氣: 中性、冷靜、公務化。
    Output Format: JSON ONLY. Example: { "fulfill": "是/否", "reason": "審查意見" }
    `;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;
        const r = await axios.post(url, { 
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: { responseMimeType: "application/json" }
        });
        
        let txt = r.data.candidates[0].content.parts[0].text;
        const jsonMatch = txt.match(/{[\s\S]*}/);
        if (jsonMatch) {
             try { res.json(JSON.parse(jsonMatch[0])); } catch (e) { res.json({ fulfill: "失敗", reason: "AI 格式錯誤" }); }
        } else { 
            res.json({ fulfill: "失敗", reason: "AI 未回傳 JSON" }); 
        }
    } catch (e) { 
        console.error("AI Error:", e.message);
        res.status(500).json({ error: "AI 連線失敗" }); 
    }
});

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Server running on ${PORT}`));