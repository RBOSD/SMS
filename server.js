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

const app = express();

const pool = new Pool({ 
    connectionString: process.env.DATABASE_URL, 
    ssl: { rejectUnauthorized: false } 
});

app.use(session({
  store: new pgSession({ pool: pool, tableName: 'session', createTableIfMissing: true }),
  secret: process.env.SESSION_SECRET || 'secret',
  resave: false, saveUninitialized: false,
  cookie: { maxAge: 30 * 24 * 60 * 60 * 1000, httpOnly: true }
}));

app.use(cors({ credentials: true, origin: true }));
app.use(bodyParser.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// 確保首頁能被讀取 (如果 index.html 也在 public 內)
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

function checkAuth(req, res, next) { if (req.session.userId) next(); else res.status(401).json({ error: 'Unauthorized' }); }
function checkAdmin(req, res, next) { if (req.session.role === 'admin') next(); else res.status(403).json({ error: 'Forbidden' }); }
function checkReviewer(req, res, next) { if (['admin', 'editor'].includes(req.session.role)) next(); else res.status(403).json({ error: 'Forbidden' }); }

async function initDB() {
    const client = await pool.connect();
    try {
        await client.query(`CREATE TABLE IF NOT EXISTS users (id SERIAL PRIMARY KEY, username VARCHAR(100) UNIQUE NOT NULL, name VARCHAR(50), password_hash VARCHAR(255) NOT NULL, role VARCHAR(20) DEFAULT 'viewer', created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);`);
        await client.query(`CREATE TABLE IF NOT EXISTS login_logs (id SERIAL PRIMARY KEY, user_id INTEGER, ip_address VARCHAR(50), login_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP);`);
        await client.query(`CREATE TABLE IF NOT EXISTS issues (id SERIAL PRIMARY KEY, title VARCHAR(255), content TEXT, status VARCHAR(50), year VARCHAR(20), unit VARCHAR(50), category VARCHAR(50), inspection_category VARCHAR(50), division VARCHAR(50), handling TEXT, review TEXT, raw_data JSONB, created_by VARCHAR(50), created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);`);
        try { await client.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS name VARCHAR(50);`); } catch(e){}
    } catch (err) { console.error(err); } finally { client.release(); }
}
initDB();

app.post('/api/auth/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        const r = await pool.query('SELECT * FROM users WHERE username = $1', [username]);
        if (r.rows.length === 0) return res.status(401).json({ error: '帳號不存在' });
        const user = r.rows[0];
        if (await bcrypt.compare(password, user.password_hash)) {
            req.session.userId = user.id; req.session.username = user.username; req.session.name = user.name; req.session.role = user.role;
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
    if (req.session.userId) res.json({ isLogin: true, username: req.session.username, name: req.session.name || req.session.username, role: req.session.role });
    else res.json({ isLogin: false });
});

app.get('/api/users', checkAuth, checkAdmin, async (req, res) => {
    const r = await pool.query('SELECT id, username, name, role, created_at FROM users ORDER BY id ASC');
    res.json(r.rows);
});
app.post('/api/users', async (req, res) => {
    const { username, name, password, role } = req.body;
    const count = await pool.query('SELECT count(*) FROM users');
    if (parseInt(count.rows[0].count) > 0) {
        if (!req.session.userId || req.session.role !== 'admin') return res.status(403).json({ error: 'Forbidden' });
    }
    try {
        const hash = await bcrypt.hash(password, 10);
        const finalRole = parseInt(count.rows[0].count) === 0 ? 'admin' : role;
        await pool.query('INSERT INTO users (username, name, password_hash, role) VALUES ($1, $2, $3, $4)', [username, name, hash, finalRole]);
        res.json({ success: true });
    } catch (e) { res.status(400).json({ error: 'Error' }); }
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
    if(parseInt(req.params.id) === req.session.userId) return res.status(400).json({ error: 'Cannot delete self' });
    await pool.query('DELETE FROM users WHERE id=$1', [req.params.id]);
    res.json({ success: true });
});
app.get('/api/admin/logs', checkAuth, checkAdmin, async(req,res)=>{
    const r = await pool.query(`SELECT l.id, l.login_time, l.ip_address, u.name, u.username FROM login_logs l LEFT JOIN users u ON l.user_id=u.id ORDER BY l.login_time DESC LIMIT 50`);
    res.json(r.rows);
});

app.get('/api/issues', checkAuth, async (req, res) => {
    try {
        const r = await pool.query('SELECT * FROM issues ORDER BY created_at DESC');
        const data = r.rows.map(row => {
            const raw = row.raw_data || {};
            return { ...raw, ...row, id: String(row.id) };
        });
        res.json(data);
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/issues/:id', checkAuth, checkReviewer, async (req, res) => {
    await pool.query('DELETE FROM issues WHERE id=$1', [req.params.id]);
    res.json({ success: true });
});

app.put('/api/issues/:id', checkAuth, checkReviewer, async (req, res) => {
    const { status, round, handling, review } = req.body;
    const id = req.params.id;
    const r = await pool.query('SELECT * FROM issues WHERE id=$1', [id]);
    let raw = r.rows[0].raw_data || {};
    raw.status = status;
    const suffix = parseInt(round) === 1 ? '' : round;
    raw['handling'+suffix] = handling;
    raw['review'+suffix] = review;
    let sql = 'UPDATE issues SET status=$1, raw_data=$2';
    let params = [status, JSON.stringify(raw)];
    if(parseInt(round) === 1) { sql += ', handling=$3, review=$4'; params.push(handling, review); }
    sql += ` WHERE id=$${params.length+1}`; params.push(id);
    await pool.query(sql, params);
    res.json({ success: true });
});

app.post('/api/issues/import', checkAuth, checkReviewer, async (req, res) => {
    const { data, round, reviewDate, actualReplyDate } = req.body;
    const targetRound = parseInt(round || 1);
    const suffix = targetRound === 1 ? '' : targetRound;
    const client = await pool.connect();
    try {
        await client.query('BEGIN');
        for (const item of data) {
            const check = await client.query('SELECT id, raw_data FROM issues WHERE title = $1', [item.number]);
            const newHandling = item.handling || '';
            const newReview = item.review || '';
            const newStatus = item.status || '';
            const newContent = item.content || '';
            if (check.rows.length > 0) {
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
                let raw = { ...item };
                raw['handling'+suffix] = newHandling;
                raw['review'+suffix] = newReview;
                raw['round'+targetRound+'Date'] = reviewDate;
                await client.query(
                    'INSERT INTO issues (title, content, status, year, unit, handling, review, raw_data) VALUES ($1, $2, $3, $4, $5, $6, $7, $8)',
                    [item.number, newContent, newStatus, item.year, item.unit, (targetRound===1?newHandling:''), (targetRound===1?newReview:''), JSON.stringify(raw)]
                );
            }
        }
        await client.query('COMMIT'); res.json({ success: true });
    } catch (e) { await client.query('ROLLBACK'); res.status(500).json({ error: e.message }); } finally { client.release(); }
});

// ★ AI FIX: 使用 gemini-2.0-flash 並設定嚴格監理機關語氣
const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
app.post('/api/gemini', checkAuth, checkReviewer, async (req, res) => {
    const { content, rounds } = req.body;
    if (!GEMINI_API_KEY) return res.status(500).json({ error: "No API Key" });

    const prompt = `Role: 監理機關審查人員 (Regulatory Authority Officer).
    Task: 審查營運機構對於「開立事項」的「辦理情形」是否合規。
    
    開立事項 (Finding): ${content}
    辦理情形 (Action): ${JSON.stringify(rounds)}

    【語氣要求 - 重要】
    1. 絕對中性、冷靜、公務化 (Strictly Neutral & Bureaucratic)。
    2. 禁止使用稱讚詞彙：不可出現「良好」、「完善」、「感謝」、「優秀」、「不錯」等語。
    3. 審查結論用語範例：
       - 合格時：請用「擬予同意解除列管」、「所陳報改善措施尚屬合規，擬予備查」、「經檢視佐證資料，擬予解除列管」。
       - 不合格時：請用「請補充...」、「說明尚欠具體」、「請提供...之佐證資料」、「尚待釐清...」。

    Output Format: Provide ONLY a JSON object (No Markdown):
    { "fulfill": "是" or "否", "reason": "以繁體中文撰寫的監理機關審查意見" }`;

    try {
        // ★ 模型名稱依照您的要求設定為 gemini-2.0-flash
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;
        const r = await axios.post(url, { contents: [{ parts: [{ text: prompt }] }] });
        
        let txt = r.data.candidates[0].content.parts[0].text;
        
        // ★ 強力清洗 JSON
        const jsonMatch = txt.match(/{[\s\S]*}/);
        if (jsonMatch) {
            try {
                const result = JSON.parse(jsonMatch[0]);
                res.json(result);
            } catch (parseError) {
                res.json({ fulfill: "解析失敗", reason: "JSON 格式錯誤: " + txt });
            }
        } else {
            res.json({ fulfill: "解析失敗", reason: "AI 未回傳 JSON: " + txt });
        }

    } catch (e) {
        const msg = e.response?.data?.error?.message || e.message;
        console.error("AI Error:", msg);
        res.status(500).json({ error: msg });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Running on ${PORT}`));