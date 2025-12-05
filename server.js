const express = require('express');
const { Pool } = require('pg'); 
const bodyParser = require('body-parser');
const path = require('path');
const session = require('express-session');
const bcrypt = require('bcrypt');
const { GoogleGenerativeAI } = require("@google/generative-ai"); // AI 套件
require('dotenv').config(); 

const app = express();

// [重要] 信任 Proxy，解決 Render 上 Session 失效問題
app.set('trust proxy', 1); 

const PORT = process.env.PORT || 3000;

// 初始化 PostgreSQL 連線池
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false
});

app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Session 設定
app.use(session({
    secret: 'sms-secret-key-pg-final-v2',
    resave: false,
    saveUninitialized: false,
    proxy: true, // 允許 Proxy
    cookie: { 
        maxAge: 24 * 60 * 60 * 1000,
        secure: process.env.NODE_ENV === 'production', // HTTPS 必開
        sameSite: 'lax' 
    } 
}));

// 權限檢查 Middleware
const requireAuth = (req, res, next) => {
    if (req.session && req.session.user) {
        next();
    } else {
        res.status(401).json({ error: 'Unauthorized' });
    }
};

// --- 資料庫初始化 ---
async function initDB() {
    const client = await pool.connect();
    try {
        console.log('Connected to PostgreSQL. Checking schema...');

        // 1. 建立主表
        await client.query(`CREATE TABLE IF NOT EXISTS issues (
            id SERIAL PRIMARY KEY,
            number TEXT UNIQUE,
            year TEXT,
            unit TEXT,
            content TEXT,
            status TEXT,
            item_kind_code TEXT,
            division_name TEXT,
            inspection_category_name TEXT,
            category TEXT,
            handling TEXT,
            review TEXT,
            plan_name TEXT, 
            issue_date TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )`);

        // 2. 建立使用者表
        await client.query(`CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE,
            password TEXT,
            name TEXT,
            role TEXT DEFAULT 'viewer',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )`);

        // 3. 建立 Log 表
        await client.query(`CREATE TABLE IF NOT EXISTS logs (
            id SERIAL PRIMARY KEY,
            username TEXT,
            action TEXT,
            details TEXT,
            ip_address TEXT,
            login_time TIMESTAMP,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )`);

        // 4. 動態欄位補強 (確保 1~20 回合欄位都存在)
        const newColumns = [];
        for (let i = 2; i <= 20; i++) {
            newColumns.push({ name: `handling${i}`, type: 'TEXT' });
            newColumns.push({ name: `review${i}`, type: 'TEXT' });
        }
        for (let i = 1; i <= 20; i++) {
            newColumns.push({ name: `reply_date_r${i}`, type: 'TEXT' });
            newColumns.push({ name: `response_date_r${i}`, type: 'TEXT' });
        }
        newColumns.push({ name: 'plan_name', type: 'TEXT' });
        newColumns.push({ name: 'issue_date', type: 'TEXT' });

        for (const col of newColumns) {
            try {
                await client.query(`ALTER TABLE issues ADD COLUMN IF NOT EXISTS ${col.name} ${col.type}`);
            } catch (e) { /* 忽略欄位已存在 */ }
        }

        // 5. 建立預設 Admin (若無使用者)
        const userRes = await client.query("SELECT count(*) as count FROM users");
        if (parseInt(userRes.rows[0].count) === 0) {
            const hash = bcrypt.hashSync('admin123', 10);
            await client.query("INSERT INTO users (username, password, name, role) VALUES ($1, $2, $3, $4)", 
                ['admin', hash, '系統管理員', 'admin']);
            console.log("Default admin created.");
        }
        
        console.log('Database initialized.');

    } catch (err) {
        console.error('Init DB Error:', err);
        throw err;
    } finally {
        client.release();
    }
}

// Log 輔助
async function logAction(username, action, details, req) {
    const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    try {
        await pool.query("INSERT INTO logs (username, action, details, ip_address, created_at) VALUES ($1, $2, $3, $4, CURRENT_TIMESTAMP)", 
            [username, action, details, ip]);
    } catch (e) { console.error("Log error:", e); }
}

// --- API Routes ---

// 1. 登入
app.post('/api/auth/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        const result = await pool.query("SELECT * FROM users WHERE username = $1", [username]);
        const user = result.rows[0];
        
        if (!user || !user.password) {
            return res.status(401).json({ error: 'Invalid credentials' });
        }

        if (bcrypt.compareSync(password, user.password)) {
            req.session.user = { id: user.id, username: user.username, role: user.role, name: user.name };
            // 手動保存 Session 以確保寫入
            req.session.save((err) => {
                if(err) return res.status(500).json({error: 'Session error'});
                logAction(user.username, 'LOGIN', 'User logged in', req).catch(()=>{});
                res.json({ success: true, user: req.session.user });
            });
        } else {
            res.status(401).json({ error: 'Invalid credentials' });
        }
    } catch (e) {
        console.error("Login Error:", e);
        res.status(500).json({ error: 'System error' });
    }
});

app.post('/api/auth/logout', (req, res) => {
    req.session.destroy();
    res.json({ success: true });
});

app.get('/api/auth/me', (req, res) => {
    if (req.session.user) res.json({ isLogin: true, ...req.session.user });
    else res.json({ isLogin: false });
});

app.put('/api/auth/profile', requireAuth, async (req, res) => {
    const { name, password } = req.body;
    const id = req.session.user.id;
    try {
        if (password) {
            const hash = bcrypt.hashSync(password, 10);
            await pool.query("UPDATE users SET name = $1, password = $2 WHERE id = $3", [name, hash, id]);
        } else {
            await pool.query("UPDATE users SET name = $1 WHERE id = $2", [name, id]);
        }
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// 2. AI 審查 (Gemini Flash 2.5)
app.post('/api/gemini', async (req, res) => {
    const { content, rounds } = req.body;

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) return res.status(500).json({ error: '後端未設定 GEMINI_API_KEY' });

    try {
        const genAI = new GoogleGenerativeAI(apiKey);
		
                      // 使用Google gemini 2.5 flash    
        const model = genAI.getGenerativeModel({ model: "gemini-2.5-" });

        const latestRound = (rounds && rounds.length > 0) ? rounds[rounds.length - 1] : { handling: '無', review: '無' };
        
        const prompt = `
        你是一個鐵道監理機關的安全改善事項的專業審查人員。
        【事項內容】：${content}
        【最新辦理情形】：${latestRound.handling || '無'}
        
        請分析上述「辦理情形」是否已具體解決「事項內容」的問題，審查的內容須保持中性、客觀的語氣，可帶有一點公務機關常用文字。
        
        請嚴格依照以下 JSON 格式回覆 (不要 Markdown 標記)：
        {
            "fulfill": "Yes 或 No (Yes代表已完成, No代表未完成)",
            "reason": "請提供簡潔且說明清楚的審查意見"
        }
        `;

        const result = await model.generateContent(prompt);
        const response = await result.response;
        let text = response.text();
        text = text.replace(/```json/g, '').replace(/```/g, '').trim();

        try {
            const json = JSON.parse(text);
            res.json(json);
        } catch (parseError) {
            console.error("JSON Parse Error, raw text:", text);
            res.json({ fulfill: text.includes("Yes")?"Yes":"No", reason: text });
        }

    } catch (e) {
        console.error("Gemini API Error:", e);
        // 回傳 500 讓前端知道失敗，而不是回傳 HTML
        res.status(500).json({ error: 'AI 分析失敗: ' + e.message });
    }
});

// 3. 事項查詢
app.get('/api/issues', requireAuth, async (req, res) => {
    const { page = 1, pageSize = 20, q, year, unit, status, itemKindCode, division, inspectionCategory, sortField, sortDir } = req.query;
    const limit = parseInt(pageSize);
    const offset = (page - 1) * limit;
    
    let where = ["1=1"], params = [], idx = 1;

    if (q) {
        where.push(`(number LIKE $${idx} OR content LIKE $${idx} OR handling LIKE $${idx} OR review LIKE $${idx} OR plan_name LIKE $${idx})`);
        params.push(`%${q}%`); idx++;
    }
    if (year) { where.push(`year = $${idx}`); params.push(year); idx++; }
    if (unit) { where.push(`unit = $${idx}`); params.push(unit); idx++; }
    if (status) { where.push(`status = $${idx}`); params.push(status); idx++; }
    if (itemKindCode) { where.push(`item_kind_code = $${idx}`); params.push(itemKindCode); idx++; }
    if (division) { where.push(`division_name = $${idx}`); params.push(division); idx++; }
    if (inspectionCategory) { where.push(`inspection_category_name = $${idx}`); params.push(inspectionCategory); idx++; }

    let orderBy = "created_at DESC";
    const validCols = ['year', 'number', 'unit', 'status', 'created_at'];
    if (sortField && validCols.includes(sortField)) {
        orderBy = `${sortField} ${sortDir === 'asc' ? 'ASC' : 'DESC'}`;
    }

    try {
        const countRes = await pool.query(`SELECT count(*) FROM issues WHERE ${where.join(" AND ")}`, params);
        const total = parseInt(countRes.rows[0].count);
        
        const dataRes = await pool.query(`SELECT * FROM issues WHERE ${where.join(" AND ")} ORDER BY ${orderBy} LIMIT $${idx} OFFSET $${idx+1}`, [...params, limit, offset]);
        
        const sRes = await pool.query("SELECT status, count(*) as count FROM issues GROUP BY status");
        const uRes = await pool.query("SELECT unit, count(*) as count FROM issues GROUP BY unit");
        const yRes = await pool.query("SELECT year, count(*) as count FROM issues GROUP BY year");
        const tRes = await pool.query("SELECT max(created_at) as latest, max(updated_at) as updated FROM issues");
        const latestTime = tRes.rows[0] ? (tRes.rows[0].updated || tRes.rows[0].latest) : null;

        res.json({
            data: dataRes.rows,
            total,
            page: parseInt(page),
            pageSize: limit,
            pages: Math.ceil(total / limit),
            latestCreatedAt: latestTime,
            globalStats: { status: sRes.rows, unit: uRes.rows, year: yRes.rows }
        });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/issues/:id', requireAuth, async (req, res) => {
    const { status, round, handling, review, replyDate, responseDate } = req.body;
    const id = req.params.id;
    const r = parseInt(round);
    const hField = r === 1 ? 'handling' : `handling${r}`;
    const rField = r === 1 ? 'review' : `review${r}`;
    const replyField = `reply_date_r${r}`;
    const respField = `response_date_r${r}`;

    try {
        await pool.query(`UPDATE issues SET status=$1, ${hField}=$2, ${rField}=$3, ${replyField}=$4, ${respField}=$5, updated_at=CURRENT_TIMESTAMP WHERE id=$6`, 
            [status, handling, review, replyDate, responseDate, id]);
        logAction(req.session.user.username, 'UPDATE', `Updated issue ${id}`, req);
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/issues/:id', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    try {
        await pool.query("DELETE FROM issues WHERE id=$1", [req.params.id]);
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/issues/batch-delete', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    const { ids } = req.body;
    try {
        await pool.query("DELETE FROM issues WHERE id = ANY($1)", [ids]);
        logAction(req.session.user.username, 'BATCH_DELETE', `Deleted ${ids.length} items`, req);
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/issues/import', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    const { data, round, reviewDate } = req.body;
    const r = parseInt(round) || 1;
    
    const client = await pool.connect();
    try {
        await client.query('BEGIN');

        for (const item of data) {
            const check = await client.query("SELECT id FROM issues WHERE number = $1", [item.number]);
            
            if (check.rows.length > 0) {
                const hCol = r===1 ? 'handling' : `handling${r}`;
                const rCol = r===1 ? 'review' : `review${r}`;
                const respCol = `response_date_r${r}`;
                
                await client.query(
                    `UPDATE issues SET status=$1, ${hCol}=$2, ${rCol}=$3, ${respCol}=$4, plan_name=COALESCE($5, plan_name), updated_at=CURRENT_TIMESTAMP WHERE number=$6`,
                    [
                        item.status, item.handling||'', item.review||'',
                        reviewDate||'', 
                        item.planName || null,
                        item.number
                    ]
                );
            } else {
                await client.query(
                    `INSERT INTO issues (
                        number, year, unit, content, status, item_kind_code, category, division_name, inspection_category_name,
                        handling, review, plan_name, issue_date, response_date_r1
                    ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14)`,
                    [
                        item.number, item.year, item.unit, item.content, item.status||'持續列管',
                        item.itemKindCode, item.category, item.divisionName, item.inspectionCategoryName,
                        item.handling||'', item.review||'', 
                        item.planName || null, 
                        item.issueDate || null,
                        reviewDate || '' 
                    ]
                );
            }
        }
        await client.query('COMMIT');
        logAction(req.session.user.username, 'IMPORT', `Imported ${data.length} items`, req);
        res.json({ success: true, count: data.length });
    } catch (e) {
        await client.query('ROLLBACK');
        res.status(500).json({ error: e.message });
    } finally {
        client.release();
    }
});

// Users
app.get('/api/users', requireAuth, async (req, res) => {
    if (req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    const { page=1, pageSize=20, q, sortField='id', sortDir='asc' } = req.query;
    const limit = parseInt(pageSize);
    const offset = (page-1)*limit;
    
    let where = ["1=1"], params = [], idx = 1;
    if(q) { where.push(`(username LIKE $${idx} OR name LIKE $${idx})`); params.push(`%${q}%`); idx++; }
    const order = `${sortField} ${sortDir==='desc'?'DESC':'ASC'}`;
    
    try {
        const cRes = await pool.query(`SELECT count(*) FROM users WHERE ${where.join(" AND ")}`, params);
        const total = parseInt(cRes.rows[0].count);
        const dRes = await pool.query(`SELECT id, username, name, role, created_at FROM users WHERE ${where.join(" AND ")} ORDER BY ${order} LIMIT $${idx} OFFSET $${idx+1}`, [...params, limit, offset]);
        res.json({data:dRes.rows, total, page: parseInt(page), pages: Math.ceil(total/limit)});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/users', requireAuth, async (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    const { username, password, name, role } = req.body;
    try {
        const hash = bcrypt.hashSync(password, 10);
        await pool.query("INSERT INTO users (username, password, name, role) VALUES ($1, $2, $3, $4)", [username, hash, name, role]);
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/users/:id', requireAuth, async (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    const { name, password, role } = req.body;
    const id = req.params.id;
    try {
        if (password) {
            const hash = bcrypt.hashSync(password, 10);
            await pool.query("UPDATE users SET name=$1, role=$2, password=$3 WHERE id=$4", [name, role, hash, id]);
        } else {
            await pool.query("UPDATE users SET name=$1, role=$2 WHERE id=$3", [name, role, id]);
        }
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/users/:id', requireAuth, async (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    if(parseInt(req.params.id) === req.session.user.id) return res.status(400).json({error:'Cannot self delete'});
    try {
        await pool.query("DELETE FROM users WHERE id=$1", [req.params.id]);
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/admin/logs', requireAuth, async (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    try {
        const { rows } = await pool.query("SELECT * FROM logs WHERE action='LOGIN' ORDER BY login_time DESC LIMIT 50");
        res.json({data:rows, total:rows.length, page:1, pages:1});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/admin/action_logs', requireAuth, async (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    try {
        const { rows } = await pool.query("SELECT * FROM logs WHERE action!='LOGIN' ORDER BY created_at DESC LIMIT 50");
        res.json({data:rows, total:rows.length, page:1, pages:1});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/admin/logs', requireAuth, async (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    await pool.query("DELETE FROM logs WHERE action='LOGIN'");
    res.json({success:true});
});

app.delete('/api/admin/action_logs', requireAuth, async (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    await pool.query("DELETE FROM logs WHERE action!='LOGIN'");
    res.json({success:true});
});

// 穩定啟動
async function startServer() {
    try {
        await initDB();
        app.listen(PORT, () => {
            console.log(`Server running on http://localhost:${PORT}`);
        });
    } catch (e) {
        console.error("Server start failed:", e);
        process.exit(1);
    }
}

startServer();