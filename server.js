// [新增] 強制 Node.js 優先使用 IPv4 解析，解決 Render 連線 Supabase 的 ENETUNREACH 問題
require('dns').setDefaultResultOrder('ipv4first');

const express = require('express');
const { Pool } = require('pg'); 
const bodyParser = require('body-parser');
const path = require('path');
const session = require('express-session');
// [新增] 引入 pg-simple session store
const pgSession = require('connect-pg-simple')(session);
const bcrypt = require('bcrypt');
const { GoogleGenerativeAI } = require("@google/generative-ai"); 
require('dotenv').config(); 

const app = express();

app.set('trust proxy', 1); 

const PORT = process.env.PORT || 3000;

// [新增] 處理 DATABASE_URL，移除 sslmode 參數以避免覆蓋 ssl 設定
let databaseUrl = process.env.DATABASE_URL;
if (databaseUrl) {
    try {
        const url = new URL(databaseUrl);
        url.searchParams.delete('sslmode');
        databaseUrl = url.toString();
    } catch (e) {
        console.warn('Warning: Could not parse DATABASE_URL:', e.message);
    }
}

// [修改] 初始化 PostgreSQL 連線池
const pool = new Pool({
    connectionString: databaseUrl,
    ssl: { rejectUnauthorized: false }, // 允許自簽憑證
    max: 20, 
    idleTimeoutMillis: 30000,
    connectionTimeoutMillis: 5000, // 稍微拉長連線逾時時間
});

app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// [修改] Session 設定
app.use(session({
    store: new pgSession({
        pool: pool,
        tableName: 'session',
        createTableIfMissing: true 
    }),
    secret: process.env.SESSION_SECRET || 'sms-secret-key-pg-final-v3',
    resave: false,
    saveUninitialized: false,
    proxy: true, 
    cookie: { 
        maxAge: 30 * 24 * 60 * 60 * 1000, 
        secure: process.env.NODE_ENV === 'production',
        sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax'
    } 
}));

const requireAuth = (req, res, next) => {
    if (req.session && req.session.user) {
        next();
    } else {
        res.status(401).json({ error: 'Unauthorized' });
    }
};

// --- 資料庫初始化 ---
async function initDB() {
    // [修改] 增加重試機制，因為第一次連線有時會因為網路冷啟動而失敗
    let retries = 5;
    while (retries > 0) {
        try {
            const client = await pool.connect();
            try {
                console.log('Connected to Supabase PostgreSQL. Checking schema...');

                await client.query(`
                    CREATE TABLE IF NOT EXISTS session (
                        sid varchar NOT NULL COLLATE "default",
                        sess json NOT NULL,
                        expire timestamp(6) NOT NULL
                    ) WITH (OIDS=FALSE);
                `);
                try {
                    await client.query(`ALTER TABLE session ADD CONSTRAINT session_pkey PRIMARY KEY (sid) NOT DEFERRABLE INITIALLY IMMEDIATE`);
                } catch (e) {}
                try {
                    await client.query(`CREATE INDEX IF NOT EXISTS "IDX_session_expire" ON session (expire)`);
                } catch (e) {}

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

                await client.query(`CREATE TABLE IF NOT EXISTS users (
                    id SERIAL PRIMARY KEY,
                    username TEXT UNIQUE,
                    password TEXT,
                    name TEXT,
                    role TEXT DEFAULT 'viewer',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )`);

                await client.query(`CREATE TABLE IF NOT EXISTS logs (
                    id SERIAL PRIMARY KEY,
                    username TEXT,
                    action TEXT,
                    details TEXT,
                    ip_address TEXT,
                    login_time TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )`);

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
                    } catch (e) { }
                }

                const userRes = await client.query("SELECT count(*) as count FROM users");
                if (parseInt(userRes.rows[0].count) === 0) {
                    const hash = bcrypt.hashSync('admin123', 10);
                    await client.query("INSERT INTO users (username, password, name, role) VALUES ($1, $2, $3, $4)", 
                        ['admin', hash, '系統管理員', 'admin']);
                    console.log("Default admin created.");
                }
                
                console.log('Database initialized successfully on Supabase.');
                return; // 初始化成功，跳出迴圈

            } catch (err) {
                console.error('Init DB Schema Error:', err);
                throw err;
            } finally {
                client.release();
            }
        } catch (connErr) {
            console.error(`Connection failed, retrying... (${retries} left)`, connErr.message);
            retries--;
            await new Promise(res => setTimeout(res, 2000)); // 等待 2 秒後重試
        }
    }
    throw new Error('Could not connect to database after multiple retries.');
}

async function logAction(username, action, details, req) {
    const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    try {
        await pool.query("INSERT INTO logs (username, action, details, ip_address, created_at) VALUES ($1, $2, $3, $4, CURRENT_TIMESTAMP)", 
            [username, action, details, ip]);
    } catch (e) { console.error("Log error:", e); }
}

// --- API Routes ---

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
            req.session.save((err) => {
                if(err) {
                    console.error("Session save error:", err);
                    return res.status(500).json({error: 'Session error'});
                }
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
    req.session.destroy((err) => {
        if (err) return res.status(500).json({ error: 'Logout failed' });
        res.clearCookie('connect.sid'); 
        res.json({ success: true });
    });
});

app.get('/api/auth/me', async (req, res) => {
    if (req.session && req.session.user) {
        try {
            const result = await pool.query("SELECT id, username, name, role FROM users WHERE id = $1", [req.session.user.id]);
            const latestUser = result.rows[0];

            if (!latestUser) {
                req.session.destroy();
                return res.json({ isLogin: false });
            }
            req.session.user = latestUser;
            res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
            res.json({ isLogin: true, ...latestUser });
        } catch (e) {
            console.error("Auth check db error:", e);
            res.json({ isLogin: false });
        }
    } else {
        res.json({ isLogin: false });
    }
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

app.post('/api/gemini', async (req, res) => {
    const { content, rounds } = req.body;
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) return res.status(500).json({ error: '後端未設定 GEMINI_API_KEY' });
    try {
        const genAI = new GoogleGenerativeAI(apiKey);
        const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });
        const latestRound = (rounds && rounds.length > 0) ? rounds[rounds.length - 1] : { handling: '無', review: '無' };
        const previousReview = (rounds && rounds.length > 1) ? rounds[rounds.length - 2].review : '無';
        const prompt = `
        你現在是【鐵道監理機關】的專業審查人員，正在審核受檢機構針對缺失事項的改善情形。
        請秉持「中立、客觀、平實」的原則進行審查。
        【待改善事項內容】：${content}
        【上一回合審查意見】：${previousReview}
        【本次機構辦理情形】：${latestRound.handling || '無'}
        【回覆格式要求】：JSON: {"fulfill": "Yes/No", "reason": "100字內簡評"}
        `;
        const result = await model.generateContent(prompt);
        const response = await result.response;
        let text = response.text().replace(/```json/g, '').replace(/```/g, '').trim();
        try {
            const json = JSON.parse(text);
            res.json(json);
        } catch (parseError) {
            res.json({ fulfill: text.includes("Yes") ? "Yes" : "No", reason: text.replace(/[{}]/g, '').trim() });
        }
    } catch (e) {
        console.error("Gemini API Error:", e);
        res.status(500).json({ error: 'AI 分析失敗: ' + e.message });
    }
});

app.get('/api/issues', requireAuth, async (req, res) => {
    const { page = 1, pageSize = 20, q, year, unit, status, itemKindCode, division, inspectionCategory, planName, sortField, sortDir } = req.query;
    const limit = parseInt(pageSize);
    const offset = (page - 1) * limit;
    let where = ["1=1"], params = [], idx = 1;

    res.set('Cache-Control', 'no-store');

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
    if (planName) { where.push(`plan_name = $${idx}`); params.push(planName); idx++; }

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
        const tRes = await pool.query("SELECT max(updated_at) as updated, max(created_at) as latest FROM issues");
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
    const { data, round, reviewDate, replyDate } = req.body;
    const r = parseInt(round) || 1;
    const client = await pool.connect();
    try {
        await client.query('BEGIN');
        for (const item of data) {
            const check = await client.query("SELECT id FROM issues WHERE number = $1", [item.number]);
            if (check.rows.length > 0) {
                const hCol = r===1 ? 'handling' : `handling${r}`;
                const rCol = r===1 ? 'review' : `review${r}`;
                const replyCol = `reply_date_r${r}`;
                const respCol = `response_date_r${r}`;
                await client.query(
                    `UPDATE issues SET 
                        status=$1, ${hCol}=$2, ${rCol}=$3, ${replyCol}=$4, ${respCol}=$5,
                        plan_name=COALESCE($6, plan_name), updated_at=CURRENT_TIMESTAMP 
                    WHERE number=$7`,
                    [item.status, item.handling||'', item.review||'', replyDate||'', reviewDate||'', item.planName || null, item.number]
                );
            } else {
                await client.query(
                    `INSERT INTO issues (
                        number, year, unit, content, status, item_kind_code, category, division_name, inspection_category_name,
                        handling, review, plan_name, issue_date, response_date_r1, reply_date_r1
                    ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15)`,
                    [
                        item.number, item.year, item.unit, item.content, item.status||'持續列管',
                        item.itemKindCode, item.category, item.divisionName, item.inspectionCategoryName,
                        item.handling||'', item.review||'', item.planName || null, item.issueDate || null, 
                        reviewDate || '', replyDate || '' 
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

app.get('/api/users', requireAuth, async (req, res) => {
    if (req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    const { page=1, pageSize=20, q, sortField='id', sortDir='asc' } = req.query;
    const limit = parseInt(pageSize);
    const offset = (page-1)*limit;
    let where = ["1=1"], params = [], idx = 1;
    if(q) { where.push(`(username LIKE $${idx} OR name LIKE $${idx})`); params.push(`%${q}%`); idx++; }
    const safeSortFields = ['id', 'username', 'name', 'role', 'created_at'];
    const safeField = safeSortFields.includes(sortField) ? sortField : 'id';
    const order = `${safeField} ${sortDir==='desc'?'DESC':'ASC'}`;
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

app.get('/api/options/plans', requireAuth, async (req, res) => {
    try {
        const result = await pool.query("SELECT DISTINCT plan_name FROM issues WHERE plan_name IS NOT NULL AND plan_name != '' ORDER BY plan_name DESC");
        res.set('Cache-Control', 'no-store');
        res.json({ data: result.rows.map(r => r.plan_name) });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

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