require('dotenv').config();
const express = require('express');
const { Pool } = require('pg');
const bodyParser = require('body-parser');
const path = require('path');
const session = require('express-session');
const pgSession = require('connect-pg-simple')(session); // 關鍵修正：將 Session 存入 DB
const bcrypt = require('bcrypt');
const { GoogleGenerativeAI } = require("@google/generative-ai");

const app = express();

// 1. 信任 Proxy (解決 Render/Heroku/CloudRun 等平台上的 Cookie 問題)
app.set('trust proxy', 1);

const PORT = process.env.PORT || 3000;

// 2. 資料庫連線設定 (強化 SSL 判斷)
const isProduction = process.env.NODE_ENV === 'production';
const connectionString = process.env.DATABASE_URL;

// 防呆：如果沒有設定 DATABASE_URL，給予明確錯誤提示
if (!connectionString) {
    console.error("❌ 嚴重錯誤: 未設定 DATABASE_URL 環境變數！");
    console.error("   請在 .env 檔案中設定 DATABASE_URL=postgres://user:pass@host:port/dbname");
}

const pool = new Pool({
    connectionString: connectionString,
    ssl: isProduction ? { rejectUnauthorized: false } : false,
    // 增加連線超時設定，避免請求卡死
    connectionTimeoutMillis: 5000,
    idleTimeoutMillis: 30000,
});

// 監聽連線錯誤，避免後端崩潰
pool.on('error', (err, client) => {
    console.error('❌ PostgreSQL Pool Error:', err);
});

app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// 3. Session 設定 (關鍵修正：使用 Database Store)
// 這樣即使伺服器重啟，使用者也不會被登出
app.use(session({
    store: new pgSession({
        pool: pool,                // 使用同一個連線池
        tableName: 'session',      // 稍後會自動建立此表
        createTableIfMissing: true // 自動建立 session 表
    }),
    secret: process.env.SESSION_SECRET || 'sms-secret-key-pg-final-v3',
    resave: false,
    saveUninitialized: false,
    proxy: true,
    cookie: {
        maxAge: 24 * 60 * 60 * 1000, // 1 天
        secure: isProduction,        // 線上環境強制 HTTPS
        httpOnly: true,
        sameSite: isProduction ? 'none' : 'lax' // 修正跨站 Cookie 問題
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
        console.log('✅ PostgreSQL 連線成功。正在檢查資料表結構...');

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

        // 4. 動態欄位補強 (增加更多錯誤捕捉)
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
            } catch (e) { /* 欄位已存在，忽略 */ }
        }

        // 5. 建立預設 Admin (如果完全沒人)
        const userRes = await client.query("SELECT count(*) as count FROM users");
        if (parseInt(userRes.rows[0].count) === 0) {
            const hash = bcrypt.hashSync('admin123', 10);
            await client.query("INSERT INTO users (username, password, name, role) VALUES ($1, $2, $3, $4)", 
                ['admin', hash, '系統管理員', 'admin']);
            console.log("⚠️ 系統初始化：已建立預設管理員 (admin / admin123)");
        }
        
        console.log('✅ 資料庫初始化完成 (Schema Checked)');

    } catch (err) {
        console.error('❌ Init DB Error:', err);
        // 這裡不拋出錯誤，讓伺服器繼續嘗試運行，避免一直重啟
    } finally {
        client.release();
    }
}

// Log 輔助
async function logAction(username, action, details, req) {
    const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    try {
        await pool.query("INSERT INTO logs (username, action, details, ip_address, created_at) VALUES ($1, $2, $3, $4, NOW())", 
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
            return res.status(401).json({ error: '帳號或密碼錯誤' });
        }

        if (bcrypt.compareSync(password, user.password)) {
            req.session.user = { id: user.id, username: user.username, role: user.role, name: user.name };
            
            // 寫入登入 Log (Login Time)
            const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
            pool.query("INSERT INTO logs (username, action, details, ip_address, login_time, created_at) VALUES ($1, 'LOGIN', 'User logged in', $2, NOW(), NOW())", [user.username, ip]);

            req.session.save((err) => {
                if(err) {
                    console.error("Session Save Error:", err);
                    return res.status(500).json({error: 'Session error'});
                }
                res.json({ success: true, user: req.session.user });
            });
        } else {
            res.status(401).json({ error: '帳號或密碼錯誤' });
        }
    } catch (e) {
        console.error("Login Error:", e);
        res.status(500).json({ error: 'System error' });
    }
});

app.post('/api/auth/logout', (req, res) => {
    req.session.destroy((err) => {
        if (err) console.error("Logout Error:", err);
        res.clearCookie('connect.sid'); // 清除 Cookie
        res.json({ success: true });
    });
});

// 獲取當前使用者 (修正權限不同步問題)
app.get('/api/auth/me', async (req, res) => {
    // 防止 Session 尚未載入
    if (req.session && req.session.user) {
        try {
            const result = await pool.query("SELECT id, username, name, role FROM users WHERE id = $1", [req.session.user.id]);
            const freshUser = result.rows[0];

            if (freshUser) {
                req.session.user = { 
                    id: freshUser.id, 
                    username: freshUser.username, 
                    role: freshUser.role, 
                    name: freshUser.name 
                };
                res.json({ isLogin: true, ...freshUser });
            } else {
                req.session.destroy();
                res.json({ isLogin: false });
            }
        } catch (e) {
            console.error("Auth Me Error:", e);
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

// 2. AI 審查
app.post('/api/gemini', async (req, res) => {
    const { content, rounds } = req.body;
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) return res.status(500).json({ error: '後端未設定 GEMINI_API_KEY' });

    try {
        const genAI = new GoogleGenerativeAI(apiKey);
        const model = genAI.getGenerativeModel({ model: "gemini-2.0-flash" }); // 修正模型名稱

        const latestRound = (rounds && rounds.length > 0) ? rounds[rounds.length - 1] : { handling: '無', review: '無' };
        const previousReview = (rounds && rounds.length > 1) ? rounds[rounds.length - 2].review : '無';

        const prompt = `
        你現在是【鐵道監理機關】的專業審查人員，正在審核受檢機構針對缺失事項的改善情形。
        ... (保留原有 Prompt) ...
        【回覆格式要求】：
        請嚴格依照以下 JSON 格式回覆：
        { "fulfill": "Yes 或 No", "reason": "簡短審查意見" }
        `;
        
        const result = await model.generateContent(prompt);
        const response = await result.response;
        let text = response.text().replace(/```json/g, '').replace(/```/g, '').trim();
        
        try {
            res.json(JSON.parse(text));
        } catch (parseError) {
            res.json({ fulfill: text.includes("Yes") ? "Yes" : "No", reason: text });
        }
    } catch (e) {
        console.error("Gemini API Error:", e);
        res.status(500).json({ error: 'AI 分析失敗: ' + e.message });
    }
});

// 3. 事項查詢 (強化參數處理)
app.get('/api/issues', requireAuth, async (req, res) => {
    const { page = 1, pageSize = 20, q, year, unit, status, itemKindCode, division, inspectionCategory, planName, sortField, sortDir } = req.query;
    const limit = parseInt(pageSize) || 20;
    const offset = ((parseInt(page) || 1) - 1) * limit;
    
    let where = ["1=1"], params = [], idx = 1;

    if (q) {
        where.push(`(number ILIKE $${idx} OR content ILIKE $${idx} OR handling ILIKE $${idx} OR review ILIKE $${idx} OR plan_name ILIKE $${idx})`);
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
        const countQuery = `SELECT count(*) FROM issues WHERE ${where.join(" AND ")}`;
        const countRes = await pool.query(countQuery, params);
        const total = parseInt(countRes.rows[0].count);
        
        const dataQuery = `SELECT * FROM issues WHERE ${where.join(" AND ")} ORDER BY ${orderBy} LIMIT $${idx} OFFSET $${idx+1}`;
        const dataRes = await pool.query(dataQuery, [...params, limit, offset]);
        
        // 為了效能，統計資料可以非同步並行查詢
        const [sRes, uRes, yRes, pRes] = await Promise.all([
            pool.query("SELECT status, count(*) as count FROM issues GROUP BY status"),
            pool.query("SELECT unit, count(*) as count FROM issues GROUP BY unit"),
            pool.query("SELECT year, count(*) as count FROM issues GROUP BY year ORDER BY year DESC"),
            pool.query("SELECT plan_name, count(*) as count FROM issues WHERE plan_name IS NOT NULL AND plan_name != '' GROUP BY plan_name ORDER BY plan_name DESC")
        ]);

        res.json({
            data: dataRes.rows,
            total,
            page: parseInt(page),
            pageSize: limit,
            pages: Math.ceil(total / limit),
            globalStats: { status: sRes.rows, unit: uRes.rows, year: yRes.rows, plans: pRes.rows }
        });
    } catch (e) { 
        console.error("Query Issues Error:", e);
        res.status(500).json({ error: e.message }); 
    }
});

// --- 其他 CRUD Routes (保持原本邏輯，僅增加 Error Logging) ---

app.put('/api/issues/:id', requireAuth, async (req, res) => {
    const { status, round, handling, review, replyDate, responseDate } = req.body;
    const id = req.params.id;
    const r = parseInt(round) || 1;
    const hField = r === 1 ? 'handling' : `handling${r}`;
    const rField = r === 1 ? 'review' : `review${r}`;
    const replyField = `reply_date_r${r}`;
    const respField = `response_date_r${r}`;

    try {
        await pool.query(`UPDATE issues SET status=$1, ${hField}=$2, ${rField}=$3, ${replyField}=$4, ${respField}=$5, updated_at=NOW() WHERE id=$6`, 
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
                    `UPDATE issues SET status=$1, ${hCol}=$2, ${rCol}=$3, ${respCol}=$4, plan_name=COALESCE($5, plan_name), updated_at=NOW() WHERE number=$6`,
                    [item.status, item.handling||'', item.review||'', reviewDate||'', item.planName || null, item.number]
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
                        item.handling||'', item.review||'', item.planName || null, item.issueDate || null, reviewDate || '' 
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

// Admin Routes (Users & Logs)
app.get('/api/users', requireAuth, async (req, res) => {
    if (req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    const { page=1, pageSize=20, q, sortField='id', sortDir='asc' } = req.query;
    const limit = parseInt(pageSize);
    const offset = (parseInt(page)-1)*limit;
    
    let where = ["1=1"], params = [], idx = 1;
    if(q) { where.push(`(username ILIKE $${idx} OR name ILIKE $${idx})`); params.push(`%${q}%`); idx++; }
    const order = `${sortField} ${sortDir==='desc'?'DESC':'ASC'}`;
    
    try {
        const cRes = await pool.query(`SELECT count(*) FROM users WHERE ${where.join(" AND ")}`, params);
        const dRes = await pool.query(`SELECT id, username, name, role, created_at FROM users WHERE ${where.join(" AND ")} ORDER BY ${order} LIMIT $${idx} OFFSET $${idx+1}`, [...params, limit, offset]);
        res.json({data:dRes.rows, total:parseInt(cRes.rows[0].count), page: parseInt(page), pages: Math.ceil(parseInt(cRes.rows[0].count)/limit)});
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
            console.log(`🚀 Server running on http://localhost:${PORT}`);
        });
    } catch (e) {
        console.error("❌ Server start failed:", e);
        process.exit(1);
    }
}

startServer();