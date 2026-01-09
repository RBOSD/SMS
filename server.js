// [Added] Force Node.js to prefer IPv4 resolution to solve ENETUNREACH issues on some platforms (like Render + Supabase)
require('dns').setDefaultResultOrder('ipv4first');

const express = require('express');
const { Pool } = require('pg'); 
const bodyParser = require('body-parser');
const path = require('path');
const session = require('express-session');
// [Added] pg-simple session store
const pgSession = require('connect-pg-simple')(session);
const bcrypt = require('bcrypt');
const { GoogleGenerativeAI } = require("@google/generative-ai");
const rateLimit = require('express-rate-limit');
require('dotenv').config(); 

const app = express();

app.set('trust proxy', 1); 

const PORT = process.env.PORT || 3000;

// [Modified] Initialize PostgreSQL Connection Pool
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false }, // Allow self-signed certs
    max: 20, 
    idleTimeoutMillis: 30000,
    connectionTimeoutMillis: 5000,
});

app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// [Modified] Session Configuration
app.use(session({
    store: new pgSession({
        pool: pool,
        tableName: 'session',
        createTableIfMissing: true 
    }),
    secret: (() => {
        const secret = process.env.SESSION_SECRET;
        if (!secret || secret === 'sms-secret-key-pg-final-v3') {
            console.error('警告: SESSION_SECRET 環境變數未設定或使用預設值！');
            console.error('請在 .env 檔案中設定一個隨機且複雜的 SESSION_SECRET');
            console.error('可以使用命令產生: openssl rand -base64 32');
            if (process.env.NODE_ENV === 'production') {
                throw new Error('SESSION_SECRET environment variable is required in production');
            }
        }
        return secret || 'sms-secret-key-pg-final-v3-dev-only';
    })(),
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

// 對所有 API 路由套用速率限制（除了已經有特定限制的路由）
// 注意：這個中間件必須放在 session 之後
app.use('/api/', (req, res, next) => {
    // 登入路由使用 loginLimiter，不需要再次限制
    if (req.path === '/auth/login') {
        return next();
    }
    // Gemini API 使用 geminiLimiter
    if (req.path === '/gemini') {
        return next();
    }
    // 其他 API 使用通用限制
    return apiLimiter(req, res, next);
});

// --- Database Initialization ---
async function initDB() {
    let retries = 5;
    while (retries > 0) {
        try {
            const client = await pool.connect();
            try {
                console.log('Connected to PostgreSQL. Checking schema...');

                // Session Table
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

                // Issues Table
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

                // Users Table
                await client.query(`CREATE TABLE IF NOT EXISTS users (
                    id SERIAL PRIMARY KEY,
                    username TEXT UNIQUE,
                    password TEXT,
                    name TEXT,
                    role TEXT DEFAULT 'viewer',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )`);

                // Logs Table
                await client.query(`CREATE TABLE IF NOT EXISTS logs (
                    id SERIAL PRIMARY KEY,
                    username TEXT,
                    action TEXT,
                    details TEXT,
                    ip_address TEXT,
                    login_time TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )`);

                // Add missing columns if they don't exist
                const newColumns = [];
                // 支持無限次審查，預先創建前 100 次欄位（如果需要更多可以動態創建）
                for (let i = 2; i <= 100; i++) {
                    newColumns.push({ name: `handling${i}`, type: 'TEXT' });
                    newColumns.push({ name: `review${i}`, type: 'TEXT' });
                }
                for (let i = 1; i <= 100; i++) {
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

                // Create Default Admin if no users exist
                const userRes = await client.query("SELECT count(*) as count FROM users");
                if (parseInt(userRes.rows[0].count) === 0) {
                    // 使用環境變數提供預設密碼，或在首次登入後強制修改
                    const defaultPassword = process.env.DEFAULT_ADMIN_PASSWORD || 
                        require('crypto').randomBytes(16).toString('hex');
                    const hash = bcrypt.hashSync(defaultPassword, 10);
                    await client.query("INSERT INTO users (username, password, name, role) VALUES ($1, $2, $3, $4)", 
                        ['admin', hash, '系統管理員', 'admin']);
                    console.log("===========================================");
                    console.log("警告: 已建立預設管理員帳號");
                    console.log("帳號: admin");
                    console.log("密碼: " + (process.env.DEFAULT_ADMIN_PASSWORD ? "使用環境變數設定" : defaultPassword));
                    console.log("請立即登入並修改密碼！");
                    console.log("===========================================");
                }
                
                console.log('Database initialized successfully.');
                return;

            } catch (err) {
                console.error('Init DB Schema Error:', err);
                throw err;
            } finally {
                client.release();
            }
        } catch (connErr) {
            console.error(`Connection failed, retrying... (${retries} left)`, connErr.message);
            retries--;
            await new Promise(res => setTimeout(res, 2000));
        }
    }
    throw new Error('Could not connect to database after multiple retries.');
}

// Rate Limiting 設定
const loginLimiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 分鐘
    max: 5, // 最多 5 次登入嘗試
    message: { error: '登入嘗試過多，請 15 分鐘後再試' },
    standardHeaders: true,
    legacyHeaders: false,
    skip: (req) => {
        // 開發環境可以放寬限制
        return process.env.NODE_ENV === 'development';
    }
});

const apiLimiter = rateLimit({
    windowMs: 1 * 60 * 1000, // 1 分鐘
    max: 100, // 最多 100 次請求
    message: { error: 'API 調用過於頻繁，請稍後再試' },
    standardHeaders: true,
    legacyHeaders: false,
});

const geminiLimiter = rateLimit({
    windowMs: 60 * 1000, // 1 分鐘
    max: 20, // 最多 20 次 AI 分析請求
    message: { error: 'AI 分析請求過於頻繁，請稍後再試' },
    standardHeaders: true,
    legacyHeaders: false,
});

async function logAction(username, action, details, req) {
    const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    try {
        await pool.query("INSERT INTO logs (username, action, details, ip_address, created_at) VALUES ($1, $2, $3, $4, CURRENT_TIMESTAMP)", 
            [username, action, details, ip]);
    } catch (e) { console.error("Log error:", e); }
}

// --- API Routes ---

app.post('/api/auth/login', loginLimiter, async (req, res) => {
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

app.post('/api/gemini', geminiLimiter, async (req, res) => {
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
        【回覆格式要求】：JSON: {"fulfill": "Yes/No", "result": "100字內簡評"}
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
        // 先查詢 issue number
        const issueRes = await pool.query("SELECT number FROM issues WHERE id=$1", [id]);
        const issueNumber = issueRes.rows[0]?.number || `ID:${id}`;
        
        // 如果超過預設欄位範圍（100次），動態創建欄位
        if (r > 100) {
            try {
                await pool.query(`ALTER TABLE issues ADD COLUMN IF NOT EXISTS ${hField} TEXT`);
                await pool.query(`ALTER TABLE issues ADD COLUMN IF NOT EXISTS ${rField} TEXT`);
                await pool.query(`ALTER TABLE issues ADD COLUMN IF NOT EXISTS ${replyField} TEXT`);
                await pool.query(`ALTER TABLE issues ADD COLUMN IF NOT EXISTS ${respField} TEXT`);
            } catch (colError) {
                // 忽略欄位已存在的錯誤
                if (!colError.message.includes('already exists')) {
                    console.error('Error creating columns:', colError);
                }
            }
        }
        
        await pool.query(`UPDATE issues SET status=$1, ${hField}=$2, ${rField}=$3, ${replyField}=$4, ${respField}=$5, updated_at=CURRENT_TIMESTAMP WHERE id=$6`, 
            [status, handling, review, replyDate, responseDate, id]);
        logAction(req.session.user.username, 'UPDATE', `更新列管事項：編號 ${issueNumber}`, req);
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/issues/:id', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    try {
        // 先查詢 issue number 再刪除
        const issueRes = await pool.query("SELECT number FROM issues WHERE id=$1", [req.params.id]);
        const issueNumber = issueRes.rows[0]?.number || `ID:${req.params.id}`;
        
        await pool.query("DELETE FROM issues WHERE id=$1", [req.params.id]);
        logAction(req.session.user.username, 'DELETE', `刪除列管事項：編號 ${issueNumber}`, req);
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/issues/batch-delete', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    const { ids } = req.body;
    try {
        // 先查詢所有要刪除的編號
        const issueRes = await pool.query("SELECT number FROM issues WHERE id = ANY($1)", [ids]);
        const numbers = issueRes.rows.map(r => r.number).filter(Boolean);
        const numberList = numbers.length > 0 ? numbers.join(', ') : `${ids.length} 筆`;
        
        await pool.query("DELETE FROM issues WHERE id = ANY($1)", [ids]);
        logAction(req.session.user.username, 'BATCH_DELETE', `批次刪除列管事項：${numberList} (共 ${ids.length} 筆)`, req);
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

// --- User Management API ---

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
    if(req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    const { username, password, name, role } = req.body;
    try {
        // Basic Validation
        if (!username || !password) return res.status(400).json({error: 'Username and password required'});
        
        const hash = bcrypt.hashSync(password, 10);
        await pool.query("INSERT INTO users (username, password, name, role) VALUES ($1, $2, $3, $4)", [username, hash, name, role]);
        logAction(req.session.user.username, 'CREATE_USER', `新增使用者：${name} (${username})，權限：${role}`, req);
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/users/:id', requireAuth, async (req, res) => {
    if(req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    const { name, password, role } = req.body;
    const id = req.params.id;
    try {
        // 先查詢使用者資訊以便記錄
        const userRes = await pool.query("SELECT username, name FROM users WHERE id=$1", [id]);
        const targetUser = userRes.rows[0];
        const targetUsername = targetUser ? targetUser.username : `ID:${id}`;
        const targetName = targetUser ? targetUser.name : '未知';
        
        if (password) {
            const hash = bcrypt.hashSync(password, 10);
            await pool.query("UPDATE users SET name=$1, role=$2, password=$3 WHERE id=$4", [name, role, hash, id]);
            logAction(req.session.user.username, 'UPDATE_USER', `修改使用者：${targetName} (${targetUsername})，已更新姓名、權限和密碼`, req);
        } else {
            await pool.query("UPDATE users SET name=$1, role=$2 WHERE id=$3", [name, role, id]);
            logAction(req.session.user.username, 'UPDATE_USER', `修改使用者：${targetName} (${targetUsername})，已更新姓名和權限`, req);
        }
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/users/:id', requireAuth, async (req, res) => {
    if(req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    if(parseInt(req.params.id) === req.session.user.id) return res.status(400).json({error:'Cannot self delete'});
    try {
        // 先查詢使用者資訊以便記錄
        const userRes = await pool.query("SELECT username, name FROM users WHERE id=$1", [req.params.id]);
        const targetUser = userRes.rows[0];
        const targetUsername = targetUser ? targetUser.username : `ID:${req.params.id}`;
        const targetName = targetUser ? targetUser.name : '未知';
        
        await pool.query("DELETE FROM users WHERE id=$1", [req.params.id]);
        logAction(req.session.user.username, 'DELETE_USER', `刪除使用者：${targetName} (${targetUsername})`, req);
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// --- Admin Logs API ---

app.get('/api/admin/logs', requireAuth, async (req, res) => {
    if(req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    try {
        const { page = 1, pageSize = 50, q } = req.query;
        const limit = parseInt(pageSize);
        const offset = (parseInt(page) - 1) * limit;
        let where = ["action='LOGIN'"];
        let params = [];
        let idx = 1;
        
        if (q) {
            // 搜尋所有欄位：username, ip_address, details, login_time, created_at
            where.push(`(
                COALESCE(username, '') LIKE $${idx} OR 
                COALESCE(ip_address, '') LIKE $${idx} OR 
                COALESCE(details, '') LIKE $${idx} OR
                COALESCE(CAST(login_time AS TEXT), '') LIKE $${idx} OR
                COALESCE(CAST(created_at AS TEXT), '') LIKE $${idx}
            )`);
            params.push(`%${q}%`);
            idx++;
        }
        
        const whereClause = where.join(' AND ');
        const countQuery = `SELECT COUNT(*) FROM logs WHERE ${whereClause}`;
        const dataQuery = `SELECT * FROM logs WHERE ${whereClause} ORDER BY login_time DESC LIMIT $${idx} OFFSET $${idx + 1}`;
        
        const countResult = await pool.query(countQuery, params);
        const total = parseInt(countResult.rows[0].count);
        const pages = Math.ceil(total / limit);
        
        params.push(limit, offset);
        const { rows } = await pool.query(dataQuery, params);
        
        res.json({data:rows, total, page:parseInt(page), pages});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/admin/action_logs', requireAuth, async (req, res) => {
    if(req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    try {
        const { page = 1, pageSize = 50, q } = req.query;
        const limit = parseInt(pageSize);
        const offset = (parseInt(page) - 1) * limit;
        let where = ["action!='LOGIN'"];
        let params = [];
        let idx = 1;
        
        if (q) {
            // 搜尋所有欄位：username, action, details, ip_address, created_at
            where.push(`(
                COALESCE(username, '') LIKE $${idx} OR 
                COALESCE(action, '') LIKE $${idx} OR 
                COALESCE(details, '') LIKE $${idx} OR
                COALESCE(ip_address, '') LIKE $${idx} OR
                COALESCE(CAST(created_at AS TEXT), '') LIKE $${idx}
            )`);
            params.push(`%${q}%`);
            idx++;
        }
        
        const whereClause = where.join(' AND ');
        const countQuery = `SELECT COUNT(*) FROM logs WHERE ${whereClause}`;
        const dataQuery = `SELECT * FROM logs WHERE ${whereClause} ORDER BY created_at DESC LIMIT $${idx} OFFSET $${idx + 1}`;
        
        const countResult = await pool.query(countQuery, params);
        const total = parseInt(countResult.rows[0].count);
        const pages = Math.ceil(total / limit);
        
        params.push(limit, offset);
        const { rows } = await pool.query(dataQuery, params);
        
        res.json({data:rows, total, page:parseInt(page), pages});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/admin/logs', requireAuth, async (req, res) => {
    if(req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    await pool.query("DELETE FROM logs WHERE action='LOGIN'");
    res.json({success:true});
});

app.delete('/api/admin/action_logs', requireAuth, async (req, res) => {
    if(req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    await pool.query("DELETE FROM logs WHERE action!='LOGIN'");
    res.json({success:true});
});

// 根據時間範圍清除舊記錄
app.post('/api/admin/logs/cleanup', requireAuth, async (req, res) => {
    if(req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    try {
        const { days } = req.body;
        if (!days || days < 1) {
            return res.status(400).json({error:'請提供有效的保留天數（至少1天）'});
        }
        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - parseInt(days));
        
        const result = await pool.query(
            "DELETE FROM logs WHERE action='LOGIN' AND login_time < $1",
            [cutoffDate]
        );
        
        logAction(req.session.user.username, 'CLEANUP_LOGS', `清除 ${days} 天前的登入紀錄，刪除 ${result.rowCount} 筆`, req);
        res.json({success:true, deleted: result.rowCount});
    } catch (e) {
        res.status(500).json({error: e.message});
    }
});

app.post('/api/admin/action_logs/cleanup', requireAuth, async (req, res) => {
    if(req.session.user.role !== 'admin') return res.status(403).json({error:'Denied'});
    try {
        const { days } = req.body;
        if (!days || days < 1) {
            return res.status(400).json({error:'請提供有效的保留天數（至少1天）'});
        }
        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - parseInt(days));
        
        const result = await pool.query(
            "DELETE FROM logs WHERE action!='LOGIN' AND created_at < $1",
            [cutoffDate]
        );
        
        logAction(req.session.user.username, 'CLEANUP_ACTION_LOGS', `清除 ${days} 天前的操作紀錄，刪除 ${result.rowCount} 筆`, req);
        res.json({success:true, deleted: result.rowCount});
    } catch (e) {
        res.status(500).json({error: e.message});
    }
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