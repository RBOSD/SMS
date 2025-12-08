const express = require('express');
const { Pool } = require('pg'); 
const bodyParser = require('body-parser');
const path = require('path');
const session = require('express-session');
const bcrypt = require('bcrypt');
const { GoogleGenerativeAI } = require("@google/generative-ai"); 
require('dotenv').config(); 

const app = express();

// 設定 Trust Proxy (為了 Render/Heroku 等環境)
app.set('trust proxy', 1); 

const PORT = process.env.PORT || 3000;

// PostgreSQL 連線池
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false
});

// Middleware
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Session 設定
app.use(session({
    secret: 'sms-secret-key-pg-final-v6', // 更新密鑰確保 Session 刷新
    resave: false,
    saveUninitialized: false,
    proxy: true, 
    cookie: { 
        maxAge: 24 * 60 * 60 * 1000, // 24小時
        secure: process.env.NODE_ENV === 'production', 
        sameSite: 'lax' 
    } 
}));

// 權限驗證 Middleware
const requireAuth = (req, res, next) => {
    if (req.session && req.session.user) {
        next();
    } else {
        res.status(401).json({ error: 'Unauthorized' });
    }
};

// 資料庫初始化
async function initDB() {
    const client = await pool.connect();
    try {
        console.log('Connected to PostgreSQL. Checking schema...');

        // 1. Issues 表
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

        // 2. Users 表
        await client.query(`CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE,
            password TEXT,
            name TEXT,
            role TEXT DEFAULT 'viewer',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )`);

        // 3. Logs 表
        await client.query(`CREATE TABLE IF NOT EXISTS logs (
            id SERIAL PRIMARY KEY,
            username TEXT,
            action TEXT,
            details TEXT,
            ip_address TEXT,
            login_time TIMESTAMP,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )`);

        // 動態新增欄位 (handling2~20, review2~20, dates, plan_name)
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
            } catch (e) { 
                // Ignore if exists
            }
        }

        // 預設管理員
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

// Log 輔助函式
async function logAction(username, action, details, req) {
    const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    try {
        await pool.query("INSERT INTO logs (username, action, details, ip_address, created_at) VALUES ($1, $2, $3, $4, CURRENT_TIMESTAMP)", 
            [username, action, details, ip]);
    } catch (e) { console.error("Log error:", e); }
}

// --- API Routes ---

// 登入
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

// 登出
app.post('/api/auth/logout', (req, res) => {
    req.session.destroy();
    res.json({ success: true });
});

// 檢查登入狀態
app.get('/api/auth/me', async (req, res) => {
    if (req.session && req.session.user) {
        try {
            // 每次檢查都從 DB 撈最新權限
            const result = await pool.query("SELECT id, username, name, role FROM users WHERE id = $1", [req.session.user.id]);
            const latestUser = result.rows[0];

            if (!latestUser) {
                req.session.destroy();
                return res.json({ isLogin: false });
            }
            // Update session info
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

// 修改個人資料
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

// Gemini AI
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
        
        請依照以下 JSON 格式回覆：
        {
            "fulfill": "Yes" 或 "No", 
            "reason": "100字以內的簡短評語，說明為何通過或為何不通過"
        }
        `;

        const result = await model.generateContent(prompt);
        const response = await result.response;
        let text = response.text();
        // 清理 markdown
        text = text.replace(/```json/g, '').replace(/```/g, '').trim();

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

// 查詢事項 (支援多種篩選)
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
        
        // 統計數據 (全域，不受分頁影響但受篩選影響太慢，這裡做全表統計)
        // 若要隨篩選變動，需帶入 where。這裡為求效能先給全表概況。
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

// 更新單筆事項
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

// 刪除單筆
app.delete('/api/issues/:id', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    try {
        await pool.query("DELETE FROM issues WHERE id=$1", [req.params.id]);
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// 批次刪除
app.post('/api/issues/batch-delete', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    const { ids } = req.body;
    try {
        await pool.query("DELETE FROM issues WHERE id = ANY($1)", [ids]);
        logAction(req.session.user.username, 'BATCH_DELETE', `Deleted ${ids.length} items`, req);
        res.json({success:true});
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// 批次匯入/新增事項 (核心邏輯)
app.post('/api/issues/import', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    const { data, round, reviewDate, replyDate } = req.body;
    const r = parseInt(round) || 1;
    
    const client = await pool.connect();
    try {
        await client.query('BEGIN');

        for (const item of data) {
            // [關鍵] 若 number 為空，可能是無效行，跳過
            if (!item.number) continue;

            // 檢查是否存在
            const check = await client.query("SELECT id FROM issues WHERE number = $1", [item.number]);
            const exists = (check.rows.length > 0);
            
            if (exists) {
                const hCol = r===1 ? 'handling' : `handling${r}`;
                const rCol = r===1 ? 'review' : `review${r}`;
                const replyCol = `reply_date_r${r}`;
                const respCol = `response_date_r${r}`;
                
                await client.query(
                    `UPDATE issues SET 
                        status=$1, ${hCol}=$2, ${rCol}=$3, 
                        ${replyCol}=$4, ${respCol}=$5,
                        plan_name=COALESCE($6, plan_name), 
                        updated_at=CURRENT_TIMESTAMP 
                    WHERE number=$7`,
                    [
                        item.status, item.handling||'', item.review||'',
                        replyDate||'', reviewDate||'', 
                        item.planName || null,
                        item.number
                    ]
                );
            } else {
                // 新增
                await client.query(
                    `INSERT INTO issues (
                        number, year, unit, content, status, item_kind_code, category, division_name, inspection_category_name,
                        handling, review, plan_name, issue_date, 
                        response_date_r1, reply_date_r1
                    ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15)`,
                    [
                        item.number || '', item.year, item.unit, item.content || '', item.status||'持續列管',
                        item.itemKindCode, item.category, item.divisionName, item.inspectionCategoryName,
                        item.handling||'', item.review||'', 
                        item.planName || null, 
                        item.issueDate || null, 
                        reviewDate || '', 
                        replyDate || ''   
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

// 使用者管理 API
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

// Logs API
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

// --- 計畫相關 API ---

// 1. 取得純計畫名稱選單
app.get('/api/options/plans', requireAuth, async (req, res) => {
    try {
        const result = await pool.query("SELECT DISTINCT plan_name FROM issues WHERE plan_name IS NOT NULL AND plan_name != '' ORDER BY plan_name DESC");
        res.set('Cache-Control', 'no-store');
        res.json({ data: result.rows.map(r => r.plan_name) });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 2. 取得計畫詳情列表
app.get('/api/options/plans_details', requireAuth, async (req, res) => {
    const { q = '', year = '', page = 1, pageSize = 15 } = req.query;
    const limit = parseInt(pageSize), offset = (page - 1) * limit;
    let where = ['plan_name IS NOT NULL AND TRIM(plan_name) != \'\''], params = [], idx = 1;
    if (q) { where.push(`plan_name ILIKE $${idx}`); params.push('%' + q + '%'); idx++; }
    if (year) { where.push(`year = $${idx}`); params.push(year); idx++; }
    try {
        const countRes = await pool.query(
            `SELECT COUNT(DISTINCT plan_name || year) AS count FROM issues WHERE ${where.join(' AND ')}`
            , params);
        const total = parseInt(countRes.rows[0]?.count || 0);
        // number != '' 排除掉純計畫佔位符
        const mainRes = await pool.query(
            `SELECT plan_name, year, COUNT(CASE WHEN number != '' THEN 1 END) as issues_count, MAX(updated_at) as updated_at
             FROM issues WHERE ${where.join(' AND ')}
             GROUP BY plan_name, year ORDER BY updated_at DESC NULLS LAST
             LIMIT $${idx} OFFSET $${idx + 1}`,
            [...params, limit, offset]
        );
        res.json({
            data: mainRes.rows,
            page: parseInt(page),
            pageSize: limit,
            total,
            pages: Math.ceil(total / limit)
        });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// 3. 單筆新增計畫
app.post('/api/options/plans', requireAuth, async (req, res) => {
    const { plan_name, year } = req.body;
    if (!plan_name || !year) return res.status(400).json({ error: '缺少計畫名稱或年度' });
    try {
        // 插入一筆佔位資料來建立計畫
        await pool.query(
            `INSERT INTO issues (plan_name, year, number, content, status)
             VALUES ($1, $2, '', '', '')`,
            [plan_name, year]
        );
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// 4. [新增] 批次新增計畫 (接收陣列)
app.post('/api/options/plans/batch', requireAuth, async (req, res) => {
    const { plans } = req.body; // [{plan_name, year}, ...]
    if (!Array.isArray(plans)) return res.status(400).json({ error: 'Data format error' });

    const client = await pool.connect();
    try {
        await client.query('BEGIN');
        for (const p of plans) {
            if (p.plan_name && p.year) {
                // 檢查該年度該計畫是否已存在 (避免重複建立佔位符)
                const check = await client.query("SELECT 1 FROM issues WHERE plan_name = $1 AND year = $2 AND number = '' LIMIT 1", [p.plan_name, p.year]);
                if (check.rows.length === 0) {
                    await client.query(
                        `INSERT INTO issues (plan_name, year, number, content, status)
                         VALUES ($1, $2, '', '', '')`,
                        [p.plan_name, p.year]
                    );
                }
            }
        }
        await client.query('COMMIT');
        res.json({ success: true, count: plans.length });
    } catch (e) {
        await client.query('ROLLBACK');
        res.status(500).json({ error: e.message });
    } finally {
        client.release();
    }
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