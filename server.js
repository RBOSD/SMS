require('dotenv').config();
const express = require('express');
const { Pool } = require('pg');
const bodyParser = require('body-parser');
const path = require('path');
const session = require('express-session');
const pgSession = require('connect-pg-simple')(session);
const bcrypt = require('bcrypt');
const { GoogleGenerativeAI } = require("@google/generative-ai");

const app = express();
const PORT = process.env.PORT || 3000;

// --- 1. 環境與連線設定 ---

// [Render 必備] 信任反向代理，確保 Cookie/Session 正常運作
app.set('trust proxy', 1);

// 資料庫連線字串檢查
const connectionString = process.env.DATABASE_URL;
if (!connectionString) {
    console.error("❌ [嚴重錯誤] 未設定 DATABASE_URL 環境變數！系統無法啟動。");
    process.exit(1);
}

const pool = new Pool({
    connectionString: connectionString,
    ssl: { rejectUnauthorized: false }, // Render 強制 SSL
    connectionTimeoutMillis: 10000,     // 放寬連線超時
    idleTimeoutMillis: 30000
});

pool.on('error', (err) => {
    console.error('❌ Database Pool Error:', err);
});

// --- 2. 中介軟體 (Middleware) ---
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// --- 3. Session 設定 (使用 PostgreSQL 儲存) ---
app.use(session({
    store: new pgSession({
        pool: pool,
        tableName: 'session',
        createTableIfMissing: true // 自動嘗試建表
    }),
    secret: process.env.SESSION_SECRET || 'sms-system-secret-key-2025-full',
    resave: false,
    saveUninitialized: false,
    proxy: true,
    cookie: {
        maxAge: 24 * 60 * 60 * 1000, // 1 天
        secure: true,                // Render 上必須為 true
        sameSite: 'none'             // 允許跨站 Cookie
    }
}));

// --- 4. 資料庫初始化 (Init DB) ---
async function initDB() {
    let client;
    try {
        client = await pool.connect();
        console.log('✅ PostgreSQL 資料庫連線成功');

        // (A) 手動確保 Session 表存在 (防止套件自動建立失敗導致白畫面)
        await client.query(`
            CREATE TABLE IF NOT EXISTS "session" (
                "sid" varchar NOT NULL COLLATE "default",
                "sess" json NOT NULL,
                "expire" timestamp(6) NOT NULL
            ) WITH (OIDS=FALSE);
        `);
        try {
            await client.query(`ALTER TABLE "session" ADD CONSTRAINT "session_pkey" PRIMARY KEY ("sid") NOT DEFERRABLE INITIALLY IMMEDIATE;`);
        } catch (e) { /* 忽略主鍵已存在錯誤 */ }
        try {
            await client.query(`CREATE INDEX IF NOT EXISTS "IDX_session_expire" ON "session" ("expire");`);
        } catch (e) { /* 忽略索引已存在錯誤 */ }

        // (B) 建立主要業務資料表
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

        // (C) 動態欄位補強 (確保歷史紀錄欄位存在)
        const newColumns = [];
        for(let i=2; i<=20; i++) { 
            newColumns.push({name:`handling${i}`,type:'TEXT'}); 
            newColumns.push({name:`review${i}`,type:'TEXT'}); 
        }
        for(let i=1; i<=20; i++) { 
            newColumns.push({name:`reply_date_r${i}`,type:'TEXT'}); 
            newColumns.push({name:`response_date_r${i}`,type:'TEXT'}); 
        }
        newColumns.push({name:'plan_name',type:'TEXT'}); 
        newColumns.push({name:'issue_date',type:'TEXT'});
        
        for(const col of newColumns) {
            try { 
                await client.query(`ALTER TABLE issues ADD COLUMN IF NOT EXISTS ${col.name} ${col.type}`); 
            } catch(e){}
        }

        // (D) 建立預設管理員
        const uRes = await client.query("SELECT count(*) FROM users");
        if(parseInt(uRes.rows[0].count) === 0) {
            const hash = bcrypt.hashSync('admin123', 10);
            await client.query("INSERT INTO users (username, password, name, role) VALUES ($1, $2, $3, $4)", 
                ['admin', hash, '系統管理員', 'admin']);
            console.log("⚠️ 初始化：已建立預設管理員 (admin / admin123)");
        }

        console.log('✅ 資料庫結構初始化完成');

    } catch (err) {
        console.error('❌ 資料庫初始化失敗:', err);
        // 不退出進程，嘗試讓伺服器繼續運行
    } finally {
        if(client) client.release();
    }
}

// Log 輔助函式
async function logAction(username, action, details, req) {
    const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    try {
        await pool.query("INSERT INTO logs (username, action, details, ip_address, created_at) VALUES ($1, $2, $3, $4, NOW())", 
            [username, action, details, ip]);
    } catch (e) { console.error("Log Error:", e); }
}

// 權限檢查 Middleware
const requireAuth = (req, res, next) => {
    if (req.session && req.session.user) {
        next();
    } else {
        res.status(401).json({ error: 'Unauthorized' });
    }
};

// --- API Routes ---

// 1. 登入與驗證
app.post('/api/auth/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        const r = await pool.query("SELECT * FROM users WHERE username=$1", [username]);
        const user = r.rows[0];
        
        if(user && bcrypt.compareSync(password, user.password)) {
            req.session.user = { 
                id: user.id, 
                username: user.username, 
                role: user.role, 
                name: user.name 
            };
            
            // 記錄登入 Log
            const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
            await pool.query("INSERT INTO logs (username, action, details, ip_address, login_time, created_at) VALUES ($1, 'LOGIN', 'User logged in', $2, NOW(), NOW())", [user.username, ip]);

            req.session.save(err => {
                if(err) { console.error(err); return res.status(500).json({error:'Session save failed'}); }
                res.json({ success: true, user: req.session.user });
            });
        } else {
            res.status(401).json({ error: '帳號或密碼錯誤' });
        }
    } catch(e) { 
        console.error("Login Error:", e);
        res.status(500).json({error: 'System Error'}); 
    }
});

app.post('/api/auth/logout', (req, res) => {
    req.session.destroy();
    res.clearCookie('connect.sid');
    res.json({ success: true });
});

app.get('/api/auth/me', async (req, res) => {
    if(req.session && req.session.user) {
        // 重新確認權限 (防止權限變更後 Session 未更新)
        try {
            const r = await pool.query("SELECT id, username, name, role FROM users WHERE id=$1", [req.session.user.id]);
            if(r.rows[0]) {
                req.session.user = { ...req.session.user, ...r.rows[0] };
                return res.json({ isLogin: true, ...req.session.user });
            }
        } catch(e) { console.error(e); }
        // 如果 DB 讀取失敗，暫時回傳 session 內的資料，避免白畫面
        res.json({ isLogin: true, ...req.session.user });
    } else {
        res.json({ isLogin: false });
    }
});

// 2. AI 審查 (Gemini - 完整監理機關版)
app.post('/api/gemini', async (req, res) => {
    const { content, rounds } = req.body;
    const apiKey = process.env.GEMINI_API_KEY;

    if (!apiKey) return res.json({ fulfill: 'No', reason: '後端未設定 GEMINI_API_KEY' });

    try {
        const genAI = new GoogleGenerativeAI(apiKey);
        const model = genAI.getGenerativeModel({ model: "gemini-2.0-flash" });

        const latestRound = (rounds && rounds.length > 0) ? rounds[rounds.length - 1] : { handling: '無', review: '無' };
        const previousReview = (rounds && rounds.length > 1) ? rounds[rounds.length - 2].review : '無';

        // [完整保留] 專業審查 Prompt
        const prompt = `
        你現在是【鐵道監理機關】的專業審查人員，正在審核受檢機構針對缺失事項的改善情形。
        請秉持「中立、客觀、平實」的原則進行審查，語氣需符合公務機關公文風格。

        【待改善事項內容】：
        ${content}

        【上一回合審查意見】(若有，請確認是否已針對此意見修正)：
        ${previousReview}

        【本次機構辦理情形】：
        ${latestRound.handling || '無'}

        ---
        【審查判斷邏輯】：
        1. **缺失事項**：必須比對「辦理情形」是否具體解決問題。若僅說明「將於未來辦理」或「納入計畫」，應評為「持續列管」，並要求具體期程。
        2. **觀察事項 (Observation)**：此類事項通常涉及潛在風險或長期趨勢。若機構僅表示「研議中」或「評估中」而無具體結論或改善措施，**不可解除列管**，應要求持續觀察或提出具體方案。
        3. **建議事項 (Recommendation)**：此類事項通常由機構**自行列管**。若機構已回復相關處置或說明理由，原則上可同意備查或維持自行列管。
        4. **佐證資料**：若機構宣稱已完成，審查意見應習慣性提及「請檢附相關佐證資料(如照片、紀錄)」。

        【回覆格式要求】：
        請嚴格依照以下 JSON 格式回覆 (不要 Markdown 標記，純 JSON)：
        {
            "fulfill": "Yes 或 No (Yes代表同意解除列管/備查, No代表不同意/持續列管)",
            "reason": "請撰寫平實的審查意見(100字內)。範例：「說明內容尚屬妥適，同意解除列管。」或「所陳改善措施尚未具體，請補充相關佐證資料後報局。」"
        }
        `;

        const result = await model.generateContent(prompt);
        const txt = result.response.text().replace(/```json|```/g, '').trim();
        
        try {
            res.json(JSON.parse(txt));
        } catch(e) {
            // 備援解析
            res.json({ 
                fulfill: txt.includes("Yes") ? "Yes" : "No", 
                reason: txt.replace(/[{}]/g, '').replace(/"reason":/g, '').replace(/"fulfill":.*/g, '').trim() 
            });
        }
    } catch(e) { 
        console.error("AI Error:", e);
        res.json({ fulfill: 'No', reason: 'AI 連線失敗: ' + e.message });
    }
});

// 3. 事項查詢 API (全功能篩選)
app.get('/api/issues', requireAuth, async (req, res) => {
    const { page=1, pageSize=20, q, planName, status, year, unit, itemKindCode, division, inspectionCategory, sortField='created_at', sortDir='desc' } = req.query;
    const limit = parseInt(pageSize);
    const offset = (page-1)*limit;
    
    let where = ["1=1"], params = [], idx = 1;
    
    // 搜尋條件
    if(q) { where.push(`(number ILIKE $${idx} OR content ILIKE $${idx} OR handling ILIKE $${idx} OR review ILIKE $${idx})`); params.push(`%${q}%`); idx++; }
    if(planName) { where.push(`plan_name = $${idx}`); params.push(planName); idx++; }
    if(status) { where.push(`status = $${idx}`); params.push(status); idx++; }
    if(year) { where.push(`year = $${idx}`); params.push(year); idx++; }
    if(unit) { where.push(`unit = $${idx}`); params.push(unit); idx++; }
    if(itemKindCode) { where.push(`item_kind_code = $${idx}`); params.push(itemKindCode); idx++; }
    if(division) { where.push(`division_name = $${idx}`); params.push(division); idx++; }
    if(inspectionCategory) { where.push(`inspection_category_name = $${idx}`); params.push(inspectionCategory); idx++; }

    // 排序
    const order = `${sortField} ${sortDir==='asc'?'ASC':'DESC'}`;

    try {
        // 總數
        const cRes = await pool.query(`SELECT count(*) FROM issues WHERE ${where.join(" AND ")}`, params);
        const total = parseInt(cRes.rows[0].count);
        
        // 資料
        const dRes = await pool.query(`SELECT * FROM issues WHERE ${where.join(" AND ")} ORDER BY ${order} LIMIT ${limit} OFFSET ${offset}`, params);
        
        // 統計 (用於前端篩選器與儀表板)
        // 使用 Promise.all 並行查詢以提升效能
        const [sRes, uRes, yRes, pRes] = await Promise.all([
            pool.query("SELECT status, count(*) as count FROM issues GROUP BY status"),
            pool.query("SELECT unit, count(*) as count FROM issues GROUP BY unit"),
            pool.query("SELECT year, count(*) as count FROM issues GROUP BY year ORDER BY year DESC"),
            pool.query("SELECT plan_name, count(*) as count FROM issues WHERE plan_name IS NOT NULL AND plan_name != '' GROUP BY plan_name ORDER BY plan_name DESC")
        ]);

        res.json({
            data: dRes.rows,
            total,
            page: parseInt(page),
            pageSize: limit,
            pages: Math.ceil(total/limit),
            globalStats: { status: sRes.rows, unit: uRes.rows, year: yRes.rows, plans: pRes.rows }
        });
    } catch(e) { 
        console.error("Query Error:", e);
        res.status(500).json({error: e.message}); 
    }
});

// 4. 匯入/新增 API (支援新版計畫綁定邏輯)
app.post('/api/issues/import', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Permission Denied'});
    
    const { data, round, reviewDate } = req.body;
    const r = parseInt(round) || 1;
    const client = await pool.connect();
    
    try {
        await client.query('BEGIN');
        
        for(const item of data) {
            // 檢查是否已存在 (依編號)
            const check = await client.query("SELECT id FROM issues WHERE number=$1", [item.number]);
            
            if(check.rows.length > 0) {
                // 更新模式 (Update)
                const hCol = r===1 ? 'handling' : `handling${r}`;
                const rCol = r===1 ? 'review' : `review${r}`;
                const dateCol = `response_date_r${r}`;
                
                await client.query(
                    `UPDATE issues SET 
                        status=$1, 
                        ${hCol}=$2, 
                        ${rCol}=$3, 
                        ${dateCol}=$4, 
                        plan_name=COALESCE($5, plan_name), 
                        updated_at=NOW() 
                    WHERE number=$6`,
                    [
                        item.status, 
                        item.handling||'', 
                        item.review||'', 
                        reviewDate||'', 
                        item.planName||null, // 若有帶入計畫名稱，則更新歸屬
                        item.number
                    ]
                );
            } else {
                // 新增模式 (Insert)
                await client.query(
                    `INSERT INTO issues (
                        number, year, unit, content, status, 
                        item_kind_code, category, division_name, inspection_category_name,
                        handling, review, plan_name, issue_date, response_date_r1
                    ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14)`,
                    [
                        item.number, 
                        item.year, 
                        item.unit, 
                        item.content, 
                        item.status || '持續列管',
                        item.itemKindCode,
                        item.category,
                        item.divisionName,
                        item.inspectionCategoryName,
                        item.handling || '', 
                        item.review || '', 
                        item.planName || null, 
                        item.issueDate || null, 
                        reviewDate || ''
                    ]
                );
            }
        }
        
        await client.query('COMMIT');
        await logAction(req.session.user.username, 'IMPORT', `Imported/Updated ${data.length} items`, req);
        res.json({ success: true, count: data.length });

    } catch(e) {
        await client.query('ROLLBACK');
        console.error("Import Error:", e);
        res.status(500).json({ error: e.message });
    } finally {
        client.release();
    }
});

// 5. 單筆編輯與刪除
app.put('/api/issues/:id', requireAuth, async (req, res) => {
    const { status, round, handling, review, replyDate, responseDate } = req.body;
    const id = req.params.id;
    const r = parseInt(round);
    const hField = r===1 ? 'handling' : `handling${r}`;
    const rField = r===1 ? 'review' : `review${r}`;
    const replyField = `reply_date_r${r}`;
    const respField = `response_date_r${r}`;

    try {
        await pool.query(`UPDATE issues SET status=$1, ${hField}=$2, ${rField}=$3, ${replyField}=$4, ${respField}=$5, updated_at=NOW() WHERE id=$6`, 
            [status, handling, review, replyDate, responseDate, id]);
        await logAction(req.session.user.username, 'UPDATE', `Updated issue ${id}`, req);
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/issues/:id', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    try {
        await pool.query("DELETE FROM issues WHERE id=$1", [req.params.id]);
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/issues/batch-delete', requireAuth, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    const { ids } = req.body;
    try {
        await pool.query("DELETE FROM issues WHERE id = ANY($1)", [ids]);
        await logAction(req.session.user.username, 'BATCH_DELETE', `Deleted ${ids.length} items`, req);
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// 6. 啟動伺服器
initDB().then(() => {
    app.listen(PORT, () => {
        console.log(`\n🚀 Server running on port ${PORT}`);
    });
});