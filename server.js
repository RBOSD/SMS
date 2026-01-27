// [Added] Force Node.js to prefer IPv4 resolution to solve ENETUNREACH issues on some platforms (like Render + Supabase)
require('dns').setDefaultResultOrder('ipv4first');

const express = require('express');
const { Pool } = require('pg'); 
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs');
const session = require('express-session');
// [Added] pg-simple session store
const pgSession = require('connect-pg-simple')(session);
const bcrypt = require('bcrypt');
const { GoogleGenerativeAI } = require("@google/generative-ai");
const rateLimit = require('express-rate-limit');
const csrf = require('csrf');
require('dotenv').config(); 

const app = express();

app.set('trust proxy', 1); 

const PORT = process.env.PORT || 3000;

// [Modified] Initialize PostgreSQL Connection Pool
// SSL 設定：
// - 預設允許自簽憑證（適用於 Render、Heroku 等雲端平台）
// - 可透過 DB_SSL_REJECT_UNAUTHORIZED=true 強制要求有效憑證
// - 可透過 DB_SSL_REJECT_UNAUTHORIZED=false 明確允許自簽憑證
const sslConfig = (() => {
    // 如果明確指定了環境變數，使用該值
    if (process.env.DB_SSL_REJECT_UNAUTHORIZED !== undefined) {
        return { rejectUnauthorized: process.env.DB_SSL_REJECT_UNAUTHORIZED === 'true' };
    }
    // 預設允許自簽憑證（適用於大多數雲端平台）
    return { rejectUnauthorized: false };
})();

// 主應用程式連線池（Supabase 使用 Supavisor/PgBouncer，Session Mode 連線數受限）
// Supabase 免費方案：直接連線 60，Pooler 200，但 Session Mode 的 pool_size 通常較小
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.DATABASE_URL ? sslConfig : false,
    max: 2, // Supabase Session Mode 建議使用較小的連線池（2-3 個）
    idleTimeoutMillis: 5000, // 快速釋放未使用的連線（5 秒）
    connectionTimeoutMillis: 2000, // 快速超時，避免等待
    allowExitOnIdle: false,
});

// Session store 使用同一個連線池（避免建立過多連線）
// Supabase Session Mode：每個連線獨佔一個底層連線，pool_size 限制了可用連線數
// 因此共用同一個連線池，總連線數限制在 2 個以內
const sessionPool = pool;

// 資料庫連線錯誤處理
pool.on('error', async (err) => {
    // Supabase 連線錯誤處理
    if (err.message && err.message.includes('MaxClientsInSessionMode')) {
        console.warn('Supabase 連線池已滿，請等待連線釋放或考慮使用 Transaction Mode (port 6543)');
    } else if (err.message && err.message.includes('Connection terminated')) {
        console.warn('資料庫連線終止（可能是暫時的）:', err.message);
    } else {
        console.error('資料庫連線錯誤:', err?.message || err);
        // 避免在連線錯誤時記錄到資料庫（可能造成循環）
        if (!err.message || !err.message.includes('MaxClients')) {
            await logError(err, 'Database connection error', null).catch(() => {});
        }
    }
});

app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true }));

// [Modified] Session Configuration（必須在路由保護之前，才能使用 req.session）
let sessionStore;
try {
    sessionStore = new pgSession({
        pool: sessionPool, // 使用獨立的 session 連線池
        tableName: 'session',
        createTableIfMissing: true,
        pruneSessionInterval: false // 手動控制清理，避免初始化問題
    });
} catch (storeError) {
    console.warn('Session store initialization warning:', storeError?.message || storeError);
    // 如果 session store 初始化失敗，使用記憶體 store（僅開發環境）
    if (process.env.NODE_ENV !== 'production') {
        console.warn('Falling back to memory store for development');
        sessionStore = undefined;
    }
}

app.use(session({
    store: sessionStore,
    secret: (() => {
        const secret = process.env.SESSION_SECRET;
        const defaultSecret = 'sms-secret-key-pg-final-v3';
        const devSecret = 'sms-secret-key-pg-final-v3-dev-only';
        
        // 生產環境必須設定 SESSION_SECRET
        if (process.env.NODE_ENV === 'production') {
            if (!secret || secret === defaultSecret || secret === devSecret) {
                console.error('===========================================');
                console.error('錯誤: 生產環境必須設定 SESSION_SECRET 環境變數！');
                console.error('請在 .env 檔案中設定一個隨機且複雜的 SESSION_SECRET');
                console.error('可以使用命令產生: openssl rand -base64 32');
                console.error('===========================================');
                throw new Error('SESSION_SECRET environment variable is required in production');
            }
            // 驗證生產環境的 SESSION_SECRET 長度（至少 32 字元）
            if (secret.length < 32) {
                console.error('警告: 生產環境的 SESSION_SECRET 長度建議至少 32 字元');
            }
        } else {
            // 開發環境警告
            if (!secret || secret === defaultSecret || secret === devSecret) {
                console.warn('警告: SESSION_SECRET 環境變數未設定或使用預設值！');
                console.warn('請在 .env 檔案中設定一個隨機且複雜的 SESSION_SECRET');
                console.warn('可以使用命令產生: openssl rand -base64 32');
            }
        }
        return secret || devSecret;
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

// 路由保護中間件：檢查 HTML 頁面訪問權限（必須在 session 之後）
const protectHtmlPages = (req, res, next) => {
    // 允許 API 路由
    if (req.path.startsWith('/api/')) {
        return next();
    }
    
    // 允許登入頁面
    if (req.path === '/login.html' || req.path === '/login') {
        return next();
    }
    
    // 允許靜態資源（CSS、JS、圖片等）
    if (req.path.match(/\.(css|js|png|jpg|jpeg|gif|ico|svg|woff|woff2|ttf|eot)$/)) {
        return next();
    }
    
    // 檢查是否需要認證的頁面（HTML 檔案或根路徑）
    if (req.path === '/' || req.path.endsWith('.html')) {
        if (!req.session || !req.session.user) {
            // 未登入，重定向到登入頁
            return res.redirect('/login.html');
        }
    }
    
    next();
};

// 應用路由保護（在靜態檔案服務之前）
app.use(protectHtmlPages);

// 靜態檔案服務
app.use(express.static(path.join(__dirname, 'public')));

// 權限檢查中間件
const requireAdmin = (req, res, next) => {
    if (!req.session || !req.session.user || req.session.user.role !== 'admin') {
        return res.status(403).json({ error: 'Denied' });
    }
    next();
};

const requireAdminOrManager = (req, res, next) => {
    if (!req.session || !req.session.user || !['admin', 'manager'].includes(req.session.user.role)) {
        return res.status(403).json({ error: 'Denied' });
    }
    next();
};

// 統一的 API 錯誤處理函數
function handleApiError(e, req, res, context) {
    // 記錄錯誤日誌
    logError(e, context, req).catch(() => {});
    
    // 根據環境決定錯誤訊息
    const errorMessage = process.env.NODE_ENV === 'production' 
        ? '伺服器錯誤，請稍後再試' 
        : e.message;
    
    res.status(500).json({ error: errorMessage });
}

// CSRF 保護設定
const csrfProtection = new csrf();
const getCsrfToken = (req, res, next) => {
    if (!req.session.csrfSecret) {
        req.session.csrfSecret = csrfProtection.secretSync();
    }
    req.csrfToken = csrfProtection.create(req.session.csrfSecret);
    next();
};

// CSRF 驗證中間件（僅用於需要保護的路由）
const verifyCsrf = (req, res, next) => {
    if (req.method === 'GET' || req.method === 'HEAD' || req.method === 'OPTIONS') {
        return next();
    }
    
    try {
        const token = req.headers['x-csrf-token'] || req.body._csrf;
        const secret = req.session.csrfSecret;
        
        if (!secret || !token) {
            return res.status(403).json({ error: 'CSRF token missing' });
        }
        
        if (!csrfProtection.verify(secret, token)) {
            return res.status(403).json({ error: 'Invalid CSRF token' });
        }
        
        next();
    } catch (e) {
        console.error('CSRF verification error:', e);
        logError(e, 'CSRF verification error', req).catch(() => {});
        return res.status(500).json({ error: 'CSRF verification failed' });
    }
};

// 為所有需要認證的路由提供 CSRF token
app.use('/api/', getCsrfToken);

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
                    must_change_password BOOLEAN DEFAULT true,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )`);
                
                // 新增 must_change_password 欄位（如果不存在）
                try {
                    await client.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS must_change_password BOOLEAN DEFAULT true`);
                } catch (e) {}

                // 檢查計畫單一表（原 inspection_plan_schedule）：月曆排程 + 取號 等，不再使用 inspection_plans
                await client.query(`CREATE TABLE IF NOT EXISTS inspection_plan_schedule (
                    id SERIAL PRIMARY KEY,
                    start_date DATE NOT NULL,
                    end_date DATE,
                    plan_name TEXT NOT NULL,
                    year TEXT NOT NULL,
                    railway TEXT NOT NULL,
                    inspection_type TEXT NOT NULL,
                    business TEXT NOT NULL,
                    inspection_seq TEXT NOT NULL,
                    plan_number TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )`);
                // 向後兼容：如果存在 scheduled_date 欄位，遷移到 start_date
                try {
                    await client.query(`ALTER TABLE inspection_plan_schedule ADD COLUMN IF NOT EXISTS start_date DATE`);
                    await client.query(`ALTER TABLE inspection_plan_schedule ADD COLUMN IF NOT EXISTS end_date DATE`);
                    await client.query(`UPDATE inspection_plan_schedule SET start_date = scheduled_date WHERE start_date IS NULL AND scheduled_date IS NOT NULL`);
                } catch (e) {}
                try {
                    await client.query(`ALTER TABLE inspection_plan_schedule DROP COLUMN IF EXISTS scheduled_date`);
                } catch (e) {}
                try {
                    await client.query(`DROP INDEX IF EXISTS idx_schedule_date`);
                } catch (e) {}
                try {
                    await client.query(`CREATE INDEX IF NOT EXISTS idx_schedule_start_date ON inspection_plan_schedule(start_date)`);
                    await client.query(`CREATE INDEX IF NOT EXISTS idx_schedule_year ON inspection_plan_schedule(year)`);
                } catch (e) {}

                // 遷移：若曾使用 inspection_plans，將僅存於該表的 (name,year) 補進 schedule 後刪除 inspection_plans
                try {
                    const hasPlans = await client.query(
                        "SELECT 1 FROM information_schema.tables WHERE table_schema='public' AND table_name='inspection_plans'"
                    );
                    if (hasPlans.rows.length > 0) {
                        const planRows = await client.query(
                            "SELECT name, year, created_at FROM inspection_plans"
                        );
                        for (const r of planRows.rows || []) {
                            const n = (r.name || '').trim();
                            const y = (r.year || '').trim();
                            if (!n) continue;
                            const ex = await client.query(
                                "SELECT 1 FROM inspection_plan_schedule WHERE plan_name = $1 AND year = $2 LIMIT 1",
                                [n, y]
                            );
                            if (ex.rows.length > 0) continue;
                            const sd = r.created_at ? new Date(r.created_at).toISOString().slice(0, 10) : '2000-01-01';
                            await client.query(
                                `INSERT INTO inspection_plan_schedule (start_date, end_date, plan_name, year, railway, inspection_type, business, inspection_seq, plan_number)
                                 VALUES ($1, NULL, $2, $3, '-', '-', '-', '00', '(手動)')`,
                                [sd, n, y]
                            );
                        }
                        await client.query(`DROP TABLE IF EXISTS inspection_plans CASCADE`);
                    }
                } catch (e) {
                    console.warn('inspection_plans migration warning:', e?.message || e);
                }

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
                // 支持無限次審查，預先創建前 30 次欄位（如果需要更多可以動態創建）
                for (let i = 2; i <= 30; i++) {
                    newColumns.push({ name: `handling${i}`, type: 'TEXT' });
                    newColumns.push({ name: `review${i}`, type: 'TEXT' });
                }
                for (let i = 1; i <= 30; i++) {
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
                    await client.query("INSERT INTO users (username, password, name, role, must_change_password) VALUES ($1, $2, $3, $4, $5)", 
                        ['admin', hash, '系統管理員', 'admin', true]);
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
        // 如果是登入動作，同時寫入 login_time
        if (action === 'LOGIN') {
            await pool.query("INSERT INTO logs (username, action, details, ip_address, login_time, created_at) VALUES ($1, $2, $3, $4, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)", 
                [username, action, details, ip]);
        } else {
            await pool.query("INSERT INTO logs (username, action, details, ip_address, created_at) VALUES ($1, $2, $3, $4, CURRENT_TIMESTAMP)", 
                [username, action, details, ip]);
        }
    } catch (e) { 
        console.error("Log error:", e);
        // 記錄錯誤到檔案
        writeToLogFile(`Error logging action: ${e.message}`, 'ERROR');
    }
}

// 錯誤日誌記錄函數（記錄到資料庫）
async function logError(error, context, req) {
    try {
        const ip = req ? (req.headers['x-forwarded-for'] || req.socket.remoteAddress) : 'system';
        const username = req?.session?.user?.username || 'system';
        const errorMessage = error instanceof Error ? error.message : String(error);
        const errorStack = error instanceof Error ? error.stack : '';
        const details = `${context}: ${errorMessage}${errorStack ? `\nStack: ${errorStack.substring(0, 500)}` : ''}`;
        
        // 嘗試記錄到資料庫，如果失敗則只記錄到檔案
        try {
            await pool.query(
                "INSERT INTO logs (username, action, details, ip_address, created_at) VALUES ($1, $2, $3, $4, CURRENT_TIMESTAMP)",
                [username, 'ERROR', details, ip]
            );
        } catch (dbError) {
            // 資料庫記錄失敗，只記錄到檔案
            console.error("Failed to log error to database:", dbError);
        }
        
        // 同時寫入檔案日誌
        writeToLogFile(`[ERROR] ${context}: ${errorMessage}`, 'ERROR');
    } catch (e) {
        // 如果整個錯誤記錄過程失敗，至少輸出到 console
        console.error("Failed to log error:", e);
        console.error("Original error:", error);
    }
}

// 寫入日誌檔案
function writeToLogFile(message, level = 'INFO') {
    try {
        const logDir = path.join(__dirname, 'logs');
        if (!fs.existsSync(logDir)) {
            fs.mkdirSync(logDir, { recursive: true });
        }
        const today = new Date().toISOString().split('T')[0];
        const logFile = path.join(logDir, `app-${today}.log`);
        const timestamp = new Date().toISOString();
        const logEntry = `[${timestamp}] [${level}] ${message}\n`;
        fs.appendFileSync(logFile, logEntry, 'utf8');
    } catch (e) {
        console.error("Write log file error:", e);
    }
}

// 密碼複雜度驗證函數
function validatePassword(password) {
    if (!password || typeof password !== 'string') {
        return { valid: false, message: '密碼不能為空' };
    }
    
    if (password.length < 8) {
        return { valid: false, message: '密碼至少需要 8 個字元' };
    }
    
    if (!/[A-Z]/.test(password)) {
        return { valid: false, message: '密碼必須包含至少一個大寫字母' };
    }
    
    if (!/[a-z]/.test(password)) {
        return { valid: false, message: '密碼必須包含至少一個小寫字母' };
    }
    
    if (!/[0-9]/.test(password)) {
        return { valid: false, message: '密碼必須包含至少一個數字' };
    }
    
    return { valid: true };
}

// 日誌輪轉機制：清理舊日誌檔案
function cleanupOldLogs() {
    try {
        const logDir = path.join(__dirname, 'logs');
        if (!fs.existsSync(logDir)) {
            return;
        }
        
        const maxAge = 30 * 24 * 60 * 60 * 1000; // 30 天
        const now = Date.now();
        
        fs.readdir(logDir, (err, files) => {
            if (err) {
                console.error('Error reading log directory:', err);
                return;
            }
            
            files.forEach(file => {
                if (!file.startsWith('app-') || !file.endsWith('.log')) {
                    return;
                }
                
                const filePath = path.join(logDir, file);
                fs.stat(filePath, (err, stats) => {
                    if (err) {
                        return;
                    }
                    
                    const fileAge = now - stats.mtime.getTime();
                    if (fileAge > maxAge) {
                        fs.unlink(filePath, (err) => {
                            if (err) {
                                console.error(`Error deleting old log file ${file}:`, err);
                            } else {
                                console.log(`Deleted old log file: ${file}`);
                            }
                        });
                    }
                });
            });
        });
    } catch (e) {
        console.error('Error in cleanupOldLogs:', e);
    }
}

// 啟動時執行一次日誌清理，然後每天執行一次
cleanupOldLogs();
setInterval(cleanupOldLogs, 24 * 60 * 60 * 1000); // 每 24 小時執行一次

// API: 取得 CSRF token
app.get('/api/csrf-token', (req, res) => {
    if (!req.session.csrfSecret) {
        req.session.csrfSecret = csrfProtection.secretSync();
    }
    const token = csrfProtection.create(req.session.csrfSecret);
    res.json({ csrfToken: token });
});

// API: 接收前端日誌
app.post('/api/log', requireAuth, verifyCsrf, (req, res) => {
    try {
        const { message, level = 'INFO' } = req.body;
        if (message) {
            writeToLogFile(message, level);
            res.json({ success: true });
        } else {
            res.status(400).json({ error: 'Message is required' });
        }
    } catch (e) {
        console.error("Log API error:", e);
        logError(e, 'Log API error', req).catch(() => {});
        res.status(500).json({ error: 'Failed to write log' });
    }
});

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
                // 檢查是否需要更新密碼
                const mustChangePassword = user.must_change_password === true || user.must_change_password === null;
                res.json({ 
                    success: true, 
                    user: req.session.user,
                    mustChangePassword: mustChangePassword
                });
            });
        } else {
            res.status(401).json({ error: 'Invalid credentials' });
        }
    } catch (e) {
        console.error("Login Error:", e);
        logError(e, 'Login error', req).catch(() => {});
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

app.put('/api/auth/profile', requireAuth, verifyCsrf, async (req, res) => {
    const { name, password } = req.body;
    const id = req.session.user.id;
    try {
        if (password) {
            // 驗證密碼複雜度
            const passwordValidation = validatePassword(password);
            if (!passwordValidation.valid) {
                return res.status(400).json({ error: passwordValidation.message });
            }
            const hash = bcrypt.hashSync(password, 10);
            await pool.query("UPDATE users SET name = $1, password = $2, must_change_password = $3 WHERE id = $4", [name, hash, false, id]);
            logAction(req.session.user.username, 'UPDATE_PROFILE', `更新個人資料：已更新姓名為「${name}」並變更密碼`, req);
        } else {
            await pool.query("UPDATE users SET name = $1 WHERE id = $2", [name, id]);
            logAction(req.session.user.username, 'UPDATE_PROFILE', `更新個人資料：已更新姓名為「${name}」`, req);
        }
        res.json({ success: true });
    } catch (e) { 
        logError(e, 'Update profile error', req).catch(() => {});
        res.status(500).json({ error: e.message }); 
    }
});

// 首次登入強制更新密碼 API
app.post('/api/auth/change-password', requireAuth, verifyCsrf, async (req, res) => {
    const { password } = req.body;
    const id = req.session.user.id;
    try {
        if (!password) {
            return res.status(400).json({ error: '密碼為必填項目' });
        }
        
        // 驗證密碼複雜度
        const passwordValidation = validatePassword(password);
        if (!passwordValidation.valid) {
            return res.status(400).json({ error: passwordValidation.message });
        }
        
        // 更新密碼並清除 must_change_password 標記
        const hash = bcrypt.hashSync(password, 10);
        await pool.query("UPDATE users SET password = $1, must_change_password = $2 WHERE id = $3", [hash, false, id]);
        
        logAction(req.session.user.username, 'CHANGE_PASSWORD', 'User changed password (first login)', req).catch(()=>{});
        res.json({ success: true });
    } catch (e) {
        handleApiError(e, req, res, 'Change password error');
    }
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
    // 修正：如果提供了計畫名稱，需要同時考慮年度來精確匹配
    // planName 參數現在可能是 "planName|||year" 格式，或者只有 planName
    if (planName) {
        const planParts = planName.split('|||');
        const actualPlanName = planParts[0];
        const planYear = planParts[1];
        
        if (planYear) {
            // 如果提供了年度，同時匹配計畫名稱和年度
            where.push(`plan_name = $${idx} AND year = $${idx+1}`);
            params.push(actualPlanName, planYear);
            idx += 2;
        } else {
            // 如果沒有提供年度，只匹配計畫名稱（向後兼容）
            where.push(`plan_name = $${idx}`);
            params.push(actualPlanName);
            idx++;
        }
    }

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
    } catch (e) { 
        handleApiError(e, req, res, 'Get issues error');
    }
});

app.put('/api/issues/:id', requireAuth, verifyCsrf, async (req, res) => {
    const { status, round, handling, review, replyDate, responseDate, content, issueDate, 
            number, year, unit, divisionName, inspectionCategoryName, itemKindCode, category, planName } = req.body;
    const id = req.params.id;
    const r = parseInt(round) || 1;
    const hField = r === 1 ? 'handling' : `handling${r}`;
    const rField = r === 1 ? 'review' : `review${r}`;
    const replyField = `reply_date_r${r}`;
    const respField = `response_date_r${r}`;
    try {
        // 先查詢 issue number
        const issueRes = await pool.query("SELECT number FROM issues WHERE id=$1", [id]);
        const issueNumber = issueRes.rows[0]?.number || `ID:${id}`;
        
        // 如果超過預設欄位範圍（30次），動態創建欄位
        if (r > 30) {
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
        
        // 構建更新語句，如果提供了 content 或 issueDate 則包含它們
        // 注意：replyDate 和 responseDate 如果提供空字符串，應該明確更新為空字符串
        // 如果未提供（undefined），則不更新該欄位
        let updateFields = [`status=$1`, `${hField}=$2`, `${rField}=$3`, `updated_at=CURRENT_TIMESTAMP`];
        let params = [status, handling || '', review || ''];
        let paramIdx = 4;
        
        // 處理 replyDate：如果提供了（即使是空字符串），也要更新
        if (replyDate !== undefined) {
            updateFields.splice(updateFields.length - 1, 0, `${replyField}=$${paramIdx}`);
            params.push(replyDate || '');
            paramIdx++;
        }
        
        // 處理 responseDate：如果提供了（即使是空字符串），也要更新
        if (responseDate !== undefined) {
            updateFields.splice(updateFields.length - 1, 0, `${respField}=$${paramIdx}`);
            params.push(responseDate || '');
            paramIdx++;
        }
        
        if (content !== undefined) {
            updateFields.push(`content=$${paramIdx}`);
            params.push(content);
            paramIdx++;
        }
        
        if (issueDate !== undefined) {
            updateFields.push(`issue_date=$${paramIdx}`);
            params.push(issueDate);
            paramIdx++;
        }
        
        // 支持更新更多字段
        if (number !== undefined) {
            updateFields.push(`number=$${paramIdx}`);
            params.push(number);
            paramIdx++;
        }
        
        if (year !== undefined) {
            updateFields.push(`year=$${paramIdx}`);
            params.push(year);
            paramIdx++;
        }
        
        if (unit !== undefined) {
            updateFields.push(`unit=$${paramIdx}`);
            params.push(unit);
            paramIdx++;
        }
        
        if (divisionName !== undefined) {
            updateFields.push(`division_name=$${paramIdx}`);
            params.push(divisionName);
            paramIdx++;
        }
        
        if (inspectionCategoryName !== undefined) {
            updateFields.push(`inspection_category_name=$${paramIdx}`);
            params.push(inspectionCategoryName);
            paramIdx++;
        }
        
        if (itemKindCode !== undefined) {
            updateFields.push(`item_kind_code=$${paramIdx}`);
            params.push(itemKindCode);
            paramIdx++;
        }
        
        if (category !== undefined) {
            updateFields.push(`category=$${paramIdx}`);
            params.push(category);
            paramIdx++;
        }
        
        if (planName !== undefined) {
            updateFields.push(`plan_name=$${paramIdx}`);
            params.push(planName);
            paramIdx++;
        }
        
        params.push(id);
        await pool.query(`UPDATE issues SET ${updateFields.join(', ')} WHERE id=$${paramIdx}`, params);
        const actionDetails = `更新開立事項：編號 ${issueNumber}，第 ${r} 次審查，狀態：${status}${content !== undefined ? '，內容已更新' : ''}${issueDate !== undefined ? '，開立日期已更新' : ''}${number !== undefined ? '，編號已更新' : ''}${year !== undefined ? '，年度已更新' : ''}${unit !== undefined ? '，機構已更新' : ''}${divisionName !== undefined ? '，分組已更新' : ''}${inspectionCategoryName !== undefined ? '，檢查種類已更新' : ''}${itemKindCode !== undefined ? '，類型已更新' : ''}${planName !== undefined ? '，檢查計畫已更新' : ''}`;
        logAction(req.session.user.username, 'UPDATE_ISSUE', actionDetails, req);
        res.json({ success: true });
    } catch (e) { 
        handleApiError(e, req, res, 'Update issue error');
    }
});

app.delete('/api/issues/:id', requireAuth, requireAdminOrManager, verifyCsrf, async (req, res) => {
    try {
        // 先查詢 issue number 再刪除
        const issueRes = await pool.query("SELECT number FROM issues WHERE id=$1", [req.params.id]);
        const issueNumber = issueRes.rows[0]?.number || `ID:${req.params.id}`;
        
        await pool.query("DELETE FROM issues WHERE id=$1", [req.params.id]);
        logAction(req.session.user.username, 'DELETE_ISSUE', `刪除開立事項：編號 ${issueNumber}`, req);
        res.json({success:true});
    } catch (e) { 
        logError(e, 'Delete issue error', req).catch(() => {});
        res.status(500).json({ error: e.message }); 
    }
});

app.post('/api/issues/batch-delete', requireAuth, verifyCsrf, async (req, res) => {
    if (!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    const { ids } = req.body;
    try {
        // 先查詢所有要刪除的編號
        const issueRes = await pool.query("SELECT number FROM issues WHERE id = ANY($1)", [ids]);
        const numbers = issueRes.rows.map(r => r.number).filter(Boolean);
        const numberList = numbers.length > 0 ? numbers.join(', ') : `${ids.length} 筆`;
        
        await pool.query("DELETE FROM issues WHERE id = ANY($1)", [ids]);
        logAction(req.session.user.username, 'BATCH_DELETE_ISSUES', `批次刪除開立事項：${numberList} (共 ${ids.length} 筆)`, req);
        res.json({success:true});
    } catch (e) { 
        handleApiError(e, req, res, 'Batch delete issues error');
    }
});

app.post('/api/issues/import', requireAuth, requireAdminOrManager, verifyCsrf, async (req, res) => {
    const { data, round, reviewDate, replyDate, allowUpdate } = req.body;
    const r = parseInt(round) || 1;
    const client = await pool.connect();
    try {
        await client.query('BEGIN');
        const duplicateNumbers = [];
        const operationResults = []; // 記錄每個項目的操作類型
        
        for (const item of data) {
            // 使用精確匹配查詢編號（區分大小寫，去除前後空格）
            const trimmedNumber = (item.number || '').trim();
            const check = await client.query("SELECT id, content FROM issues WHERE TRIM(number) = $1", [trimmedNumber]);
            if (check.rows.length > 0) {
                // 如果是新增事項（round=1）且不允許更新，檢查內容是否相同
                // 只有在編號已存在、內容不同、且現有內容不為空時，才視為重複編號錯誤
                if (r === 1 && !allowUpdate) {
                    const existingContent = (check.rows[0].content || '').trim();
                    const newContent = (item.content || '').trim();
                    // 只有在明確是重複編號且內容不同時才報錯
                    // 如果現有內容為空，視為可以更新（可能是之前新增失敗留下的空記錄）
                    if (existingContent !== '' && newContent !== '' && existingContent !== newContent) {
                        duplicateNumbers.push({
                            number: trimmedNumber,
                            existingContent: existingContent
                        });
                        continue; // 跳過這個項目，不進行更新
                    }
                    // 如果內容相同或現有內容為空，允許更新（視為正常的新增/更新操作）
                }
                
                // 允許更新：更新現有記錄
                const hCol = r===1 ? 'handling' : `handling${r}`;
                const rCol = r===1 ? 'review' : `review${r}`;
                const replyCol = `reply_date_r${r}`;
                const respCol = `response_date_r${r}`;
                
                // 如果是新增事項（round=1），也更新內容和其他欄位
                // 優先使用 item.replyDate，如果沒有則使用統一的 replyDate
                const itemReplyDate = item.replyDate || replyDate || '';
                if (r === 1) {
                    await client.query(
                        `UPDATE issues SET 
                            status=$1, content=$2, ${hCol}=$3, ${rCol}=$4, ${replyCol}=$5, ${respCol}=$6,
                            plan_name=COALESCE($7, plan_name), issue_date=COALESCE($8, issue_date),
                            year=COALESCE($9, year), unit=COALESCE($10, unit),
                            division_name=COALESCE($11, division_name),
                            inspection_category_name=COALESCE($12, inspection_category_name),
                            item_kind_code=COALESCE($13, item_kind_code),
                            updated_at=CURRENT_TIMESTAMP 
                        WHERE TRIM(number)=$14`,
                        [
                            item.status, item.content, item.handling||'', item.review||'', 
                            itemReplyDate, reviewDate||'', item.planName || null, item.issueDate || null,
                            item.year || null, item.unit || null,
                            item.divisionName || null, item.inspectionCategoryName || null,
                            item.itemKindCode || null, trimmedNumber
                        ]
                    );
                    // 記錄為更新操作
                    operationResults.push({ number: trimmedNumber, action: 'updated' });
                } else {
                    // 更新輪次資料
                    await client.query(
                        `UPDATE issues SET 
                            status=$1, ${hCol}=$2, ${rCol}=$3, ${replyCol}=$4, ${respCol}=$5,
                            plan_name=COALESCE($6, plan_name), updated_at=CURRENT_TIMESTAMP 
                        WHERE TRIM(number)=$7`,
                        [item.status, item.handling||'', item.review||'', itemReplyDate, reviewDate||'', item.planName || null, trimmedNumber]
                    );
                    operationResults.push({ number: trimmedNumber, action: 'updated' });
                }
            } else {
                // 新增記錄（使用trimmedNumber確保編號沒有前後空格）
                // 優先使用 item.replyDate，如果沒有則使用統一的 replyDate
                const itemReplyDate = item.replyDate || replyDate || '';
                await client.query(
                    `INSERT INTO issues (
                        number, year, unit, content, status, item_kind_code, category, division_name, inspection_category_name,
                        handling, review, plan_name, issue_date, response_date_r1, reply_date_r1
                    ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15)`,
                    [
                        trimmedNumber, item.year, item.unit, item.content, item.status||'持續列管',
                        item.itemKindCode, item.category, item.divisionName, item.inspectionCategoryName,
                        item.handling||'', item.review||'', item.planName || null, item.issueDate || null, 
                        reviewDate || '', itemReplyDate
                    ]
                );
                // 記錄為新增操作
                operationResults.push({ number: trimmedNumber, action: 'created' });
            }
        }
        
        // 如果有重複編號且內容不同，回滾事務並返回錯誤
        if (duplicateNumbers.length > 0) {
            await client.query('ROLLBACK');
            return res.status(400).json({
                error: '編號重複',
                message: `以下編號已存在且內容不同：${duplicateNumbers.map(d => d.number).join(', ')}`,
                duplicates: duplicateNumbers
            });
        }
        
        await client.query('COMMIT');
        
        // 統計新增和更新的項目（使用操作記錄）
        let newCount = 0, updateCount = 0;
        const results = operationResults.map(op => {
            if (op.action === 'created') {
                newCount++;
            } else {
                updateCount++;
            }
            return op;
        });
        
        const roundInfo = r > 1 ? `，第 ${r} 次審查` : '，初次開立';
        const planInfo = data[0]?.planName ? `，檢查計畫：${data[0].planName}` : '';
        logAction(req.session.user.username, 'IMPORT_ISSUES', `匯入開立事項：共 ${data.length} 筆（新增 ${newCount} 筆，更新 ${updateCount} 筆）${roundInfo}${planInfo}`, req);
        res.json({ 
            success: true, 
            count: data.length,
            newCount: newCount,
            updateCount: updateCount,
            results: results
        });
    } catch (e) {
        await client.query('ROLLBACK');
        handleApiError(e, req, res, 'Import issues error');
    } finally {
        client.release();
    }
});

// --- User Management API ---

app.get('/api/users', requireAuth, requireAdmin, async (req, res) => {
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
    } catch (e) { 
        handleApiError(e, req, res, 'Get users error');
    }
});

app.post('/api/users', requireAuth, requireAdmin, verifyCsrf, async (req, res) => {
    const { username, password, name, role } = req.body;
    try {
        // Basic Validation
        if (!username || !password) return res.status(400).json({error: 'Username and password required'});
        
        // 驗證密碼複雜度
        const passwordValidation = validatePassword(password);
        if (!passwordValidation.valid) {
            return res.status(400).json({ error: passwordValidation.message });
        }
        
        const hash = bcrypt.hashSync(password, 10);
        await pool.query("INSERT INTO users (username, password, name, role, must_change_password) VALUES ($1, $2, $3, $4, $5)", [username, hash, name, role, true]);
        logAction(req.session.user.username, 'CREATE_USER', `新增使用者：${name} (${username})，權限：${role}`, req);
        res.json({success:true});
    } catch (e) { 
        handleApiError(e, req, res, 'Create user error');
    }
});

app.put('/api/users/:id', requireAuth, requireAdmin, verifyCsrf, async (req, res) => {
    const { name, password, role } = req.body;
    const id = req.params.id;
    try {
        // 先查詢使用者資訊以便記錄
        const userRes = await pool.query("SELECT username, name FROM users WHERE id=$1", [id]);
        const targetUser = userRes.rows[0];
        const targetUsername = targetUser ? targetUser.username : `ID:${id}`;
        const targetName = targetUser ? targetUser.name : '未知';
        
        if (password) {
            // 驗證密碼複雜度
            const passwordValidation = validatePassword(password);
            if (!passwordValidation.valid) {
                return res.status(400).json({ error: passwordValidation.message });
            }
            const hash = bcrypt.hashSync(password, 10);
            await pool.query("UPDATE users SET name=$1, role=$2, password=$3, must_change_password=$4 WHERE id=$5", [name, role, hash, true, id]);
            logAction(req.session.user.username, 'UPDATE_USER', `修改使用者：${targetName} (${targetUsername})，已更新姓名、權限和密碼`, req);
        } else {
            await pool.query("UPDATE users SET name=$1, role=$2 WHERE id=$3", [name, role, id]);
            logAction(req.session.user.username, 'UPDATE_USER', `修改使用者：${targetName} (${targetUsername})，已更新姓名和權限`, req);
        }
        res.json({success:true});
    } catch (e) { 
        handleApiError(e, req, res, 'Update user error');
    }
});

app.delete('/api/users/:id', requireAuth, requireAdmin, verifyCsrf, async (req, res) => {
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
    } catch (e) { 
        handleApiError(e, req, res, 'Delete user error');
    }
});

// 帳號匯入 API
app.post('/api/users/import', requireAuth, requireAdmin, verifyCsrf, async (req, res) => {
    const { data } = req.body;
    if (!data || !Array.isArray(data)) return res.status(400).json({error: '無效的資料格式'});
    
    const results = { success: 0, failed: 0, errors: [] };
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const { name, username, role, password } = row;
        
        // 驗證必填欄位
        if (!name || !username || !role) {
            results.failed++;
            results.errors.push(`第 ${i + 2} 行：姓名、帳號和權限為必填`);
            continue;
        }
        
        // 驗證權限值
        const validRoles = ['admin', 'manager', 'editor', 'viewer'];
        if (!validRoles.includes(role.toLowerCase())) {
            results.failed++;
            results.errors.push(`第 ${i + 2} 行（${name}）：無效的權限值 "${role}"，應為：${validRoles.join(', ')}`);
            continue;
        }
        
        try {
            // 檢查是否已存在相同帳號
            const checkRes = await pool.query("SELECT id FROM users WHERE username = $1", [username]);
            const exists = checkRes.rows.length > 0;
            
            if (exists) {
                // 如果已存在，更新資料（但不更新密碼，除非有提供）
                if (password) {
                    // 驗證密碼複雜度
                    const passwordValidation = validatePassword(password);
                    if (!passwordValidation.valid) {
                        results.failed++;
                        results.errors.push(`第 ${i + 2} 行（${name}）：${passwordValidation.message}`);
                        continue;
                    }
                    const hash = bcrypt.hashSync(password, 10);
                    await pool.query(
                        "UPDATE users SET name=$1, role=$2, password=$3, must_change_password=$4 WHERE username=$5",
                        [name, role.toLowerCase(), hash, true, username]
                    );
                } else {
                    await pool.query(
                        "UPDATE users SET name=$1, role=$2 WHERE username=$3",
                        [name, role.toLowerCase(), username]
                    );
                }
                results.success++;
            } else {
                // 如果不存在，新增帳號
                // 如果沒有提供密碼，使用預設密碼（建議在匯入時提供）
                let hash;
                if (password) {
                    // 驗證密碼複雜度
                    const passwordValidation = validatePassword(password);
                    if (!passwordValidation.valid) {
                        results.failed++;
                        results.errors.push(`第 ${i + 2} 行（${name}）：${passwordValidation.message}`);
                        continue;
                    }
                    hash = bcrypt.hashSync(password, 10);
                } else {
                    // 預設密碼為 username@123456（建議匯入時提供密碼）
                    hash = bcrypt.hashSync(`${username}@123456`, 10);
                }
                
                await pool.query(
                    "INSERT INTO users (name, username, role, password, must_change_password) VALUES ($1, $2, $3, $4, $5) RETURNING id",
                    [name, username, role.toLowerCase(), hash, true]
                );
                results.success++;
            }
        } catch (e) {
            results.failed++;
            const errorMsg = `第 ${i + 2} 行（${name}）：${e.message}`;
            results.errors.push(errorMsg);
            logError(e, `Import user error - row ${i + 2}`, req).catch(() => {});
        }
    }
    
    if (results.success > 0) {
        logAction(req.session.user.username, 'IMPORT_USERS', `匯入帳號：成功 ${results.success} 筆，失敗 ${results.failed} 筆`, req);
    }
    
    res.json({
        success: true,
        successCount: results.success,
        failed: results.failed,
        errors: results.errors
    });
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
    } catch (e) { 
        handleApiError(e, req, res, 'Get admin logs error');
    }
});

app.get('/api/admin/action_logs', requireAuth, requireAdmin, async (req, res) => {
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
    } catch (e) { 
        handleApiError(e, req, res, 'Get admin logs error');
    }
});

app.delete('/api/admin/logs', requireAuth, requireAdmin, verifyCsrf, async (req, res) => {
    await pool.query("DELETE FROM logs WHERE action='LOGIN'");
    res.json({success:true});
});

app.delete('/api/admin/action_logs', requireAuth, requireAdmin, verifyCsrf, async (req, res) => {
    await pool.query("DELETE FROM logs WHERE action!='LOGIN'");
    res.json({success:true});
});

// 根據時間範圍清除舊記錄
app.post('/api/admin/logs/cleanup', requireAuth, requireAdmin, verifyCsrf, async (req, res) => {
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
        handleApiError(e, req, res, 'Cleanup logs error');
    }
});

app.post('/api/admin/action_logs/cleanup', requireAuth, requireAdmin, verifyCsrf, async (req, res) => {
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
        handleApiError(e, req, res, 'Cleanup logs error');
    }
});

app.get('/api/options/plans', requireAuth, async (req, res) => {
    try {
        const { withIssues } = req.query;
        
        let planResult;
        try {
            if (withIssues === 'true') {
                planResult = await pool.query(`
                    SELECT DISTINCT s.plan_name AS name, s.year 
                    FROM inspection_plan_schedule s
                    INNER JOIN issues i ON i.plan_name = s.plan_name AND i.year = s.year
                    WHERE s.plan_name IS NOT NULL AND s.plan_name != ''
                        AND i.plan_name IS NOT NULL AND i.plan_name != ''
                        AND i.year IS NOT NULL AND i.year != ''
                        AND s.year IS NOT NULL AND s.year != ''
                    ORDER BY s.year DESC, s.plan_name ASC
                `);
            } else {
                planResult = await pool.query(`
                    SELECT DISTINCT plan_name AS name, year 
                    FROM inspection_plan_schedule 
                    WHERE plan_name IS NOT NULL AND plan_name != ''
                        AND year IS NOT NULL AND year != ''
                    ORDER BY year DESC, plan_name ASC
                `);
            }
        } catch (queryError) {
            console.error('Database query error in /api/options/plans:', queryError);
            return res.status(500).json({ error: '查詢資料庫時發生錯誤', details: queryError.message });
        }
        
        res.set('Cache-Control', 'no-store');
        const plans = (planResult?.rows || [])
            .filter(r => r && r.name && String(r.name).trim() !== '')
            .map(r => {
                const name = String(r.name || '').trim();
                const year = String(r.year || '').trim();
                return {
                    name,
                    year,
                    display: `${name}${year ? ` (${year})` : ''}`,
                    value: `${name}|||${year}`
                };
            });
        res.json({ data: plans });
    } catch (e) {
        console.error('Get plan options error:', e);
        handleApiError(e, req, res, 'Get plan options error');
    }
});

// --- Inspection Plans Management API ---

app.get('/api/plans', requireAuth, requireAdminOrManager, async (req, res) => {
    const { page=1, pageSize=20, q, year, sortField='id', sortDir='desc' } = req.query;
    const limit = parseInt(pageSize);
    const offset = (page-1)*limit;
    let where = ["1=1"], params = [], idx = 1;
    if(q) { where.push(`s.plan_name LIKE $${idx}`); params.push(`%${q}%`); idx++; }
    if(year) { where.push(`s.year = $${idx}`); params.push(year); idx++; }
    const safeSortFields = ['id', 'name', 'year', 'created_at', 'updated_at'];
    const safeField = safeSortFields.includes(sortField) ? sortField : 'id';
    const safeSortDir = sortDir === 'asc' ? 'ASC' : 'DESC';
    const orderCol = safeField === 'id' ? 'g.min_id' : safeField === 'name' ? 'g.name' : `g.${safeField}`;
    const order = `${orderCol} ${safeSortDir}`;
    try {
        const countQuery = `
            SELECT count(*) FROM (
                SELECT plan_name, year FROM inspection_plan_schedule s WHERE ${where.join(" AND ")}
                GROUP BY plan_name, year
            ) g`;
        const cRes = await pool.query(countQuery, params);
        const total = parseInt(cRes.rows[0].count);
        
        const dataQuery = `
            WITH g AS (
                SELECT plan_name AS name, year, MIN(id) AS min_id,
                    MIN(created_at) AS created_at, MAX(updated_at) AS updated_at
                FROM inspection_plan_schedule s WHERE ${where.join(" AND ")}
                GROUP BY plan_name, year
            )
            SELECT g.min_id AS id, g.name, g.year, g.created_at, g.updated_at,
                   COALESCE(COUNT(DISTINCT i.id), 0) AS issue_count
            FROM g
            LEFT JOIN issues i ON i.plan_name = g.name AND i.year = g.year
            GROUP BY g.min_id, g.name, g.year, g.created_at, g.updated_at
            ORDER BY ${order}
            LIMIT $${idx} OFFSET $${idx+1}
        `;
        const dRes = await pool.query(dataQuery, [...params, limit, offset]);
        
        const plansWithCounts = dRes.rows.map(row => ({
            id: row.id,
            name: row.name,
            year: row.year,
            created_at: row.created_at,
            updated_at: row.updated_at,
            issue_count: parseInt(row.issue_count) || 0
        }));
        
        res.json({data: plansWithCounts, total, page: parseInt(page), pages: Math.ceil(total/limit)});
    } catch (e) { 
        handleApiError(e, req, res, 'Get plans error');
    }
});

app.get('/api/plans/:id', requireAuth, requireAdminOrManager, async (req, res) => {
    try {
        const result = await pool.query(
            "SELECT id, plan_name AS name, year, created_at, updated_at FROM inspection_plan_schedule WHERE id = $1",
            [req.params.id]
        );
        if (result.rows.length === 0) return res.status(404).json({error: 'Plan not found'});
        res.json(result.rows[0]);
    } catch (e) { 
        handleApiError(e, req, res, 'Get plan by id error');
    }
});

app.get('/api/plans/:id/issues', requireAuth, requireAdminOrManager, async (req, res) => {
    try {
        const planResult = await pool.query(
            "SELECT plan_name AS name, year FROM inspection_plan_schedule WHERE id = $1",
            [req.params.id]
        );
        if (planResult.rows.length === 0) return res.status(404).json({error: 'Plan not found'});
        const planName = planResult.rows[0].name;
        const planYear = planResult.rows[0].year || '';
        
        const { page=1, pageSize=20 } = req.query;
        const limit = parseInt(pageSize);
        const offset = (page-1)*limit;
        
        // 修正：加入年度條件，確保只查詢相同名稱且年度匹配的事項
        const countRes = await pool.query("SELECT count(*) FROM issues WHERE plan_name = $1 AND year = $2", [planName, planYear]);
        const total = parseInt(countRes.rows[0].count);
        const dataRes = await pool.query("SELECT * FROM issues WHERE plan_name = $1 AND year = $2 ORDER BY created_at DESC LIMIT $3 OFFSET $4", [planName, planYear, limit, offset]);
        
        res.json({data: dataRes.rows, total, page: parseInt(page), pages: Math.ceil(total/limit)});
    } catch (e) { 
        handleApiError(e, req, res, 'Get plan issues error');
    }
});

app.get('/api/plans/:id/schedules', requireAuth, requireAdminOrManager, async (req, res) => {
    try {
        const planResult = await pool.query(
            "SELECT plan_name AS name, year FROM inspection_plan_schedule WHERE id = $1",
            [req.params.id]
        );
        if (planResult.rows.length === 0) return res.status(404).json({error: 'Plan not found'});
        const planName = planResult.rows[0].name;
        const planYear = planResult.rows[0].year || '';
        
        const scheduleRes = await pool.query(
            `SELECT id, start_date, end_date, plan_number, inspection_seq, railway, inspection_type, business 
             FROM inspection_plan_schedule 
             WHERE plan_name = $1 AND year = $2 
             ORDER BY start_date ASC, id ASC`,
            [planName, planYear]
        );
        
        res.json({data: scheduleRes.rows || []});
    } catch (e) { 
        handleApiError(e, req, res, 'Get plan schedules error');
    }
});

app.post('/api/plans', requireAuth, requireAdminOrManager, verifyCsrf, async (req, res) => {
    const { name, year } = req.body;
    try {
        if (!name || !year) return res.status(400).json({error: '計畫名稱和年度為必填'});
        const n = name.trim();
        const y = year.trim();
        const today = new Date().toISOString().slice(0, 10);
        const exists = await pool.query(
            "SELECT 1 FROM inspection_plan_schedule WHERE plan_name = $1 AND year = $2 LIMIT 1",
            [n, y]
        );
        if (exists.rows.length > 0) {
            return res.status(400).json({ error: `計畫名稱「${n}」在年度「${y}」已存在` });
        }
        await pool.query(
            `INSERT INTO inspection_plan_schedule (start_date, end_date, plan_name, year, railway, inspection_type, business, inspection_seq, plan_number)
             VALUES ($1, NULL, $2, $3, '-', '-', '-', '00', '(手動)')`,
            [today, n, y]
        );
        logAction(req.session.user.username, 'CREATE_PLAN', `新增檢查計畫：${n} (年度：${y})`, req);
        res.json({success:true});
    } catch (e) { 
        handleApiError(e, req, res, 'Create plan error');
    }
});

app.put('/api/plans/:id', requireAuth, requireAdminOrManager, verifyCsrf, async (req, res) => {
    const { name, year } = req.body;
    const id = req.params.id;
    try {
        const planRes = await pool.query(
            "SELECT plan_name AS name, year FROM inspection_plan_schedule WHERE id = $1",
            [id]
        );
        if (planRes.rows.length === 0) return res.status(404).json({error: 'Plan not found'});
        const oldName = planRes.rows[0].name;
        const oldYear = planRes.rows[0].year || '';
        
        if (!name || !year) return res.status(400).json({error: '計畫名稱和年度為必填'});
        const n = name.trim();
        const y = year.trim();
        
        if (n !== oldName || y !== oldYear) {
            const conflict = await pool.query(
                "SELECT 1 FROM inspection_plan_schedule WHERE plan_name = $1 AND year = $2 LIMIT 1",
                [n, y]
            );
            if (conflict.rows.length > 0) {
                return res.status(400).json({ error: '計畫名稱與年度組合已存在' });
            }
            await pool.query(
                "UPDATE issues SET plan_name = $1, year = $2 WHERE plan_name = $3 AND year = $4",
                [n, y, oldName, oldYear]
            );
            await pool.query(
                "UPDATE inspection_plan_schedule SET plan_name = $1, year = $2, updated_at = CURRENT_TIMESTAMP WHERE plan_name = $3 AND year = $4",
                [n, y, oldName, oldYear]
            );
        }
        logAction(req.session.user.username, 'UPDATE_PLAN', `修改檢查計畫：${oldName} → ${n} (年度：${y})`, req);
        res.json({success:true});
    } catch (e) { 
        handleApiError(e, req, res, 'Update plan error');
    }
});

app.delete('/api/plans/:id', requireAuth, requireAdminOrManager, verifyCsrf, async (req, res) => {
    try {
        const planRes = await pool.query(
            "SELECT plan_name AS name, year FROM inspection_plan_schedule WHERE id = $1",
            [req.params.id]
        );
        if (planRes.rows.length === 0) return res.status(404).json({error: 'Plan not found'});
        const planName = planRes.rows[0].name;
        const planYear = planRes.rows[0].year || '';
        
        const issueCount = await pool.query("SELECT count(*) FROM issues WHERE plan_name = $1 AND year = $2", [planName, planYear]);
        const count = parseInt(issueCount.rows[0].count);
        if (count > 0) {
            return res.status(400).json({error: `無法刪除計畫，因為尚有 ${count} 筆相關開立事項。請先刪除或轉移相關事項。`});
        }
        
        if (planYear) {
            await pool.query("DELETE FROM inspection_plan_schedule WHERE plan_name = $1 AND year = $2", [planName, planYear]);
        } else {
            await pool.query("DELETE FROM inspection_plan_schedule WHERE plan_name = $1", [planName]);
        }
        logAction(req.session.user.username, 'DELETE_PLAN', `刪除檢查計畫：${planName}${planYear ? ` (年度：${planYear})` : ''}`, req);
        res.json({success:true});
    } catch (e) { 
        handleApiError(e, req, res, 'Delete plan error');
    }
});

// 檢查計畫 CSV 匯入 API
app.post('/api/plans/import', requireAuth, requireAdminOrManager, verifyCsrf, async (req, res) => {
    const { data } = req.body; // 接收解析後的 CSV 資料
    if (!data || !Array.isArray(data)) return res.status(400).json({error: '無效的資料格式'});
    
    // 收到匯入資料（日誌已移除，只在需要時記錄錯誤）
    
    const results = { success: 0, failed: 0, errors: [], skipped: 0 };
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        let name = '', year = '', start_date = '', end_date = '', railway = '', inspection_type = '', business = '';
        
        for (const key in row) {
            const cleanKey = key.trim();
            if (cleanKey === '計畫名稱' || cleanKey === 'name' || cleanKey === 'planName' || cleanKey === '計劃名稱') {
                name = String(row[key] || '').trim();
            } else if (cleanKey === '年度' || cleanKey === 'year') {
                year = String(row[key] || '').trim();
            } else if (cleanKey === '開始日期' || cleanKey === 'start_date' || cleanKey === 'startDate') {
                start_date = String(row[key] || '').trim();
            } else if (cleanKey === '結束日期' || cleanKey === 'end_date' || cleanKey === 'endDate') {
                end_date = String(row[key] || '').trim();
            } else if (cleanKey === '鐵路機構' || cleanKey === 'railway') {
                railway = String(row[key] || '').trim().toUpperCase();
            } else if (cleanKey === '檢查類別' || cleanKey === 'inspection_type' || cleanKey === 'inspectionType') {
                inspection_type = String(row[key] || '').trim();
            } else if (cleanKey === '業務類別' || cleanKey === 'business') {
                business = String(row[key] || '').trim().toUpperCase();
            }
        }
        
        if (!name && !year && !start_date) {
            results.skipped++;
            continue;
        }
        
        if (!name || !start_date) {
            results.failed++;
            results.errors.push(`第 ${i + 2} 行：計畫名稱和開始日期為必填`);
            continue;
        }
        
        if (!year) {
            const adYear = parseInt(start_date.slice(0, 4), 10);
            year = String(adYear - 1911).padStart(3, '0');
        } else if (!/^\d{3}$/.test(year)) {
            results.failed++;
            results.errors.push(`第 ${i + 2} 行（${name}）：年度格式錯誤，應為3位數字（例如：113）`);
            continue;
        }
        
        if (!railway) railway = '-';
        if (!inspection_type) inspection_type = '-';
        if (!business) business = '-';
        
        try {
            const y = String(year).replace(/\D/g, '').slice(-3).padStart(3, '0');
            const r = railway === '-' ? '-' : String(railway).toUpperCase();
            const it = inspection_type === '-' ? '-' : String(inspection_type);
            const b = business === '-' ? '-' : String(business).toUpperCase();
            
            let inspection_seq = '00';
            let plan_number = '(匯入)';
            
            if (r !== '-' && it !== '-' && b !== '-') {
                const maxRes = await pool.query(
                    `SELECT COALESCE(MAX(CAST(inspection_seq AS INTEGER)), 0) AS mx 
                     FROM inspection_plan_schedule 
                     WHERE year = $1 AND railway = $2 AND inspection_type = $3 AND business = $4`,
                    [y, r, it, b]
                );
                const next = (parseInt(maxRes.rows[0]?.mx || 0, 10) + 1);
                inspection_seq = String(next).padStart(2, '0');
                plan_number = `${y}${r}${it}-${inspection_seq}-${b}`;
            }
            
            await pool.query(
                `INSERT INTO inspection_plan_schedule (start_date, end_date, plan_name, year, railway, inspection_type, business, inspection_seq, plan_number)
                 VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)`,
                [start_date, end_date || null, name, y, r, it, b, inspection_seq, plan_number]
            );
            results.success++;
        } catch (e) {
            results.failed++;
            results.errors.push(`第 ${i + 2} 行（${name}）：${e.message}`);
        }
    }
    
    // 完成資訊已在 logAction 中記錄
    
    if (results.success > 0) {
        logAction(req.session.user.username, 'IMPORT_PLANS', `匯入檢查計畫：成功 ${results.success} 筆，失敗 ${results.failed} 筆，跳過 ${results.skipped || 0} 筆`, req);
    }
    
    // 返回結果，使用 successCount 來避免與 success 布林值衝突
    res.json({ 
        success: true, 
        successCount: results.success,
        failed: results.failed, 
        errors: results.errors, 
        skipped: results.skipped 
    });
});

// --- 檢查計畫規劃（月曆排程）API ---
// 取號規則：年度(3碼)-鐵路機構-檢查類別-檢查次數-業務類別；檢查次數由 01 起自動加 1
const RAILWAY_CODES = { T: '臺鐵', H: '高鐵', A: '林鐵', S: '糖鐵' };
const INSPECTION_CODES = { '1': '年度定期檢查', '2': '特別檢查', '3': '例行性檢查', '4': '臨時檢查', '5': '調查' };
const BUSINESS_CODES = { OP: '運轉', CV: '土建', ME: '機務', EL: '電務', SM: '安全管理', AD: '營運', OT: '其他' };

app.get('/api/plan-schedule', requireAuth, async (req, res) => {
    const { year, month } = req.query;
    try {
        if (!year || !month) {
            return res.status(400).json({ error: '請提供 year 與 month 參數（西元年、月）' });
        }
        const y = parseInt(String(year), 10);
        const m = parseInt(String(month), 10);
        const start = `${y}-${String(m).padStart(2, '0')}-01`;
        const lastDay = new Date(y, m, 0).getDate();
        const end = `${y}-${String(m).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
        const rows = await pool.query(
            `SELECT id, start_date, end_date, plan_name, year, railway, inspection_type, business, inspection_seq, plan_number, created_at 
             FROM inspection_plan_schedule 
             WHERE (start_date <= $2::date AND (end_date IS NULL OR end_date >= $1::date))
             ORDER BY start_date ASC, id ASC`,
            [start, end]
        );
        res.json({ data: rows.rows || [] });
    } catch (e) {
        handleApiError(e, req, res, 'Get plan schedule error');
    }
});

app.get('/api/plan-schedule/next-number', requireAuth, async (req, res) => {
    const { year, railway, inspectionType, business } = req.query;
    try {
        if (!year || !railway || !inspectionType || !business) {
            return res.status(400).json({ error: '請提供 year, railway, inspectionType, business' });
        }
        const y = String(year).replace(/\D/g, '').slice(-3).padStart(3, '0');
        const r = String(railway).toUpperCase();
        const it = String(inspectionType);
        const b = String(business).toUpperCase();
        const maxRes = await pool.query(
            `SELECT COALESCE(MAX(CAST(inspection_seq AS INTEGER)), 0) AS mx 
             FROM inspection_plan_schedule 
             WHERE year = $1 AND railway = $2 AND inspection_type = $3 AND business = $4`,
            [y, r, it, b]
        );
        const next = (parseInt(maxRes.rows[0]?.mx || 0, 10) + 1);
        const seq = String(next).padStart(2, '0');
        const planNumber = `${y}${r}${it}-${seq}-${b}`;
        res.json({ nextSeq: seq, planNumber });
    } catch (e) {
        handleApiError(e, req, res, 'Get next plan number error');
    }
});

app.post('/api/plan-schedule', requireAuth, requireAdminOrManager, verifyCsrf, async (req, res) => {
    const { plan_name, start_date, end_date, year, railway, inspection_type, business } = req.body;
    try {
        if (!plan_name || !start_date || !year || !railway || !inspection_type || !business) {
            return res.status(400).json({ error: '計畫名稱、開始日期、年度、鐵路機構、檢查類別、業務類別為必填' });
        }
        const y = String(year).replace(/\D/g, '').slice(-3).padStart(3, '0');
        const r = String(railway).toUpperCase();
        const it = String(inspection_type);
        const b = String(business).toUpperCase();
        const maxRes = await pool.query(
            `SELECT COALESCE(MAX(CAST(inspection_seq AS INTEGER)), 0) AS mx 
             FROM inspection_plan_schedule 
             WHERE year = $1 AND railway = $2 AND inspection_type = $3 AND business = $4`,
            [y, r, it, b]
        );
        const next = (parseInt(maxRes.rows[0]?.mx || 0, 10) + 1);
        const seq = String(next).padStart(2, '0');
        const planNumber = `${y}${r}${it}-${seq}-${b}`;
        const name = String(plan_name).trim();
        await pool.query(
            `INSERT INTO inspection_plan_schedule (start_date, end_date, plan_name, year, railway, inspection_type, business, inspection_seq, plan_number) 
             VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)`,
            [start_date, end_date || null, name, y, r, it, b, seq, planNumber]
        );
        const dateRange = end_date ? `${start_date} ~ ${end_date}` : start_date;
        logAction(req.session.user.username, 'CREATE_PLAN_SCHEDULE', `新增檢查計畫規劃：${name}，取號 ${planNumber}，日期 ${dateRange}`, req);
        res.json({ success: true, planNumber, inspectionSeq: seq });
    } catch (e) {
        handleApiError(e, req, res, 'Create plan schedule error');
    }
});

app.get('/api/plan-schedule/all', requireAuth, requireAdminOrManager, async (req, res) => {
    try {
        const rows = await pool.query(
            `SELECT id, start_date, end_date, plan_name, year, railway, inspection_type, business, inspection_seq, plan_number, created_at, updated_at 
             FROM inspection_plan_schedule 
             ORDER BY year DESC, start_date ASC, id ASC`
        );
        res.json({ data: rows.rows || [] });
    } catch (e) {
        handleApiError(e, req, res, 'Get all plan schedules error');
    }
});

// 假日資料來源：GitHub ruyut/TaiwanCalendar（中華民國政府行政機關辦公日曆）
// https://github.com/ruyut/TaiwanCalendar | CDN: cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/{year}.json
app.get('/api/holidays/:year', requireAuth, async (req, res) => {
    try {
        const year = parseInt(req.params.year);
        if (!year || year < 2000 || year > 2100) {
            return res.status(400).json({ error: '無效的年份' });
        }
        
        const https = require('https');
        const url = `https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/${year}.json`;
        
        return new Promise((resolve) => {
            const request = https.get(url, { timeout: 8000 }, (response) => {
                if (response.statusCode !== 200) {
                    res.json({ data: [] });
                    return resolve();
                }
                let data = '';
                response.setEncoding('utf8');
                response.on('data', (chunk) => { data += chunk; });
                response.on('end', () => {
                    try {
                        const rawData = JSON.parse(data);
                        const arr = Array.isArray(rawData) ? rawData : [];
                        const holidays = arr.map(h => {
                            const d = String(h.date || '').trim();
                            const dateStr = d.match(/^\d{8}$/)
                                ? `${d.slice(0, 4)}-${d.slice(4, 6)}-${d.slice(6, 8)}`
                                : d;
                            return {
                                date: dateStr,
                                name: (h.description || '').trim() || '假日',
                                isHoliday: h.isHoliday === true
                            };
                        });
                        res.json({ data: holidays });
                        resolve();
                    } catch (e) {
                        res.json({ data: [] });
                        resolve();
                    }
                });
            });
            request.on('error', () => {
                res.json({ data: [] });
                resolve();
            });
            request.on('timeout', () => {
                request.destroy();
                res.json({ data: [] });
                resolve();
            });
        });
    } catch (e) {
        res.json({ data: [] });
    }
});

app.put('/api/plan-schedule/:id', requireAuth, requireAdminOrManager, verifyCsrf, async (req, res) => {
    const { plan_name, start_date, end_date, year, railway, inspection_type, business } = req.body;
    try {
        if (!plan_name || !start_date || !year || !railway || !inspection_type || !business) {
            return res.status(400).json({ error: '計畫名稱、開始日期、年度、鐵路機構、檢查類別、業務類別為必填' });
        }
        const r = await pool.query('SELECT * FROM inspection_plan_schedule WHERE id = $1', [req.params.id]);
        if (r.rows.length === 0) return res.status(404).json({ error: '找不到該筆排程' });
        
        const y = String(year).replace(/\D/g, '').slice(-3).padStart(3, '0');
        const rCode = String(railway).toUpperCase();
        const it = String(inspection_type);
        const b = String(business).toUpperCase();
        
        let inspection_seq = r.rows[0].inspection_seq;
        let plan_number = r.rows[0].plan_number;
        
        if (rCode !== r.rows[0].railway || it !== r.rows[0].inspection_type || b !== r.rows[0].business) {
            const maxRes = await pool.query(
                `SELECT COALESCE(MAX(CAST(inspection_seq AS INTEGER)), 0) AS mx 
                 FROM inspection_plan_schedule 
                 WHERE year = $1 AND railway = $2 AND inspection_type = $3 AND business = $4`,
                [y, rCode, it, b]
            );
            const next = (parseInt(maxRes.rows[0]?.mx || 0, 10) + 1);
            inspection_seq = String(next).padStart(2, '0');
            plan_number = `${y}${rCode}${it}-${inspection_seq}-${b}`;
        }
        
        await pool.query(
            `UPDATE inspection_plan_schedule 
             SET plan_name = $1, start_date = $2, end_date = $3, year = $4, railway = $5, 
                 inspection_type = $6, business = $7, inspection_seq = $8, plan_number = $9, updated_at = CURRENT_TIMESTAMP
             WHERE id = $10`,
            [plan_name.trim(), start_date, end_date || null, y, rCode, it, b, inspection_seq, plan_number, req.params.id]
        );
        
        const dateRange = end_date ? `${start_date} ~ ${end_date}` : start_date;
        logAction(req.session.user.username, 'UPDATE_PLAN_SCHEDULE', `更新檢查計畫規劃：${plan_name}，取號 ${plan_number}，日期 ${dateRange}`, req);
        res.json({ success: true, planNumber: plan_number });
    } catch (e) {
        handleApiError(e, req, res, 'Update plan schedule error');
    }
});

app.delete('/api/plan-schedule/:id', requireAuth, requireAdminOrManager, verifyCsrf, async (req, res) => {
    try {
        const r = await pool.query(
            'SELECT plan_name, plan_number FROM inspection_plan_schedule WHERE id = $1',
            [req.params.id]
        );
        if (r.rows.length === 0) return res.status(404).json({ error: '找不到該筆排程' });
        await pool.query('DELETE FROM inspection_plan_schedule WHERE id = $1', [req.params.id]);
        logAction(req.session.user.username, 'DELETE_PLAN_SCHEDULE', `刪除檢查計畫規劃：${r.rows[0].plan_name}（${r.rows[0].plan_number}）`, req);
        res.json({ success: true });
    } catch (e) {
        handleApiError(e, req, res, 'Delete plan schedule error');
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