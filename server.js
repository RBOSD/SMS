const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const bodyParser = require('body-parser');
const path = require('path');
const session = require('express-session');
const bcrypt = require('bcrypt');
// const { GoogleGenerativeAI } = require("@google/generative-ai"); // 若無使用 AI 可註解

const app = express();
const PORT = process.env.PORT || 3000;
const DB_PATH = path.join(__dirname, 'issues.db');

// Middleware
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

app.use(session({
    secret: 'sms-secret-key',
    resave: false,
    saveUninitialized: false,
    cookie: { maxAge: 24 * 60 * 60 * 1000 }
}));

// 初始化資料庫
const db = new sqlite3.Database(DB_PATH, (err) => {
    if (err) console.error('Error opening database:', err.message);
    else {
        console.log('Connected to SQLite database.');
        initDB();
    }
});

// 權限檢查
const requireAuth = (req, res, next) => {
    if (req.session && req.session.user) next();
    else res.status(401).json({ error: 'Unauthorized' });
};

function initDB() {
    db.serialize(() => {
        // 確保基本資料表存在
        db.run(`CREATE TABLE IF NOT EXISTS issues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number TEXT UNIQUE, year TEXT, unit TEXT, content TEXT, status TEXT,
            item_kind_code TEXT, division_name TEXT, inspection_category_name TEXT,
            category TEXT, handling TEXT, review TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE, password TEXT, name TEXT, role TEXT DEFAULT 'viewer',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        db.run(`CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT, action TEXT, details TEXT, ip_address TEXT,
            login_time DATETIME, created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        // [重要修正]：這裡執行「無痛遷移」，嘗試新增新功能需要的欄位
        // 如果欄位已存在，SQLite 會報錯，我們直接忽略錯誤即可，保證舊資料安全
        const newColumns = [
            'plan_name TEXT', 
            'issue_date TEXT'
        ];
        // 20回合的欄位
        for (let i = 1; i <= 20; i++) {
            if(i > 1) {
                newColumns.push(`handling${i} TEXT`);
                newColumns.push(`review${i} TEXT`);
            }
            newColumns.push(`reply_date_r${i} TEXT`);    // 機構回復日期
            newColumns.push(`response_date_r${i} TEXT`); // 函復日期
        }

        newColumns.forEach(colDef => {
            db.run(`ALTER TABLE issues ADD COLUMN ${colDef}`, (err) => {
                // 忽略 "duplicate column name" 錯誤，這代表欄位已經有了，不用擔心
            });
        });

        // 預設 Admin (只在完全沒人時才建立)
        db.get("SELECT count(*) as count FROM users", (err, row) => {
            if (!err && row.count === 0) {
                const hash = bcrypt.hashSync('admin123', 10);
                db.run("INSERT INTO users (username, password, name, role) VALUES (?, ?, ?, ?)", 
                    ['admin', hash, '系統管理員', 'admin']);
            }
        });
    });
}

function logAction(username, action, details, req) {
    const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    db.run("INSERT INTO logs (username, action, details, ip_address) VALUES (?, ?, ?, ?)", [username, action, details, ip]);
}

// --- API Routes ---

app.post('/api/auth/login', (req, res) => {
    const { username, password } = req.body;
    db.get("SELECT * FROM users WHERE username = ?", [username], (err, user) => {
        if (err || !user) return res.status(401).json({ error: 'Invalid credentials' });
        if (bcrypt.compareSync(password, user.password)) {
            req.session.user = { id: user.id, username: user.username, role: user.role, name: user.name };
            logAction(user.username, 'LOGIN', 'User logged in', req);
            res.json({ success: true, user: req.session.user });
        } else {
            res.status(401).json({ error: 'Invalid credentials' });
        }
    });
});

app.post('/api/auth/logout', (req, res) => {
    req.session.destroy();
    res.json({ success: true });
});

app.get('/api/auth/me', (req, res) => {
    if (req.session.user) res.json({ isLogin: true, ...req.session.user });
    else res.json({ isLogin: false });
});

app.put('/api/auth/profile', requireAuth, (req, res) => {
    const { name, password } = req.body;
    const id = req.session.user.id;
    let sql = "UPDATE users SET name = ? WHERE id = ?";
    let params = [name, id];
    if(password) {
        sql = "UPDATE users SET name = ?, password = ? WHERE id = ?";
        params = [name, bcrypt.hashSync(password, 10), id];
    }
    db.run(sql, params, (err) => {
        if(err) return res.status(500).json({ error: err.message });
        res.json({ success: true });
    });
});

app.get('/api/issues', requireAuth, (req, res) => {
    const { page = 1, pageSize = 20, q, year, unit, status, itemKindCode, division, inspectionCategory, sortField, sortDir } = req.query;
    const offset = (page - 1) * pageSize;
    let where = ["1=1"], params = [];

    if (q) {
        where.push("(number LIKE ? OR content LIKE ? OR handling LIKE ? OR review LIKE ? OR plan_name LIKE ?)");
        const l = `%${q}%`; params.push(l, l, l, l, l);
    }
    if (year) { where.push("year = ?"); params.push(year); }
    if (unit) { where.push("unit = ?"); params.push(unit); }
    if (status) { where.push("status = ?"); params.push(status); }
    if (itemKindCode) { where.push("item_kind_code = ?"); params.push(itemKindCode); }
    if (division) { where.push("division_name = ?"); params.push(division); }
    if (inspectionCategory) { where.push("inspection_category_name = ?"); params.push(inspectionCategory); }

    let order = "created_at DESC";
    if(sortField && ['year','number','unit','status','created_at'].includes(sortField)) {
        order = `${sortField} ${sortDir === 'asc' ? 'ASC' : 'DESC'}`;
    }

    db.get(`SELECT count(*) as total FROM issues WHERE ${where.join(" AND ")}`, params, (err, row) => {
        if (err) return res.status(500).json({ error: err.message });
        const total = row.total;
        db.all(`SELECT * FROM issues WHERE ${where.join(" AND ")} ORDER BY ${order} LIMIT ? OFFSET ?`, [...params, pageSize, offset], (err, rows) => {
            if (err) return res.status(500).json({ error: err.message });
            
            // 簡單統計
            db.all(`SELECT status, count(*) as count FROM issues GROUP BY status 
                    UNION ALL SELECT unit, count(*) FROM issues GROUP BY unit
                    UNION ALL SELECT year, count(*) FROM issues GROUP BY year`, [], (err, statsRows) => {
                // 這裡簡化回傳，前端會自己處理分類，或者如果前端依賴特定格式，維持基本的
                // 為求相容舊版邏輯，重新組裝 stats
                const stats = { status: [], unit: [], year: [] };
                // (前端其實比較依賴 specific key，但舊版 server.js 在這裡邏輯比較複雜)
                // 為了確保 100% 相容，我們做一個簡單的查詢分離
                const p1 = new Promise(r => db.all("SELECT status, count(*) as count FROM issues GROUP BY status", (e,d)=>r(d||[])));
                const p2 = new Promise(r => db.all("SELECT unit, count(*) as count FROM issues GROUP BY unit", (e,d)=>r(d||[])));
                const p3 = new Promise(r => db.all("SELECT year, count(*) as count FROM issues GROUP BY year", (e,d)=>r(d||[])));
                
                Promise.all([p1, p2, p3]).then(([s, u, y]) => {
                    res.json({
                        data: rows, total, page: parseInt(page), pages: Math.ceil(total/pageSize),
                        globalStats: { status: s, unit: u, year: y }
                    });
                });
            });
        });
    });
});

// 更新 (包含日期)
app.put('/api/issues/:id', requireAuth, (req, res) => {
    const { status, round, handling, review, replyDate, responseDate } = req.body;
    const id = req.params.id;
    const r = parseInt(round);
    
    const hField = r === 1 ? 'handling' : `handling${r}`;
    const rField = r === 1 ? 'review' : `review${r}`;
    // 新增日期欄位更新
    const replyField = `reply_date_r${r}`;
    const respField = `response_date_r${r}`;

    const sql = `UPDATE issues SET status=?, ${hField}=?, ${rField}=?, ${replyField}=?, ${respField}=?, updated_at=CURRENT_TIMESTAMP WHERE id=?`;
    
    db.run(sql, [status, handling, review, replyDate, responseDate, id], function(err) {
        if(err) return res.status(500).json({ error: err.message });
        logAction(req.session.user.username, 'UPDATE', `Updated issue ${id}`, req);
        res.json({ success: true });
    });
});

app.delete('/api/issues/:id', requireAuth, (req, res) => {
    if(!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    db.run("DELETE FROM issues WHERE id=?", [req.params.id], (err) => {
        if(err) return res.status(500).json({error:err.message});
        res.json({success:true});
    });
});

app.post('/api/issues/batch-delete', requireAuth, (req, res) => {
    if(!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    const { ids } = req.body;
    const ph = ids.map(()=>'?').join(',');
    db.run(`DELETE FROM issues WHERE id IN (${ph})`, ids, (err) => {
        if(err) return res.status(500).json({error:err.message});
        res.json({success:true});
    });
});

// 匯入 (包含新欄位)
app.post('/api/issues/import', requireAuth, (req, res) => {
    if(!['admin','manager'].includes(req.session.user.role)) return res.status(403).json({error:'Denied'});
    const { data, round, reviewDate, mode } = req.body; // data now has planName, issueDate
    const r = parseInt(round) || 1;

    db.serialize(() => {
        const check = db.prepare("SELECT id FROM issues WHERE number = ?");
        
        // 插入: 包含 plan_name, issue_date, 以及第一回合的 response_date
        const insert = db.prepare(`INSERT INTO issues (
            number, year, unit, content, status, item_kind_code, category, division_name, inspection_category_name,
            handling, review, plan_name, issue_date, response_date_r1
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);

        // 更新: 包含 plan_name
        const hCol = r===1?'handling':`handling${r}`;
        const rCol = r===1?'review':`review${r}`;
        const respCol = `response_date_r${r}`;
        
        const update = db.prepare(`UPDATE issues SET status=?, ${hCol}=?, ${rCol}=?, ${respCol}=?, plan_name=COALESCE(?, plan_name), updated_at=CURRENT_TIMESTAMP WHERE number=?`);

        let processed = 0;
        data.forEach(item => {
            check.get([item.number], (err, row) => {
                if(row) {
                    // update
                    update.run([
                        item.status, item.handling||'', item.review||'', 
                        reviewDate||'', // map reviewDate to this round's response date
                        item.planName, 
                        item.number
                    ]);
                } else {
                    // insert (只在第一次帶入 issueDate)
                    insert.run([
                        item.number, item.year, item.unit, item.content, item.status||'持續列管',
                        item.itemKindCode, item.category, item.divisionName, item.inspectionCategoryName,
                        item.handling||'', item.review||'', 
                        item.planName, item.issueDate,
                        reviewDate||'' // response_date_r1
                    ]);
                }
                processed++;
                if(processed === data.length) res.json({success:true, count:data.length});
            });
        });
        check.finalize();
        insert.finalize();
        update.finalize();
    });
});

// Users
app.get('/api/users', requireAuth, (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    const { page=1, pageSize=20, q, sortField='id', sortDir='asc' } = req.query;
    const offset = (page-1)*pageSize;
    let sql = "SELECT id, username, name, role, created_at FROM users WHERE 1=1";
    let params = [];
    if(q) { sql+=" AND (username LIKE ? OR name LIKE ?)"; params.push(`%${q}%`,`%${q}%`); }
    sql += ` ORDER BY ${sortField} ${sortDir==='desc'?'DESC':'ASC'} LIMIT ? OFFSET ?`;
    
    db.get("SELECT count(*) as t FROM users", (err,row)=>{
        db.all(sql, [...params, pageSize, offset], (err, rows)=>{
            res.json({data:rows, total:row.t, page, pages:Math.ceil(row.t/pageSize)});
        });
    });
});

app.post('/api/users', requireAuth, (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    const { username, password, name, role } = req.body;
    db.run("INSERT INTO users (username, password, name, role) VALUES (?,?,?,?)", 
        [username, bcrypt.hashSync(password, 10), name, role], (err)=>{
        if(err) return res.status(400).json({error:err.message});
        res.json({success:true});
    });
});

app.put('/api/users/:id', requireAuth, (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    const { name, password, role } = req.body;
    let sql = "UPDATE users SET name=?, role=? WHERE id=?";
    let p = [name, role, req.params.id];
    if(password) { sql="UPDATE users SET name=?, role=?, password=? WHERE id=?"; p=[name, role, bcrypt.hashSync(password,10), req.params.id]; }
    db.run(sql, p, (err)=>{
        if(err) return res.status(500).json({error:err.message});
        res.json({success:true});
    });
});

app.delete('/api/users/:id', requireAuth, (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    if(parseInt(req.params.id) === req.session.user.id) return res.status(400).json({error:'Cannot self delete'});
    db.run("DELETE FROM users WHERE id=?", [req.params.id], (err)=>{
        res.json({success:true});
    });
});

app.get('/api/admin/logs', requireAuth, (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    const { page=1, pageSize=20, q } = req.query;
    const offset = (page-1)*pageSize;
    let sql = "SELECT * FROM logs WHERE action='LOGIN'";
    let p = [];
    if(q) { sql+=" AND (username LIKE ? OR ip_address LIKE ?)"; p.push(`%${q}%`,`%${q}%`); }
    sql += " ORDER BY login_time DESC LIMIT ? OFFSET ?";
    db.all(sql, [...p, pageSize, offset], (err, rows) => res.json({data:rows, total:0, page, pages:1})); // 簡化total
});

app.get('/api/admin/action_logs', requireAuth, (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    const { page=1, pageSize=20, q } = req.query;
    const offset = (page-1)*pageSize;
    let sql = "SELECT * FROM logs WHERE action!='LOGIN'";
    let p = [];
    if(q) { sql+=" AND (username LIKE ? OR action LIKE ?)"; p.push(`%${q}%`,`%${q}%`); }
    sql += " ORDER BY created_at DESC LIMIT ? OFFSET ?";
    db.all(sql, [...p, pageSize, offset], (err, rows) => res.json({data:rows, total:0, page, pages:1}));
});

app.delete('/api/admin/logs', requireAuth, (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    db.run("DELETE FROM logs WHERE action='LOGIN'", ()=>res.json({success:true}));
});

app.delete('/api/admin/action_logs', requireAuth, (req, res) => {
    if(req.session.user.role!=='admin') return res.status(403).json({error:'Denied'});
    db.run("DELETE FROM logs WHERE action!='LOGIN'", ()=>res.json({success:true}));
});

app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});