        // 全域狀態
        let rawData = [], currentData = [], currentUser = null, charts = {}, currentEditItem = null, userList = [], sortState = { field: null, dir: 'asc' }, stagedImportData = [];
        let autoLogoutTimer;
        let currentLogs = { login: [], action: [] };
        let cachedGlobalStats = null;
        let csrfToken = null; // CSRF token 快取
        
        // 開發模式檢測（用於條件輸出 console.warn）
        const isDevelopment = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1' || window.location.hostname.includes('dev');
        
        // 取得 CSRF token
        async function getCsrfToken() {
            if (csrfToken) return csrfToken;
            try {
                const res = await fetch('/api/csrf-token', {
                    credentials: 'include'
                });
                if (res.ok) {
                    const data = await res.json();
                    csrfToken = data.csrfToken;
                    return csrfToken;
                }
            } catch (e) {
                console.error('Failed to get CSRF token:', e);
            }
            return null;
        }
        
        // 密碼複雜度驗證函數（前端）
        function validatePasswordFrontend(password) {
            if (!password || password.length < 8) {
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
        
        // 統一的 API 請求包裝函數，自動處理認證錯誤和 CSRF token
        // 生成類別標籤 HTML
        function getKindLabel(kindCode) {
            if (!kindCode) return '';
            const labels = {
                'N': '<span class="kind-tag N">缺失</span>',
                'O': '<span class="kind-tag O">觀察</span>',
                'R': '<span class="kind-tag R">建議</span>'
            };
            return labels[kindCode] || '';
        }

        // 生成狀態標籤 HTML
        function getStatusBadge(status) {
            if (!status || status === 'Open') return '';
            const statusClass = status === '持續列管' ? 'active' : (status === '解除列管' ? 'resolved' : 'self');
            return `<span class="badge ${statusClass}">${status}</span>`;
        }

        // 驗證日期格式（6或7位數字，例如：1130601 或 1141001）
        function validateDateFormat(dateStr, fieldName = '日期') {
            if (!dateStr || !/^\d{6,7}$/.test(dateStr)) {
                showToast(`${fieldName}格式錯誤，應為6或7位數字（例如：1130601 或 1141001）`, 'error');
                return false;
            }
            return true;
        }

        async function apiFetch(url, options = {}) {
            try {
                // 對於需要 CSRF 保護的請求（POST, PUT, DELETE），自動加入 token
                const needsCsrf = ['POST', 'PUT', 'DELETE', 'PATCH'].includes(options.method);
                if (needsCsrf) {
                    const token = await getCsrfToken();
                    if (!token) {
                        console.error('Failed to get CSRF token');
                        throw new Error('無法取得 CSRF token，請重新整理頁面');
                    }
                    options.headers = options.headers || {};
                    options.headers['X-CSRF-Token'] = token;
                }
                
                const response = await fetch(url, {
                    ...options,
                    credentials: 'include', // 確保包含 cookies
                    headers: {
                        'Content-Type': 'application/json',
                        ...options.headers
                    }
                });
                
                // 處理認證錯誤
                if (response.status === 401 || response.status === 403) {
                    // 如果是 CSRF token 錯誤，清除快取並重試一次
                    if (response.status === 403 && needsCsrf) {
                        const errorData = await response.json().catch(() => ({}));
                        if (errorData.error && errorData.error.includes('CSRF')) {
                            csrfToken = null; // 清除快取的 token
                            const newToken = await getCsrfToken();
                            if (newToken) {
                                options.headers['X-CSRF-Token'] = newToken;
                                const retryResponse = await fetch(url, {
                                    ...options,
                                    credentials: 'include',
                                    headers: {
                                        'Content-Type': 'application/json',
                                        ...options.headers
                                    }
                                });
                                if (retryResponse.ok || retryResponse.status !== 403) {
                                    return retryResponse;
                                }
                            }
                        }
                    }
                    if (isDevelopment) console.warn('認證失敗，重定向到登入頁');
                    // 清除 sessionStorage
                    sessionStorage.clear();
                    // 重定向到登入頁
                    window.location.href = '/login.html';
                    throw new Error('Unauthorized');
                }
                
                return response;
            } catch (error) {
                // 如果是我們自己拋出的認證錯誤，直接重新拋出
                if (error.message === 'Unauthorized') {
                    throw error;
                }
                // 其他錯誤正常處理
                throw error;
            }
        }
        
        // 日誌記錄函數（寫入檔案，不在控制台顯示）
        async function writeLog(message, level = 'INFO') {
            try {
                await apiFetch('/api/log', {
                    method: 'POST',
                    body: JSON.stringify({ message, level })
                }).catch(() => {}); // 靜默失敗，不影響主流程
            } catch (e) {
                // 靜默處理錯誤
            }
        }
        let issuesPage = 1, issuesPageSize = 20, issuesTotal = 0, issuesPages = 1;
        let usersPage = 1, usersPageSize = 20, usersTotal = 0, usersPages = 1, usersSortField = 'id', usersSortDir = 'asc';
        let plansPage = 1, plansPageSize = 20, plansTotal = 0, plansPages = 1, plansSortField = 'year', plansSortDir = 'desc';
        let planList = [];
        // 目前在計畫管理編輯視窗中正在查看的該計畫所有排程
        let currentPlanSchedules = [];
        let logsPage = 1, logsPageSize = 20, logsTotal = 0, logsPages = 1;
        let actionsPage = 1, actionsPageSize = 20, actionsTotal = 0, actionsPages = 1;
        // Current import mode: 'word' (uses param) or 'backup' (ignores param)
        let currentImportMode = 'word';

        function resetAutoLogout() { clearTimeout(autoLogoutTimer); autoLogoutTimer = setTimeout(() => { showToast("您已閒置過久，系統將自動登出。", 'warning'); setTimeout(() => logout(), 2000); }, 1800000); }
        window.onload = resetAutoLogout; document.onmousemove = resetAutoLogout; document.onkeypress = resetAutoLogout;

        function toggleDashboard(btn) { 
            const d = document.getElementById('dashboardSection'); 
            const c = d.classList.contains('collapsed'); 
            d.classList.toggle('collapsed', !c); 
            const icon = btn.querySelector('.toggle-icon');
            if (icon) {
                icon.textContent = c ? '▲' : '▼';
            }
            btn.title = c ? '收合統計圖表' : '展開統計圖表';
        }
        function toggleUserMenu() { document.getElementById('userDropdown').classList.toggle('show'); }
        window.addEventListener('click', function (e) { if (!e.target.closest('.user-menu-container')) { document.getElementById('userDropdown').classList.remove('show'); } });

        function togglePwdVisibility(inputId, btn) { const input = document.getElementById(inputId); if (input.type === 'password') { input.type = 'text'; btn.innerText = '🚫'; } else { input.type = 'password'; btn.innerText = '👁️'; } }

        // [New] Toggle Advanced Filters
        function toggleAdvancedFilters(btn) {
            const panel = document.getElementById('advancedFilters');
            const isShown = panel.classList.contains('show');
            if (isShown) {
                panel.classList.remove('show');
                btn.innerText = '⬇️ 顯示更多篩選條件';
            } else {
                panel.classList.add('show');
                btn.innerText = '⬆️ 收合篩選條件';
            }
        }

        // --- Helper functions (Safe Versions) ---
        function normalizeCodeString(str) {
            if (!str) return "";
            var s = String(str);
            s = (s.normalize ? s.normalize("NFKC") : s);
            s = s.replace(/[\u200B-\u200D\uFEFF]/g, "");
            s = s.replace(/[\u2010-\u2015\u2212\uFE63\uFF0D]/g, "-");
            s = s.replace(/[ \t]+/g, " ").replace(/\s*-\s*/g, "-");
            return s.trim();
        }
        function stripHtml(h) {
            if (!h) return '';
            let t = document.createElement("DIV");
            t.innerHTML = String(h);
            return t.textContent || t.innerText || "";
        }
        function getLatest(i, p) { 
            // 支持無限次，動態查找（從200開始向下找，實際應該不會超過這個數字）
            for (let k = 200; k >= 1; k--) { 
                const key = k === 1 ? p : `${p}${k}`; 
                if (i[key]) return i[key]; 
            } 
            return null; 
        }
        
        // [Added] 獲取最新的審查或辦理情形（比較輪次）
        function getLatestReviewOrHandling(item) {
            let latestReviewRound = 0;
            let latestHandlingRound = 0;
            let latestReview = null;
            let latestHandling = null;
            
            // 查找最新的審查意見
            for (let k = 200; k >= 1; k--) {
                const key = k === 1 ? 'review' : `review${k}`;
                if (item[key] && item[key].trim()) {
                    latestReviewRound = k;
                    latestReview = item[key];
                    break;
                }
            }
            
            // 查找最新的辦理情形
            for (let k = 200; k >= 1; k--) {
                const key = k === 1 ? 'handling' : `handling${k}`;
                if (item[key] && item[key].trim()) {
                    latestHandlingRound = k;
                    latestHandling = item[key];
                    break;
                }
            }
            
            // 比較輪次，選擇輪次更高的
            if (latestReviewRound > latestHandlingRound) {
                return { type: 'review', content: latestReview, round: latestReviewRound };
            } else if (latestHandlingRound > latestReviewRound) {
                return { type: 'handling', content: latestHandling, round: latestHandlingRound };
            } else if (latestReviewRound > 0 && latestReviewRound === latestHandlingRound) {
                // 輪次相同，優先顯示審查（因為審查在辦理之後）
                return { type: 'review', content: latestReview, round: latestReviewRound };
            } else if (latestReview) {
                return { type: 'review', content: latestReview, round: latestReviewRound };
            } else if (latestHandling) {
                return { type: 'handling', content: latestHandling, round: latestHandlingRound };
            }
            
            return null;
        }
        function getRoleName(r) { const map = { 'admin': '系統管理員', 'manager': '資料管理者', 'editor': '審查人員', 'viewer': '檢視人員' }; return map[r] || r; }
        // [Enhanced] 改進編號提取，支持從帶換行的儲存格中提取編號
        function extractNumberFromCell(cell) {
            if (!cell) return "";
            var whole = normalizeCodeString(cell.innerText || cell.textContent || "");
            
            // 1. 先嘗試直接提取 TRC-v2 格式 (123-TRC-1-7-OP-N12)
            var mB = whole.match(/(\d{3}-[A-Za-z]{3}-[1-4]-\d+-[A-Za-z]{2,3}-[NORnor]\d{1,3})/);
            if (mB) return (mB[1] || "").toUpperCase();
            
            // 2. 嘗試 THAS-v2 格式（新提案）有分隔符 (113T1-01-OP-N01)
            var mC = whole.match(/(\d{3}[THASthas][1-5]-\d{2}-[A-Za-z]{2,3}-[NORnor]\d{2})/);
            if (mC) return (mC[1] || "").toUpperCase();
            
            // 3. 嘗試 THAS-v2 格式（新提案）無分隔符 (113T101OPN01)
            var mC2 = whole.match(/(\d{3}[THASthas][1-5]\d{2}[A-Za-z]{2,3}[NORnor]\d{2})/);
            if (mC2) return (mC2[1] || "").toUpperCase();
            
            // 4. 嘗試 THAS-v1 格式 (13T1-A01-N01)
            var mA = whole.match(/(\d{2}[THASthas][1-4]-[A-Ga-g]\d{2}-[NORnor]\d{2})/);
            if (mA) return (mA[1] || "").toUpperCase();
            
            // 5. 處理帶 <br> 的情況，分行匹配
            var rawHtml = cell.innerHTML || "";
            var lines = normalizeCodeString(rawHtml.replace(/<br\s*\/?>/gi, "\n").replace(/<\/p>/gi, "\n").replace(/<[^>]*>/g, "")).split("\n");
            for (var i = 0; i < lines.length; i++) {
                var line = (lines[i] || "").trim();
                if (!line) continue;
                var m1 = line.match(/(\d{3}-[A-Za-z]{3}-[1-4]-\d+-[A-Za-z]{2,3}-[NORnor]\d{1,3})/);
                if (m1) return (m1[1] || "").toUpperCase();
                var m2 = line.match(/(\d{3}[THASthas][1-5]-\d{2}-[A-Za-z]{2,3}-[NORnor]\d{2})/);
                if (m2) return (m2[1] || "").toUpperCase();
                var m3 = line.match(/(\d{3}[THASthas][1-5]\d{2}[A-Za-z]{2,3}[NORnor]\d{2})/);
                if (m3) return (m3[1] || "").toUpperCase();
                var m4 = line.match(/(\d{2}[THASthas][1-4]-[A-Ga-g]\d{2}-[NORnor]\d{2})/);
                if (m4) return (m4[1] || "").toUpperCase();
            }
            
            return whole.trim();
        }

        // [Updated] Map & Parser
        const ORG_MAP = { "T": "臺鐵", "H": "高鐵", "A": "林鐵", "S": "糖鐵", "TRC": "臺鐵", "HSR": "高鐵", "AFR": "林鐵", "TSC": "糖鐵" };
        // [Added] 機構交叉映射表（THAS-v1 ↔ TRC-v2）
        const ORG_CROSSWALK = { "T": "TRC", "H": "HSR", "A": "AFR", "S": "TSC", "TRC": "TRC", "HSR": "HSR", "AFR": "AFR", "TSC": "TSC" };
        const INSPECTION_MAP = { "1": "定期檢查", "2": "例行性檢查", "3": "特別檢查", "4": "臨時檢查", "5": "調查" };
        // [Verified] Division Map includes all requested codes
        const DIVISION_MAP = { 
            "A": "運務", "B": "工務", "C": "機務", "D": "電務", "E": "安全", "F": "審核", "G": "災防", 
            "OP": "運轉", "CV": "土建", "ME": "機務", "EL": "電務", "SM": "安全管理", "AD": "營運", "OT": "其他",
            "CP": "土木", "EM": "機電" 
        };
        const KIND_MAP = { "N": "缺失事項", "O": "觀察事項", "R": "建議事項" };
        const FILLED_MARKS = ["■", "☑", "☒", "✔", "✅", "●", "◉", "✓"]; var EMPTY_MARKS = ["□", "☐", "◻", "○", "◯", "◇", "△"];

        // [Enhanced] 改進編號解析，支持 scheme 和 period 字段
        function parseItemNumber(numberStr) {
            var raw = normalizeCodeString(numberStr || "");
            if (!raw) return null;
            
            // 1. THAS-v1 格式：13T1-A01-N01 (2位年+T+类别-部门+序号-类型+序号)
            var m = raw.match(/^(\d{2})([THAS])([1-4])\-([A-G])(\d{2})\-([NOR])(\d{2})$/i);
            if (m) {
                var yy = parseInt(m[1], 10);
                var rocYear = 100 + yy;
                var orgCode = m[2].toUpperCase();
                var itemSeq = m[7];
                var divisionSeq = m[5];
                return {
                    scheme: "THAS-v1",
                    raw: raw,
                    yearRoc: rocYear,
                    orgCode: orgCode,
                    orgCodeRaw: orgCode,
                    inspectCode: m[3],
                    divCode: m[4].toUpperCase(),
                    divisionCode: m[4].toUpperCase(),
                    divisionSeq: divisionSeq,
                    kindCode: m[6].toUpperCase(),
                    itemSeq: itemSeq,
                    period: ""
                };
            }
            
            // 2. TRC-v2 格式：123-TRC-1-7-OP-N12 (3位年-机构-类别-期数-部门-类型序号)
            m = raw.match(/^(\d{3})-([A-Z]{3})-([1-4])-(\d+)-([A-Z]{2,3})-([NOR])(\d{1,3})$/i);
            if (m) {
                var rocYear2 = parseInt(m[1], 10);
                var orgCode2 = m[2].toUpperCase();
                var period = m[4];
                var itemSeq2 = m[7];
                return {
                    scheme: "TRC-v2",
                    raw: raw,
                    yearRoc: rocYear2,
                    orgCode: orgCode2,
                    orgCodeRaw: orgCode2,
                    inspectCode: m[3],
                    divCode: m[5].toUpperCase(),
                    divisionCode: m[5].toUpperCase(),
                    divisionSeq: "",
                    kindCode: m[6].toUpperCase(),
                    itemSeq: itemSeq2,
                    period: period
                };
            }
            
            // 3. THAS-v2 格式（新提案）：113T1-01-OP-N01 (3位年+机构+类别-检查次数-业务类别-类型+流水号)
            // 支援有分隔符和無分隔符兩種格式
            m = raw.match(/^(\d{3})([THAS])([1-5])\-(\d{2})\-([A-Z]{2,3})\-([NOR])(\d{2})$/i);
            if (m) {
                var rocYear3 = parseInt(m[1], 10);
                var orgCode3 = m[2].toUpperCase();
                var period3 = m[4];
                var itemSeq3 = m[7];
                return {
                    scheme: "THAS-v2",
                    raw: raw,
                    yearRoc: rocYear3,
                    orgCode: orgCode3,
                    orgCodeRaw: orgCode3,
                    inspectCode: m[3],
                    divCode: m[5].toUpperCase(),
                    divisionCode: m[5].toUpperCase(),
                    divisionSeq: "",
                    kindCode: m[6].toUpperCase(),
                    itemSeq: itemSeq3,
                    period: period3
                };
            }
            
            // 3b. THAS-v2 無分隔符格式：113T101OPN01
            m = raw.match(/^(\d{3})([THAS])([1-5])(\d{2})([A-Z]{2,3})([NOR])(\d{2})$/i);
            if (m) {
                var rocYear3b = parseInt(m[1], 10);
                var orgCode3b = m[2].toUpperCase();
                var period3b = m[4];
                var itemSeq3b = m[7];
                return {
                    scheme: "THAS-v2",
                    raw: raw,
                    yearRoc: rocYear3b,
                    orgCode: orgCode3b,
                    orgCodeRaw: orgCode3b,
                    inspectCode: m[3],
                    divCode: m[5].toUpperCase(),
                    divisionCode: m[5].toUpperCase(),
                    divisionSeq: "",
                    kindCode: m[6].toUpperCase(),
                    itemSeq: itemSeq3b,
                    period: period3b
                };
            }
            
            // 3c. THAS-v2 寬鬆格式（支援缺少檢查發現分類的情況）：115T2-01-OP-001
            // 這種格式最後是數字而不是 N/O/R+數字，我們嘗試解析但 kindCode 會是空
            m = raw.match(/^(\d{3})([THAS])([1-5])\-(\d{2})\-([A-Z]{2,3})\-(\d{2,3})$/i);
            if (m) {
                var rocYear3c = parseInt(m[1], 10);
                var orgCode3c = m[2].toUpperCase();
                var period3c = m[4];
                var itemSeq3c = m[6];
                return {
                    scheme: "THAS-v2",
                    raw: raw,
                    yearRoc: rocYear3c,
                    orgCode: orgCode3c,
                    orgCodeRaw: orgCode3c,
                    inspectCode: m[3],
                    divCode: m[5].toUpperCase(),
                    divisionCode: m[5].toUpperCase(),
                    divisionSeq: "",
                    kindCode: "", // 無法確定檢查發現分類
                    itemSeq: itemSeq3c,
                    period: period3c
                };
            }
            
            // 3d. THAS-v2 寬鬆格式無分隔符：115T201OP001
            m = raw.match(/^(\d{3})([THAS])([1-5])(\d{2})([A-Z]{2,3})(\d{2,3})$/i);
            if (m) {
                var rocYear3d = parseInt(m[1], 10);
                var orgCode3d = m[2].toUpperCase();
                var period3d = m[4];
                var itemSeq3d = m[6];
                return {
                    scheme: "THAS-v2",
                    raw: raw,
                    yearRoc: rocYear3d,
                    orgCode: orgCode3d,
                    orgCodeRaw: orgCode3d,
                    inspectCode: m[3],
                    divCode: m[5].toUpperCase(),
                    divisionCode: m[5].toUpperCase(),
                    divisionSeq: "",
                    kindCode: "", // 無法確定檢查發現分類
                    itemSeq: itemSeq3d,
                    period: period3d
                };
            }
            
            // 4. 長格式（兼容舊格式）：123-TRC-1-7-OP-N12 (支持 3-4 位機構代碼)
            var cleanRaw = raw.replace(/[^a-zA-Z0-9\-]/g, "");
            var mLong = cleanRaw.match(/^(\d{3})-([A-Z]{3,4})-([0-9])-(\d+)-([A-Z]{2,4})-([NOR])(\d+)$/i);
            if (mLong) {
                return {
                    scheme: "TRC-v2",
                    raw: mLong[0],
                    yearRoc: parseInt(mLong[1], 10),
                    orgCode: mLong[2].toUpperCase(),
                    orgCodeRaw: mLong[2].toUpperCase(),
                    inspectCode: mLong[3],
                    divCode: mLong[5].toUpperCase(),
                    divisionCode: mLong[5].toUpperCase(),
                    divisionSeq: "",
                    kindCode: mLong[6].toUpperCase(),
                    itemSeq: mLong[7],
                    period: mLong[4]
                };
            }
            
            // 5. 短格式（兼容舊格式）：13T1-A01-N01 (支持 2-3 位年份)
            var mShort = cleanRaw.match(/^(\d{2,3})([A-Z])([0-9])-([A-Z])(\d{2})-([NOR])(\d{2})$/i);
            if (mShort) {
                var yy = parseInt(mShort[1], 10);
                var rocYear = (yy < 1000) ? (yy + (yy < 100 ? 100 : 0)) : (yy - 1911);
                return {
                    scheme: "THAS-v1",
                    raw: mShort[0],
                    yearRoc: rocYear,
                    orgCode: mShort[2].toUpperCase(),
                    orgCodeRaw: mShort[2].toUpperCase(),
                    inspectCode: mShort[3],
                    divCode: mShort[4].toUpperCase(),
                    divisionCode: mShort[4].toUpperCase(),
                    divisionSeq: mShort[5],
                    kindCode: mShort[6].toUpperCase(),
                    itemSeq: mShort[7],
                    period: ""
                };
            }
            
            // 6. 寬鬆匹配（fallback）
            var mLoose = cleanRaw.match(/(\d{2,3}).*([NOR])\d+/i);
            if (mLoose) {
                return {
                    scheme: "",
                    raw: mLoose[0],
                    yearRoc: parseInt(mLoose[1], 10),
                    orgCode: "?",
                    orgCodeRaw: "?",
                    inspectCode: "?",
                    divCode: "?",
                    divisionCode: "?",
                    divisionSeq: "",
                    kindCode: mLoose[2].toUpperCase(),
                    itemSeq: "",
                    period: ""
                };
            }
            
            return {
                scheme: "",
                raw: cleanRaw,
                yearRoc: "",
                orgCode: "",
                orgCodeRaw: "",
                inspectCode: "",
                divCode: "",
                divisionCode: "",
                divisionSeq: "",
                kindCode: "",
                itemSeq: "",
                period: ""
            };
        }

        function normalizeMultiline(s) { s = String(s || ""); return s.replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/\u00A0/g, " ").replace(/\u3000/g, " ").replace(/\r/g, "").replace(/[ \t]+\n/g, "\n").replace(/\n[ \t]+/g, "\n").trim(); }
        
        // [Added] 編號規範化函數（參考轉換工具）
        function canonicalNumber(info) {
            if (!info) return "";
            if (info.scheme === "TRC-v2") {
                // [修正] 保留原始序號，不要去掉前導零
                var seq = info.itemSeq || "0";
                return (info.yearRoc + "-" + info.orgCodeRaw + "-" + 
                        info.inspectCode + "-" + (info.period || "") + "-" + 
                        info.divisionCode + "-" + info.kindCode + seq).toUpperCase();
            }
            if (info.scheme === "THAS-v2") {
                // THAS-v2 格式：113T1-01-OP-N01 (3位年+机构+类别-检查次数-业务类别-类型+流水号)
                var period = String(info.period || "0");
                period = ("0" + period).slice(-2); // 確保2碼
                var seq = String(info.itemSeq || "0");
                seq = ("0" + seq).slice(-2); // 確保2碼
                return (info.yearRoc + info.orgCodeRaw + info.inspectCode + "-" + 
                        period + "-" + info.divisionCode + "-" + 
                        info.kindCode + seq).toUpperCase();
            }
            if (info.scheme === "THAS-v1") {
                var yy = String(info.yearRoc - 100);
                yy = ("0" + yy).slice(-2);
                var seq2 = String(parseInt(info.itemSeq || "0", 10));
                seq2 = ("0" + seq2).slice(-2);
                return (yy + info.orgCodeRaw + info.inspectCode + "-" + 
                        info.divisionCode + (info.divisionSeq || "") + "-" + 
                        info.kindCode + seq2).toUpperCase();
            }
            return (info.raw || "").toUpperCase();
        }

        // [修正與增強] 內容清理與切割：只抓取最新的回覆內容
        function sanitizeContent(html) {
            if (!html) return "";
            var s = String(html);

            // 1. 先做基礎清理，移除多餘的樣式標籤，保留換行結構
            s = s.replace(/<\s*br\s*\/?>/gi, "\n")
                .replace(/<\s*\/p\s*>/gi, "\n")
                .replace(/<\s*p[^>]*>/gi, "")
                .replace(/<[^>]+>/g, ""); // 移除剩餘所有 HTML 標籤

            // 2. 正規化空白與特殊字元
            s = s.replace(/&nbsp;/g, " ")
                .replace(/[\u200B-\u200D\uFEFF]/g, "")
                .trim();

            // 3. [關鍵邏輯] 智慧切割：抓取「最上面」的內容
            var lines = s.split('\n');
            var resultLines = [];
            var hasContent = false;

            for (var i = 0; i < lines.length; i++) {
                var line = lines[i].trim();

                // 略過開頭的空行
                if (!hasContent && line.length === 0) continue;

                // 遇到常見的分隔線符號，視為舊資料開始，直接結束截取
                if (/^[-=_]{3,}/.test(line)) {
                    break;
                }

                // 遇到明顯的「日期標籤」且不是第一行時，視為舊資料的開始
                if (hasContent && /^(\d{2,3}[./-]\d{1,2}[./-]\d{1,2})/.test(line)) {
                    break;
                }

                // 遇到「前次」、「上次」關鍵字開頭，視為舊資料
                if (hasContent && /^(前次|上次|第\d+次)(辦理|審查|回復|說明)/.test(line)) {
                    break;
                }

                // 加入有效行
                resultLines.push(line);
                if (line.length > 0) hasContent = true;
            }

            return resultLines.join("\n").trim();
        }

        function parseStatusFromResultCell(cell) { if (!cell) return ""; var src = normalizeMultiline((cell.innerText || cell.textContent || "") + "\n" + (cell.innerHTML || "").replace(/<[^>]+>/g, "")); if (!src) return ""; var allMarks = FILLED_MARKS.concat(EMPTY_MARKS).join(""); allMarks = allMarks.replace(/[-\\^$*+?.()|[\]{}]/g, "\\$&"); var reFront = new RegExp("([" + allMarks + "])\\s*(?:[:：﹕-]?\\s*)?(解除列管|持續列管|自行列管)", "g"); var reBack = new RegExp("(解除列管|持續列管|自行列管)\\s*(?:[:：﹕-]?\\s*)?([" + allMarks + "])", "g"); var hits = [], m; while ((m = reFront.exec(src)) !== null) { hits.push({ idx: m.index, label: m[2], mark: m[1], filled: FILLED_MARKS.indexOf(m[1]) >= 0 }); } while ((m = reBack.exec(src)) !== null) { hits.push({ idx: m.index, label: m[1], mark: m[2], filled: FILLED_MARKS.indexOf(m[2]) >= 0 }); } var filled = hits.filter(function (h) { return h.filled; }).sort(function (a, b) { return a.idx - b.idx; }); if (filled.length) return filled[filled.length - 1].label; var labels = ["解除列管", "持續列管", "自行列管"]; var present = labels.filter(function (l) { return src.indexOf(l) >= 0; }); if (present.length === 1) return present[0]; return ""; }
        function formatHtmlToText(html) { if (!html) return ""; let temp = String(html).replace(/<li[^>]*>/gi, "\n• ").replace(/<\/li>/gi, "").replace(/<ul[^>]*>/gi, "").replace(/<\/ul>/gi, "").replace(/<ol[^>]*>/gi, "").replace(/<\/ol>/gi, "").replace(/<br\s*\/?>/gi, "\n").replace(/<\/p>/gi, "\n").replace(/<p[^>]*>/gi, ""); let div = document.createElement("div"); div.innerHTML = temp; return (div.textContent || div.innerText || "").replace(/\n\s*\n/g, "\n").trim(); }

        function showToast(message, type = 'success') {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            let icon, title;
            if (type === 'success') {
                icon = '✅';
                title = '成功';
            } else if (type === 'warning') {
                icon = '⚠️';
                title = '警告';
            } else if (type === 'info') {
                icon = 'ℹ️';
                title = '資訊';
            } else {
                icon = '❌';
                title = '錯誤';
            }
            toast.className = `toast ${type}`;
            toast.innerHTML = `<div class="toast-icon">${icon}</div><div class="toast-content"><div class="toast-title">${title}</div><div class="toast-msg">${message}</div></div>`;
            container.appendChild(toast);
            requestAnimationFrame(() => { toast.classList.add('show'); });
            setTimeout(() => { toast.classList.remove('show'); toast.addEventListener('transitionend', () => toast.remove()); }, 3000);
        }

        function showPreview(html, title) { document.getElementById('previewTitle').innerText = title || '內容預覽'; document.getElementById('previewContent').innerHTML = html || '(無內容)'; document.getElementById('previewModal').classList.add('open'); }
        function closePreview() { document.getElementById('previewModal').classList.remove('open'); }
        
        // 自訂確認對話框（Promise 版本）
        let confirmModalResolve = null;
        let confirmModalHandler = null;
        
        function showConfirmModal(message, confirmText = '確認', cancelText = '取消') {
            return new Promise((resolve) => {
                const modal = document.getElementById('confirmModal');
                const messageEl = document.getElementById('confirmModalMessage');
                const confirmBtn = document.getElementById('confirmModalConfirmBtn');
                
                if (!modal || !messageEl || !confirmBtn) {
                    // 如果 modal 不存在，回退到原生 confirm
                    resolve(confirm(message));
                    return;
                }
                
                // 清除舊的事件處理器
                if (confirmModalHandler) {
                    confirmBtn.removeEventListener('click', confirmModalHandler);
                }
                
                // 重置狀態
                confirmModalResolve = resolve;
                
                messageEl.textContent = message;
                confirmBtn.textContent = confirmText;
                
                // 設置新的確認按鈕點擊事件
                confirmModalHandler = function handleConfirm() {
                    modal.style.display = 'none';
                    if (confirmModalResolve) {
                        confirmModalResolve(true);
                        confirmModalResolve = null;
                    }
                };
                confirmBtn.addEventListener('click', confirmModalHandler);
                
                modal.style.display = 'flex';
            });
        }
        
        function closeConfirmModal() {
            const modal = document.getElementById('confirmModal');
            if (modal) {
                modal.style.display = 'none';
                if (confirmModalResolve) {
                    // 取消時 resolve(false)
                    confirmModalResolve(false);
                    confirmModalResolve = null;
                }
            }
        }
        
        // 點擊 modal 背景關閉
        document.addEventListener('DOMContentLoaded', () => {
            const confirmModal = document.getElementById('confirmModal');
            if (confirmModal) {
                confirmModal.addEventListener('click', (e) => {
                    if (e.target === confirmModal) {
                        closeConfirmModal();
                    }
                });
            }
        });

        // 載入計畫選項（資料管理頁面使用：顯示所有計畫）
        async function loadPlanOptions() {
            try {
                const res = await fetch('/api/options/plans?t=' + Date.now(), {
                    cache: 'no-store',
                    headers: {
                        'Cache-Control': 'no-cache'
                    }
                });
                
                if (!res.ok) {
                    console.error('載入計畫選項失敗：', res.status, res.statusText);
                    return;
                }
                
                const json = await res.json();
                if (!json.data || json.data.length === 0) {
                    if (isDevelopment) console.warn('沒有找到任何檢查計畫');
                    // 即使沒有計畫，也要嘗試載入查詢看板的計畫選項
                    await loadFilterPlanOptions();
                    return;
                }
                
                // 更新資料管理頁面的計畫選擇下拉選單（顯示所有計畫）
                const selectIds = ['importPlanName', 'batchPlanName', 'manualPlanName', 'createPlanName'];
                selectIds.forEach(selectId => {
                    const select = document.getElementById(selectId);
                    if (select) {
                        const currentValue = select.value;
                        // 保留第一個選項（通常是「全部計畫」或「請選擇計畫」）
                        const firstOption = select.options[0] ? select.options[0].outerHTML : '';
                        
                        // 處理新的資料格式，按年度分組
                        const yearGroups = new Map(); // key: 年度, value: 該年度下的所有計畫
                        const existingValues = new Set();
                        
                        if (firstOption) {
                            const tempDiv = document.createElement('div');
                            tempDiv.innerHTML = firstOption;
                            const firstOpt = tempDiv.querySelector('option');
                            if (firstOpt && firstOpt.value) {
                                existingValues.add(firstOpt.value);
                            }
                        }
                        
                        // 將計畫按年度分組
                        json.data.forEach(p => {
                            let planName, planYear, planValue, planDisplay;
                            
                            if (typeof p === 'object' && p !== null) {
                                planName = p.name || '';
                                planYear = p.year || '';
                                planValue = p.value || `${planName}|||${planYear}`;
                                // 因為已經用年度分組，所以只顯示計畫名稱，不顯示年度
                                planDisplay = planName;
                            } else {
                                // 舊格式（字串），向後兼容
                                planName = p;
                                planYear = '';
                                planValue = p;
                                planDisplay = p;
                            }
                            
                            if (!existingValues.has(planValue) && planName) {
                                existingValues.add(planValue);
                                // 使用年度作為分組鍵，如果沒有年度則使用「未分類」
                                const groupKey = planYear || '未分類';
                                if (!yearGroups.has(groupKey)) {
                                    yearGroups.set(groupKey, []);
                                }
                                yearGroups.get(groupKey).push({ 
                                    value: planValue, 
                                    display: planDisplay, 
                                    name: planName, 
                                    year: planYear 
                                });
                            }
                        });
                        
                        // 建立選項 HTML
                        let allOptions = '';
                        
                        // 將年度分組按年度降序排序（最新的在前）
                        const sortedYears = Array.from(yearGroups.keys()).sort((a, b) => {
                            // 「未分類」放在最後
                            if (a === '未分類') return 1;
                            if (b === '未分類') return -1;
                            const yearA = parseInt(a) || 0;
                            const yearB = parseInt(b) || 0;
                            return yearB - yearA;
                        });
                        
                        sortedYears.forEach(year => {
                            const plans = yearGroups.get(year);
                            // 按計畫名稱排序（同一年度內的計畫按名稱排序）
                            plans.sort((a, b) => {
                                return (a.name || '').localeCompare(b.name || '', 'zh-TW');
                            });
                            
                            // 使用 optgroup 按年度分組
                            const yearLabel = year === '未分類' ? '未分類' : `${year} 年度`;
                            allOptions += `<optgroup label="${yearLabel}">`;
                            plans.forEach(plan => {
                                allOptions += `<option value="${plan.value}">${plan.display}</option>`;
                            });
                            allOptions += `</optgroup>`;
                        });
                        
                        // 完全重建選項列表
                        select.innerHTML = firstOption + allOptions;
                        
                        // 恢復之前選擇的值
                        if (currentValue && Array.from(select.options).some(opt => opt.value === currentValue)) {
                            select.value = currentValue;
                        }
                    }
                });
                
                // 同時更新查詢看板的計畫選項（只顯示有關聯開立事項的計畫）
                await loadFilterPlanOptions();
            } catch (e) {
                console.error("Load plans failed", e);
            }
        }
        
        // 載入查詢看板的計畫選項（只顯示有關聯開立事項的計畫）
        async function loadFilterPlanOptions() {
            try {
                const res = await fetch('/api/options/plans?withIssues=true&t=' + Date.now(), {
                    cache: 'no-store',
                    headers: {
                        'Cache-Control': 'no-cache'
                    }
                });
                
                if (!res.ok) {
                    console.error('載入查詢看板計畫選項失敗：', res.status, res.statusText);
                    return;
                }
                
                const json = await res.json();
                const select = document.getElementById('filterPlan');
                if (!select) {
                    if (isDevelopment) console.warn('找不到 filterPlan 元素');
                    return;
                }
                
                const currentValue = select.value;
                // 保留第一個選項（「全部計畫」）
                const firstOption = select.options[0] ? select.options[0].outerHTML : '';
                
                if (!json.data || json.data.length === 0) {
                    // 如果沒有資料，只保留第一個選項
                    writeLog('查詢看板：沒有找到有關聯開立事項的計畫');
                    select.innerHTML = firstOption;
                    return;
                }
                
                writeLog(`查詢看板：找到 ${json.data.length} 個有關聯開立事項的計畫`);
                
                // 處理新的資料格式，按年度分組
                const yearGroups = new Map();
                const existingValues = new Set();
                
                if (firstOption) {
                    const tempDiv = document.createElement('div');
                    tempDiv.innerHTML = firstOption;
                    const firstOpt = tempDiv.querySelector('option');
                    if (firstOpt && firstOpt.value) {
                        existingValues.add(firstOpt.value);
                    }
                }
                
                // 將計畫按年度分組
                json.data.forEach(p => {
                    let planName, planYear, planValue, planDisplay;
                    
                    if (typeof p === 'object' && p !== null) {
                        planName = p.name || '';
                        planYear = p.year || '';
                        planValue = p.value || `${planName}|||${planYear}`;
                        // 因為已經用年度分組，所以只顯示計畫名稱，不顯示年度
                        planDisplay = planName;
                    } else {
                        planName = p;
                        planYear = '';
                        planValue = p;
                        planDisplay = p;
                    }
                    
                    if (!existingValues.has(planValue) && planName) {
                        existingValues.add(planValue);
                        const groupKey = planYear || '未分類';
                        if (!yearGroups.has(groupKey)) {
                            yearGroups.set(groupKey, []);
                        }
                        yearGroups.get(groupKey).push({ 
                            value: planValue, 
                            display: planDisplay, 
                            name: planName, 
                            year: planYear 
                        });
                    }
                });
                
                // 建立選項 HTML
                let allOptions = '';
                
                // 將年度分組按年度降序排序（最新的在前）
                const sortedYears = Array.from(yearGroups.keys()).sort((a, b) => {
                    if (a === '未分類') return 1;
                    if (b === '未分類') return -1;
                    const yearA = parseInt(a) || 0;
                    const yearB = parseInt(b) || 0;
                    return yearB - yearA;
                });
                
                sortedYears.forEach(year => {
                    const plans = yearGroups.get(year);
                    plans.sort((a, b) => {
                        return (a.name || '').localeCompare(b.name || '', 'zh-TW');
                    });
                    
                    const yearLabel = year === '未分類' ? '未分類' : `${year} 年度`;
                    allOptions += `<optgroup label="${yearLabel}">`;
                    plans.forEach(plan => {
                        allOptions += `<option value="${plan.value}">${plan.display}</option>`;
                    });
                    allOptions += `</optgroup>`;
                });
                
                // 完全重建選項列表
                select.innerHTML = firstOption + allOptions;
                
                // 恢復之前選擇的值
                if (currentValue && Array.from(select.options).some(opt => opt.value === currentValue)) {
                    select.value = currentValue;
                }
            } catch (e) {
                console.error("Load filter plan options failed", e);
            }
        }
        
        // 輔助函數：從計畫選項值中提取計畫名稱和年度
        function parsePlanValue(value) {
            if (!value) return { name: '', year: '' };
            // 新格式：使用 ||| 分隔符
            if (value.includes('|||')) {
                const parts = value.split('|||');
                return { name: parts[0] || '', year: parts[1] || '' };
            }
            // 舊格式：直接是計畫名稱
            return { name: value, year: '' };
        }
        
        // 共用函數：載入計畫下的所有事項
        async function loadIssuesByPlan(planValue, options = {}) {
            const { showError = true, returnEmpty = false } = options;
            try {
                const res = await fetch(`/api/issues?page=1&pageSize=1000&planName=${encodeURIComponent(planValue)}&_t=${Date.now()}`);
                if (!res.ok) {
                    if (showError) throw new Error('載入事項列表失敗');
                    return null;
                }
                
                const json = await res.json();
                const issueList = json.data || [];
                
                if (issueList.length === 0) {
                    if (showError) {
                        showToast('該檢查計畫下尚無開立事項', 'error');
                    }
                    return returnEmpty ? [] : null;
                }
                
                return issueList;
            } catch (e) {
                if (showError) {
                    console.error('載入計畫事項失敗:', e);
                    showToast('載入事項列表失敗', 'error');
                }
                return null;
            }
        }
        
        // 共用函數：從編號提取類別代碼
        function extractKindCodeFromNumber(numberStr) {
            if (!numberStr) return null;
            const m = numberStr.match(/-([NOR])\d+$/i);
            return m ? m[1].toUpperCase() : null;
        }
        
        // 批次建檔：當選擇計畫時，自動帶入年度
        async function handleBatchPlanChange() {
            const planValue = this.value;
            const yearInput = document.getElementById('batchYear');
            if (!planValue || !yearInput) return;
            
            const { name, year } = parsePlanValue(planValue);
            if (year) {
                yearInput.value = year;
            } else if (name) {
                // 如果沒有年度資訊，嘗試從計畫名稱中提取年度
                const yearMatch = name.match(/(\d{3})年度/);
                if (yearMatch) yearInput.value = yearMatch[1];
            }
        }
        
        // 手動新增：當選擇計畫時，自動帶入年度
        async function handleManualPlanChange() {
            const planValue = this.value;
            const yearDisplay = document.getElementById('manualYearDisplay');
            if (!planValue || !yearDisplay) return;
            
            const { name, year } = parsePlanValue(planValue);
            if (year) {
                yearDisplay.value = year;
            } else if (name) {
                // 如果沒有年度資訊，嘗試從計畫名稱中提取年度
                const yearMatch = name.match(/(\d{3})年度/);
                if (yearMatch) yearDisplay.value = yearMatch[1];
            }
        }

        function initImportRoundOptions() {
            const s = document.getElementById('importRoundSelect');
            if (!s) return;
            s.innerHTML = '';
            // 支援無限次審查，先建立前 30 次選項
            for (let i = 1; i <= 30; i++) {
                s.innerHTML += `<option value="${i}">第 ${i} 次審查</option>`;
            }
        }

        async function switchView(viewId) {
            // 保存當前視圖到 sessionStorage
            sessionStorage.setItem('currentView', viewId);
            
            document.querySelectorAll('.view-section').forEach(el => {
                el.classList.remove('active');
            });
            const viewElement = document.getElementById(viewId);
            if (!viewElement) {
                console.error('View element not found:', viewId);
                return;
            }
            viewElement.classList.add('active');
            
            document.querySelectorAll('.sidebar-btn').forEach(btn => {
                btn.classList.remove('active');
            });
            const btn = document.getElementById('btn-' + viewId);
            if(btn) btn.classList.add('active');
    // [新增] 切換頁面時滾動到頂部
    window.scrollTo(0, 0);
	// 隱藏/顯示 dashboard
const dashboard = document.getElementById('dashboardSection');
if (dashboard) {
    dashboard.style.display = (viewId === 'searchView') ? 'block' : 'none';
}
    const mainContent = document. querySelector('.main-content');
    if (mainContent) mainContent.scrollTop = 0;

    // [新增] 關閉側邊欄
    const panel = document.getElementById('filtersPanel');
    if (panel && panel.classList.contains('open')) {
        onToggleSidebar();
    }

            const viewMap = {
                'planCalendarView': '/views/calendar-dashboard-view.html',
                'importView': '/views/import-view.html',
                'usersView': '/views/users-view.html'
            };
            
            if (viewMap[viewId] && !viewElement.dataset.loaded) {
                try {
                    const response = await fetch(viewMap[viewId]);
                    if (response.ok) {
                        const html = await response.text();
                        viewElement.innerHTML = html;
                        viewElement.dataset.loaded = 'true';
                        if (viewId === 'planCalendarView') {
                            setTimeout(() => {
                                loadDashboardYearOptions();
                            }, 100);
                        } else if (viewId === 'importView') {
                            setupAdminElements();
                            setTimeout(() => setupImportListeners(), 100);
                            loadPlanOptions();
                            const openPlansSchedule = sessionStorage.getItem('openPlansSchedule');
                            if (openPlansSchedule) {
                                sessionStorage.removeItem('openPlansSchedule');
                                setTimeout(() => { switchDataTab('plans'); switchPlansSubTab('schedule'); }, 250);
                            } else {
                                const savedTab = sessionStorage.getItem('currentDataTab');
                                if (savedTab) setTimeout(() => switchDataTab(savedTab), 200);
                            }
                        } else if (viewId === 'usersView') {
                            setTimeout(() => setupCleanupDaysSelect(), 100);
                        }
                    }
                } catch (error) {}
            } else if (viewId === 'planCalendarView' && viewElement.dataset.loaded) {
                loadCalendarDashboardStats();
            } else if (viewId === 'usersView' && viewElement.dataset.loaded) {
                setTimeout(() => setupCleanupDaysSelect(), 100);
            }

            if(viewId === 'searchView') {
                // 恢復查詢看板的狀態
                restoreSearchViewState();
                // 等待篩選選項載入完成後再載入資料
                setTimeout(() => {
                    loadIssuesPage(issuesPage || 1);
                    updateSortUI();
                }, 100);
            } else if (viewId === 'usersView') {
                // 恢復帳號管理頁面的狀態
                restoreUsersViewState();
                // 恢復 tab
                const savedTab = sessionStorage.getItem('currentUsersTab') || 'users';
                setTimeout(() => {
                    switchAdminTab(savedTab);
                    setupCleanupDaysSelect();
                }, 200);
            } else if (viewId === 'planCalendarView' && viewElement.dataset.loaded) {
                loadDashboardYearOptions();
            } else if (viewId === 'importView' && viewElement.dataset.loaded) {
                const openPlansSchedule = sessionStorage.getItem('openPlansSchedule');
                loadPlanOptions();
                if (openPlansSchedule) {
                    sessionStorage.removeItem('openPlansSchedule');
                    setTimeout(() => { switchDataTab('plans'); switchPlansSubTab('schedule'); }, 100);
                } else {
                    const savedTab = sessionStorage.getItem('currentDataTab');
                    if (savedTab) {
                        setTimeout(() => {
                            switchDataTab(savedTab);
                            if (savedTab === 'plans') {
                                restorePlansViewState();
                                setTimeout(() => loadPlansPage(plansPage || 1), 300);
                            }
                        }, 200);
                    }
                }
            }
        }
        
        async function loadDashboardYearOptions() {
            try {
                const res = await fetch(`/api/plans/dashboard-stats/years?t=${Date.now()}`, { credentials: 'include' });
                if (!res.ok) return;
                const data = await res.json();
                const years = data.years || [];
                const select = document.getElementById('dashboardYearSelect');
                if (!select) return;
                const currentValue = select.value || dashboardSelectedYear;
                select.innerHTML = '<option value="">請選擇年度</option>';
                years.forEach(y => {
                    const opt = document.createElement('option');
                    opt.value = y;
                    opt.textContent = `${y}年`;
                    select.appendChild(opt);
                });
                if (currentValue) {
                    select.value = currentValue;
                    dashboardSelectedYear = currentValue;
                } else if (years.length > 0) {
                    select.value = years[0];
                    dashboardSelectedYear = years[0];
                }
                if (dashboardSelectedYear) {
                    loadCalendarDashboardStats();
                    initDashboardCalendar();
                }
            } catch (e) {
                console.error('載入年度選單失敗:', e);
            }
        }

        function onDashboardYearChange() {
            const select = document.getElementById('dashboardYearSelect');
            if (!select) return;
            const year = select.value;
            if (!year) {
                dashboardSelectedYear = '';
                const statPlans = document.getElementById('dashboardStatPlans');
                const statSchedules = document.getElementById('dashboardStatSchedules');
                const statByType = document.getElementById('dashboardStatByType');
                const progressBody = document.getElementById('dashboardPlanProgressBody');
                if (statPlans) statPlans.textContent = '-';
                if (statSchedules) statSchedules.textContent = '-';
                const statPlanned = document.getElementById('dashboardStatPlanned');
                if (statPlanned) statPlanned.textContent = '-';
                if (statByType) statByType.innerHTML = '<span style="color:#64748b;">請先選擇年度</span>';
                if (progressBody) progressBody.innerHTML = '<tr><td colspan="5" style="padding:12px;color:#64748b;">請先選擇年度</td></tr>';
                dashboardMonthData = [];
                renderDashboardCalendar();
                const box = document.getElementById('dashboardScheduleDayListBody');
                if (box) box.innerHTML = '請先選擇年度';
                return;
            }
            dashboardSelectedYear = year;
            loadCalendarDashboardStats();
            loadDashboardScheduleForMonth();
            renderDashboardCalendar();
        }

        async function loadCalendarDashboardStats() {
            if (!dashboardSelectedYear) return;
            const statPlans = document.getElementById('dashboardStatPlans');
            const statSchedules = document.getElementById('dashboardStatSchedules');
            const statByType = document.getElementById('dashboardStatByType');
            const progressBody = document.getElementById('dashboardPlanProgressBody');
            if (!statPlans && !statSchedules) return;
            const typeNames = { '1': '年度定期檢查', '2': '特別檢查', '3': '例行性檢查', '4': '臨時檢查' };
            try {
                const res = await fetch(`/api/plans/dashboard-stats?year=${encodeURIComponent(dashboardSelectedYear)}&t=${Date.now()}`, { credentials: 'include' });
                if (!res.ok) throw new Error('無法載入統計');
                const data = await res.json();
                const totalPlans = data.totalPlans != null ? data.totalPlans : 0;
                const totalSchedules = data.totalSchedules != null ? data.totalSchedules : 0;
                const totalPlanned = data.totalPlanned != null ? data.totalPlanned : 0;
                const byType = data.byType || {};
                const planProgress = data.planProgress || [];
                if (statPlans) statPlans.textContent = totalPlans;
                if (statSchedules) statSchedules.textContent = totalSchedules;
                const statPlanned = document.getElementById('dashboardStatPlanned');
                if (statPlanned) statPlanned.textContent = totalPlanned;
                if (statByType) {
                    statByType.innerHTML = ['1', '2', '3', '4'].map(t => {
                        const count = byType[t] || 0;
                        return `<span style="background:#f1f5f9; padding:6px 12px; border-radius:8px; font-size:13px;">${typeNames[t] || t}：${count}</span>`;
                    }).join('');
                }
                if (progressBody) {
                    progressBody.innerHTML = planProgress.map(p => {
                        const planned = p.planned_count != null ? p.planned_count : 0;
                        const done = p.schedule_count != null ? p.schedule_count : 0;
                        const pct = planned > 0 ? Math.min(100, Math.round((done / planned) * 100)) : (done > 0 ? 100 : 0);
                        return `<tr><td data-label="年度" style="padding:10px;">${p.year || '-'}</td><td data-label="計畫名稱" style="padding:10px;">${(p.name || '-')}</td><td data-label="規劃次數" style="padding:10px;">${planned}</td><td data-label="已檢查次數" style="padding:10px;">${done}</td><td data-label="進度" style="padding:10px;"><span style="display:inline-block;width:80px;height:8px;background:#e2e8f0;border-radius:4px;overflow:hidden;"><span style="display:block;height:100%;width:${pct}%;background:#2563eb;border-radius:4px;"></span></span> ${pct}%</td></tr>`;
                    }).join('') || '<tr><td data-label="說明" colspan="5" style="padding:12px;color:#64748b;">尚無資料</td></tr>';
                }
            } catch (e) {
                if (statPlans) statPlans.textContent = '—';
                if (statSchedules) statSchedules.textContent = '—';
                const statPlanned = document.getElementById('dashboardStatPlanned');
                if (statPlanned) statPlanned.textContent = '—';
                if (statByType) statByType.innerHTML = '<span style="color:#94a3b8;">載入失敗</span>';
                if (progressBody) progressBody.innerHTML = '<tr><td data-label="說明" colspan="5" style="padding:12px;color:#ef4444;">載入失敗</td></tr>';
            }
        }

        document.addEventListener('DOMContentLoaded', async () => {
            // 首先確保 body 可見，避免空白頁面
            document.body.style.display = 'flex';
            
            try {
                await checkAuth();
                if (currentUser) {
                    // 檢查是否需要更新密碼（首次登入）
                    const mustChangePassword = sessionStorage.getItem('mustChangePassword') === 'true';
                    if (mustChangePassword) {
                        // 顯示密碼更新模態框
                        const modal = document.getElementById('changePasswordModal');
                        if (modal) {
                            modal.style.display = 'flex';
                            // 清除 sessionStorage 中的標記
                            sessionStorage.removeItem('mustChangePassword');
                            // 阻止其他操作，直到密碼更新完成
                            return;
                        }
                    }
                    
                    // 確保 body 可見（再次確認）
                    document.body.style.display = 'flex';
                    
                    // 嘗試恢復上次的視圖
                    const savedView = sessionStorage.getItem('currentView');
                    let targetView = savedView || 'searchView';
                    
                    // 確保視圖存在
                    const viewElement = document.getElementById(targetView);
                    if (!viewElement) {
                        targetView = 'searchView';
                    }
                    
                    // 切換到目標視圖（添加錯誤處理）
                    try {
                        await switchView(targetView);
                    } catch (viewError) {
                        console.error('切換視圖錯誤:', viewError);
                        // 如果切換失敗，至少顯示 searchView
                        const searchViewEl = document.getElementById('searchView');
                        if (searchViewEl) {
                            searchViewEl.classList.add('active');
                            document.querySelectorAll('.view-section').forEach(el => {
                                if (el.id !== 'searchView') el.classList.remove('active');
                            });
                        }
                    }
                    
                    // 初始化其他功能（每個都添加錯誤處理）
                    try {
                        initListeners();
                    } catch (e) {
                        console.error('初始化監聽器錯誤:', e);
                    }
                    
                    try {
                        initEditForm();
                    } catch (e) {
                        console.error('初始化編輯表單錯誤:', e);
                    }
                    
                    try {
                        initCharts();
                    } catch (e) {
                        console.error('初始化圖表錯誤:', e);
                    }
                    
                    try {
                        loadPlanOptions(); // 這會自動調用 loadFilterPlanOptions()
                        // 確保查詢看板的計畫選項也被載入
                        loadFilterPlanOptions();
                    } catch (e) {
                        console.error('載入計畫選項錯誤:', e);
                    }
                    
                    try {
                        initImportRoundOptions();
                    } catch (e) {
                        console.error('初始化匯入輪次選項錯誤:', e);
                    }
                    
                    // 如果目標視圖是 searchView，載入資料
                    if (targetView === 'searchView') {
                        try {
                            await loadIssuesPage(1);
                        } catch (e) {
                            console.error('載入事項資料錯誤:', e);
                            // 即使載入失敗，也要顯示錯誤訊息
                            const emptyMsg = document.getElementById('emptyMsg');
                            if (emptyMsg) {
                                emptyMsg.innerText = '載入資料時發生錯誤，請重新整理頁面';
                                emptyMsg.style.display = 'block';
                            }
                        }
                    }
                    // Preload users if needed
                    if(currentUser.role === 'admin' && targetView === 'usersView') {
                        try {
                            loadUsersPage(1);
                        } catch (e) {
                            console.error('載入使用者資料錯誤:', e);
                        }
                    }
                } else {
                    // 如果沒有 currentUser，應該是重定向到登入頁
                    // 但如果重定向失敗，至少顯示 body
                    if (isDevelopment) console.warn('未檢測到登入狀態，嘗試重定向到登入頁');
                }
            } catch (error) {
                console.error('初始化錯誤:', error);
                // 即使出錯也嘗試顯示頁面
                document.body.style.display = 'flex';
                
                // 顯示錯誤訊息給用戶
                const appBody = document.getElementById('appBody');
                if (appBody) {
                    appBody.innerHTML = `
                        <div style="padding: 40px; text-align: center; color: #ef4444;">
                            <h2>初始化錯誤</h2>
                            <p>頁面載入時發生錯誤，請重新整理頁面或聯絡系統管理員。</p>
                            <button onclick="window.location.reload()" class="btn btn-primary" style="margin-top: 20px;">
                                重新整理頁面
                            </button>
                        </div>
                    `;
                }
            }
        });

        async function checkAuth() {
            try {
                // 使用 Promise.race 實現超時處理
                const timeoutPromise = new Promise((_, reject) => {
                    setTimeout(() => reject(new Error('TIMEOUT')), 10000); // 10秒超時
                });
                
                const fetchPromise = apiFetch('/api/auth/me?t=' + Date.now(), { 
                    headers: { 'Cache-Control': 'no-cache' }
                });
                
                const res = await Promise.race([fetchPromise, timeoutPromise]);
                
                if (!res.ok) {
                    console.error('認證檢查失敗:', res.status, res.statusText);
                    // 如果認證失敗，重定向到登入頁
                    sessionStorage.clear();
                    window.location.href = '/login.html';
                    return;
                }
                
                const data = await res.json();
                if (data.isLogin && data.id && data.username) {
                    currentUser = data;
                    const nameEl = document.getElementById('headerUserName');
                    const roleEl = document.getElementById('headerUserRole');
                    if (nameEl) nameEl.innerText = data.name || data.username;
                    if (roleEl) roleEl.innerText = getRoleName(data.role);
                    
                    const btnCalendar = document.getElementById('btn-planCalendarView');
                    if (btnCalendar) btnCalendar.classList.remove('hidden');
                    if (['admin', 'manager'].includes(data.role)) {
                        const btnImport = document.getElementById('btn-importView');
                        if (btnImport) btnImport.classList.remove('hidden');
                        if (data.role === 'admin') {
                            const btnUsers = document.getElementById('btn-usersView');
                            if (btnUsers) btnUsers.classList.remove('hidden');
                        }
                    } else {
                        const btnImport = document.getElementById('btn-importView');
                        const btnUsers = document.getElementById('btn-usersView');
                        if (btnImport) btnImport.classList.add('hidden');
                        if (btnUsers) btnUsers.classList.add('hidden');
                    }
                } else {
                    // 未登入或資料不完整，重定向到登入頁
                    if (isDevelopment) console.warn('認證資料不完整，重定向到登入頁');
                    sessionStorage.clear();
                    window.location.href = '/login.html';
                }
            } catch (e) {
                // 如果是超時錯誤，顯示錯誤訊息
                if (e.message === 'TIMEOUT') {
                    console.error('認證檢查超時');
                    // 超時時顯示錯誤訊息，不直接重定向
                    document.body.style.display = 'flex';
                    const appBody = document.getElementById('appBody');
                    if (appBody) {
                        appBody.innerHTML = `
                            <div style="padding: 40px; text-align: center; color: #ef4444;">
                                <h2>連線逾時</h2>
                                <p>無法連線到伺服器，請檢查網路連線後重新整理頁面。</p>
                                <button onclick="window.location.reload()" class="btn btn-primary" style="margin-top: 20px;">
                                    重新整理頁面
                                </button>
                            </div>
                        `;
                    }
                } else if (e.message === 'Unauthorized') {
                    // 認證錯誤已在 apiFetch 中處理，這裡不需要再做什麼
                    return;
                } else {
                    console.error('認證檢查錯誤:', e);
                    // 其他錯誤，重定向到登入頁
                    sessionStorage.clear();
                    window.location.href = '/login.html';
                }
            }
        }
        
        // 在視圖載入後設置 admin 專屬元素的函數
        function setupAdminElements() {
            if (!currentUser || currentUser.role !== 'admin') return;
            
            const uploadCardBackup = document.getElementById('uploadCardBackup');
            if (uploadCardBackup) {
                uploadCardBackup.classList.remove('hidden');
            }
            
            const exportJsonOption = document.getElementById('exportJsonOption');
            if (exportJsonOption) {
                exportJsonOption.style.display = 'flex';
                exportJsonOption.style.alignItems = 'center';
            }
        }
        
        // [Added] 設置導入視圖的事件監聽器
        function setupImportListeners() {
            const wordInputEl = document.getElementById('wordInput');
            const importIssueDateEl = document.getElementById('importIssueDate');
            
            if (wordInputEl) {
                // [修正] 確保文件選擇框是啟用的
                wordInputEl.disabled = false;
                // 移除舊的事件監聽器（如果有的話），然後添加新的
                wordInputEl.removeEventListener('change', checkImportReady);
                wordInputEl.addEventListener('change', checkImportReady);
            }
            
            if (importIssueDateEl) {
                importIssueDateEl.removeEventListener('input', checkImportReady);
                importIssueDateEl.removeEventListener('keyup', checkImportReady);
                importIssueDateEl.addEventListener('input', checkImportReady);
                importIssueDateEl.addEventListener('keyup', checkImportReady);
            }
            
            // [Added] 初始化審查次數選項
            initImportRoundOptions();
            
            // 初始化按鈕狀態（但不禁用文件選擇框）
            checkImportReady();
        }

        function renderPagination(containerId, currentPage, totalPages, onPageChange) {
            const containerTop = document.getElementById(containerId + 'Top'); const containerBottom = document.getElementById(containerId + 'Bottom'); let html = '';
            html += `<button class="page-btn" ${currentPage === 1 ? 'disabled' : ''} onclick="${onPageChange}(${currentPage - 1})">◀</button>`;
            const delta = 2, range = [];
            for (let i = Math.max(2, currentPage - delta); i <= Math.min(totalPages - 1, currentPage + delta); i++) { range.push(i); }
            if (currentPage - delta > 2) range.unshift('...'); if (currentPage + delta < totalPages - 1) range.push('...');
            range.unshift(1); if (totalPages > 1) range.push(totalPages);
            range.forEach(i => { if (i === '...') html += `<div class="page-dots">...</div>`; else html += `<button class="page-btn ${i === currentPage ? 'active' : ''}" onclick="${onPageChange}(${i})">${i}</button>`; });
            html += `<button class="page-btn" ${currentPage === totalPages ? 'disabled' : ''} onclick="${onPageChange}(${currentPage + 1})">▶</button>`;
            if (containerTop) containerTop.innerHTML = html; if (containerBottom) containerBottom.innerHTML = html;
        }

        async function loadIssuesPage(page = 1) {
            issuesPage = page; 
            if (document.getElementById('issuesPageSizeTop')) document.getElementById('issuesPageSizeTop').value = issuesPageSize; 
            if (document.getElementById('issuesPageSizeBottom')) document.getElementById('issuesPageSizeBottom').value = issuesPageSize;
            saveSearchViewState();
            const q = document.getElementById('filterKeyword').value || '', year = document.getElementById('filterYear').value || '', unit = document.getElementById('filterUnit').value || '', status = document.getElementById('filterStatus').value || '', kind = document.getElementById('filterKind').value || '';
            const division = document.getElementById('filterDivision') ? document.getElementById('filterDivision').value : '';
            const inspection = document.getElementById('filterInspection') ? document.getElementById('filterInspection').value : '';
            const planValue = document.getElementById('filterPlan') ? document.getElementById('filterPlan').value : '';
            // 從計畫選項值中提取計畫名稱和年度（用於查詢）
            // 傳遞完整值（包含年度資訊）給後端，格式為 "planName|||year"
            const planName = planValue || '';

            // 預設以年度最新排序（降序）
            let sortField = 'year', sortDir = 'desc';
            if (sortState.field) { 
                if (sortState.field === 'number') sortField = 'title'; 
                else if (sortState.field === 'year') sortField = 'year'; 
                else if (sortState.field === 'unit') sortField = 'unit'; 
                else if (sortState.field === 'status') sortField = 'status';
                else if (sortState.field === 'content') sortField = 'content';
                else if (sortState.field === 'latest') sortField = 'updated_at'; // 最新辦理/審查情形使用更新時間排序
                sortDir = sortState.dir || 'asc'; 
            }

            const params = new URLSearchParams({ page: issuesPage, pageSize: issuesPageSize, q, year, unit, status, itemKindCode: kind, division, inspectionCategory: inspection, planName, sortField, sortDir, _t: Date.now() });

            try {
                const res = await apiFetch('/api/issues?' + params.toString());
                if (!res.ok) {
                    const errJson = await res.json().catch(() => ({}));
                    console.error("Server Error:", errJson);
                    showToast('載入資料失敗: ' + (errJson.error || res.statusText), 'error');
                    return;
                }
                const j = await res.json(); currentData = j.data || []; issuesTotal = j.total || 0; issuesPages = j.pages || 1;
                if (j.latestCreatedAt) { const d = new Date(j.latestCreatedAt); document.getElementById('dataTimestamp').innerText = `資料庫更新時間：${d.toLocaleDateString('zh-TW')} ${d.toLocaleTimeString('zh-TW', { hour: '2-digit', minute: '2-digit' })}`; } else { document.getElementById('dataTimestamp').innerText = ''; }
                if (document.getElementById('filterYear').options.length === 0 && j.globalStats) { const years = [...new Set(j.globalStats.year.map(x => x.year).filter(Boolean))].sort().reverse(); document.getElementById('filterYear').innerHTML = '<option value="">全部年度</option>' + years.map(v => `<option value="${v}">${v}</option>`).join(''); const units = [...new Set(j.globalStats.unit.map(x => x.unit).filter(Boolean))].sort(); document.getElementById('filterUnit').innerHTML = '<option value="">全部機構</option>' + units.map(v => `<option value="${v}">${v}</option>`).join(''); }
                if (j.globalStats) { cachedGlobalStats = j.globalStats; updateChartsData(j.globalStats); renderStats(j.globalStats); }
                renderTable(); renderPagination('issuesPagination', issuesPage, issuesPages, 'loadIssuesPage'); document.getElementById('issuesTotalCount').innerText = issuesTotal;
            } catch (e) { console.error(e); showToast('載入資料錯誤 (請檢查 Console)', 'error'); }
        }

        function applyFilters() { 
            issuesPage = 1; 
            saveSearchViewState();
            loadIssuesPage(1); 
        }
        
        // 保存查詢看板的狀態
        function saveSearchViewState() {
            const state = {
                keyword: document.getElementById('filterKeyword')?.value || '',
                year: document.getElementById('filterYear')?.value || '',
                plan: document.getElementById('filterPlan')?.value || '',
                unit: document.getElementById('filterUnit')?.value || '',
                status: document.getElementById('filterStatus')?.value || '',
                kind: document.getElementById('filterKind')?.value || '',
                division: document.getElementById('filterDivision')?.value || '',
                inspection: document.getElementById('filterInspection')?.value || '',
                page: issuesPage,
                pageSize: issuesPageSize,
                sortField: sortState.field || '',
                sortDir: sortState.dir || 'asc'
            };
            sessionStorage.setItem('searchViewState', JSON.stringify(state));
        }
        
        // 恢復查詢看板的狀態
        function restoreSearchViewState() {
            const saved = sessionStorage.getItem('searchViewState');
            if (!saved) return;
            
            try {
                const state = JSON.parse(saved);
                // 每次重新載入後，所有篩選條件都恢復為預設值（清空）
                if (document.getElementById('filterKeyword')) document.getElementById('filterKeyword').value = '';
                if (document.getElementById('filterYear')) document.getElementById('filterYear').value = '';
                if (document.getElementById('filterPlan')) document.getElementById('filterPlan').value = '';
                if (document.getElementById('filterUnit')) document.getElementById('filterUnit').value = '';
                if (document.getElementById('filterStatus')) document.getElementById('filterStatus').value = '';
                if (document.getElementById('filterKind')) document.getElementById('filterKind').value = '';
                if (document.getElementById('filterDivision')) document.getElementById('filterDivision').value = '';
                if (document.getElementById('filterInspection')) document.getElementById('filterInspection').value = '';
                
                // 保留分頁和排序狀態（這些是瀏覽狀態，不是篩選條件）
                if (state.page) issuesPage = state.page;
                if (state.pageSize) issuesPageSize = state.pageSize;
                if (state.sortField) sortState.field = state.sortField;
                if (state.sortDir) sortState.dir = state.sortDir;
            } catch (e) {
                // 忽略解析錯誤
            }
        }
        function resetFilters() { document.querySelectorAll('.filter-input,.filter-select').forEach(e => e.value = ''); applyFilters(); }
        function sortData(field) { 
            if (sortState.field === field) {
                sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc'; 
            } else { 
                sortState.field = field; 
                sortState.dir = 'asc'; 
            } 
            saveSearchViewState();
            loadIssuesPage(1); 
            updateSortUI(); 
        }
        function updateSortUI() { document.querySelectorAll('th').forEach(th => { th.classList.remove('sort-asc', 'sort-desc'); if (th.getAttribute('onclick') && th.getAttribute('onclick').includes(`'${sortState.field}'`)) th.classList.add(sortState.dir === 'asc' ? 'sort-asc' : 'sort-desc'); }); }
        function renderStats(stats) { const s = stats.status; const total = s.reduce((sum, item) => sum + parseInt(item.count), 0); const active = s.find(x => x.status === '持續列管')?.count || 0; const resolved = s.filter(x => ['解除列管', '自行列管'].includes(x.status)).reduce((sum, x) => sum + parseInt(x.count), 0); document.getElementById('countTotal').innerText = total; document.getElementById('countActive').innerText = active; document.getElementById('countResolved').innerText = resolved; }
        
        function updateBatchUI() {
            const checkboxes = document.querySelectorAll('.issue-check:checked');
            const count = checkboxes.length;
            const container = document.getElementById('batchActionContainer');
            const badge = document.getElementById('selectedCountBadge');
            
            if (count > 0) {
                container.style.display = 'block';
                badge.textContent = `(${count})`;
            } else {
                container.style.display = 'none';
                badge.textContent = '';
            }
        }
        
        function toggleAllCheckboxes() {
            const selectAll = document.getElementById('selectAll');
            const checkboxes = document.querySelectorAll('.issue-check');
            checkboxes.forEach(cb => cb.checked = selectAll.checked);
            updateBatchUI();
        }
        
        async function batchDeleteIssues() {
            const checkboxes = document.querySelectorAll('.issue-check:checked');
            if (checkboxes.length === 0) {
                showToast('請至少選擇一筆資料', 'error');
                return;
            }
            
            const ids = Array.from(checkboxes).map(cb => cb.value);
            const confirmed = await showConfirmModal(`確定要刪除 ${ids.length} 筆資料嗎？\n\n此操作無法復原！`, '確定刪除', '取消');
            if (!confirmed) {
                return;
            }
            
            try {
                const res = await apiFetch('/api/issues/batch-delete', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ ids })
                });
                
                if (res.ok) {
                    // 清除所有選中的checkbox
                    checkboxes.forEach(cb => cb.checked = false);
                    // 更新批次操作UI
                    updateBatchUI();
                    // 重新載入資料
                    await loadIssuesPage(issuesPage);
                    showToast(`成功刪除 ${ids.length} 筆資料`, 'success');
                } else {
                    const j = await res.json();
                    showToast('刪除失敗: ' + (j.error || '不明錯誤'), 'error');
                }
            } catch (e) {
                showToast('刪除失敗: ' + e.message, 'error');
            }
        }

        function renderTable() {
            const tbody = document.getElementById('dataBody'); tbody.innerHTML = '';
            if (!currentData || currentData.length === 0) { document.getElementById('emptyMsg').style.display = 'block'; return; }
            document.getElementById('emptyMsg').style.display = 'none';
            const canManage = currentUser && ['admin', 'manager'].includes(currentUser.role);
            const canEdit = currentUser && ['admin', 'manager', 'editor'].includes(currentUser.role);
            const isViewer = currentUser && currentUser.role === 'viewer';
            document.getElementById('batchActionContainer').style.display = 'none'; document.getElementById('selectedCountBadge').innerText = ''; document.getElementById('selectAll').checked = false;
            document.querySelectorAll('.manager-col').forEach(el => el.style.display = canManage ? 'table-cell' : 'none');

            let html = '';
            currentData.forEach(item => {
                try {
                    let badge = '';
                    const st = String(item.status || 'Open');
                    if (st !== 'Open') {
                        const stClass = st === '持續列管' ? 'active' : (st === '解除列管' ? 'resolved' : 'self');
                        badge = `<span class="badge ${stClass}">${st}</span>`;
                    }
                    // [修正] 顯示最新的審查或辦理情形（比較輪次）
                    let updateTxt = '-';
                    const latest = getLatestReviewOrHandling(item);
                    if (latest) {
                        const prefix = latest.type === 'review' ? '[審]' : '[回]';
                        updateTxt = `${prefix} ${stripHtml(latest.content).slice(0, 80)}`;
                    }
                    let aiContent = ''; if (item.aiResult && item.aiResult.status === 'done') { const f = String(item.aiResult.fulfill || ''); const isYes = f.includes('是') || f.includes('Yes'); aiContent = `<div class="ai-tag ${isYes ? 'yes' : 'no'}">${isYes ? '✅' : '⚠️'} ${f}</div>`; }
                    // 檢視人員顯示「查看詳情」按鈕（不顯示「審查」字樣），其他權限顯示「審查/查看詳情」
                    const btnText = isViewer ? '✏️ 查看詳情' : '✏️ 審查/查看詳情';
                    const editBtn = (canEdit || isViewer) ? `<button class="badge" style="background:#fff;border:1px solid #ddd;cursor:pointer;margin-top:4px;" onclick="event.stopPropagation();openDetail('${item.id}',false)">${btnText}</button>` : '';
                    const checkbox = canManage ? `<td class="manager-col"><input type="checkbox" class="issue-check" value="${item.id}" onclick="event.stopPropagation(); updateBatchUI()"></td>` : `<td class="manager-col" style="display:none"></td>`;

                    let k = item.itemKindCode;
                    if (!k) {
                        k = extractKindCodeFromNumber(item.number);
                    }

                    let kindLabel = getKindLabel(k);

                    const statusHtml = `<div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">${kindLabel}${badge}</div>`;
                    const snippet = stripHtml(item.content || '').slice(0, 180);
                    const fullHtml = String(item.content || '');

                    html += `<tr onclick="openDetail('${item.id}',false)"> ${checkbox} <td data-label="年度">${item.year}</td><td data-label="編號" style="font-weight:600;color:var(--primary);">${item.number}</td><td data-label="機構">${item.unit}</td><td data-label="狀態與類型">${statusHtml}</td><td data-label="事項內容"><div class="text-content">${snippet}${(stripHtml(item.content || '').length > 180 ? ` <a href='javascript:void(0)' onclick="event.stopPropagation();showPreview(${JSON.stringify(fullHtml)}, '編號 ${item.number} 內容')">...更多</a>` : '')}</div></td><td data-label="最新辦理/審查情形"><div class="text-content">${stripHtml(updateTxt)}</div></td><td data-label="操作"><div style="display:flex;flex-direction:column;gap:4px;align-items:flex-start;">${aiContent}${editBtn}</div></td></tr>`;
                } catch (err) {
                    console.error("Skipping bad row:", item, err);
                }
            });
            tbody.innerHTML = html;
        }

        function onIssuesPageSizeChange(val) { 
            issuesPageSize = parseInt(val, 10); 
            saveSearchViewState();
            loadIssuesPage(1); 
        }
        // 保存帳號管理頁面的狀態
        function saveUsersViewState() {
            const state = {
                search: document.getElementById('userSearch')?.value || '',
                page: usersPage,
                pageSize: usersPageSize,
                sortField: usersSortField,
                sortDir: usersSortDir,
                tab: sessionStorage.getItem('currentUsersTab') || 'users'
            };
            sessionStorage.setItem('usersViewState', JSON.stringify(state));
        }
        
        // 恢復帳號管理頁面的狀態
        function restoreUsersViewState() {
            const saved = sessionStorage.getItem('usersViewState');
            if (!saved) return;
            
            try {
                const state = JSON.parse(saved);
                if (document.getElementById('userSearch')) document.getElementById('userSearch').value = state.search || '';
                if (state.page) usersPage = state.page;
                if (state.pageSize) usersPageSize = state.pageSize;
                if (state.sortField) usersSortField = state.sortField;
                if (state.sortDir) usersSortDir = state.sortDir;
                if (state.tab) sessionStorage.setItem('currentUsersTab', state.tab);
            } catch (e) {
                // 忽略解析錯誤
            }
        }
        
        async function loadUsersPage(page = 1) { 
            // 檢查是否在 usersView 中
            const usersView = document.getElementById('usersView');
            if (!usersView || !usersView.classList.contains('active')) {
                // 如果不在 usersView，不執行載入
                return;
            }
            
            usersPage = page; 
            const usersPageSizeEl = document.getElementById('usersPageSize');
            if (!usersPageSizeEl) {
                usersPageSize = 20;
            } else {
                usersPageSize = parseInt(usersPageSizeEl.value, 10) || 20;
            }
            const userSearchEl = document.getElementById('userSearch');
            const q = userSearchEl ? (userSearchEl.value || '') : ''; 
            saveUsersViewState();
            const params = new URLSearchParams({ page: usersPage, pageSize: usersPageSize, q, sortField: usersSortField, sortDir: usersSortDir, _t: Date.now() }); 
            try { 
                const res = await apiFetch('/api/users?' + params.toString()); 
                if (!res.ok) { showToast('載入使用者失敗', 'error'); return; } 
                const j = await res.json(); 
                userList = j.data || []; 
                usersTotal = j.total || 0; 
                usersPages = j.pages || 1; 
                renderUsers(); 
                const usersPaginationEl = document.getElementById('usersPagination');
                if (usersPaginationEl) {
                    renderPagination('usersPagination', usersPage, usersPages, 'loadUsersPage'); 
                }
            } catch (e) { 
                showToast('載入使用者錯誤', 'error'); 
            } 
        }
        function renderUsers() { 
            const tbody = document.getElementById('usersTableBody');
            if (!tbody) {
                if (isDevelopment) console.warn('usersTableBody element not found');
                return;
            }
            tbody.innerHTML = userList.map(u => `<tr><td data-label="姓名" style="padding:12px;">${u.name || '-'}</td><td data-label="帳號">${u.username}</td><td data-label="權限">${getRoleName(u.role)}</td><td data-label="註冊時間">${new Date(u.created_at).toLocaleDateString()}</td><td data-label="操作">${u.id !== currentUser.userId ? `<button class="btn btn-outline" style="padding:2px 6px;margin-right:4px;" onclick="openUserModal('edit', ${u.id})">✏️</button><button class="btn btn-danger" style="padding:2px 6px;" onclick="deleteUser(${u.id})">🗑️</button>` : '-'}</td></tr>`).join(''); 
        }
        function usersSortBy(field) { 
            if (usersSortField === field) {
                usersSortDir = usersSortDir === 'asc' ? 'desc' : 'asc'; 
            } else { 
                usersSortField = field; 
                usersSortDir = 'asc'; 
            } 
            saveUsersViewState();
            loadUsersPage(1); 
        }

        // 保存登入紀錄頁面的狀態
        function saveLogsViewState() {
            const state = {
                search: document.getElementById('loginSearch')?.value || '',
                page: logsPage,
                pageSize: logsPageSize
            };
            sessionStorage.setItem('logsViewState', JSON.stringify(state));
        }
        
        // 恢復登入紀錄頁面的狀態
        function restoreLogsViewState() {
            const saved = sessionStorage.getItem('logsViewState');
            if (!saved) return;
            
            try {
                const state = JSON.parse(saved);
                if (document.getElementById('loginSearch')) document.getElementById('loginSearch').value = state.search || '';
                if (state.page) logsPage = state.page;
                if (state.pageSize) logsPageSize = state.pageSize;
            } catch (e) {
                // 忽略解析錯誤
            }
        }
        
        async function loadLogsPage(page = 1) {
            const loginSearchEl = document.getElementById('loginSearch');
            if (!loginSearchEl) {
                if (isDevelopment) console.warn('loginSearch element not found');
                return;
            }
            logsPage = page;
            const q = loginSearchEl.value || '';
            saveLogsViewState();
            const params = new URLSearchParams({ page: logsPage, pageSize: logsPageSize, q, _t: Date.now() });
            const logsLoadingEl = document.getElementById('logsLoading');
            if (logsLoadingEl) logsLoadingEl.style.display = 'block';
            try {
                const res = await apiFetch('/api/admin/logs?' + params.toString());
                if (!res.ok) {
                    showToast('載入登入紀錄失敗', 'error');
                    return;
                }
                const j = await res.json();
                currentLogs.login = j.data || [];
                logsTotal = j.total || 0;
                logsPages = j.pages || 1;
                const logsTableBody = document.getElementById('logsTableBody');
                if (logsTableBody) {
                    logsTableBody.innerHTML = currentLogs.login.map(l => `<tr><td data-label="時間" style="padding:12px;">${new Date(l.login_time).toLocaleString('zh-TW')}</td><td data-label="帳號">${l.username}</td><td data-label="IP">${l.ip_address || '-'}</td></tr>`).join('');
                }
                renderPagination('logsPagination', logsPage, logsPages, 'loadLogsPage');
            } catch (e) {
                console.error(e);
                showToast('載入登入紀錄錯誤', 'error');
            } finally {
                if (logsLoadingEl) logsLoadingEl.style.display = 'none';
            }
        }
        
        // 保存操作歷程頁面的狀態
        function saveActionsViewState() {
            const state = {
                search: document.getElementById('actionSearch')?.value || '',
                page: actionsPage,
                pageSize: actionsPageSize
            };
            sessionStorage.setItem('actionsViewState', JSON.stringify(state));
        }
        
        // 恢復操作歷程頁面的狀態
        function restoreActionsViewState() {
            const saved = sessionStorage.getItem('actionsViewState');
            if (!saved) return;
            
            try {
                const state = JSON.parse(saved);
                if (document.getElementById('actionSearch')) document.getElementById('actionSearch').value = state.search || '';
                if (state.page) actionsPage = state.page;
                if (state.pageSize) actionsPageSize = state.pageSize;
            } catch (e) {
                // 忽略解析錯誤
            }
        }
        
        async function loadActionsPage(page = 1) {
            const actionSearchEl = document.getElementById('actionSearch');
            if (!actionSearchEl) {
                if (isDevelopment) console.warn('actionSearch element not found');
                return;
            }
            actionsPage = page;
            const q = actionSearchEl.value || '';
            saveActionsViewState();
            const params = new URLSearchParams({ page: actionsPage, pageSize: actionsPageSize, q, _t: Date.now() });
            const logsLoadingEl = document.getElementById('logsLoading');
            if (logsLoadingEl) logsLoadingEl.style.display = 'block';
            try {
                const res = await apiFetch('/api/admin/action_logs?' + params.toString());
                if (!res.ok) {
                    showToast('載入操作紀錄失敗', 'error');
                    return;
                }
                const j = await res.json();
                currentLogs.action = j.data || [];
                actionsTotal = j.total || 0;
                actionsPages = j.pages || 1;
                const actionsTableBody = document.getElementById('actionsTableBody');
                if (actionsTableBody) {
                    actionsTableBody.innerHTML = currentLogs.action.map(l => `<tr><td data-label="時間" style="padding:12px;white-space:nowrap;">${new Date(l.created_at).toLocaleString('zh-TW')}</td><td data-label="帳號">${l.username}</td><td data-label="動作"><span class="badge new">${l.action}</span></td><td data-label="詳細內容"><div style="font-size:12px;color:#666;">${l.details}</div></td></tr>`).join('');
                }
                renderPagination('actionsPagination', actionsPage, actionsPages, 'loadActionsPage');
            } catch (e) {
                console.error(e);
                showToast('載入操作紀錄錯誤', 'error');
            } finally {
                if (logsLoadingEl) logsLoadingEl.style.display = 'none';
            }
        }

        function exportLogs(type) { const data = type === 'login' ? currentLogs.login : currentLogs.action; if (!data || data.length === 0) return showToast('無資料可匯出', 'error'); let csvContent = '\uFEFF'; if (type === 'login') { csvContent += "時間,帳號,IP位址\n"; data.forEach(row => { csvContent += `"${new Date(row.login_time).toLocaleString('zh-TW')}","${row.username}","${row.ip_address}"\n`; }); } else { csvContent += "時間,帳號,動作,詳細內容\n"; data.forEach(row => { csvContent += `"${new Date(row.created_at).toLocaleString('zh-TW')}","${row.username}","${row.action}","${(row.details || '').replace(/"/g, '""')}"\n`; }); } const link = document.createElement("a"); link.setAttribute("href", URL.createObjectURL(new Blob([csvContent], { type: 'text/csv;charset=utf-8;' }))); link.setAttribute("download", `${type}_logs_${new Date().toISOString().slice(0, 10)}.csv`); document.body.appendChild(link); link.click(); document.body.removeChild(link); }
        // 刪除資料庫記錄（根據選擇：刪除舊記錄或全部）
        async function deleteLogsFromDB(type) {
            const daysSelect = document.getElementById(type === 'login' ? 'loginCleanupDays' : 'actionCleanupDays');
            const customDaysInput = document.getElementById(type === 'login' ? 'loginCustomDays' : 'actionCustomDays');
            const logTypeName = type === 'login' ? '登入' : '操作';
            
            // 如果選擇"刪除全部"
            if (daysSelect.value === 'all') {
                const confirmed = await showConfirmModal(`確定要刪除資料庫中所有「${logTypeName}」紀錄嗎？\n\n此動作無法復原！`, '確定刪除', '取消');
                if (!confirmed) {
                    return;
                }
                
                const endpoint = type === 'login' ? '/api/admin/logs' : '/api/admin/action_logs';
                try {
                    const res = await apiFetch(endpoint, { method: 'DELETE' });
                    if (res.ok) {
                        showToast('資料庫記錄已全部刪除');
                        if (type === 'login') loadLogsPage(1);
                        else loadActionsPage(1);
                    } else {
                        showToast('刪除失敗', 'error');
                    }
                } catch (e) {
                    showToast('Error: ' + e.message, 'error');
                }
                return;
            }
            
            // 刪除指定天數前的記錄
            let days = parseInt(daysSelect.value);
            
            if (daysSelect.value === 'custom') {
                days = parseInt(customDaysInput.value);
                if (!days || days < 1) {
                    showToast('請輸入有效的保留天數（至少1天）', 'error');
                    return;
                }
            }
            
            const confirmed = await showConfirmModal(`確定要刪除資料庫中 ${days} 天前的「${logTypeName}」紀錄嗎？\n\n將保留最近 ${days} 天的記錄，刪除更早的記錄。\n\n此動作無法復原！`, '確定刪除', '取消');
            if (!confirmed) {
                return;
            }
            
            const endpoint = type === 'login' ? '/api/admin/logs/cleanup' : '/api/admin/action_logs/cleanup';
            try {
                const res = await apiFetch(endpoint, {
                    method: 'POST',
                    body: JSON.stringify({ days })
                });
                const data = await res.json();
                if (res.ok) {
                    showToast(`已刪除資料庫中 ${data.deleted || 0} 筆 ${days} 天前的${logTypeName}紀錄`);
                    if (type === 'login') loadLogsPage(1);
                    else loadActionsPage(1);
                } else {
                    showToast(data.error || '刪除失敗', 'error');
                }
            } catch (e) {
                showToast('Error: ' + e.message, 'error');
            }
        }
        
        // 處理自訂天數輸入框的顯示/隱藏
        function setupCleanupDaysSelect() {
            const loginSelect = document.getElementById('loginCleanupDays');
            const actionSelect = document.getElementById('actionCleanupDays');
            const loginCustom = document.getElementById('loginCustomDays');
            const actionCustom = document.getElementById('actionCustomDays');
            
            if (loginSelect) {
                loginSelect.addEventListener('change', function() {
                    loginCustom.classList.toggle('hidden', this.value !== 'custom');
                    if (this.value !== 'custom') loginCustom.value = '';
                });
            }
            if (actionSelect) {
                actionSelect.addEventListener('change', function() {
                    actionCustom.classList.toggle('hidden', this.value !== 'custom');
                    if (this.value !== 'custom') actionCustom.value = '';
                });
            }
        }

        function switchAdminTab(tab) {
            // 保存當前 tab 到 sessionStorage
            sessionStorage.setItem('currentAdminTab', tab); 
            // 保存當前 tab
            sessionStorage.setItem('currentUsersTab', tab);
            saveUsersViewState();
            
            document.querySelectorAll('.admin-tab-btn').forEach(b => b.classList.remove('active')); 
            if (event && event.target) {
                event.target.classList.add('active');
            } else {
                // 如果沒有 event，找到對應的按鈕
                const buttons = document.querySelectorAll('.admin-tab-btn');
                buttons.forEach(btn => {
                    if (btn.getAttribute('onclick') && btn.getAttribute('onclick').includes(`'${tab}'`)) {
                        btn.classList.add('active');
                    }
                });
            }
            document.getElementById('tab-users').classList.toggle('hidden', tab !== 'users'); 
            document.getElementById('tab-import-export').classList.toggle('hidden', tab !== 'import-export');
            document.getElementById('tab-logs').classList.toggle('hidden', tab !== 'logs'); 
            document.getElementById('tab-actions').classList.toggle('hidden', tab !== 'actions'); 
            if (tab === 'logs') {
                restoreLogsViewState();
                loadLogsPage(logsPage || 1); 
            }
            if (tab === 'actions') {
                restoreActionsViewState();
                loadActionsPage(actionsPage || 1); 
            }
            if (tab === 'users') {
                loadUsersPage(usersPage || 1);
            }
        }

        // [修正與增強] HTML 解析核心：提升對 Word 表格的容錯率
        function parseFromHTML(html) {
            var items = [];
            try {
                var doc = new DOMParser().parseFromString(html, "text/html");
                var tables = doc.querySelectorAll("table");

                tables.forEach(function (table) {
                    var rows = table.querySelectorAll("tr");
                    var headerRow = -1, dataStart = -1;

                    for (var i = 0; i < Math.min(rows.length, 10); i++) {
                        var t = (rows[i].innerText || rows[i].textContent || "").replace(/\s+/g, "");
                        if ((/編號|項次|序號/).test(t) && (/內容|摘要/).test(t)) {
                            headerRow = i;
                            dataStart = i + 1;
                            break;
                        }
                    }

                    if (headerRow === -1) return;

                    var headerCells = rows[headerRow].querySelectorAll("td,th");
                    var col = { number: -1, content: -1, handling: -1, result: -1 };
                    var reviewCols = [];

                    headerCells.forEach(function (cell, idx) {
                        var text = (cell.innerText || cell.textContent || "").replace(/\s+/g, "");
                        if ((/編號|項次|序號/).test(text)) col.number = idx;
                        else if ((/事項內容|缺失內容|觀察內容|內容/).test(text)) col.content = idx;
                        else if ((/辦理情形|改善情形/).test(text)) col.handling = idx;
                        else if ((/結果|狀態|列管/).test(text)) col.result = idx;

                        var mm = text.match(/第(\d+)次.*(審查|意見)/);
                        if (mm) {
                            reviewCols.push({ idx: idx, round: parseInt(mm[1], 10) });
                        } else if ((/審查意見|意見審查/).test(text)) {
                            reviewCols.push({ idx: idx, round: 1 });
                        }
                    });

                    if (col.number === -1) col.number = 0;
                    if (col.content === -1) col.content = (col.number === 0) ? 1 : 0;

                    for (var r = dataStart; r < rows.length; r++) {
                        var cells = rows[r].querySelectorAll("td,th");
                        if (cells.length < 2) continue;

                        var rawNumText = extractNumberFromCell(cells[col.number]);
                        var info = parseItemNumber(rawNumText);
                        if (!info || !info.raw) continue;

                        var orgUnifiedCode = ORG_CROSSWALK[info.orgCodeRaw] || info.orgCodeRaw || info.orgCode || "";
                        var orgCodeToUse = info.orgCodeRaw || info.orgCode || "";
                        var unitName = ORG_MAP[orgCodeToUse] || orgCodeToUse || "";
                        var inspectName = INSPECTION_MAP[info.inspectCode] || info.inspectCode || "";
                        var divCodeToUse = info.divisionCode || info.divCode || "";
                        var divName = DIVISION_MAP[divCodeToUse] || divCodeToUse || "";
                        var kindName = KIND_MAP[info.kindCode] || "其他";
                        
                        // 使用規範化編號（如果解析成功），否則使用原始編號
                        var canonicalNum = canonicalNumber(info);
                        var finalNumber = canonicalNum || info.raw.toUpperCase();

                        var item = {
                            number: finalNumber,
                            rawNumber: info.raw.toUpperCase(),
                            scheme: info.scheme || "",
                            year: String(info.yearRoc || ""),
                            yearRoc: info.yearRoc || "",
                            unit: unitName,
                            orgCodeRaw: orgCodeToUse,
                            orgUnifiedCode: orgUnifiedCode,
                            orgName: unitName,
                            itemKindCode: info.kindCode || "",
                            category: kindName,
                            inspectionCategoryCode: info.inspectCode || "",
                            inspectionCategoryName: inspectName,
                            divisionCode: divCodeToUse,
                            divisionName: divName,
                            divisionSeq: info.divisionSeq || "",
                            itemSeq: info.itemSeq || "",
                            period: info.period || "",
                            content: "",
                            handling: "",
                            status: "持續列管"
                        };

                        if (col.content !== -1 && cells[col.content]) item.content = sanitizeContent(cells[col.content].innerHTML);
                        if (col.handling !== -1 && cells[col.handling]) item.handling = sanitizeContent(cells[col.handling].innerHTML);
                        if (info.kindCode === "R") item.status = "自行列管"; else if (col.result !== -1 && cells[col.result]) item.status = parseStatusFromResultCell(cells[col.result]) || "持續列管";

                        reviewCols.forEach(function (rc) {
                            var key = (rc.round === 1 ? "review" : ("review" + rc.round));
                            if (cells[rc.idx]) item[key] = sanitizeContent(cells[rc.idx].innerHTML);
                        });

                        items.push(item);
                    }
                });
            } catch (e) {
                console.error("Parse error:", e);
                showToast("解析 Word 表格時發生錯誤，請確認表格格式是否包含「編號」與「內容」欄位。", 'error');
            }
            return items;
        }

        function onImportStageChange() {
            const stage = document.querySelector('input[name="importStage"]:checked').value;
            const roundContainer = document.getElementById('importRoundContainer');
            const planNameContainer = document.getElementById('importPlanNameContainer');

            if (stage === 'initial') {
                roundContainer.style.display = 'none';
                planNameContainer.style.gridColumn = 'span 2';
                document.getElementById('importDateGroup_Initial').style.display = 'block';
                document.getElementById('importDateGroup_Review').style.display = 'none';
                document.getElementById('importStatusWord').innerText = '';
            } else {
                roundContainer.style.display = 'block';
                planNameContainer.style.gridColumn = 'auto';
                document.getElementById('importDateGroup_Initial').style.display = 'none';
                document.getElementById('importDateGroup_Review').style.display = 'block';
            }
            checkImportReady();
        }

        function checkImportReady() {
            const wordInputEl = document.getElementById('wordInput');
            const btnParseWordEl = document.getElementById('btnParseWord');
            if (!wordInputEl || !btnParseWordEl) return;
            
            const f = wordInputEl.files[0];
            if (currentImportMode === 'backup') return;

            const stageRadio = document.querySelector('input[name="importStage"]:checked');
            if (!stageRadio) return;

            const stage = stageRadio.value;
            let valid = false;

            if (stage === 'initial') {
                const importIssueDateEl = document.getElementById('importIssueDate');
                const d = importIssueDateEl ? importIssueDateEl.value.trim() : '';
                valid = (d.length > 0);
            } else {
                valid = true;
            }

            // [修正] 允許先選擇文件，不限制文件選擇框
            // wordInputEl.disabled = !valid;  // 移除這行，允許隨時選擇文件
            // [修正] 只有在日期未填寫且沒有文件時才禁用按鈕
            btnParseWordEl.disabled = !valid || !f;
        }

        async function previewWord() {
            const f = document.getElementById('wordInput').files[0], round = document.getElementById('importRoundSelect') ? document.getElementById('importRoundSelect').value : 1, msg = document.getElementById('importStatusWord');
            if (!f) return showToast('請先選擇 Word 檔案', 'error');
            msg.innerText = 'Word 解析中...';
            currentImportMode = 'word';
            try {
                const b = await f.arrayBuffer();
                const r = await mammoth.convertToHtml({ arrayBuffer: b });
                const items = parseFromHTML(r.value);
                processParsedItems(items, round, msg);
            } catch (e) { console.error(e); msg.innerText = 'Word 解析錯誤: ' + e.message; }
        }

        function parseHistoryField(text) {
            if (!text || typeof text !== 'string') return {};
            const chunks = {};
            const matches = [...text.matchAll(/\[第(\d+)次\]/g)];
            if (matches.length === 0) return {};
            matches.forEach((m, i) => {
                const round = parseInt(m[1], 10);
                const start = m.index + m[0].length;
                const end = (i + 1 < matches.length) ? matches[i + 1].index : text.length;
                let content = text.substring(start, end).trim();
                content = content.replace(/^-+\s*|\s*-+$/g, '');
                if (content) chunks[round] = content;
            });
            return chunks;
        }

        async function previewBackup() {
            const f = document.getElementById('backupInput').files[0];
            const msg = document.getElementById('importStatusBackup');

            if (!f) return showToast('請先選擇備份檔案', 'error');
            if (!msg) { showToast("系統錯誤：找不到狀態顯示區域", 'error'); return; }

            msg.innerText = '備份檔解析中...';
            currentImportMode = 'backup';
            const ext = f.name.split('.').pop().toLowerCase();

            try {
                let items = [];
                if (ext === 'json') {
                    const text = await f.text();
                    const json = JSON.parse(text);
                    const rawItems = Array.isArray(json) ? json : (json.data || []);
                    items = rawItems.map(i => {
                        const newItem = {
                            number: i.number || i['編號'] || '',
                            year: i.year || i['年度'] || '',
                            unit: i.unit || i['機構'] || '',
                            content: i.content || i['內容'] || i['事項內容'] || i['內容摘要'] || '',
                            status: i.status || i['狀態'] || '持續列管',
                            handling: i.handling || i['辦理情形'] || i['最新辦理情形'] || '',
                            review: i.review || i['審查意見'] || i['最新審查意見'] || '',
                            itemKindCode: i.itemKindCode,
                            category: i.category,
                            divisionName: i.division,
                            inspectionCategoryName: i.inspection_category,
                            planName: i.planName,
                            issueDate: i.issueDate
                        };

                        // 支持無限次，動態查找（從1到200，實際應該不會超過這個數字）
                        for (let k = 1; k <= 200; k++) {
                            const suffix = k === 1 ? '' : k;
                            if (i[`handling${suffix}`]) newItem[`handling${suffix}`] = i[`handling${suffix}`];
                            if (i[`review${suffix}`]) newItem[`review${suffix}`] = i[`review${suffix}`];
                        }

                        const potentialHandling = i['完整辦理情形歷程'] || i.fullHandling || i.handling || i['辦理情形'] || '';
                        const potentialReview = i['完整審查意見歷程'] || i.fullReview || i.review || i['審查意見'] || '';

                        const hChunks = parseHistoryField(potentialHandling);
                        const rChunks = parseHistoryField(potentialReview);

                        Object.keys(hChunks).forEach(r => { const key = parseInt(r) === 1 ? 'handling' : `handling${r}`; newItem[key] = hChunks[r]; });
                        Object.keys(rChunks).forEach(r => { const key = parseInt(r) === 1 ? 'review' : `review${r}`; newItem[key] = rChunks[r]; });

                        return newItem;
                    });
                    processParsedItems(items, 0, msg);
                } else if (ext === 'csv') {
                    Papa.parse(f, {
                        header: true,
                        skipEmptyLines: true,
                        encoding: "UTF-8",
                        complete: function (results) {
                            const msgInside = document.getElementById('importStatusBackup');
                            try {
                                if (results.errors.length && results.data.length === 0) { if (msgInside) msgInside.innerText = 'CSV 解析錯誤'; return; }
                                const mapped = results.data.map(i => {
                                    let item = {
                                        number: i['編號'] || i.number || '',
                                        year: i['年度'] || i.year || '',
                                        unit: i['機構'] || i.unit || '',
                                        content: i['內容'] || i['事項內容'] || i['內容摘要'] || i.content || '',
                                        status: i['狀態'] || i.status || '持續列管',
                                        handling: i['最新辦理情形'] || i['辦理情形'] || i.handling || '',
                                        review: i['最新審查意見'] || i['審查意見'] || i.review || ''
                                    };
                                    const fullH = i['完整辦理情形歷程'] || i.handling || '';
                                    const fullR = i['完整審查意見歷程'] || i.review || '';
                                    const hChunks = parseHistoryField(fullH);
                                    const rChunks = parseHistoryField(fullR);
                                    Object.keys(hChunks).forEach(r => { const key = parseInt(r) === 1 ? 'handling' : `handling${r}`; item[key] = hChunks[r]; });
                                    Object.keys(rChunks).forEach(r => { const key = parseInt(r) === 1 ? 'review' : `review${r}`; item[key] = rChunks[r]; });
                                    return item;
                                });
                                const validRows = mapped.filter(r => r.number || r.content);
                                if (validRows.length === 0) { if (msgInside) msgInside.innerText = '錯誤：未解析到有效資料'; return; }
                                processParsedItems(mapped, 0, msgInside);
                            } catch (err) { console.error(err); if (msgInside) msgInside.innerText = 'CSV 處理錯誤: ' + err.message; }
                        }
                    });
                } else { throw new Error('不支援的檔案格式 (僅限 JSON 或 CSV)'); }
            } catch (e) { console.error(e); msg.innerText = '解析錯誤: ' + e.message; }
        }

        function processParsedItems(items, round, msgElement) {
            if (msgElement && items.length === 0) { msgElement.innerText = '錯誤：未解析到有效資料'; return; }
            stagedImportData = items.map(item => ({ ...item, _importStatus: 'new' }));

            if (currentImportMode === 'word') {
                const stageRadio = document.querySelector('input[name="importStage"]:checked');
                const stageText = stageRadio && stageRadio.value === 'initial' ? '初次開立' : `第 ${round} 次審查`;
                const badgeClass = stageRadio && stageRadio.value === 'initial' ? 'new' : 'update';

                document.getElementById('previewModeBadge').innerHTML = `<span class="badge ${badgeClass}">Word 匯入 (${stageText})</span>`;
                document.getElementById('uploadCardWord').classList.add('hidden');
                document.getElementById('uploadCardBackup').classList.add('hidden');
            } else {
                document.getElementById('previewModeBadge').innerHTML = `<span class="badge active">⚠️ 災難復原模式</span>`;
                document.getElementById('uploadCardWord').classList.add('hidden');
                document.getElementById('uploadCardBackup').classList.add('hidden');
            }
            renderPreviewTable();
            document.getElementById('previewContainer').classList.remove('hidden');
            if (msgElement) msgElement.innerText = '';
        }

        function renderPreviewTable() {
            document.getElementById('previewCount').innerText = stagedImportData.length;
            const tbody = document.getElementById('previewBody');
            tbody.innerHTML = stagedImportData.map(item => {
                const statusBadge = item._importStatus === 'new' ? `<span class="badge new">新增</span>` : `<span class="badge update">更新</span>`;
                let progress = `[審查] ${item.review || '-'}<br>[辦理] ${item.handling || '-'}`;
                return `<tr>
                    <td>${statusBadge}</td>
                    <td style="font-weight:600;color:var(--primary);">${item.number}</td>
                    <td>${item.unit}</td>
                    <td><div class="preview-content-box">${stripHtml(item.content)}</div></td>
                    <td><div class="preview-content-box">${stripHtml(progress)}</div></td>
                </tr>`;
            }).join('');
        }

        function cancelImport() {
            stagedImportData = [];
            document.getElementById('previewContainer').classList.add('hidden');
            document.getElementById('uploadCardWord').classList.remove('hidden');
            if (currentUser && currentUser.role === 'admin') { document.getElementById('uploadCardBackup').classList.remove('hidden'); }
            document.getElementById('wordInput').value = '';
            document.getElementById('backupInput').value = '';
            document.getElementById('importStatusWord').innerText = '';
            document.getElementById('importStatusBackup').innerText = '';
        }

        async function confirmImport() {
            const count = stagedImportData.length;
            const isBackup = currentImportMode === 'backup';
            const msg = isBackup ? `⚠️ 警告：即將進行「災難復原」，這將覆蓋或新增 ${count} 筆資料。\n確定要執行嗎？` : `確定要匯入 ${count} 筆資料嗎？`;
            const confirmed = await showConfirmModal(msg, '確認', '取消');
            if (!confirmed) return;

            let round = 1;
            let issueDate = '';
            let replyDate = '';
            let responseDate = '';

            if (!isBackup) {
                const stage = document.querySelector('input[name="importStage"]:checked').value;

                if (stage === 'initial') {
                    round = 1;
                    issueDate = document.getElementById('importIssueDate').value;
                } else {
                    round = document.getElementById('importRoundSelect').value;
                    replyDate = document.getElementById('importReplyDate').value;
                    responseDate = document.getElementById('importResponseDate').value;
                }
            }

            const planValue = isBackup ? '' : document.getElementById('importPlanName').value;
            if (!isBackup && !planValue) {
                return showToast('請選擇檢查計畫', 'error');
            }
            // 從計畫選項值中提取計畫名稱和年度
            const selectedPlan = isBackup ? { name: '', year: '' } : parsePlanValue(planValue);
            
            // 取得所有計畫選項，用於根據年度匹配計畫
            let allPlans = [];
            if (!isBackup) {
                try {
                    const plansRes = await fetch('/api/options/plans?t=' + Date.now());
                    if (plansRes.ok) {
                        const plansJson = await plansRes.json();
                        allPlans = plansJson.data || [];
                        writeLog(`載入的計畫選項：${allPlans.length} 個`);
                        writeLog(`選擇的計畫：${selectedPlan.name} (${selectedPlan.year || '無年度'})`);
                    }
                } catch (e) {
                    if (isDevelopment) console.warn('無法載入計畫選項，將使用選擇的計畫名稱', e);
                    writeLog(`無法載入計畫選項：${e.message}`, 'WARN');
                }
            }

            let cleanData = stagedImportData.map(({ _importStatus, ...item }) => {
                if (currentImportMode === 'word') {
                    // 根據開立事項的年度，自動匹配到相同名稱但對應年度的計畫
                    // 例如：選擇「上半年定期檢查 (113)」，113年度的事項綁定到「上半年定期檢查 (113)」
                    // 114年度的事項應該綁定到「上半年定期檢查 (114)」（如果存在）
                    if (!item.planName && selectedPlan.name) {
                        const itemYear = String(item.year || '').trim();
                        
                        if (itemYear) {
                            // 查找相同名稱且年度匹配的計畫
                            const matchedPlan = allPlans.find(p => {
                                const planName = typeof p === 'object' ? String(p.name || '').trim() : String(p || '').trim();
                                const planYear = typeof p === 'object' ? String(p.year || '').trim() : '';
                                // 計畫名稱必須與選擇的計畫名稱相同，且年度必須與開立事項的年度匹配
                                return planName === selectedPlan.name && planYear === itemYear;
                            });
                            
                            if (matchedPlan) {
                                // 找到匹配的計畫，使用該計畫的名稱
                                item.planName = typeof matchedPlan === 'object' ? matchedPlan.name : matchedPlan;
                                const planName = typeof matchedPlan === 'object' ? matchedPlan.name : matchedPlan;
                                const planYear = typeof matchedPlan === 'object' ? matchedPlan.year : '';
                                writeLog(`找到匹配的計畫：事項年度=${itemYear}，計畫名稱="${planName}"，計畫年度="${planYear}"`);
                                writeLog(`使用匹配的計畫：${planName}`);
                            } else if (selectedPlan.year && selectedPlan.year === itemYear) {
                                // 選擇的計畫年度與事項年度匹配，使用選擇的計畫名稱
                                item.planName = selectedPlan.name;
                                writeLog(`使用選擇的計畫（年度匹配）：${selectedPlan.name}`);
                            } else {
                                // 沒找到匹配的計畫，且年度不匹配
                                // 使用選擇的計畫名稱（這會導致不同年度的事項被歸類到同一計畫）
                                item.planName = selectedPlan.name;
                                const warnMsg = `找不到匹配的計畫：選擇的計畫名稱="${selectedPlan.name}"，選擇的計畫年度="${selectedPlan.year}"，事項年度="${itemYear}"。使用選擇的計畫名稱。`;
                                if (isDevelopment) console.warn(`⚠️ ${warnMsg}`);
                                writeLog(warnMsg, 'WARN');
                            }
                        } else {
                            // 開立事項沒有年度，使用選擇的計畫名稱
                            item.planName = selectedPlan.name;
                        }
                    }
                    const stage = document.querySelector('input[name="importStage"]:checked') ? document.querySelector('input[name="importStage"]:checked').value : 'initial';
                    if (!item.issueDate && stage === 'initial') item.issueDate = issueDate;
                }
                return item;
            });

            try {
                const res = await apiFetch('/api/issues/import', {
                    method: 'POST',
                    body: JSON.stringify({
                        data: cleanData,
                        round: round,
                        reviewDate: responseDate,
                        replyDate: replyDate,
                        mode: currentImportMode
                    })
                });
                if (res.ok) { 
                    showToast('匯入成功！'); 
                    cancelImport(); 
                    // 使用 try-catch 包裹後續操作，避免影響成功訊息的顯示
                    try {
                        await loadIssuesPage(1); 
                        await loadPlanOptions(); 
                    } catch (e) {
                        console.error('載入資料時發生錯誤（匯入已成功）：', e);
                    }
                } else { 
                    const errorData = await res.json().catch(() => ({}));
                    showToast(errorData.error || '匯入失敗', 'error'); 
                }
            } catch (e) { 
                // 只有在真正的網路錯誤時才顯示
                if (e.message && (e.message.includes('Failed to fetch') || e.message.includes('NetworkError'))) {
                    showToast('匯入錯誤：網路連線失敗', 'error'); 
                } else {
                    console.error('匯入時發生未預期錯誤：', e);
                    showToast('匯入時發生錯誤：' + e.message, 'error'); 
                }
            }
        }

        function switchDataTab(tab) { 
            // 保存當前 tab 到 sessionStorage
            sessionStorage.setItem('currentDataTab', tab);
            
            document.querySelectorAll('#importView .admin-tab-btn').forEach(b => b.classList.remove('active')); 
            if (event && event.target) {
                event.target.classList.add('active');
            } else {
                // 如果沒有 event，找到對應的按鈕
                const buttons = document.querySelectorAll('#importView .admin-tab-btn');
                buttons.forEach(btn => {
                    if (btn.getAttribute('onclick') && btn.getAttribute('onclick').includes(`'${tab}'`)) {
                        btn.classList.add('active');
                    }
                });
            }
            
            // 主要 tab 切換
            document.getElementById('tab-data-issues').classList.toggle('hidden', tab !== 'issues'); 
            document.getElementById('tab-data-plans').classList.toggle('hidden', tab !== 'plans');
            document.getElementById('tab-data-export').classList.toggle('hidden', tab !== 'export');
            
            // 處理各 tab 的初始化
            if (tab === 'issues') {
                // 恢復開立事項子 tab
                const savedSubTab = sessionStorage.getItem('currentIssuesSubTab') || 'import';
                setTimeout(() => switchIssuesSubTab(savedSubTab), 100);
                // 確保檢查計畫選項已載入（用於資料管理頁面的其他功能）
                loadPlanOptions();
            }
            if (tab === 'plans') {
                // 恢復檢查計畫子 tab
                const savedSubTab = sessionStorage.getItem('currentPlansSubTab') || 'schedule';
                setTimeout(() => switchPlansSubTab(savedSubTab), 100);
                loadPlanOptions();
            }
            if (tab === 'export') {
                // 設置匯出選項的顯示/隱藏
                setTimeout(() => setupExportOptions(), 100);
            }
        }
        
        // 檢查計畫的子 tab 切換
        function switchPlansSubTab(subTab) {
            sessionStorage.setItem('currentPlansSubTab', subTab);
            
            document.querySelectorAll('#tab-data-plans .admin-tab-btn').forEach(b => b.classList.remove('active'));
            if (typeof event !== 'undefined' && event && event.target) {
                event.target.classList.add('active');
            } else {
                const buttons = document.querySelectorAll('#tab-data-plans .admin-tab-btn');
                buttons.forEach(btn => {
                    if (btn.getAttribute('onclick') && btn.getAttribute('onclick').includes(`'${subTab}'`)) {
                        btn.classList.add('active');
                    }
                });
            }
            
            document.getElementById('subtab-plans-manage').classList.toggle('hidden', subTab !== 'manage');
            document.getElementById('subtab-plans-schedule').classList.toggle('hidden', subTab !== 'schedule');
            
            if (subTab === 'manage') {
                restorePlansViewState();
                setTimeout(() => loadPlansPage(plansPage || 1), 200);
            }
            if (subTab === 'schedule') {
                initScheduleCalendar();
            }
        }
        
        // 開立事項的子 tab 切換
        function switchIssuesSubTab(subTab) {
            sessionStorage.setItem('currentIssuesSubTab', subTab);
            
            // 更新子 tab 按鈕狀態
            document.querySelectorAll('#tab-data-issues .admin-tab-btn').forEach(b => b.classList.remove('active'));
            if (event && event.target) {
                event.target.classList.add('active');
            } else {
                const buttons = document.querySelectorAll('#tab-data-issues .admin-tab-btn');
                buttons.forEach(btn => {
                    if (btn.getAttribute('onclick') && btn.getAttribute('onclick').includes(`'${subTab}'`)) {
                        btn.classList.add('active');
                    }
                });
            }
            
            // 切換子 tab 內容
            document.getElementById('subtab-issues-import').classList.toggle('hidden', subTab !== 'import');
            document.getElementById('subtab-issues-create').classList.toggle('hidden', subTab !== 'create');
            document.getElementById('subtab-issues-year-edit').classList.toggle('hidden', subTab !== 'year-edit');
            
            // 向後兼容：batch 和 manual 都指向 create
            if (subTab === 'batch' || subTab === 'manual') {
                document.getElementById('subtab-issues-create').classList.remove('hidden');
                if (subTab === 'batch') {
                    switchCreateMode('batch');
                } else {
                    switchCreateMode('single');
                }
            }
            
            if (subTab === 'create') {
                // 初始化開立事項建檔頁面
                createMode = 'batch'; // 固定為批次模式
                initCreateIssuePage();
                // 確保計畫選項已載入
                loadPlanOptions();
            }
            
            if (subTab === 'year-edit') {
                // 重置事項修正頁面
                yearEditIssue = null;
                yearEditIssueList = [];
                hideYearEditIssueContent();
                hideYearEditIssueList();
                document.getElementById('yearEditEmpty').style.display = 'block';
                document.getElementById('yearEditNotFound').style.display = 'none';
                // 載入有開立事項的檢查計畫選項到下拉選單
                setTimeout(() => {
                    loadYearEditPlanOptions();
                }, 100);
            }
        }
        
        function setupExportOptions() {
            const exportDataTypeRadios = document.querySelectorAll('input[name="exportDataType"]');
            const exportIssuesOptions = document.getElementById('exportIssuesOptions');
            
            if (exportDataTypeRadios.length > 0 && exportIssuesOptions) {
                exportDataTypeRadios.forEach(radio => {
                    // 移除舊的事件監聽器（如果有的話）
                    const newRadio = radio.cloneNode(true);
                    radio.parentNode.replaceChild(newRadio, radio);
                    
                    newRadio.addEventListener('change', function() {
                        if (this.value === 'plans') {
                            exportIssuesOptions.style.display = 'none';
                        } else {
                            exportIssuesOptions.style.display = 'block';
                        }
                    });
                });
                
                // 初始化顯示狀態
                const checked = document.querySelector('input[name="exportDataType"]:checked');
                if (checked && checked.value === 'plans') {
                    exportIssuesOptions.style.display = 'none';
                } else {
                    exportIssuesOptions.style.display = 'block';
                }
            }
        }

        // --- Batch Edit Logic ---
        function initBatchGrid() {
            const tbody = document.getElementById('batchGridBody');
            tbody.innerHTML = '';
            for (let i = 0; i < 5; i++) addBatchRow();
        }

        function addBatchRow() {
            const tbody = document.getElementById('batchGridBody');
            const rowIdx = tbody.children.length + 1;
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="text-align:center;color:#94a3b8;font-size:12px;">${rowIdx}</td>
                <td><input type="text" class="filter-input batch-number" placeholder="編號..." onchange="handleBatchNumberChange(this)" style="font-family:monospace;"></td>
                <td><textarea class="filter-input batch-content" rows="1" placeholder="內容..." style="resize:vertical;"></textarea></td>
                <td><input type="text" class="filter-input batch-year" style="background:#f1f5f9;color:#64748b;" readonly></td>
                <td><input type="text" class="filter-input batch-unit" style="background:#f1f5f9;color:#64748b;" readonly></td>
                <td><select class="filter-select batch-division"><option value="">-</option><option value="運務">運務</option><option value="工務">工務</option><option value="機務">機務</option><option value="電務">電務</option><option value="安全">安全</option><option value="審核">審核</option><option value="災防">災防</option><option value="運轉">運轉</option><option value="土木">土木</option><option value="機電">機電</option><option value="土建">土建</option><option value="安全管理">安全管理</option><option value="營運">營運</option><option value="其他">其他</option></select></td>
                <td><select class="filter-select batch-inspection"><option value="">-</option><option value="定期檢查">定期檢查</option><option value="例行性檢查">例行性檢查</option><option value="特別檢查">特別檢查</option><option value="臨時檢查">臨時檢查</option></select></td>
                <td><select class="filter-select batch-kind"><option value="">-</option><option value="N">缺失</option><option value="O">觀察</option><option value="R">建議</option></select></td>
                <td><select class="filter-select batch-status"><option value="持續列管">持續列管</option><option value="解除列管">解除列管</option><option value="自行列管">自行列管</option></select></td>
                <td style="text-align:center;"><button class="btn btn-danger btn-sm" onclick="removeBatchRow(this)" style="padding:4px 8px;">×</button></td>
            `;
            tbody.appendChild(tr);
        }

        function removeBatchRow(btn) {
            const tr = btn.closest('tr');
            if (document.querySelectorAll('#batchGridBody tr').length > 1) {
                tr.remove();
                // Re-index
                document.querySelectorAll('#batchGridBody tr').forEach((row, idx) => {
                    row.cells[0].innerText = idx + 1;
                });
            } else {
                showToast('至少需保留一列', 'error');
            }
        }

        function handleBatchNumberChange(input) {
            const tr = input.closest('tr');
            const val = input.value.trim();
            if (!val) return;

            const info = parseItemNumber(val);
            if (info) {
                if (info.yearRoc) tr.querySelector('.batch-year').value = info.yearRoc;
                if (info.orgCode) {
                    const name = ORG_MAP[info.orgCode] || info.orgCode;
                    if (name && name !== '?') tr.querySelector('.batch-unit').value = name;
                }
                if (info.divCode) {
                    const divName = DIVISION_MAP[info.divCode];
                    if (divName) tr.querySelector('.batch-division').value = divName;
                }
                if (info.inspectCode) {
                    const inspectName = INSPECTION_MAP[info.inspectCode];
                    if (inspectName) tr.querySelector('.batch-inspection').value = inspectName;
                }
                if (info.kindCode) {
                    tr.querySelector('.batch-kind').value = info.kindCode;
                }
            }
        }

        async function saveBatchItems() {
            const planValue = document.getElementById('batchPlanName').value.trim();
            const issueDate = document.getElementById('batchIssueDate').value.trim();
            const batchYear = document.getElementById('batchYear') ? document.getElementById('batchYear').value.trim() : '';

            if (!planValue) return showToast('請選擇檢查計畫', 'error');
            // 從計畫選項值中提取計畫名稱
            const planName = parsePlanValue(planValue).name;
            if (!issueDate) return showToast('請填寫初次發函日期', 'error');

            const rows = document.querySelectorAll('#batchGridBody tr');
            const items = [];
            let hasError = false;

            rows.forEach((tr, idx) => {
                const number = tr.querySelector('.batch-number').value.trim();
                const content = tr.querySelector('.batch-content').value.trim();

                // Skip empty rows
                if (!number && !content) return;

                if (!number) {
                    showToast(`第 ${idx + 1} 列缺少編號`, 'error');
                    hasError = true;
                    return;
                }

                const year = tr.querySelector('.batch-year').value.trim();
                const unit = tr.querySelector('.batch-unit').value.trim();

                if (!year || !unit) {
                    showToast(`第 ${idx + 1} 列的年度或機構未能自動判別，請確認編號格式`, 'error');
                    hasError = true;
                    return;
                }

                items.push({
                    number,
                    year,
                    unit,
                    content,
                    status: tr.querySelector('.batch-status').value,
                    itemKindCode: tr.querySelector('.batch-kind').value,
                    divisionName: tr.querySelector('.batch-division').value,
                    inspectionCategoryName: tr.querySelector('.batch-inspection').value,
                    planName: planName,
                    issueDate: issueDate,
                    scheme: 'BATCH'
                });
            });

            if (hasError) return;
            if (items.length === 0) return showToast('請至少輸入一筆有效資料', 'error');

            const confirmed = await showConfirmModal(`確定要批次新增 ${items.length} 筆資料嗎？\n\n計畫：${planName}`, '確定新增', '取消');
            if (!confirmed) return;

            try {
                const res = await apiFetch('/api/issues/import', {
                    method: 'POST',
                    body: JSON.stringify({
                        data: items,
                        round: 1,
                        reviewDate: '',
                        replyDate: ''
                    })
                });

                if (res.ok) {
                    showToast('批次新增成功！');
                    initBatchGrid(); // Reset grid
                    document.getElementById('batchPlanName').value = '';
                    document.getElementById('batchIssueDate').value = '';
                    loadIssuesPage(1);
                    loadPlanOptions();
                } else {
                    const j = await res.json();
                    showToast('新增失敗: ' + (j.error || '不明錯誤'), 'error');
                }
            } catch (e) {
                showToast('Error: ' + e.message, 'error');
            }
        }

        // --- 開立事項建檔功能（已移除單筆模式，只保留批次模式） ---
        let createMode = 'batch'; // 固定為批次模式
        
        // 初始化開立事項建檔頁面
        function initCreateIssuePage() {
            const batchMode = document.getElementById('createBatchMode');
            if (batchMode) {
                batchMode.style.display = 'block';
            }
            
            if (document.querySelectorAll('#createBatchGridBody tr').length === 0) {
                initCreateBatchGrid();
            }
            
            // 初始化批次設定函復日期的選項
            initBatchResponseRoundOptions();
            // 初始化批次設定回復日期的選項
            initBatchReplyRoundOptions();
            
            // 顯示載入現有事項按鈕（如果已選擇計畫）
            const planSelect = document.getElementById('createPlanName');
            const loadContainer = document.getElementById('createLoadExistingContainer');
            if (loadContainer && planSelect && planSelect.value) {
                loadContainer.style.display = 'block';
            }
            
            // 重置批次設定函復日期的勾選狀態
            const toggleCheckbox = document.getElementById('createBatchResponseDateToggle');
            if (toggleCheckbox) {
                toggleCheckbox.checked = false;
                toggleBatchResponseDateSetting();
            }
        }
        
        // 保留 switchCreateMode 函數以向後兼容，但只執行批次模式的邏輯
        function switchCreateMode(mode) {
            createMode = 'batch'; // 強制為批次模式
            initCreateIssuePage();
        }
        
        // 切換批次設定函復日期的顯示
        function toggleBatchResponseDateSetting() {
            const checkbox = document.getElementById('createBatchResponseDateToggle');
            const container = document.getElementById('createBatchResponseDateContainer');
            if (checkbox && container) {
                container.style.display = checkbox.checked ? 'block' : 'none';
            }
        }
        
        // 切換批次設定回復日期的顯示
        function toggleBatchReplyDateSetting() {
            const checkbox = document.getElementById('createBatchReplyDateToggle');
            const container = document.getElementById('createBatchReplyDateContainer');
            if (checkbox && container) {
                container.style.display = checkbox.checked ? 'block' : 'none';
            }
        }
        
        // 批次設定回復日期（為所有事項的辦理情形）- 比照審查函覆日期的處理流程
        async function batchSetReplyDateForAll() {
            const roundSelect = document.getElementById('createBatchReplyRound');
            const roundManualInput = document.getElementById('createBatchReplyRoundManual');
            const dateInput = document.getElementById('createBatchReplyDate');
            const planSelect = document.getElementById('createPlanName');
            
            if (!roundSelect || !roundManualInput || !dateInput || !planSelect) return;
            
            // 優先使用下拉選單的值，如果沒有則使用手動輸入
            let round = parseInt(roundSelect.value);
            if (!round || round < 1) {
                round = parseInt(roundManualInput.value);
            }
            
            const replyDate = dateInput.value.trim();
            const planValue = planSelect.value.trim();
            
            if (!planValue) {
                showToast('請先選擇檢查計畫', 'error');
                return;
            }
            
            if (!round || round < 1) {
                showToast('請選擇或輸入回復輪次', 'error');
                return;
            }
            
            if (round > 200) {
                showToast('回復輪次不能超過200次', 'error');
                return;
            }
            
            if (!replyDate) {
                showToast('請輸入回復日期', 'error');
                return;
            }
            
            // 驗證日期格式
            if (!validateDateFormat(replyDate, '日期')) {
                return;
            }
            
            const { name: planName } = parsePlanValue(planValue);
            
            try {
                // 載入該計畫下的所有事項
                const issueList = await loadIssuesByPlan(planValue);
                if (!issueList) return;
                
                const confirmed = await showConfirmModal(
                    `確定要批次設定第 ${round} 次辦理情形的回復日期為 ${replyDate} 嗎？\n\n將更新 ${issueList.length} 筆事項。`,
                    '確認設定',
                    '取消'
                );
                
                if (!confirmed) {
                    return;
                }
                
                // 移除批次設定中的提示訊息，只保留錯誤訊息
                
                let successCount = 0;
                let errorCount = 0;
                const errors = [];
                
                // 批次更新所有事項
                for (let i = 0; i < issueList.length; i++) {
                    const issue = issueList[i];
                    const issueId = issue.id;
                    
                    if (!issueId) {
                        errorCount++;
                        errors.push(`${issue.number || '未知編號'}: 缺少事項ID`);
                        continue;
                    }
                    
                    try {
                        // 讀取該輪次的現有資料
                        const suffix = round === 1 ? '' : round;
                        const handling = issue['handling' + suffix] || '';
                        const review = issue['review' + suffix] || '';
                        const existingReplyDate = issue['reply_date_r' + round] || '';
                        
                        // 檢查是否有辦理情形內容，沒有辦理情形內容則跳過
                        if (!handling || !handling.trim()) {
                            errorCount++;
                            errors.push(`${issue.number || '未知編號'}: 第 ${round} 次尚無辦理情形，無法設定回復日期`);
                            continue;
                        }
                        
                        // 更新該輪次的回復日期
                        // 注意：只更新 replyDate（辦理情形回復日期），不更新 responseDate（審查函復日期）
                        const updateRes = await apiFetch(`/api/issues/${issueId}`, {
                            method: 'PUT',
                            body: JSON.stringify({
                                status: issue.status || '持續列管',
                                round: round,
                                handling: handling,
                                review: review,
                                // 只發送 replyDate，不發送 responseDate，讓後端保持原有的審查函復日期不變
                                replyDate: replyDate
                            })
                        });
                        
                        if (updateRes.ok) {
                            const result = await updateRes.json();
                            if (result.success) {
                                successCount++;
                            } else {
                                errorCount++;
                                errors.push(`${issue.number || '未知編號'}: 更新失敗`);
                            }
                        } else {
                            errorCount++;
                            const errorData = await updateRes.json().catch(() => ({}));
                            errors.push(`${issue.number || '未知編號'}: ${errorData.error || '更新失敗'}`);
                        }
                    } catch (e) {
                        errorCount++;
                        errors.push(`${issue.number || '未知編號'}: ${e.message}`);
                    }
                }
                
                // 顯示資料庫操作結果（成功或警告）
                if (errorCount > 0) {
                    showToast(`批次設定完成，但有 ${errorCount} 筆失敗${successCount > 0 ? `，成功 ${successCount} 筆` : ''}`, 'warning');
                    
                    // 如果有錯誤，顯示詳細資訊
                    if (errors.length > 0) {
                        console.error('批次設定回復日期錯誤:', errors);
                    }
                } else if (successCount > 0) {
                    // 完全成功時顯示成功訊息（資料庫操作結果）
                    showToast(`批次設定完成！成功 ${successCount} 筆`, 'success');
                }
                
                // 清空輸入欄位並重置為預設模式
                if (successCount > 0 || errorCount === 0) {
                    roundSelect.value = '';
                    roundManualInput.value = '';
                    dateInput.value = '';
                    
                    // 取消勾選並隱藏設定區塊
                    const toggleCheckbox = document.getElementById('createBatchReplyDateToggle');
                    if (toggleCheckbox) {
                        toggleCheckbox.checked = false;
                        toggleBatchReplyDateSetting();
                    }
                } else {
                    showToast('批次設定失敗，所有事項都無法更新', 'error');
                    if (errors.length > 0) {
                        console.error('批次設定回復日期錯誤:', errors);
                    }
                }
            } catch (e) {
                showToast('批次設定失敗: ' + e.message, 'error');
            }
        }
        
        // 回復日期輪次選擇改變時的處理
        function onBatchReplyRoundChange() {
            const roundSelect = document.getElementById('createBatchReplyRound');
            const roundManualInput = document.getElementById('createBatchReplyRoundManual');
            
            if (!roundSelect || !roundManualInput) return;
            
            if (roundSelect.value) {
                roundManualInput.value = '';
            }
        }
        
        // 回復日期輪次手動輸入改變時的處理
        function onBatchReplyRoundManualChange() {
            const roundSelect = document.getElementById('createBatchReplyRound');
            const roundManualInput = document.getElementById('createBatchReplyRoundManual');
            
            if (!roundSelect || !roundManualInput) return;
            
            if (roundManualInput.value) {
                const manualValue = parseInt(roundManualInput.value);
                if (manualValue >= 1 && manualValue <= 200) {
                    // 如果在選單範圍內，同步到選單
                    roundSelect.value = manualValue;
                } else {
                    // 如果超出範圍，清空選單
                    roundSelect.value = '';
                }
            }
        }
        
        // 從檢查計畫查詢並預填辦理情形回復輪次
        async function updateBatchReplyRoundFromPlan() {
            const planSelect = document.getElementById('createPlanName');
            const roundSelect = document.getElementById('createBatchReplyRound');
            const roundManualInput = document.getElementById('createBatchReplyRoundManual');
            
            if (!planSelect || !roundSelect || !roundManualInput) return;
            
            const planValue = planSelect.value.trim();
            if (!planValue) {
                // 清空選項
                roundSelect.value = '';
                roundManualInput.value = '';
                return;
            }
            
            try {
                // 載入該計畫下的所有事項
                const issueList = await loadIssuesByPlan(planValue, { showError: false, returnEmpty: true });
                if (!issueList || issueList.length === 0) {
                    // 沒有事項，預設為第1次
                    roundSelect.value = '1';
                    roundManualInput.value = '';
                    return;
                }
                
                // 找出第一個「有辦理情形內容但沒有回復日期」的輪次
                // 如果所有輪次都有日期，則找下一個需要填寫的輪次
                let foundIncompleteRound = null;
                let maxRound = 0;
                
                issueList.forEach(issue => {
                    // 檢查所有可能的辦理情形輪次（最多200次）
                    for (let i = 1; i <= 200; i++) {
                        const suffix = i === 1 ? '' : i;
                        const handling = issue['handling' + suffix] || '';
                        const replyDate = issue['reply_date_r' + i] || '';
                        
                        // 如果有辦理情形，記錄最高輪次
                        if (handling.trim()) {
                            if (i > maxRound) {
                                maxRound = i;
                            }
                            
                            // 如果有辦理情形內容但沒有回復日期，這是需要填寫的輪次
                            if (handling.trim() && !replyDate) {
                                if (!foundIncompleteRound || i < foundIncompleteRound) {
                                    foundIncompleteRound = i;
                                }
                            }
                        }
                    }
                });
                
                // 如果找到有辦理情形內容但無日期的輪次，使用該輪次
                // 否則使用最高輪次 + 1（如果最高輪次是0，則為第1次）
                const suggestedRound = foundIncompleteRound || (maxRound + 1);
                
                if (suggestedRound <= 200) {
                    roundSelect.value = suggestedRound;
                    roundManualInput.value = '';
                    // 移除自動預填的提示訊息，只保留錯誤訊息
                } else {
                    // 如果超過200次，使用手動輸入
                    roundSelect.value = '';
                    roundManualInput.value = suggestedRound;
                    // 移除自動預填的提示訊息，只保留錯誤訊息
                }
            } catch (e) {
                console.error('查詢辦理情形輪次失敗:', e);
            }
        }
        
        // 初始化批次設定函復日期的選項（動態生成，最多200次）
        function initBatchResponseRoundOptions() {
            const select = document.getElementById('createBatchResponseRound');
            if (!select) return;
            
            // 清空現有選項（保留第一個「請選擇」選項）
            select.innerHTML = '<option value="">請選擇</option>';
            
            // 動態生成選項（最多200次）
            for (let i = 1; i <= 200; i++) {
                const option = document.createElement('option');
                option.value = i;
                option.textContent = `第 ${i} 次`;
                select.appendChild(option);
            }
        }
        
        // 初始化批次設定回復日期的選項（動態生成，最多200次）
        function initBatchReplyRoundOptions() {
            const select = document.getElementById('createBatchReplyRound');
            if (!select) return;
            
            // 清空現有選項（保留第一個「請選擇」選項）
            select.innerHTML = '<option value="">請選擇</option>';
            
            // 動態生成選項（最多200次）
            for (let i = 1; i <= 200; i++) {
                const option = document.createElement('option');
                option.value = i;
                option.textContent = `第 ${i} 次`;
                select.appendChild(option);
            }
        }
        
        // 檢查計畫改變時（年度已包含在計畫中，無需額外處理）
        function onCreatePlanChange() {
            // 當選擇計畫時，自動帶入計畫的年度
            const planValue = document.getElementById('createPlanName').value.trim();
            if (planValue) {
                const { name, year } = parsePlanValue(planValue);
                if (year) {
                    const yearDisplay = document.getElementById('createYearDisplay');
                    if (yearDisplay) {
                        const oldYear = yearDisplay.value;
                        yearDisplay.value = year;
                        // 移除年度變更的提示訊息，只保留錯誤訊息
                    }
                }
            }
            
            // 查詢並預填審查函復輪次和辦理情形回復輪次
            updateBatchResponseRoundFromPlan();
            updateBatchReplyRoundFromPlan();
            // 顯示/隱藏載入現有事項按鈕
            const loadContainer = document.getElementById('createLoadExistingContainer');
            if (loadContainer) {
                loadContainer.style.display = planValue ? 'block' : 'none';
            }
        }
        
        // 載入現有事項到批次表格
        async function loadExistingIssuesToBatch() {
            const planSelect = document.getElementById('createPlanName');
            if (!planSelect) return;
            
            const planValue = planSelect.value.trim();
            if (!planValue) {
                showToast('請先選擇檢查計畫', 'error');
                return;
            }
            
            try {
                // 載入該計畫下的所有事項
                const issueList = await loadIssuesByPlan(planValue, { showError: false, returnEmpty: true });
                if (!issueList || issueList.length === 0) {
                    return;
                }
                
                // 確認是否要載入（如果表格中已有資料）
                const tbody = document.getElementById('createBatchGridBody');
                if (!tbody) return;
                
                const existingRows = tbody.querySelectorAll('tr');
                const hasExistingData = Array.from(existingRows).some(tr => {
                    const number = tr.querySelector('.create-batch-number')?.value.trim();
                    const contentTextarea = tr.querySelector('.create-batch-content-textarea');
                    const content = contentTextarea ? contentTextarea.value.trim() : '';
                    return number || content;
                });
                
                if (hasExistingData) {
                    const confirmed = await showConfirmModal(`表格中已有資料，載入現有事項將會清空現有資料。\n\n確定要載入 ${issueList.length} 筆事項嗎？`, '確定載入', '取消');
                    if (!confirmed) {
                        return;
                    }
                }
                
                // 清空現有表格
                tbody.innerHTML = '';
                batchHandlingData = {};
                
                // 載入事項資料到表格
                issueList.forEach((issue, index) => {
                    const rowIdx = index;
                    const tr = document.createElement('tr');
                    
                    // 取得類型代碼
                    let kindCode = issue.item_kind_code || issue.itemKindCode || '';
                    if (!kindCode) {
                        const numStr = String(issue.number || '');
                        const m = numStr.match(/-([NOR])\d+$/i);
                        if (m) kindCode = m[1].toUpperCase();
                    }
                    
                    // 取得分組名稱
                    const divisionName = issue.division_name || issue.divisionName || '';
                    
                    // 取得檢查種類
                    const inspectionName = issue.inspection_category_name || issue.inspectionCategoryName || '';
                    
                    // 取得狀態
                    const status = issue.status || '持續列管';
                    
                    tr.innerHTML = `
                        <td style="text-align:center;color:#94a3b8;font-size:12px;">${rowIdx + 1}</td>
                        <td><input type="text" class="filter-input create-batch-number" value="${escapeHtml(issue.number || '')}" onchange="handleCreateBatchNumberChange(this)" style="font-family:monospace;"></td>
                        <td style="position:relative;">
                            <textarea class="filter-input create-batch-content-textarea" rows="3" style="resize:vertical;min-height:60px;max-height:120px;font-size:13px;line-height:1.6;padding:8px 10px;">${escapeHtml(issue.content || '')}</textarea>
                        </td>
                        <td><input type="text" class="filter-input create-batch-year" value="${escapeHtml(issue.year || '')}" style="background:#f1f5f9;color:#64748b;" readonly></td>
                        <td><input type="text" class="filter-input create-batch-unit" value="${escapeHtml(issue.unit || '')}" style="background:#f1f5f9;color:#64748b;" readonly></td>
                        <td><select class="filter-select create-batch-division"><option value="">-</option><option value="運務" ${divisionName === '運務' ? 'selected' : ''}>運務</option><option value="工務" ${divisionName === '工務' ? 'selected' : ''}>工務</option><option value="機務" ${divisionName === '機務' ? 'selected' : ''}>機務</option><option value="電務" ${divisionName === '電務' ? 'selected' : ''}>電務</option><option value="安全" ${divisionName === '安全' ? 'selected' : ''}>安全</option><option value="審核" ${divisionName === '審核' ? 'selected' : ''}>審核</option><option value="災防" ${divisionName === '災防' ? 'selected' : ''}>災防</option><option value="運轉" ${divisionName === '運轉' ? 'selected' : ''}>運轉</option><option value="土木" ${divisionName === '土木' ? 'selected' : ''}>土木</option><option value="機電" ${divisionName === '機電' ? 'selected' : ''}>機電</option><option value="土建" ${divisionName === '土建' ? 'selected' : ''}>土建</option><option value="安全管理" ${divisionName === '安全管理' ? 'selected' : ''}>安全管理</option><option value="營運" ${divisionName === '營運' ? 'selected' : ''}>營運</option><option value="其他" ${divisionName === '其他' ? 'selected' : ''}>其他</option></select></td>
                        <td><select class="filter-select create-batch-inspection"><option value="">-</option><option value="定期檢查" ${inspectionName === '定期檢查' ? 'selected' : ''}>定期檢查</option><option value="例行性檢查" ${inspectionName === '例行性檢查' ? 'selected' : ''}>例行性檢查</option><option value="特別檢查" ${inspectionName === '特別檢查' ? 'selected' : ''}>特別檢查</option><option value="臨時檢查" ${inspectionName === '臨時檢查' ? 'selected' : ''}>臨時檢查</option></select></td>
                        <td><select class="filter-select create-batch-kind"><option value="">-</option><option value="N" ${kindCode === 'N' ? 'selected' : ''}>缺失</option><option value="O" ${kindCode === 'O' ? 'selected' : ''}>觀察</option><option value="R" ${kindCode === 'R' ? 'selected' : ''}>建議</option></select></td>
                        <td><select class="filter-select create-batch-status"><option value="持續列管" ${status === '持續列管' ? 'selected' : ''}>持續列管</option><option value="解除列管" ${status === '解除列管' ? 'selected' : ''}>解除列管</option><option value="自行列管" ${status === '自行列管' ? 'selected' : ''}>自行列管</option></select></td>
                        <td style="text-align:center;">
                            <button class="btn btn-outline btn-sm create-batch-handling-btn" onclick="openBatchHandlingModal(${rowIdx})" data-row-index="${rowIdx}" style="padding:6px 12px; font-size:12px; width:100%;" title="點擊新增或管理辦理情形">
                                <span class="create-batch-handling-status">新增辦理情形</span>
                            </button>
                        </td>
                        <td style="text-align:center;">
                            <button class="btn btn-danger btn-sm" onclick="removeCreateBatchRow(this)" style="padding:4px 8px;">×</button>
                        </td>
                    `;
                    tbody.appendChild(tr);
                    
                    // 保存事項 ID 到表格行（如果事項已存在於資料庫）
                    if (issue.id) {
                        tr.setAttribute('data-issue-id', issue.id);
                    }
                    
                    // 載入現有事項時，內容已經在 textarea 中，不需要額外處理
                    
                    // 載入現有的辦理情形資料（如果有）
                    const handlingRounds = [];
                    for (let i = 1; i <= 200; i++) {
                        const suffix = i === 1 ? '' : i;
                        const handling = issue['handling' + suffix] || '';
                        const replyDate = issue['reply_date_r' + i] || '';
                        
                        if (handling && handling.trim()) {
                            handlingRounds.push({
                                round: i,
                                handling: stripHtml(handling.trim()), // 移除 HTML 標籤
                                replyDate: replyDate || ''
                            });
                        }
                    }
                    
                    if (handlingRounds.length > 0) {
                        batchHandlingData[rowIdx] = handlingRounds;
                        updateBatchHandlingStatus(rowIdx);
                    } else {
                        updateBatchHandlingStatus(rowIdx);
                    }
                });
                
                // 移除成功消息，只保留錯誤消息
            } catch (e) {
                showToast('載入事項失敗: ' + e.message, 'error');
            }
        }
        
        // HTML 轉義函數（防止 XSS）
        function escapeHtml(text) {
            if (!text) return '';
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }
        
        // 從檢查計畫查詢並預填審查函復輪次
        async function updateBatchResponseRoundFromPlan() {
            const planSelect = document.getElementById('createPlanName');
            const roundSelect = document.getElementById('createBatchResponseRound');
            const roundManualInput = document.getElementById('createBatchResponseRoundManual');
            
            if (!planSelect || !roundSelect || !roundManualInput) return;
            
            const planValue = planSelect.value.trim();
            if (!planValue) {
                // 清空選項
                roundSelect.value = '';
                roundManualInput.value = '';
                return;
            }
            
            try {
                // 載入該計畫下的所有事項
                const issueList = await loadIssuesByPlan(planValue, { showError: false, returnEmpty: true });
                if (!issueList || issueList.length === 0) {
                    // 沒有事項，預設為第1次
                    roundSelect.value = '1';
                    roundManualInput.value = '';
                    return;
                }
                
                // 找出第一個「有審查內容但沒有函復日期」的輪次
                // 如果所有輪次都有日期，則找下一個需要填寫的輪次
                let foundIncompleteRound = null;
                let maxRound = 0;
                
                issueList.forEach(issue => {
                    // 檢查所有可能的審查輪次（最多200次）
                    for (let i = 1; i <= 200; i++) {
                        const suffix = i === 1 ? '' : i;
                        const review = issue['review' + suffix] || '';
                        const responseDate = issue['response_date_r' + i] || '';
                        
                        // 如果有審查意見，記錄最高輪次
                        if (review.trim()) {
                            if (i > maxRound) {
                                maxRound = i;
                            }
                            
                            // 如果有審查內容但沒有函復日期，這是需要填寫的輪次
                            if (review.trim() && !responseDate) {
                                if (!foundIncompleteRound || i < foundIncompleteRound) {
                                    foundIncompleteRound = i;
                                }
                            }
                        }
                    }
                });
                
                // 如果找到有審查內容但無日期的輪次，使用該輪次
                // 否則使用最高輪次 + 1（如果最高輪次是0，則為第1次）
                const suggestedRound = foundIncompleteRound || (maxRound + 1);
                
                if (suggestedRound <= 200) {
                    roundSelect.value = suggestedRound;
                    roundManualInput.value = '';
                    // 移除自動預填的提示訊息，只保留錯誤訊息
                } else {
                    // 如果超過200次，使用手動輸入
                    roundSelect.value = '';
                    roundManualInput.value = suggestedRound;
                    // 移除自動預填的提示訊息，只保留錯誤訊息
                }
            } catch (e) {
                console.error('查詢審查輪次失敗:', e);
            }
        }
        
        // 當下拉選單改變時，同步到手動輸入欄位
        function onBatchResponseRoundChange() {
            const roundSelect = document.getElementById('createBatchResponseRound');
            const roundManualInput = document.getElementById('createBatchResponseRoundManual');
            
            if (!roundSelect || !roundManualInput) return;
            
            if (roundSelect.value) {
                roundManualInput.value = '';
            }
        }
        
        // 當手動輸入改變時，同步到下拉選單
        function onBatchResponseRoundManualChange() {
            const roundSelect = document.getElementById('createBatchResponseRound');
            const roundManualInput = document.getElementById('createBatchResponseRoundManual');
            
            if (!roundSelect || !roundManualInput) return;
            
            if (roundManualInput.value) {
                const manualValue = parseInt(roundManualInput.value);
                if (manualValue >= 1 && manualValue <= 200) {
                    // 如果在選單範圍內，同步到選單
                    roundSelect.value = manualValue;
                } else {
                    // 如果超過範圍，清空選單
                    roundSelect.value = '';
                }
            } else {
                // 如果手動輸入為空，不清空選單（保留選單選擇）
            }
        }
        
        // 批次模式：當選擇計畫時，更新所有行的年度（不管是否有編號）
        function handleCreateBatchPlanChange() {
            const planValue = document.getElementById('createPlanName')?.value.trim();
            if (!planValue) return;
            
            const { year: planYear } = parsePlanValue(planValue);
            if (!planYear) return;
            
            // 更新所有行的年度為計畫的年度（不管是否有編號）
            const rows = document.querySelectorAll('#createBatchGridBody tr');
            let updatedCount = 0;
            rows.forEach(tr => {
                const yearInput = tr.querySelector('.create-batch-year');
                if (yearInput) {
                    yearInput.value = planYear;
                    updatedCount++;
                }
            });
            
            // 移除年度同步更新的提示訊息，只保留錯誤訊息
        }
        
        // 從編號自動填入欄位（單筆模式）
        function autoFillFromNumberCreate() {
            const val = document.getElementById('createNumber').value;
            const info = parseItemNumber(val);
            if (info) {
                if (info.yearRoc) {
                    const yearDisplay = document.getElementById('createYearDisplay');
                    if (yearDisplay) yearDisplay.value = info.yearRoc;
                }
                if (info.orgCode) {
                    const name = ORG_MAP[info.orgCode] || info.orgCode;
                    if (name && name !== '?') document.getElementById('createUnit').value = name;
                }
                if (info.divCode) {
                    const divName = DIVISION_MAP[info.divCode];
                    if (divName) document.getElementById('createDivision').value = divName;
                }
                if (info.inspectCode) {
                    const inspectName = INSPECTION_MAP[info.inspectCode];
                    if (inspectName) document.getElementById('createInspection').value = inspectName;
                }
                if (info.kindCode) {
                    document.getElementById('createKind').value = info.kindCode;
                }
            }
        }
        
        // 向後兼容：保留舊函數名稱
        function autoFillFromNumber() {
            autoFillFromNumberCreate();
        }

        // 辦理情形輪次管理（用於新增事項）
        let createHandlingRounds = []; // 儲存辦理情形輪次資料
        
        // 初始化辦理情形輪次（可選，預設為空，用戶可以選擇新增）
        function initCreateHandlingRounds() {
            createHandlingRounds = [];
            // 不再預設新增第一次辦理情形，讓用戶可以選擇是否要新增
            renderCreateHandlingRounds();
        }
        
        // 新增辦理情形輪次
        function addCreateHandlingRound() {
            const round = createHandlingRounds.length + 1;
                createHandlingRounds.push({
                round: round,
                handling: '',
                replyDate: ''
            });
            renderCreateHandlingRounds();
        }
        
        // 移除辦理情形輪次
        function removeCreateHandlingRound(index) {
            createHandlingRounds.splice(index, 1);
            // 重新編號
            createHandlingRounds.forEach((r, i) => {
                r.round = i + 1;
            });
            renderCreateHandlingRounds();
        }
        
        // 渲染辦理情形輪次
        function renderCreateHandlingRounds() {
            const container = document.getElementById('createHandlingRoundsContainer');
            if (!container) return;
            
            if (createHandlingRounds.length === 0) {
                container.innerHTML = '';
                return;
            }
            
            let html = '';
            createHandlingRounds.forEach((roundData, index) => {
                const isFirst = index === 0;
                html += `
                    <div class="create-handling-round" data-index="${index}" style="background:white; padding:16px; border-radius:8px; border:${isFirst ? '2px solid #10b981' : '1px solid #e2e8f0'}; margin-bottom:12px; ${isFirst ? 'border-left:4px solid #10b981;' : ''}">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                            <div style="font-weight:700; color:${isFirst ? '#047857' : '#334155'}; font-size:14px;">
                                第 ${roundData.round} 次機構辦理情形 ${isFirst ? '<span style="color:#64748b; font-size:12px;">(選填)</span>' : ''}
                            </div>
                            ${!isFirst ? `<button type="button" class="btn btn-danger btn-sm" onclick="removeCreateHandlingRound(${index})" style="padding:4px 12px; font-size:12px;">刪除</button>` : ''}
                        </div>
                        <div style="margin-bottom:12px;">
                            <label style="display:block; font-weight:600; color:#475569; font-size:13px; margin-bottom:6px;">
                                辦理情形
                            </label>
                            <textarea class="filter-input create-handling-text" data-index="${index}" 
                                placeholder="請輸入機構辦理情形..." 
                                style="width:100%; min-height:120px; padding:12px; font-size:14px; line-height:1.6; resize:vertical; background:white;"
                                oninput="updateCreateHandlingRound(${index}, 'handling', this.value)">${roundData.handling}</textarea>
                        </div>
                        <div>
                            <label style="display:block; font-weight:600; color:#475569; font-size:12px; margin-bottom:6px;">鐵路機構回復日期</label>
                            <input type="text" class="filter-input create-handling-reply-date" data-index="${index}" 
                                value="${roundData.replyDate}" placeholder="例如: 1130601" 
                                style="width:100%; background:white;"
                                oninput="updateCreateHandlingRound(${index}, 'replyDate', this.value)">
                        </div>
                    </div>
                `;
            });
            container.innerHTML = html;
        }
        
        // 更新辦理情形輪次資料
        function updateCreateHandlingRound(index, field, value) {
            if (createHandlingRounds[index]) {
                createHandlingRounds[index][field] = value;
            }
        }
        
        // 單筆新增事項
        async function submitCreateIssue() {
            const number = document.getElementById('createNumber').value.trim();
            const yearDisplay = document.getElementById('createYearDisplay');
            let year = yearDisplay ? yearDisplay.value.trim() : '';
            const unit = document.getElementById('createUnit').value.trim();
            const division = document.getElementById('createDivision').value;
            const inspection = document.getElementById('createInspection').value;
            const kind = document.getElementById('createKind').value;

            const planValue = document.getElementById('createPlanName').value.trim();
            const issueDate = document.getElementById('createIssueDate').value.trim();
            const continuousMode = document.getElementById('createContinuousMode').checked;

            const status = document.getElementById('createStatus').value;
            const content = document.getElementById('createContent').value.trim();
            
            if (!number || !unit || !content) return showToast('請填寫所有必填欄位', 'error');
            if (!planValue) return showToast('請選擇檢查計畫', 'error');
            if (!issueDate) return showToast('請填寫初次發函日期', 'error');
            
            // 從計畫選項值中提取計畫名稱和年度
            const { name: planName, year: planYear } = parsePlanValue(planValue);
            
            // 優先使用計畫的年度，如果計畫沒有年度才使用從編號解析出來的年度
            if (planYear) {
                year = planYear;
                // 更新顯示欄位
                if (yearDisplay) {
                    yearDisplay.value = year;
                }
            }
            
            // 如果還是沒有年度，嘗試從編號解析
            if (!year) {
                const info = parseItemNumber(number);
                if (info && info.yearRoc) {
                    year = info.yearRoc;
                    if (yearDisplay) {
                        yearDisplay.value = year;
                    }
                }
            }
            
            if (!year) return showToast('無法確定年度，請確認編號格式或選擇有年度的檢查計畫', 'error');
            
            // 辦理情形為選填，可以稍後再新增
            // 如果有辦理情形，使用第一個；如果沒有，使用空值
            const firstHandling = createHandlingRounds.length > 0 && createHandlingRounds[0].handling.trim() 
                ? createHandlingRounds[0] 
                : { handling: '', replyDate: '' };
            const payload = {
                data: [{
                    number, year, unit, content, status,
                    itemKindCode: kind,
                    divisionName: division,
                    inspectionCategoryName: inspection,
                    planName: planName,
                    issueDate: issueDate,
                    handling: firstHandling.handling ? firstHandling.handling.trim() : '',
                    scheme: 'MANUAL'
                }],
                round: 1, 
                reviewDate: '', 
                replyDate: firstHandling.replyDate ? firstHandling.replyDate.trim() : ''
            };

            try {
                const res = await apiFetch('/api/issues/import', { 
                    method: 'POST', 
                    body: JSON.stringify(payload) 
                });
                
                // 先檢查HTTP狀態碼
                if (!res.ok) {
                    const errorData = await res.json().catch(() => ({}));
                    // 檢查是否有編號重複的錯誤
                    if (res.status === 400 && errorData.error === '編號重複') {
                        showToast(`編號 "${number}" 已存在且內容不同，無法新增。請使用不同的編號或修改現有事項。`, 'error');
                        // 不清理表單，讓用戶可以修改編號
                        return;
                    }
                    showToast('新增失敗: ' + (errorData.error || res.statusText), 'error');
                    return;
                }
                
                const result = await res.json();
                
                // 確認是新增成功（newCount > 0）或更新成功（updateCount > 0）
                if (result.newCount > 0 || result.updateCount > 0) {
                    // 如果有多次辦理情形，需要逐一更新
                    if (createHandlingRounds.length > 0) {
                        // 驗證數據是否真的寫入資料庫
                        const verifyRes = await fetch(`/api/issues?page=1&pageSize=100&q=${encodeURIComponent(number)}&_t=${Date.now()}`);
                        if (verifyRes.ok) {
                            const verifyData = await verifyRes.json();
                            const exactMatch = verifyData.data?.find(item => String(item.number) === String(number));
                            if (exactMatch) {
                                const issueId = exactMatch.id;
                                
                                // 更新後續的辦理情形輪次（從第二次開始）
                                let updateSuccess = true;
                                let updateCount = 0;
                                for (let i = 1; i < createHandlingRounds.length; i++) {
                                    const roundData = createHandlingRounds[i];
                                    if (roundData.handling && roundData.handling.trim()) {
                                        const round = i + 1;
                                        try {
                                            const updateRes = await apiFetch(`/api/issues/${issueId}`, {
                                                method: 'PUT',
                                                body: JSON.stringify({
                                                    status: status,
                                                    round: round,
                                                    handling: roundData.handling.trim(),
                                                    review: '',
                                                    replyDate: roundData.replyDate ? roundData.replyDate.trim() : null,
                                                    responseDate: null // 辦理情形階段不需要函復日期
                                                })
                                            });
                                            if (updateRes.ok) {
                                                updateCount++;
                                            } else {
                                                updateSuccess = false;
                                                console.error(`更新第 ${round} 次辦理情形失敗`);
                                            }
                                        } catch (e) {
                                            updateSuccess = false;
                                            console.error(`更新第 ${round} 次辦理情形錯誤:`, e);
                                        }
                                    }
                                }
                                
                                if (createHandlingRounds.length > 1) {
                                    if (updateSuccess && updateCount === createHandlingRounds.length - 1) {
                                        showToast(`新增成功！已新增事項及 ${createHandlingRounds.length} 次辦理情形`);
                                    } else if (updateCount > 0) {
                                        showToast(`新增成功！已新增事項及 ${updateCount + 1} 次辦理情形（部分更新失敗）`, 'warning');
                                    } else {
                                        showToast('新增成功，但辦理情形更新失敗', 'warning');
                                    }
                                } else if (createHandlingRounds.length === 1 && createHandlingRounds[0].handling.trim()) {
                                    showToast('新增成功！已新增事項及 1 次辦理情形');
                                } else if (createHandlingRounds.length > 0 && createHandlingRounds.some(r => r.handling.trim())) {
                                    showToast('新增成功！已新增事項及辦理情形');
                                } else {
                                    showToast('新增成功！已新增事項（可稍後再新增辦理情形）');
                                }
                            } else {
                                // 驗證失敗，但後端已返回成功，仍然顯示成功
                                showToast('新增成功，資料已確認寫入資料庫');
                            }
                        } else {
                            // verifyRes 失敗，但後端已返回成功，仍然顯示成功
                            showToast('新增成功，資料已確認寫入資料庫');
                        }
                        
                        // 清理表單
                        if (continuousMode) {
                            document.getElementById('createNumber').value = '';
                            document.getElementById('createKind').value = '';
                            document.getElementById('createContent').value = '';
                            // 重置辦理情形（保留第一次）
                            createHandlingRounds = [{
                                round: 1,
                                handling: '',
                                replyDate: '',
                                responseDate: ''
                            }];
                            renderCreateHandlingRounds();
                            document.getElementById('createNumber').focus();
                        } else {
                            document.getElementById('createNumber').value = '';
                            if (yearDisplay) yearDisplay.value = '';
                            document.getElementById('createUnit').value = '';
                            document.getElementById('createDivision').value = '';
                            document.getElementById('createInspection').value = '';
                            document.getElementById('createKind').value = '';
                            document.getElementById('createContent').value = '';
                            document.getElementById('createPlanName').value = '';
                            document.getElementById('createIssueDate').value = '';
                            // 重置辦理情形
                            initCreateHandlingRounds();
                        }
                    } else {
                        // 沒有辦理情形輪次，直接顯示成功並清理表單
                        showToast('新增成功，資料已確認寫入資料庫');
                        
                        // 清理表單
                        if (continuousMode) {
                            document.getElementById('createNumber').value = '';
                            document.getElementById('createKind').value = '';
                            document.getElementById('createContent').value = '';
                            // 重置辦理情形（保留第一次）
                            createHandlingRounds = [{
                                round: 1,
                                handling: '',
                                replyDate: '',
                                responseDate: ''
                            }];
                            renderCreateHandlingRounds();
                            document.getElementById('createNumber').focus();
                        } else {
                            document.getElementById('createNumber').value = '';
                            if (yearDisplay) yearDisplay.value = '';
                            document.getElementById('createUnit').value = '';
                            document.getElementById('createDivision').value = '';
                            document.getElementById('createInspection').value = '';
                            document.getElementById('createKind').value = '';
                            document.getElementById('createContent').value = '';
                            document.getElementById('createPlanName').value = '';
                            document.getElementById('createIssueDate').value = '';
                            // 重置辦理情形
                            initCreateHandlingRounds();
                        }
                    }

                    loadIssuesPage(1);
                    loadPlanOptions();
                    return;
                } else {
                    // newCount 和 updateCount 都是 0，表示沒有資料被寫入
                    showToast('儲存失敗：沒有資料被寫入資料庫', 'error');
                }
            } catch (e) { 
                showToast('Error: ' + e.message, 'error'); 
            }
        }
        
        // 批次模式：初始化表格（快速新增模式：預設只顯示一列）
        function initCreateBatchGrid() {
            const tbody = document.getElementById('createBatchGridBody');
            if (!tbody) return;
            tbody.innerHTML = '';
            batchHandlingData = {}; // 重置辦理情形資料
            // 快速新增模式：預設只顯示一列
            addCreateBatchRow();
            // 初始化後更新所有行的辦理情形狀態
            setTimeout(() => {
                updateAllBatchHandlingStatus();
            }, 100);
        }
        
        // 批次模式：新增一列（改為直接使用 textarea 輸入事項內容）
        function addCreateBatchRow() {
            const tbody = document.getElementById('createBatchGridBody');
            if (!tbody) return;
            const rowIdx = tbody.children.length;
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="text-align:center;color:#94a3b8;font-size:12px;">${rowIdx + 1}</td>
                <td><input type="text" class="filter-input create-batch-number" placeholder="編號..." onchange="handleCreateBatchNumberChange(this)" style="font-family:monospace;"></td>
                <td style="position:relative;">
                    <textarea class="filter-input create-batch-content-textarea" rows="3" placeholder="請輸入事項內容..." style="resize:vertical;min-height:60px;max-height:120px;font-size:13px;line-height:1.6;padding:8px 10px;"></textarea>
                </td>
                <td><input type="text" class="filter-input create-batch-year" style="background:#f1f5f9;color:#64748b;" readonly></td>
                <td><input type="text" class="filter-input create-batch-unit" style="background:#f1f5f9;color:#64748b;" readonly></td>
                <td><select class="filter-select create-batch-division"><option value="">-</option><option value="運務">運務</option><option value="工務">工務</option><option value="機務">機務</option><option value="電務">電務</option><option value="安全">安全</option><option value="審核">審核</option><option value="災防">災防</option><option value="運轉">運轉</option><option value="土木">土木</option><option value="機電">機電</option><option value="土建">土建</option><option value="安全管理">安全管理</option><option value="營運">營運</option><option value="其他">其他</option></select></td>
                <td><select class="filter-select create-batch-inspection"><option value="">-</option><option value="定期檢查">定期檢查</option><option value="例行性檢查">例行性檢查</option><option value="特別檢查">特別檢查</option><option value="臨時檢查">臨時檢查</option></select></td>
                <td><select class="filter-select create-batch-kind"><option value="">-</option><option value="N">缺失</option><option value="O">觀察</option><option value="R">建議</option></select></td>
                <td><select class="filter-select create-batch-status"><option value="持續列管">持續列管</option><option value="解除列管">解除列管</option><option value="自行列管">自行列管</option></select></td>
                <td style="text-align:center;">
                    <button class="btn btn-outline btn-sm create-batch-handling-btn" onclick="openBatchHandlingModal(${rowIdx})" data-row-index="${rowIdx}" style="padding:6px 12px; font-size:12px; width:100%;" title="點擊新增或管理辦理情形">
                        <span class="create-batch-handling-status">新增辦理情形</span>
                    </button>
                </td>
                <td style="text-align:center;">
                    <button class="btn btn-danger btn-sm" onclick="removeCreateBatchRow(this)" style="padding:4px 8px;">×</button>
                </td>
            `;
            tbody.appendChild(tr);
            // 更新該行的辦理情形狀態顯示
            updateBatchHandlingStatus(rowIdx);
        }
        
        // 批次模式：移除一列
        function removeCreateBatchRow(btn) {
            const tr = btn.closest('tr');
            const tbody = document.getElementById('createBatchGridBody');
            if (tbody && tbody.children.length > 1) {
                const rowIndex = Array.from(tbody.children).indexOf(tr);
                tr.remove();
                
                // 移除該行的辦理情形資料
                if (batchHandlingData[rowIndex]) {
                    delete batchHandlingData[rowIndex];
                }
                
                // 重新索引辦理情形資料（因為行號改變了）
                const newBatchHandlingData = {};
                tbody.querySelectorAll('tr').forEach((row, idx) => {
                    const oldIndex = Array.from(tbody.children).indexOf(row);
                    if (batchHandlingData[oldIndex]) {
                        newBatchHandlingData[idx] = batchHandlingData[oldIndex];
                    }
                });
                batchHandlingData = newBatchHandlingData;
                
                // Re-index
                tbody.querySelectorAll('tr').forEach((row, idx) => {
                    row.cells[0].innerText = idx + 1;
                    // 更新辦理情形按鈕的 onclick 和 data-row-index
                    const handlingBtn = row.querySelector('.create-batch-handling-btn');
                    if (handlingBtn) {
                        handlingBtn.setAttribute('onclick', `openBatchHandlingModal(${idx})`);
                        handlingBtn.setAttribute('data-row-index', idx);
                    }
                });
                // 更新所有行的辦理情形狀態顯示
                updateAllBatchHandlingStatus();
            } else {
                showToast('至少需保留一列', 'error');
            }
        }
        
        // 批次模式：調整textarea寬度和高度（已棄用，保留以備不時之需）
        function adjustTextareaWidth(textarea) {
            // 根據內容長度動態調整寬度和高度
            const content = textarea.value;
            const contentLength = content.length;
            
            // 計算行數（假設每行約50個字符）
            const lines = Math.max(1, Math.ceil(contentLength / 50));
            const maxLines = 5; // 最多顯示5行
            textarea.rows = Math.min(lines, maxLines);
            
            // 調整寬度：根據內容長度和行數
            const minWidth = 200;
            const maxWidth = 600;
            const charWidth = 7; // 估算每個字符的寬度（px）
            const padding = 24; // 左右padding
            
            // 如果是多行，使用較大的寬度
            if (lines > 1) {
                textarea.style.width = Math.min(maxWidth, Math.max(minWidth, 400)) + 'px';
            } else {
                // 單行時根據內容長度調整
                const calculatedWidth = Math.max(minWidth, Math.min(maxWidth, contentLength * charWidth + padding));
                textarea.style.width = calculatedWidth + 'px';
            }
        }
        
        // 批次模式：事項內容編輯模態框管理
        let currentBatchContentRowIndex = null;
        
        function openBatchContentModal(rowIndex) {
            // 已改為直接在表格中輸入，此函數不再使用
            // 如果需要，可以聚焦到該行的 textarea
            const tbody = document.getElementById('createBatchGridBody');
            if (!tbody) return;
            const tr = tbody.children[rowIndex];
            if (!tr) return;
            const textarea = tr.querySelector('.create-batch-content-textarea');
            if (textarea) {
                textarea.focus();
            }
        }
        
        function closeBatchContentModal() {
            // 已改為直接在表格中輸入，此函數不再使用
            return;
        }
        
        function saveBatchContent() {
            // 已改為直接在表格中輸入，此函數不再使用
            return;
        }
        
        // 點擊模態框背景關閉（在DOMContentLoaded中初始化）
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initBatchContentModal);
        } else {
            initBatchContentModal();
        }
        
        function initBatchContentModal() {
            const modal = document.getElementById('batchContentModal');
            if (modal) {
                modal.addEventListener('click', (e) => {
                    if (e.target === modal) {
                        closeBatchContentModal();
                    }
                });
            }
        }
        
        // 批次模式：處理編號變更
        function handleCreateBatchNumberChange(input) {
            const tr = input.closest('tr');
            const val = input.value.trim();
            if (!val) return;

            const info = parseItemNumber(val);
            if (info) {
                // 優先使用計畫的年度，如果計畫沒有年度才使用從編號解析出來的年度
                const planValue = document.getElementById('createPlanName')?.value.trim();
                if (planValue) {
                    const { year: planYear } = parsePlanValue(planValue);
                    if (planYear) {
                        tr.querySelector('.create-batch-year').value = planYear;
                    } else if (info.yearRoc) {
                        tr.querySelector('.create-batch-year').value = info.yearRoc;
                    }
                } else if (info.yearRoc) {
                    tr.querySelector('.create-batch-year').value = info.yearRoc;
                }
                
                if (info.orgCode && info.orgCode !== '?') {
                    const name = ORG_MAP[info.orgCode] || info.orgCode;
                    if (name && name !== '?') {
                        const unitInput = tr.querySelector('.create-batch-unit');
                        if (unitInput) unitInput.value = name;
                    }
                }
                if (info.divCode && info.divCode !== '?') {
                    const divName = DIVISION_MAP[info.divCode];
                    if (divName) {
                        const divisionSelect = tr.querySelector('.create-batch-division');
                        if (divisionSelect) divisionSelect.value = divName;
                    }
                }
                if (info.inspectCode && info.inspectCode !== '?') {
                    const inspectName = INSPECTION_MAP[info.inspectCode];
                    if (inspectName) {
                        const inspectionSelect = tr.querySelector('.create-batch-inspection');
                        if (inspectionSelect) inspectionSelect.value = inspectName;
                    }
                }
                if (info.kindCode && info.kindCode !== '?') {
                    const kindSelect = tr.querySelector('.create-batch-kind');
                    if (kindSelect) kindSelect.value = info.kindCode;
                }
            }
        }
        
        // 批次模式辦理情形管理
        let batchHandlingData = {}; // 儲存每筆事項的辦理情形 { rowIndex: [rounds...] }
        let currentBatchHandlingRowIndex = -1; // 當前正在編輯的行索引
        
        // 開啟批次辦理情形管理 Modal
        function openBatchHandlingModal(rowIndex) {
            const rows = document.querySelectorAll('#createBatchGridBody tr');
            if (rowIndex < 0 || rowIndex >= rows.length) return;
            
            const row = rows[rowIndex];
            const number = row.querySelector('.create-batch-number').value.trim();
            
            if (!number) {
                showToast('請先填寫編號', 'error');
                return;
            }
            
            currentBatchHandlingRowIndex = rowIndex;
            document.getElementById('batchHandlingModalNumber').textContent = number || `第 ${rowIndex + 1} 列`;
            
            // 載入該行的辦理情形資料（如果有的話）
            if (!batchHandlingData[rowIndex]) {
                batchHandlingData[rowIndex] = [];
            }
            
            renderBatchHandlingRounds();
            document.getElementById('batchHandlingModal').classList.add('open');
        }
        
        // 初始化時更新所有行的辦理情形狀態
        function updateAllBatchHandlingStatus() {
            const rows = document.querySelectorAll('#createBatchGridBody tr');
            rows.forEach((row, idx) => {
                updateBatchHandlingStatus(idx);
            });
        }
        
        // 關閉批次辦理情形管理 Modal
        function closeBatchHandlingModal() {
            document.getElementById('batchHandlingModal').classList.remove('open');
            currentBatchHandlingRowIndex = -1;
        }
        
        
        // 新增批次辦理情形輪次
        function addBatchHandlingRound() {
            if (currentBatchHandlingRowIndex === -1) return;
            if (!batchHandlingData[currentBatchHandlingRowIndex]) {
                batchHandlingData[currentBatchHandlingRowIndex] = [];
            }
            
            const round = batchHandlingData[currentBatchHandlingRowIndex].length + 1;
            batchHandlingData[currentBatchHandlingRowIndex].push({
                round: round,
                handling: '',
                replyDate: ''
            });
            renderBatchHandlingRounds();
        }
        
        // 移除批次辦理情形輪次
        function removeBatchHandlingRound(index) {
            if (currentBatchHandlingRowIndex === -1) return;
            if (!batchHandlingData[currentBatchHandlingRowIndex]) return;
            
            batchHandlingData[currentBatchHandlingRowIndex].splice(index, 1);
            // 重新編號
            batchHandlingData[currentBatchHandlingRowIndex].forEach((r, i) => {
                r.round = i + 1;
            });
            renderBatchHandlingRounds();
        }
        
        // 渲染批次辦理情形輪次
        function renderBatchHandlingRounds() {
            const container = document.getElementById('batchHandlingRoundsContainer');
            if (!container || currentBatchHandlingRowIndex === -1) return;
            
            const rounds = batchHandlingData[currentBatchHandlingRowIndex] || [];
            
            if (rounds.length === 0) {
                container.innerHTML = '<div style="text-align:center; padding:40px; color:#94a3b8; font-size:14px;">尚未新增辦理情形，點擊「新增辦理情形」開始新增</div>';
                return;
            }
            
            let html = '';
            rounds.forEach((roundData, index) => {
                html += `
                    <div class="batch-handling-round" data-index="${index}" style="background:white; padding:16px; border-radius:8px; border:1px solid #e2e8f0; margin-bottom:12px;">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                            <div style="font-weight:700; color:#334155; font-size:14px;">
                                第 ${roundData.round} 次機構辦理情形
                            </div>
                            <button type="button" class="btn btn-danger btn-sm" onclick="removeBatchHandlingRound(${index})" style="padding:4px 12px; font-size:12px;">刪除</button>
                        </div>
                        <div style="margin-bottom:12px;">
                            <label style="display:block; font-weight:600; color:#475569; font-size:13px; margin-bottom:6px;">
                                辦理情形
                            </label>
                            <textarea class="filter-input batch-handling-text" data-index="${index}" 
                                placeholder="請輸入機構辦理情形..." 
                                style="width:100%; min-height:120px; padding:12px; font-size:14px; line-height:1.6; resize:vertical; background:white;"
                                oninput="updateBatchHandlingRound(${index}, 'handling', this.value)">${roundData.handling}</textarea>
                        </div>
                    </div>
                `;
            });
            container.innerHTML = html;
        }
        
        // 更新批次辦理情形輪次資料
        function updateBatchHandlingRound(index, field, value) {
            if (currentBatchHandlingRowIndex === -1) return;
            if (batchHandlingData[currentBatchHandlingRowIndex] && batchHandlingData[currentBatchHandlingRowIndex][index]) {
                batchHandlingData[currentBatchHandlingRowIndex][index][field] = value;
            }
        }
        
        // 儲存批次辦理情形
        async function saveBatchHandlingRounds() {
            if (currentBatchHandlingRowIndex === -1) return;
            
            const rows = document.querySelectorAll('#createBatchGridBody tr');
            if (currentBatchHandlingRowIndex < 0 || currentBatchHandlingRowIndex >= rows.length) return;
            
            const row = rows[currentBatchHandlingRowIndex];
            const number = row.querySelector('.create-batch-number')?.value.trim();
            const issueId = row.getAttribute('data-issue-id');
            const handlingRounds = batchHandlingData[currentBatchHandlingRowIndex] || [];
            
            // 如果事項已存在於資料庫（有 ID），則立即儲存到資料庫
            if (issueId && number) {
                try {
                    // 移除儲存中的提示訊息，只保留錯誤訊息
                    
                    // 先更新第一次辦理情形（如果有的話）
                    if (handlingRounds.length > 0 && handlingRounds[0].handling && handlingRounds[0].handling.trim()) {
                        const firstRound = handlingRounds[0];
                        const updateRes = await apiFetch(`/api/issues/${issueId}`, {
                            method: 'PUT',
                            body: JSON.stringify({
                                handling: firstRound.handling.trim(),
                                replyDate: firstRound.replyDate ? firstRound.replyDate.trim() : null,
                                responseDate: null
                            })
                        });
                        
                        if (!updateRes.ok) {
                            throw new Error('更新第一次辦理情形失敗');
                        }
                    }
                    
                    // 更新後續的辦理情形輪次（從第2次開始）
                    for (let i = 1; i < handlingRounds.length; i++) {
                        const roundData = handlingRounds[i];
                        if (roundData.handling && roundData.handling.trim()) {
                            const round = i + 1;
                            try {
                                const updateRes = await apiFetch(`/api/issues/${issueId}`, {
                                    method: 'PUT',
                                    body: JSON.stringify({
                                        round: round,
                                        handling: roundData.handling.trim(),
                                        review: '',
                                        replyDate: roundData.replyDate ? roundData.replyDate.trim() : null,
                                        responseDate: null
                                    })
                                });
                                
                                if (!updateRes.ok) {
                                    console.error(`更新第 ${round} 次辦理情形失敗`);
                                }
                            } catch (e) {
                                console.error(`更新第 ${round} 次辦理情形錯誤:`, e);
                            }
                        }
                    }
                    
                    // 保留儲存成功的提示訊息（資料庫操作結果）
                    showToast('辦理情形已成功儲存至資料庫', 'success');
                    // 更新辦理情形狀態顯示
                    updateBatchHandlingStatus(currentBatchHandlingRowIndex);
                    closeBatchHandlingModal();
                } catch (e) {
                    showToast('儲存辦理情形失敗: ' + e.message, 'error');
                }
            } else {
                // 如果事項尚未存在於資料庫（新建立的事項），則保持現有行為
                // 保留儲存成功的提示訊息（資料庫操作結果）
                showToast('辦理情形已儲存（將在批次新增時一併保存）', 'success');
                // 更新辦理情形狀態顯示
                updateBatchHandlingStatus(currentBatchHandlingRowIndex);
                closeBatchHandlingModal();
            }
        }
        
        // 更新批次辦理情形狀態顯示
        function updateBatchHandlingStatus(rowIndex) {
            const rows = document.querySelectorAll('#createBatchGridBody tr');
            if (rowIndex < 0 || rowIndex >= rows.length) return;
            
            const row = rows[rowIndex];
            const btn = row.querySelector('.create-batch-handling-btn');
            const statusSpan = row.querySelector('.create-batch-handling-status');
            
            if (!btn || !statusSpan) return;
            
            const handlingRounds = batchHandlingData[rowIndex] || [];
            const hasHandling = handlingRounds.length > 0 && handlingRounds.some(r => r.handling && r.handling.trim());
            
            if (hasHandling) {
                const count = handlingRounds.filter(r => r.handling && r.handling.trim()).length;
                statusSpan.textContent = `已填寫 (${count}次)`;
                btn.style.backgroundColor = '#ecfdf5';
                btn.style.borderColor = '#10b981';
                btn.style.color = '#047857';
            } else {
                statusSpan.textContent = '新增辦理情形';
                btn.style.backgroundColor = '';
                btn.style.borderColor = '';
                btn.style.color = '';
            }
        }
        
        // 批次模式：儲存所有項目
        async function saveCreateBatchItems() {
            const planValue = document.getElementById('createPlanName').value.trim();
            const issueDate = document.getElementById('createIssueDate').value.trim();

            if (!planValue) return showToast('請選擇檢查計畫', 'error');
            const { name: planName, year: planYear } = parsePlanValue(planValue);
            if (!issueDate) return showToast('請填寫初次發函日期', 'error');

            const rows = document.querySelectorAll('#createBatchGridBody tr');
            const items = [];
            let hasError = false;

            rows.forEach((tr, idx) => {
                const number = tr.querySelector('.create-batch-number').value.trim();
                // 改為從 textarea 讀取內容
                const contentTextarea = tr.querySelector('.create-batch-content-textarea');
                const content = contentTextarea ? contentTextarea.value.trim() : '';

                if (!number && !content) return;

                if (!number) {
                    showToast(`第 ${idx + 1} 列缺少編號`, 'error');
                    hasError = true;
                    return;
                }
                
                // 檢查編號是否為空
                if (!number.trim()) {
                    showToast(`第 ${idx + 1} 列編號不能為空`, 'error');
                    hasError = true;
                    return;
                }

                // 優先使用計畫的年度，如果計畫沒有年度才使用表格中的年度
                let year = tr.querySelector('.create-batch-year').value.trim();
                if (planYear && year !== planYear) {
                    // 如果計畫有年度且與表格中的年度不同，使用計畫的年度
                    year = planYear;
                    tr.querySelector('.create-batch-year').value = year;
                }
                
                const unit = tr.querySelector('.create-batch-unit').value.trim();

                if (!year || !unit) {
                    showToast(`第 ${idx + 1} 列的年度或機構未能自動判別，請確認編號格式或選擇有年度的檢查計畫`, 'error');
                    hasError = true;
                    return;
                }

                // 取得該行的辦理情形（第一次）
                const handlingRounds = batchHandlingData[idx] || [];
                const firstHandling = handlingRounds.length > 0 ? handlingRounds[0] : { handling: '', replyDate: '' };

                items.push({
                    number,
                    year,
                    unit,
                    content,
                    status: tr.querySelector('.create-batch-status').value,
                    itemKindCode: tr.querySelector('.create-batch-kind').value,
                    divisionName: tr.querySelector('.create-batch-division').value,
                    inspectionCategoryName: tr.querySelector('.create-batch-inspection').value,
                    planName: planName,
                    issueDate: issueDate,
                    handling: firstHandling.handling ? firstHandling.handling.trim() : '',
                    replyDate: firstHandling.replyDate ? firstHandling.replyDate.trim() : '',
                    scheme: 'BATCH',
                    handlingRounds: handlingRounds // 保存所有辦理情形輪次，用於後續更新
                });
            });

            if (hasError) return;
            if (items.length === 0) return showToast('請至少輸入一筆有效資料', 'error');

            // 檢查是否有重複編號
            const numberSet = new Set();
            const duplicateNumbers = [];
            items.forEach((item, idx) => {
                if (item.number && item.number.trim()) {
                    if (numberSet.has(item.number)) {
                        duplicateNumbers.push({ number: item.number, row: idx + 1 });
                    } else {
                        numberSet.add(item.number);
                    }
                }
            });
            
            if (duplicateNumbers.length > 0) {
                const duplicateList = duplicateNumbers.map(d => `第 ${d.row} 列：${d.number}`).join('\n');
                showToast(`發現重複編號，請修正後再儲存：\n${duplicateList}`, 'error');
                return;
            }

            const confirmed = await showConfirmModal(`確定要批次新增 ${items.length} 筆資料嗎？\n\n計畫：${planName}`, '確定新增', '取消');
            if (!confirmed) return;

            try {
                // 先新增所有事項（第一次辦理情形）
                // 注意：每個事項可能有不同的回復日期，需要在服務器端使用 item.replyDate
                const itemsForImport = items.map(item => {
                    const { handlingRounds, ...itemData } = item;
                    return itemData;
                });
                
                const res = await apiFetch('/api/issues/import', {
                    method: 'POST',
                    body: JSON.stringify({
                        data: itemsForImport,
                        round: 1,
                        reviewDate: ''
                        // 不再使用統一的 replyDate，改為使用每個 item 的 replyDate
                    })
                });

                if (res.ok) {
                    const result = await res.json();
                    
                    // 如果有多次辦理情形，需要逐一更新
                    if (result.newCount > 0 || result.updateCount > 0) {
                        // 驗證並更新後續辦理情形
                        let totalHandlingCount = 0;
                        let updateSuccessCount = 0;
                        
                        for (let i = 0; i < items.length; i++) {
                            const item = items[i];
                            const handlingRounds = item.handlingRounds || [];
                            
                            if (handlingRounds.length > 1) {
                                // 查詢該事項的 ID
                                const verifyRes = await fetch(`/api/issues?page=1&pageSize=100&q=${encodeURIComponent(item.number)}&_t=${Date.now()}`);
                                if (verifyRes.ok) {
                                    const verifyData = await verifyRes.json();
                                    const exactMatch = verifyData.data?.find(issue => String(issue.number) === String(item.number));
                                    
                                    if (exactMatch) {
                                        const issueId = exactMatch.id;
                                        totalHandlingCount += handlingRounds.length - 1;
                                        
                                        // 更新後續的辦理情形輪次
                                        for (let j = 1; j < handlingRounds.length; j++) {
                                            const roundData = handlingRounds[j];
                                            if (roundData.handling && roundData.handling.trim()) {
                                                const round = j + 1;
                                                try {
                                            const updateRes = await apiFetch(`/api/issues/${issueId}`, {
                                                method: 'PUT',
                                                body: JSON.stringify({
                                                    status: item.status,
                                                    round: round,
                                                    handling: roundData.handling.trim(),
                                                    review: '',
                                                    replyDate: roundData.replyDate ? roundData.replyDate.trim() : null,
                                                    responseDate: null // 辦理情形階段不需要函復日期
                                                })
                                            });
                                                    if (updateRes.ok) {
                                                        updateSuccessCount++;
                                                    }
                                                } catch (e) {
                                                    console.error(`更新第 ${i + 1} 筆事項的第 ${round} 次辦理情形錯誤:`, e);
                                                }
                                            }
                                        }
                                    }
                                }
                            } else if (handlingRounds.length === 1 && handlingRounds[0].handling.trim()) {
                                totalHandlingCount++;
                            }
                        }
                        
                        if (totalHandlingCount > 0) {
                            showToast(`批次新增成功！已新增 ${items.length} 筆事項，其中 ${updateSuccessCount + items.filter(item => (item.handlingRounds || []).length > 0 && (item.handlingRounds || [])[0].handling.trim()).length} 筆包含辦理情形`);
                        } else {
                            showToast('批次新增成功！');
                        }
                    } else {
                        showToast('批次新增成功！');
                    }
                    
                    // 檢查是否啟用連續新增模式
                    const continuousMode = document.getElementById('createBatchContinuousMode')?.checked || false;
                    
                    if (continuousMode) {
                        // 連續新增模式：清空已儲存的列，保留計畫和機構設定，自動新增新列
                        const savedRows = document.querySelectorAll('#createBatchGridBody tr');
                        savedRows.forEach((tr, idx) => {
                            if (idx < items.length) {
                                // 只清空編號、類型、事項內容，保留其他欄位
                                const numberInput = tr.querySelector('.create-batch-number');
                                const contentTextarea = tr.querySelector('.create-batch-content-textarea');
                                const kindSelect = tr.querySelector('.create-batch-kind');
                                
                                if (numberInput) numberInput.value = '';
                                if (contentTextarea) contentTextarea.value = '';
                                if (kindSelect) kindSelect.value = '';
                                
                                // 清空該行的辦理情形資料
                                if (batchHandlingData[idx]) {
                                    delete batchHandlingData[idx];
                                }
                                updateBatchHandlingStatus(idx);
                            }
                        });
                        
                        // 如果只有一列，確保該列被清空並聚焦到編號欄位
                        if (savedRows.length === 1) {
                            const firstRow = savedRows[0];
                            const numberInput = firstRow.querySelector('.create-batch-number');
                            if (numberInput) {
                                setTimeout(() => numberInput.focus(), 100);
                            }
                        }
                    } else {
                        // 非連續新增模式：清空所有列並重新初始化
                        initCreateBatchGrid();
                        batchHandlingData = {};
                        document.getElementById('createPlanName').value = '';
                        document.getElementById('createIssueDate').value = '';
                    }
                    
                    loadIssuesPage(1);
                    loadPlanOptions();
                } else {
                    const j = await res.json();
                    showToast('新增失敗: ' + (j.error || '不明錯誤'), 'error');
                }
            } catch (e) {
                showToast('Error: ' + e.message, 'error');
            }
        }
        
        // 向後兼容：保留舊函數名稱
        async function submitManualIssue() {
            return submitCreateIssue();
        }
        
        function initBatchGrid() {
            initCreateBatchGrid();
        }
        
        function addBatchRow() {
            addCreateBatchRow();
        }
        
        function removeBatchRow(btn) {
            removeCreateBatchRow(btn);
        }
        
        function handleBatchNumberChange(input) {
            handleCreateBatchNumberChange(input);
        }
        
        async function saveBatchItems() {
            return saveCreateBatchItems();
        }

        // 保留舊函數名稱以向後兼容
        async function exportAllIssues() {
            return exportAllData();
        }

        async function exportAllData() {
            try {
                const exportDataType = document.querySelector('input[name="exportDataType"]:checked')?.value || 'issues';
                const exportScope = document.querySelector('input[name="exportScope"]:checked')?.value || 'latest';
                const exportFormat = document.querySelector('input[name="exportFormat"]:checked')?.value || 'excel';
                showToast('準備匯出中，請稍候...', 'info');
                
                let issuesData = [];
                let planSchedulesData = [];
                
                // 根據選擇的資料類型獲取資料
                if (exportDataType === 'issues' || exportDataType === 'both') {
                    const res = await fetch('/api/issues?page=1&pageSize=10000&sortField=created_at&sortDir=desc');
                    if (!res.ok) throw new Error('取得開立事項資料失敗');
                    const json = await res.json();
                    issuesData = json.data || [];
                }
                
                if (exportDataType === 'plans' || exportDataType === 'both') {
                    const scheduleRes = await fetch('/api/plan-schedule/all', { credentials: 'include' });
                    if (scheduleRes.ok) {
                        const scheduleJson = await scheduleRes.json();
                        planSchedulesData = scheduleJson.data || [];
                    }
                }
                
                if (exportDataType === 'issues' && issuesData.length === 0) {
                    return showToast('無開立事項資料可匯出', 'error');
                }
                if (exportDataType === 'plans' && planSchedulesData.length === 0) {
                    return showToast('無檢查計畫資料可匯出', 'error');
                }
                if (exportDataType === 'both' && issuesData.length === 0 && planSchedulesData.length === 0) {
                    return showToast('無資料可匯出', 'error');
                }

                // JSON 格式匯出
                if (exportFormat === 'json') {
                    const exportData = {};
                    if (exportDataType === 'issues' || exportDataType === 'both') {
                        exportData.issues = issuesData;
                    }
                    if (exportDataType === 'plans' || exportDataType === 'both') {
                        exportData.plans = planSchedulesData;
                    }
                    const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
                    const link = document.createElement("a");
                    link.href = URL.createObjectURL(blob);
                    const dataTypeLabel = exportDataType === 'issues' ? 'Issues' : (exportDataType === 'plans' ? 'Plans' : 'All');
                    link.download = `SMS_Backup_${dataTypeLabel}_${new Date().toISOString().slice(0, 10)}.json`;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    showToast('JSON 匯出完成', 'success');
                    return;
                }

                // Excel 格式匯出
                if (exportFormat === 'excel') {
                    const wb = XLSX.utils.book_new();
                    
                    // 如果選擇合併匯出，創建多個工作表
                    if (exportDataType === 'both') {
                        if (planSchedulesData.length > 0) {
                            const schedulesWSData = [
                                ['計畫名稱', '年度', '業務類型', '規劃檢查次數', '開始日期', '結束日期', '地點', '檢查人員', '鐵路機構', '檢查類別', '檢查次數', '取號編碼', '建立時間']
                            ];
                            const getBusinessTypeNameExport = (code) => ({ 'OP': '運轉', 'CV': '土建', 'ME': '機務', 'EL': '電務', 'SM': '安全管理', 'AD': '營運／災防審核', 'OT': '其他／產管規劃' }[String(code)] || (code || ''));
                            planSchedulesData.forEach(s => {
                                schedulesWSData.push([
                                    s.plan_name || '',
                                    s.year || '',
                                    getBusinessTypeNameExport(s.business),
                                    s.planned_count != null ? s.planned_count : '',
                                    s.start_date ? s.start_date.slice(0, 10) : '',
                                    s.end_date ? s.end_date.slice(0, 10) : '',
                                    s.location || '',
                                    s.inspector || '',
                                    s.railway || '',
                                    s.inspection_type || '',
                                    s.inspection_seq || '',
                                    s.plan_number || '',
                                    s.created_at ? new Date(s.created_at).toLocaleString('zh-TW') : ''
                                ]);
                            });
                            const schedulesWS = XLSX.utils.aoa_to_sheet(schedulesWSData);
                            XLSX.utils.book_append_sheet(wb, schedulesWS, '檢查計畫');
                        }
                        
                        // 工作表：開立事項
                        if (issuesData.length > 0) {
                            const issuesWSData = [];
                            if (exportScope === 'latest') {
                                issuesWSData.push(['編號', '年度', '機構', '分組', '檢查種類', '類型', '狀態', '事項內容', '最新辦理情形', '最新審查意見']);
                                issuesData.forEach(item => {
                                    let latestH = '', latestR = '';
                                    for (let i = 200; i >= 1; i--) { 
                                        const suffix = i === 1 ? '' : i;
                                        if (!latestH && (item[`handling${suffix}`])) latestH = stripHtml(item[`handling${suffix}`] || ''); 
                                        if (!latestR && (item[`review${suffix}`])) latestR = stripHtml(item[`review${suffix}`] || ''); 
                                    }
                                    issuesWSData.push([
                                        item.number || '',
                                        item.year || '',
                                        item.unit || '',
                                        item.divisionName || '',
                                        item.inspectionCategoryName || '',
                                        item.category || '',
                                        item.status || '',
                                        stripHtml(item.content || ''),
                                        latestH,
                                        latestR
                                    ]);
                                });
                            } else {
                                issuesWSData.push(['編號', '年度', '機構', '分組', '檢查種類', '類型', '狀態', '事項內容', '完整辦理情形歷程', '完整審查意見歷程']);
                                issuesData.forEach(item => {
                                    let fullH = [], fullR = [];
                                    for (let i = 1; i <= 200; i++) {
                                        const suffix = i === 1 ? '' : i;
                                        const valH = item[`handling${suffix}`], valR = item[`review${suffix}`];
                                        if (valH) fullH.push(`[第${i}次] ${stripHtml(valH)}`); 
                                        if (valR) fullR.push(`[第${i}次] ${stripHtml(valR)}`);
                                    }
                                    const joinedH = fullH.length > 0 ? fullH.join("\n-------------------\n") : "";
                                    const joinedR = fullR.length > 0 ? fullR.join("\n-------------------\n") : "";
                                    issuesWSData.push([
                                        item.number || '',
                                        item.year || '',
                                        item.unit || '',
                                        item.divisionName || '',
                                        item.inspectionCategoryName || '',
                                        item.category || '',
                                        item.status || '',
                                        stripHtml(item.content || ''),
                                        joinedH,
                                        joinedR
                                    ]);
                                });
                            }
                            const issuesWS = XLSX.utils.aoa_to_sheet(issuesWSData);
                            XLSX.utils.book_append_sheet(wb, issuesWS, '開立事項');
                        }
                    } else if (exportDataType === 'plans') {
                        if (planSchedulesData.length > 0) {
                            const schedulesWSData = [
                                ['計畫名稱', '年度', '業務類型', '規劃檢查次數', '開始日期', '結束日期', '地點', '檢查人員', '鐵路機構', '檢查類別', '檢查次數', '取號編碼', '建立時間']
                            ];
                            const getBusinessTypeNameExport = (code) => ({ 'OP': '運轉', 'CV': '土建', 'ME': '機務', 'EL': '電務', 'SM': '安全管理', 'AD': '營運／災防審核', 'OT': '其他／產管規劃' }[String(code)] || (code || ''));
                            planSchedulesData.forEach(s => {
                                schedulesWSData.push([
                                    s.plan_name || '',
                                    s.year || '',
                                    getBusinessTypeNameExport(s.business),
                                    s.planned_count != null ? s.planned_count : '',
                                    s.start_date ? s.start_date.slice(0, 10) : '',
                                    s.end_date ? s.end_date.slice(0, 10) : '',
                                    s.location || '',
                                    s.inspector || '',
                                    s.railway || '',
                                    s.inspection_type || '',
                                    s.inspection_seq || '',
                                    s.plan_number || '',
                                    s.created_at ? new Date(s.created_at).toLocaleString('zh-TW') : ''
                                ]);
                            });
                            const schedulesWS = XLSX.utils.aoa_to_sheet(schedulesWSData);
                            XLSX.utils.book_append_sheet(wb, schedulesWS, '檢查計畫');
                        }
                    } else {
                        // 僅匯出開立事項
                        const issuesWSData = [];
                        if (exportScope === 'latest') {
                            issuesWSData.push(['編號', '年度', '機構', '分組', '檢查種類', '類型', '狀態', '事項內容', '最新辦理情形', '最新審查意見']);
                            issuesData.forEach(item => {
                                let latestH = '', latestR = '';
                                for (let i = 200; i >= 1; i--) { 
                                    const suffix = i === 1 ? '' : i;
                                    if (!latestH && (item[`handling${suffix}`])) latestH = stripHtml(item[`handling${suffix}`] || ''); 
                                    if (!latestR && (item[`review${suffix}`])) latestR = stripHtml(item[`review${suffix}`] || ''); 
                                }
                                issuesWSData.push([
                                    item.number || '',
                                    item.year || '',
                                    item.unit || '',
                                    item.divisionName || '',
                                    item.inspectionCategoryName || '',
                                    item.category || '',
                                    item.status || '',
                                    stripHtml(item.content || ''),
                                    latestH,
                                    latestR
                                ]);
                            });
                        } else {
                            issuesWSData.push(['編號', '年度', '機構', '分組', '檢查種類', '類型', '狀態', '事項內容', '完整辦理情形歷程', '完整審查意見歷程']);
                            issuesData.forEach(item => {
                                let fullH = [], fullR = [];
                                for (let i = 1; i <= 200; i++) {
                                    const suffix = i === 1 ? '' : i;
                                    const valH = item[`handling${suffix}`], valR = item[`review${suffix}`];
                                    if (valH) fullH.push(`[第${i}次] ${stripHtml(valH)}`); 
                                    if (valR) fullR.push(`[第${i}次] ${stripHtml(valR)}`);
                                }
                                const joinedH = fullH.length > 0 ? fullH.join("\n-------------------\n") : "";
                                const joinedR = fullR.length > 0 ? fullR.join("\n-------------------\n") : "";
                                issuesWSData.push([
                                    item.number || '',
                                    item.year || '',
                                    item.unit || '',
                                    item.divisionName || '',
                                    item.inspectionCategoryName || '',
                                    item.category || '',
                                    item.status || '',
                                    stripHtml(item.content || ''),
                                    joinedH,
                                    joinedR
                                ]);
                            });
                        }
                        const issuesWS = XLSX.utils.aoa_to_sheet(issuesWSData);
                        XLSX.utils.book_append_sheet(wb, issuesWS, '開立事項');
                    }
                    
                    // 生成 Excel 檔案
                    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    const link = document.createElement("a");
                    link.href = URL.createObjectURL(blob);
                    let fileName = '';
                    if (exportDataType === 'issues') {
                        const typeLabel = exportScope === 'latest' ? 'Latest' : 'FullHistory';
                        fileName = `SMS_Issues_${typeLabel}_${new Date().toISOString().slice(0, 10)}.xlsx`;
                    } else if (exportDataType === 'plans') {
                        fileName = `SMS_Plans_${new Date().toISOString().slice(0, 10)}.xlsx`;
                    } else {
                        fileName = `SMS_AllData_${new Date().toISOString().slice(0, 10)}.xlsx`;
                    }
                    link.download = fileName;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    showToast('Excel 匯出完成', 'success');
                    return;
                }

                // 如果格式不是 Excel 或 JSON，預設使用 Excel
                if (exportFormat !== 'excel' && exportFormat !== 'json') {
                    showToast('不支援的匯出格式，將使用 Excel 格式', 'warning');
                }
            } catch (e) { 
                showToast('匯出失敗: ' + e.message, 'error'); 
            }
        }

        // --- User modal submit & password strength ---
        document.getElementById('uPwd')?.addEventListener('input', updatePwdStrength); document.getElementById('uPwdConfirm')?.addEventListener('input', updatePwdStrength);
        function updatePwdStrength() { 
            const p = document.getElementById('uPwd').value || ''; 
            const conf = document.getElementById('uPwdConfirm').value || ''; 
            let score = 0; 
            let issues = [];
            
            if (p.length >= 8) score++; else issues.push('至少8字元');
            if (/[A-Z]/.test(p)) score++; else issues.push('大寫字母');
            if (/[a-z]/.test(p)) score++; else issues.push('小寫字母');
            if (/[0-9]/.test(p)) score++; else issues.push('數字');
            if (/[^A-Za-z0-9]/.test(p)) score++;
            
            const texts = ['弱', '偏弱', '一般', '良好', '強']; 
            const strengthText = texts[Math.min(score, 4)];
            const mismatchText = conf && p !== conf ? ' (密碼不相符)' : '';
            const issuesText = issues.length > 0 && p.length > 0 ? ` - 缺少: ${issues.join(', ')}` : '';
            
            document.getElementById('pwdStrength').innerText = `密碼強度: ${strengthText}${mismatchText}${issuesText}`; 
        }

        // User CRUD
        async function openUserModal(mode, id) { const m = document.getElementById('userModal'), t = document.getElementById('userModalTitle'), e = document.getElementById('uEmail'); if (mode === 'create') { t.innerText = '新增'; document.getElementById('targetUserId').value = ''; document.getElementById('uName').value = ''; e.value = ''; e.disabled = false; document.getElementById('uPwd').value = ''; document.getElementById('uPwdConfirm').value = ''; document.getElementById('pwdStrength').innerText = '密碼強度: -'; document.getElementById('pwdHint').innerText = ''; document.getElementById('uRole').value = 'viewer'; } else { const u = userList.find(x => x.id === id) || {}; t.innerText = '編輯'; document.getElementById('targetUserId').value = u.id || ''; document.getElementById('uName').value = u.name || ''; e.value = u.username || ''; e.disabled = true; document.getElementById('uPwd').value = ''; document.getElementById('uPwdConfirm').value = ''; document.getElementById('pwdHint').innerText = '(留空不改)'; document.getElementById('pwdStrength').innerText = '密碼強度: -'; document.getElementById('uRole').value = u.role || 'viewer'; } m.classList.add('open'); }
        async function submitUser() { 
            const id = document.getElementById('targetUserId').value, 
                name = document.getElementById('uName').value, 
                email = document.getElementById('uEmail').value, 
                pwd = document.getElementById('uPwd').value, 
                pwdConfirm = document.getElementById('uPwdConfirm').value, 
                role = document.getElementById('uRole').value; 
            
            if (!id) { 
                if (!email) return showToast('請輸入帳號', 'error'); 
                if (!pwd) return showToast('請輸入密碼', 'error'); 
                if (pwd !== pwdConfirm) return showToast('密碼與確認密碼不符', 'error'); 
                
                // 使用前端驗證函數
                const validation = validatePasswordFrontend(pwd);
                if (!validation.valid) return showToast(validation.message, 'error');
                
                const res = await apiFetch('/api/users', { 
                    method: 'POST', 
                    body: JSON.stringify({ username: email, name, password: pwd, role }) 
                }); 
                const j = await res.json(); 
                if (res.ok) { 
                    showToast('新增成功'); 
                    document.getElementById('userModal').classList.remove('open'); 
                    loadUsersPage(1); 
                } else showToast(j.error || '新增失敗', 'error'); 
            } else { 
                const payload = { name, role }; 
                if (pwd) { 
                    if (pwd !== pwdConfirm) return showToast('密碼與確認密碼不符', 'error'); 
                    
                    // 使用前端驗證函數
                    const validation = validatePasswordFrontend(pwd);
                    if (!validation.valid) return showToast(validation.message, 'error');
                    
                    payload.password = pwd; 
                } 
                const res = await apiFetch(`/api/users/${id}`, { 
                    method: 'PUT', 
                    body: JSON.stringify(payload) 
                }); 
                const j = await res.json(); 
                if (res.ok) { 
                    showToast('更新成功'); 
                    document.getElementById('userModal').classList.remove('open'); 
                    loadUsersPage(usersPage); 
                } else showToast(j.error || '更新失敗', 'error'); 
            } 
        }
        async function deleteUser(id) { 
            const confirmed = await showConfirmModal('確定要刪除此帳號嗎？\n\n此操作無法復原！', '確定刪除', '取消');
            if (!confirmed) return; 
            try {
                const res = await apiFetch(`/api/users/${id}`, { method: 'DELETE' }); 
                const data = await res.json();
                if (res.ok) { 
                    showToast('刪除成功'); 
                    loadUsersPage(1); 
                } else {
                    showToast(data.error || '刪除失敗', 'error'); 
                }
            } catch (e) {
                showToast('刪除失敗: ' + e.message, 'error');
            }
        }
        
        // 帳號匯出功能
        async function exportUsers() {
            try {
                showToast('準備匯出中，請稍候...', 'info');
                // 取得所有帳號資料
                const res = await fetch('/api/users?page=1&pageSize=10000');
                if (!res.ok) throw new Error('取得帳號資料失敗');
                const json = await res.json();
                const users = json.data || [];
                
                if (users.length === 0) {
                    return showToast('無帳號資料可匯出', 'error');
                }
                
                // 從頁面取得匯出格式
                const formatRadio = document.querySelector('input[name="userExportFormat"]:checked');
                const format = formatRadio ? formatRadio.value : 'csv';
                
                if (format === 'json') {
                    const blob = new Blob([JSON.stringify(users, null, 2)], { type: 'application/json' });
                    const link = document.createElement("a");
                    link.href = URL.createObjectURL(blob);
                    link.download = `Users_Backup_${new Date().toISOString().slice(0, 10)}.json`;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    showToast('JSON 匯出完成', 'success');
                } else {
                    // CSV 格式（使用英文權限代碼，與匯入格式一致）
                    let csvContent = '\uFEFF';
                    csvContent += "姓名,帳號,權限,建立時間\n";
                    users.forEach(user => {
                        const clean = (t) => `"${String(t || '').replace(/"/g, '""').trim()}"`;
                        // 使用英文權限代碼，與匯入格式一致
                        csvContent += `${clean(user.name)},${clean(user.username)},${clean(user.role)},${clean(new Date(user.created_at).toLocaleString('zh-TW'))}\n`;
                    });
                    
                    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                    const link = document.createElement("a");
                    link.href = URL.createObjectURL(blob);
                    link.download = `Users_${new Date().toISOString().slice(0, 10)}.csv`;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    showToast('CSV 匯出完成', 'success');
                }
            } catch (e) {
                showToast('匯出失敗: ' + e.message, 'error');
            }
        }
        
        // 帳號匯入功能
        function openUserImportModal() {
            const modal = document.getElementById('userImportModal');
            if (modal) modal.classList.add('open');
        }
        
        function closeUserImportModal() {
            const modal = document.getElementById('userImportModal');
            if (modal) {
                modal.classList.remove('open');
                const fileInput = document.getElementById('userImportFile');
                if (fileInput) fileInput.value = '';
            }
        }
        
        function downloadUserCSVTemplate() {
            // 範例檔格式：姓名,帳號,權限,密碼（選填）
            // 權限值：admin（系統管理員）、manager（資料管理者）、editor（審查人員）、viewer（檢視人員）
            const csv = '姓名,帳號,權限,密碼\n張三,zhang@example.com,editor,password123\n李四,li@example.com,manager,password123\n王五,wang@example.com,viewer,\n趙六,zhao@example.com,admin,admin123';
            const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = '帳號匯入範例.csv';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
        
        async function importUsersCSV() {
            const fileInput = document.getElementById('userImportFile');
            if (!fileInput) return showToast('找不到檔案選擇器', 'error');
            const file = fileInput.files[0];
            if (!file) return showToast('請選擇 CSV 檔案', 'error');
            
            const reader = new FileReader();
            reader.onload = async function(e) {
                try {
                    const csv = e.target.result;
                    Papa.parse(csv, {
                        header: true,
                        skipEmptyLines: true,
                        encoding: "UTF-8",
                        transformHeader: function(header) {
                            return header.trim();
                        },
                        transform: function(value) {
                            return value ? value.trim() : '';
                        },
                        complete: async function(results) {
                            if (results.errors.length && results.data.length === 0) {
                                return showToast('CSV 解析錯誤：' + (results.errors[0]?.message || '未知錯誤'), 'error');
                            }
                            
                            const validData = [];
                            const invalidRows = [];
                            
                            results.data.forEach((row, index) => {
                                // 支援多種欄位名稱
                                let name = '';
                                let username = '';
                                let role = '';
                                let password = '';
                                
                                for (const key in row) {
                                    const cleanKey = key.trim();
                                    if (cleanKey === '姓名' || cleanKey === 'name') {
                                        name = String(row[key] || '').trim();
                                    }
                                    if (cleanKey === '帳號' || cleanKey === 'username' || cleanKey === 'email') {
                                        username = String(row[key] || '').trim();
                                    }
                                    if (cleanKey === '權限' || cleanKey === 'role') {
                                        role = String(row[key] || '').trim();
                                    }
                                    if (cleanKey === '密碼' || cleanKey === 'password') {
                                        password = String(row[key] || '').trim();
                                    }
                                }
                                
                                // 驗證必填欄位
                                if (!name || !username || !role) {
                                    invalidRows.push({
                                        row: index + 2,
                                        name: name || '(空白)',
                                        username: username || '(空白)',
                                        role: role || '(空白)'
                                    });
                                    return;
                                }
                                
                                // 驗證權限值（支援英文代碼和中文名稱）
                                const roleMap = {
                                    'admin': 'admin',
                                    'manager': 'manager',
                                    'editor': 'editor',
                                    'viewer': 'viewer',
                                    '系統管理員': 'admin',
                                    '資料管理者': 'manager',
                                    '審查人員': 'editor',
                                    '檢視人員': 'viewer'
                                };
                                
                                const normalizedRole = roleMap[role] || roleMap[role.toLowerCase()];
                                if (!normalizedRole) {
                                    invalidRows.push({
                                        row: index + 2,
                                        error: `無效的權限值：${role}（應為：admin/系統管理員, manager/資料管理者, editor/審查人員, viewer/檢視人員）`
                                    });
                                    return;
                                }
                                
                                validData.push({ name, username, role: normalizedRole, password });
                            });
                            
                            if (validData.length === 0) {
                                let errorMsg = 'CSV 檔案中沒有有效的資料';
                                if (invalidRows.length > 0) {
                                    errorMsg += `\n發現 ${invalidRows.length} 筆資料格式錯誤`;
                                    console.error('無效行詳情：', invalidRows);
                                }
                                return showToast(errorMsg, 'error');
                            }
                            
                            try {
                                const res = await apiFetch('/api/users/import', {
                                    method: 'POST',
                                    body: JSON.stringify({ data: validData })
                                });
                                
                                if (res.status === 401) {
                                    return showToast('匯入錯誤：請先登入系統', 'error');
                                } else if (res.status === 403) {
                                    return showToast('匯入錯誤：您沒有權限執行此操作', 'error');
                                }
                                
                                let j;
                                try {
                                    j = await res.json();
                                } catch (parseError) {
                                    if (res.ok) {
                                        showToast('匯入可能已完成，但無法解析伺服器回應。請重新整理頁面確認結果。', 'warning');
                                        closeUserImportModal();
                                        await loadUsersPage(1);
                                        return;
                                    } else {
                                        return showToast('匯入錯誤：伺服器回應格式錯誤（狀態碼：' + res.status + '）', 'error');
                                    }
                                }
                                
                                if (res.ok && j.success === true) {
                                    const successCount = j.successCount || 0;
                                    let msg = `匯入完成：成功 ${successCount} 筆`;
                                    if (j.failed > 0) {
                                        msg += `，失敗 ${j.failed} 筆`;
                                        if (j.errors && j.errors.length > 0) {
                                            const errorPreview = j.errors.slice(0, 3).join('；');
                                            if (j.errors.length > 3) {
                                                msg += `\n（前3個錯誤：${errorPreview}...）`;
                                            } else {
                                                msg += `\n（錯誤：${errorPreview}）`;
                                            }
                                        }
                                    }
                                    
                                    if (successCount < validData.length) {
                                        msg += `\n⚠️ 注意：前端解析到 ${validData.length} 筆有效資料，但只成功匯入 ${successCount} 筆。可能是因為資料庫中已有重複的帳號。`;
                                    }
                                    
                                    showToast(msg, j.failed > 0 ? 'warning' : 'success');
                                    // 清除檔案選擇
                                    const fileInput = document.getElementById('userImportFile');
                                    if (fileInput) fileInput.value = '';
                                    await loadUsersPage(1);
                                    return;
                                } else {
                                    showToast(j.error || '匯入失敗', 'error');
                                    return;
                                }
                            } catch (e) {
                                if (e.message && (e.message.includes('Failed to fetch') || e.message.includes('NetworkError'))) {
                                    showToast('匯入錯誤：網路連線失敗', 'error');
                                } else {
                                    console.error('匯入時發生未預期錯誤：', e);
                                    showToast('匯入錯誤：' + e.message, 'error');
                                }
                            }
                        }
                    });
                } catch (e) {
                    showToast('讀取檔案錯誤：' + e.message, 'error');
                }
            };
            reader.readAsText(file, 'UTF-8');
        }

        // Plan Management
        // 保存檢查計畫管理頁面的狀態
        function savePlansViewState() {
            const state = {
                search: document.getElementById('planSearch')?.value || '',
                year: document.getElementById('planYearFilter')?.value || '',
                page: plansPage,
                pageSize: plansPageSize,
                sortField: plansSortField,
                sortDir: plansSortDir
            };
            sessionStorage.setItem('plansViewState', JSON.stringify(state));
        }
        
        // 恢復檢查計畫管理頁面的狀態
        function restorePlansViewState() {
            const saved = sessionStorage.getItem('plansViewState');
            if (!saved) return;
            
            try {
                const state = JSON.parse(saved);
                if (document.getElementById('planSearch')) document.getElementById('planSearch').value = state.search || '';
                if (document.getElementById('planYearFilter')) document.getElementById('planYearFilter').value = state.year || '';
                if (state.page) plansPage = state.page;
                if (state.pageSize) plansPageSize = state.pageSize;
                if (state.sortField) plansSortField = state.sortField;
                if (state.sortDir) plansSortDir = state.sortDir;
            } catch (e) {
                // 忽略解析錯誤
            }
        }
        
        async function loadPlansPage(page = 1) {
            plansPage = page;
            const plansPageSizeEl = document.getElementById('plansPageSize');
            if (plansPageSizeEl) {
                plansPageSize = parseInt(plansPageSizeEl.value, 10);
            }
            const q = document.getElementById('planSearch')?.value || '';
            const year = document.getElementById('planYearFilter')?.value || '';
            savePlansViewState();
            const params = new URLSearchParams({ page: plansPage, pageSize: plansPageSize, q, year, sortField: plansSortField, sortDir: plansSortDir, _t: Date.now() });
            try {
                const res = await fetch('/api/plans?' + params.toString());
                if (!res.ok) { 
                    const errorText = await res.text();
                    console.error('載入計畫失敗:', res.status, errorText);
                    showToast('載入計畫失敗: ' + (res.status === 500 ? '伺服器錯誤' : '請求失敗'), 'error'); 
                    return; 
                }
                const j = await res.json();
                planList = j.data || [];
                plansTotal = j.total || 0;
                plansPages = j.pages || 1;
                await renderPlans();
                renderPagination('plansPagination', plansPage, plansPages, 'loadPlansPage');
                // 更新年度選項
                updatePlanYearOptions();
            } catch (e) {
                // 錯誤已在伺服器 log 中記錄
                showToast('載入計畫錯誤: ' + e.message, 'error');
            }
        }
        function getInspectionTypeName(type) {
            const typeMap = { '1': '年度定期檢查', '2': '特別檢查', '3': '例行性檢查', '4': '臨時檢查', '5': '調查' };
            return typeMap[String(type)] || '-';
        }
        function getBusinessTypeName(code) {
            const typeMap = { 'OP': '運轉', 'CV': '土建', 'ME': '機務', 'EL': '電務', 'SM': '安全管理', 'AD': '營運／災防審核', 'OT': '其他／產管規劃' };
            return typeMap[String(code)] || '-';
        }
        
        async function renderPlans() {
            const tbody = document.getElementById('plansTableBody');
            if (!tbody) return;
            const rows = [];
            for (const p of planList) {
                let codesHtml = '';
                let datesHtml = '';
                let locationsHtml = '';
                let inspectorsHtml = '';
                let inspectionTypeHtml = '';
                try {
                    const scheduleRes = await fetch(`/api/plans/${p.id}/schedules?t=${Date.now()}`, { credentials: 'include' });
                    if (scheduleRes.ok) {
                        const scheduleData = await scheduleRes.json();
                        const schedules = scheduleData.data || [];
                        // 過濾掉沒有實際日期的排程（plan_number = '(手動)' 且 start_date 為 NULL 的記錄）
                        const validSchedules = schedules.filter(s => 
                            s.start_date && s.plan_number && s.plan_number !== '(手動)'
                        );
                        if (validSchedules.length > 0) {
                            codesHtml = validSchedules.map(s => `<div style="margin:2px 0; font-size:12px;">${s.plan_number || '-'}</div>`).join('');
                            datesHtml = validSchedules.map(s => {
                                const startDate = s.start_date ? s.start_date.slice(0, 10) : '-';
                                const endDate = s.end_date ? s.end_date.slice(0, 10) : null;
                                const range = endDate && endDate !== startDate ? `${startDate} ~ ${endDate}` : startDate;
                                return `<div style="margin:2px 0; font-size:12px;">${range}</div>`;
                            }).join('');
                            // 取得地點和人員資訊
                            locationsHtml = validSchedules.map(s => `<div style="margin:2px 0; font-size:12px;">${s.location || '-'}</div>`).join('');
                            inspectorsHtml = validSchedules.map(s => `<div style="margin:2px 0; font-size:12px;">${s.inspector || '-'}</div>`).join('');
                            // 取得檢查類別（取第一個有效的檢查類別）
                            const firstType = validSchedules[0]?.inspection_type;
                            inspectionTypeHtml = firstType ? `<div style="margin:2px 0; font-size:12px;">${getInspectionTypeName(firstType)}</div>` : '<span style="color:#94a3b8; font-size:12px;">—</span>';
                        } else {
                            inspectionTypeHtml = '<span style="color:#94a3b8; font-size:12px;">—</span>';
                        }
                    } else {
                        inspectionTypeHtml = '<span style="color:#94a3b8; font-size:12px;">—</span>';
                    }
                } catch (e) {
                    inspectionTypeHtml = '<span style="color:#94a3b8; font-size:12px;">—</span>';
                }
                const createdDate = p.created_at ? new Date(p.created_at).toISOString().slice(0, 10) : '-';
                const businessHtml = p.business ? getBusinessTypeName(p.business) : '<span style="color:#94a3b8;">—</span>';
                const plannedCountHtml = p.planned_count != null ? String(p.planned_count) : '<span style="color:#94a3b8;">—</span>';
                const scheduleCountVal = p.schedule_count != null ? p.schedule_count : 0;
                rows.push(`<tr>
                    <td data-label="選擇" style="padding:12px;text-align:center;">
                        <input type="checkbox" class="plan-check" value="${p.id}" onchange="updatePlansBatchDeleteBtn()">
                    </td>
                    <td data-label="年度" style="padding:12px;font-weight:600;">${p.year || '-'}</td>
                    <td data-label="檢查類別" style="padding:12px;">${inspectionTypeHtml}</td>
                    <td data-label="檢查計畫名稱" style="padding:12px;font-weight:600;">${p.name || '-'}</td>
                    <td data-label="業務類型" style="padding:12px;">${businessHtml}</td>
                    <td data-label="規劃次數" style="padding:12px;text-align:center;">${plannedCountHtml}</td>
                    <td data-label="已檢查次數" style="padding:12px;text-align:center;">${scheduleCountVal}</td>
                    <td data-label="檢查起訖日期" style="padding:12px;">${datesHtml || '<span style="color:#94a3b8; font-size:12px;">—</span>'}</td>
                    <td data-label="地點" style="padding:12px;">${locationsHtml || '<span style="color:#94a3b8; font-size:12px;">—</span>'}</td>
                    <td data-label="檢查人員" style="padding:12px;">${inspectorsHtml || '<span style="color:#94a3b8; font-size:12px;">—</span>'}</td>
                    <td data-label="取號編碼" style="padding:12px;">${codesHtml || '<span style="color:#94a3b8; font-size:12px;">無</span>'}</td>
                    <td data-label="開立事項數量" style="padding:12px;text-align:center;">${p.issue_count || 0}</td>
                    <td data-label="建立日期" style="padding:12px;">${createdDate}</td>
                    <td data-label="操作" style="padding:12px;">
                        <button class="btn btn-outline" style="padding:2px 6px;margin-right:4px;" onclick="openPlanModal('edit', ${p.id})">✏️</button>
                        <button class="btn btn-danger" style="padding:2px 6px;" onclick="deletePlan(${p.id})">🗑️</button>
                    </td>
                </tr>`);
            }
            tbody.innerHTML = rows.join('');
            updatePlansBatchDeleteBtn();
        }
        
        function toggleSelectAllPlans() {
            const selectAll = document.getElementById('selectAllPlans');
            const checkboxes = document.querySelectorAll('.plan-check');
            const isChecked = selectAll ? selectAll.checked : false;
            
            checkboxes.forEach(cb => cb.checked = isChecked);
            if (selectAll) selectAll.checked = isChecked;
            updatePlansBatchDeleteBtn();
        }
        
        function updatePlansBatchDeleteBtn() {
            const checkboxes = document.querySelectorAll('.plan-check:checked');
            const count = checkboxes.length;
            const container = document.getElementById('plansBatchActionContainer');
            const badge = document.getElementById('selectedPlansCountBadge');
            const selectAll = document.getElementById('selectAllPlans');
            
            if (container) {
                container.style.display = count > 0 ? 'block' : 'none';
            }
            if (badge) {
                badge.textContent = count > 0 ? `(${count})` : '';
            }
            if (selectAll) {
                const allChecked = checkboxes.length > 0 && checkboxes.length === document.querySelectorAll('.plan-check').length;
                selectAll.checked = allChecked;
            }
        }
        
        async function batchDeletePlans() {
            const checkboxes = document.querySelectorAll('.plan-check:checked');
            if (checkboxes.length === 0) {
                showToast('請至少選擇一筆資料', 'error');
                return;
            }
            
            const ids = Array.from(checkboxes).map(cb => parseInt(cb.value));
            const planNames = ids.map(id => {
                const plan = planList.find(p => p.id === id);
                return plan ? `${plan.name}${plan.year ? ` (${plan.year})` : ''}` : '';
            }).filter(Boolean);
            
            const confirmed = await showConfirmModal(`確定要刪除以下 ${ids.length} 筆檢查計畫嗎？\n\n${planNames.slice(0, 5).join('\n')}${planNames.length > 5 ? '\n...' : ''}\n\n此操作無法復原！`, '確定刪除', '取消');
            if (!confirmed) {
                return;
            }
            
            try {
                // 逐一刪除（因為需要記錄每個計畫的名稱）
                let successCount = 0;
                let failCount = 0;
                const errors = [];
                
                for (const id of ids) {
                    try {
                        const res = await apiFetch(`/api/plans/${id}`, { method: 'DELETE' });
                        const j = await res.json().catch(() => ({}));
                        
                        if (res.ok) {
                            successCount++;
                        } else {
                            failCount++;
                            const plan = planList.find(p => p.id === id);
                            const planName = plan ? `${plan.name}${plan.year ? ` (${plan.year})` : ''}` : `ID:${id}`;
                            errors.push(`${planName}: ${j.error || '刪除失敗'}`);
                        }
                    } catch (e) {
                        failCount++;
                        const plan = planList.find(p => p.id === id);
                        const planName = plan ? `${plan.name}${plan.year ? ` (${plan.year})` : ''}` : `ID:${id}`;
                        errors.push(`${planName}: ${e.message}`);
                    }
                }
                
                if (successCount > 0) {
                    let msg = `成功刪除 ${successCount} 筆`;
                    if (failCount > 0) {
                        msg += `，失敗 ${failCount} 筆`;
                        if (errors.length > 0) {
                            if (isDevelopment) console.warn('刪除錯誤詳情：', errors);
                        }
                    }
                    showToast(msg, failCount > 0 ? 'warning' : 'success');
                    loadPlansPage(plansPage);
                    loadPlanOptions();
                    // 如果當前在計畫規劃頁面，強制重新載入月曆（清除快取）
                    const scheduleTab = document.getElementById('subtab-plans-schedule');
                    if (scheduleTab && !scheduleTab.classList.contains('hidden')) {
                        scheduleMonthData = []; // 清除快取資料
                        loadScheduleForMonth();
                    }
                } else {
                    showToast(`刪除失敗：${errors.length > 0 ? errors[0] : '未知錯誤'}`, 'error');
                }
            } catch (e) {
                showToast('刪除時發生錯誤: ' + e.message, 'error');
            }
        }
        function plansSortBy(field) {
            if (plansSortField === field) {
                plansSortDir = plansSortDir === 'asc' ? 'desc' : 'asc';
            } else { 
                plansSortField = field; 
                plansSortDir = 'asc'; 
            }
            savePlansViewState();
            loadPlansPage(1);
        }
        function updatePlanYearOptions() {
            const yearSet = new Set();
            planList.forEach(p => { if (p.year) yearSet.add(p.year); });
            const years = Array.from(yearSet).sort((a, b) => b.localeCompare(a));
            const select = document.getElementById('planYearFilter');
            if (select) {
                const currentValue = select.value;
                const firstOption = select.options[0].outerHTML;
                select.innerHTML = firstOption + years.map(y => `<option value="${y}">${y}年</option>`).join('');
                if (currentValue) select.value = currentValue;
            }
        }

        // --- 檢查計畫規劃（月曆）---
        let scheduleCalendarYear = new Date().getFullYear();
        let scheduleCalendarMonth = new Date().getMonth() + 1;
        let scheduleMonthData = [];
        let holidayData = {};
        // --- 檢查計畫月曆看板（所有人可查看）---
        let dashboardCalendarYear = new Date().getFullYear();
        let dashboardCalendarMonth = new Date().getMonth() + 1;
        let dashboardMonthData = [];
        let dashboardSelectedYear = String(new Date().getFullYear() - 1911).replace(/\D/g, '').slice(-3).padStart(3, '0');

        function initScheduleCalendar() {
            const now = new Date();
            scheduleCalendarYear = now.getFullYear();
            scheduleCalendarMonth = now.getMonth() + 1;
            renderScheduleCalendar();
            loadScheduleForMonth();
            loadSchedulePlanOptions();
            const startDateInput = document.getElementById('scheduleStartDate');
            const endDateInput = document.getElementById('scheduleEndDate');
            if (startDateInput) {
                startDateInput.removeEventListener('change', scheduleOnDateChange);
                startDateInput.addEventListener('change', scheduleOnDateChange);
            }
            if (endDateInput) {
                endDateInput.removeEventListener('change', scheduleOnDateChange);
                endDateInput.addEventListener('change', scheduleOnDateChange);
            }
            scheduleClearForm();
        }

        function scheduleUpdateYearFromStartDate() {
            const v = (document.getElementById('scheduleStartDate') || {}).value || '';
            const wrap = document.getElementById('scheduleYearDisplayWrap');
            const display = document.getElementById('scheduleYearDisplay');
            if (!wrap || !display) return;
            if (!v) {
                wrap.style.display = 'none';
                display.textContent = '由開始日期自動換算';
                return;
            }
            const y = parseInt(v.slice(0, 4), 10);
            const roc = y - 1911;
            display.textContent = `民國 ${roc} 年`;
            wrap.style.display = 'block';
        }

        async function scheduleOnDateChange() {
            const startDateInput = document.getElementById('scheduleStartDate');
            const endDateInput = document.getElementById('scheduleEndDate');
            const v = startDateInput?.value || '';
            const sel = document.getElementById('scheduleSelectedDate');
            if (sel) sel.value = v;
            scheduleUpdateYearFromStartDate();
            
            // 限制結束日期必須在開始日期的月份內
            if (v && endDateInput) {
                const date = new Date(v);
                const year = date.getFullYear();
                const month = date.getMonth();
                const firstDay = `${year}-${String(month + 1).padStart(2, '0')}-01`;
                const lastDay = new Date(year, month + 1, 0).getDate();
                const lastDayStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
                
                // 設定結束日期的最小值和最大值（限制在同一個月）
                endDateInput.setAttribute('min', firstDay);
                endDateInput.setAttribute('max', lastDayStr);
                
                // 如果結束日期不在該月份範圍內，清空結束日期
                const currentEndDate = endDateInput.value;
                if (currentEndDate && (currentEndDate < firstDay || currentEndDate > lastDayStr)) {
                    endDateInput.value = '';
                }
            } else if (!v && endDateInput) {
                endDateInput.value = '';
                endDateInput.removeAttribute('min');
                endDateInput.removeAttribute('max');
            }
            
            if (!v) {
                scheduleRenderDayList('');
                hideSchedulePlanNumber();
                return;
            }
            const parts = v.split('-').map(Number);
            const y = parts[0], m = parts[1];
            if (y && m && (y !== scheduleCalendarYear || m !== scheduleCalendarMonth)) {
                scheduleCalendarYear = y;
                scheduleCalendarMonth = m;
                await loadScheduleForMonth();
                renderScheduleCalendar();
            }
            scheduleRenderDayList(v);
            
            // 如果有計畫資訊，自動計算取號編號
            await updateSchedulePlanNumber();
        }
        
        async function updateSchedulePlanNumber() {
            // 檢查是否有必要的資訊（不再需要 business）
            if (!schedulePlanDetails.railway || !schedulePlanDetails.inspection_type) {
                hideSchedulePlanNumber();
                return;
            }
            
            const startDateInput = document.getElementById('scheduleStartDate');
            const startDateVal = startDateInput?.value || '';
            if (!startDateVal) {
                hideSchedulePlanNumber();
                return;
            }
            
            try {
                // 計算年度（民國年）
                const adYear = parseInt(startDateVal.slice(0, 4), 10);
                const rocYear = adYear - 1911;
                const yr = String(rocYear).replace(/\D/g, '').slice(-3).padStart(3, '0');
                
                // 呼叫 API 取得下一個編號（不再需要 business 參數）
                const url = `/api/plan-schedule/next-number?year=${encodeURIComponent(yr)}&railway=${encodeURIComponent(schedulePlanDetails.railway)}&inspectionType=${encodeURIComponent(schedulePlanDetails.inspection_type)}&t=${Date.now()}`;
                const res = await fetch(url, {
                    credentials: 'include',
                    cache: 'no-store'
                });
                
                if (res.ok) {
                    const data = await res.json();
                    if (data.planNumber) {
                        showSchedulePlanNumber(data.planNumber);
                    } else {
                        hideSchedulePlanNumber();
                    }
                } else {
                    hideSchedulePlanNumber();
                }
            } catch (e) {
                console.error('取得取號編號失敗:', e);
                hideSchedulePlanNumber();
            }
        }
        
        function showSchedulePlanNumber(planNumber) {
            const displayDiv = document.getElementById('schedulePlanNumberDisplay');
            const valueInput = document.getElementById('schedulePlanNumberValue');
            if (displayDiv && valueInput) {
                valueInput.value = planNumber || '';
                displayDiv.style.display = 'block';
            }
        }
        
        function hideSchedulePlanNumber() {
            const displayDiv = document.getElementById('schedulePlanNumberDisplay');
            const valueInput = document.getElementById('schedulePlanNumberValue');
            if (displayDiv) displayDiv.style.display = 'none';
            if (valueInput) valueInput.value = '';
        }

        function schedulePrevMonth() {
            if (scheduleCalendarMonth === 1) {
                scheduleCalendarYear--;
                scheduleCalendarMonth = 12;
            } else {
                scheduleCalendarMonth--;
            }
            renderScheduleCalendar();
            loadScheduleForMonth();
        }

        function scheduleNextMonth() {
            if (scheduleCalendarMonth === 12) {
                scheduleCalendarYear++;
                scheduleCalendarMonth = 1;
            } else {
                scheduleCalendarMonth++;
            }
            renderScheduleCalendar();
            loadScheduleForMonth();
        }

        const SCHEDULE_PLAN_COLORS = ['#dbeafe', '#dcfce7', '#fef3c7', '#fce7f3', '#e0e7ff', '#d1fae5', '#fed7aa', '#e9d5ff'];
        const SCHEDULE_PLAN_TEXT_COLORS = ['#1e40af', '#166534', '#92400e', '#9d174d', '#3730a3', '#065f46', '#c2410c', '#6b21a8'];

        // 根據檢查類別決定顏色索引
        function schedulePlanColorIndex(inspectionType) {
            // 檢查類別：1=年度定期檢查, 2=特別檢查, 3=例行性檢查, 4=臨時檢查, 5=調查
            const typeMap = {
                '1': 0, // 年度定期檢查 - 藍色
                '2': 1, // 特別檢查 - 綠色
                '3': 2, // 例行性檢查 - 黃色
                '4': 3, // 臨時檢查 - 粉色
                '5': 4  // 調查 - 紫色
            };
            const type = String(inspectionType || '').trim();
            return typeMap[type] !== undefined ? typeMap[type] : 5; // 預設為第6個顏色（橙色）
        }

        function renderScheduleCalendar() {
            const title = document.getElementById('scheduleMonthTitle');
            const cal = document.getElementById('scheduleCalendar');
            if (!title || !cal) return;
            title.textContent = `${scheduleCalendarYear} 年 ${scheduleCalendarMonth} 月`;
            const y = scheduleCalendarYear;
            const m = scheduleCalendarMonth;
            const first = new Date(y, m - 1, 1);
            const last = new Date(y, m, 0);
            const startPad = first.getDay();
            const days = last.getDate();
            const pad = Array(startPad).fill(0).map((_, i) => `<div class="schedule-cal-day schedule-cal-pad"></div>`).join('');
            const dayCells = [];
            for (let d = 1; d <= days; d++) {
                const dateStr = `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                const holidayVal = holidayData[dateStr];
                const isHoliday = !!holidayVal;
                const holidayName = (typeof holidayVal === 'string' && holidayVal) ? holidayVal : '假日';
                const holidayClass = isHoliday ? 'schedule-cal-holiday' : '';
                const plansForDay = scheduleMonthData.filter(s => {
                    const startStr = (s.start_date || '').slice(0, 10);
                    const endStr = (s.end_date || '').slice(0, 10) || startStr;
                    // 確保單日也能顯示（開始日期=結束日期）
                    return startStr && dateStr >= startStr && dateStr <= endStr;
                });
                const hasPlan = plansForDay.length > 0;
                const colorIndices = plansForDay.map(s => schedulePlanColorIndex(s.inspection_type));
                const colorDots = colorIndices.map((idx, i) => {
                    const name = (plansForDay[i].plan_name || '').trim() || '未命名';
                    return `<span class="schedule-cal-color-dot" style="background:${SCHEDULE_PLAN_COLORS[idx]};" title="${name}"></span>`;
                }).join('');
                const colorDotsHtml = colorDots ? `<div class="schedule-cal-dots">${colorDots}</div>` : '';
                // 顯示計畫名稱、地點及人員
                const planItems = plansForDay.map((s, i) => {
                    const idx = colorIndices[i];
                    const tc = SCHEDULE_PLAN_TEXT_COLORS[idx] || '#1e3a8a';
                    const name = (s.plan_name || '').trim() || '未命名';
                    const location = (s.location || '').trim() || '';
                    const inspector = (s.inspector || '').trim() || '';
                    const planInfo = [];
                    planInfo.push(`<div class="schedule-cal-plan-name" style="color:${tc}; font-weight:600; font-size:12px; margin-bottom:2px;">${name}</div>`);
                    if (location) planInfo.push(`<div class="schedule-cal-plan-detail" style="color:#64748b; font-size:11px; margin-bottom:1px;">📍 ${location}</div>`);
                    if (inspector) planInfo.push(`<div class="schedule-cal-plan-detail" style="color:#64748b; font-size:11px;">👤 ${inspector}</div>`);
                    return `<div class="schedule-cal-plan-item" style="margin-bottom:3px; padding:2px 0;">${planInfo.join('')}</div>`;
                });
                const planText = planItems.length > 0 ? `<div class="schedule-cal-plan-names">${planItems.join('')}</div>` : '';
                const primaryColorIdx = hasPlan ? colorIndices[0] : 0;
                const colorClass = hasPlan ? `schedule-cal-plan-${primaryColorIdx}` : '';
                const finalHolidayClass = isHoliday ? 'schedule-cal-holiday' : '';
                const bgStyle = isHoliday ? 'background:#fef2f2 !important;' : '';
                const numColor = isHoliday ? 'color:#dc2626;' : '';
                const holidayTag = isHoliday ? `<span class="schedule-cal-holiday-tag" title="${holidayName}">假日</span>` : '';
                const dayNumWrap = hasPlan
                    ? `<div class="schedule-cal-day-num-wrap"><div class="schedule-cal-day-num" style="${numColor}">${d}</div><span class="schedule-cal-day-count">共 ${plansForDay.length} 筆</span></div>`
                    : `<div class="schedule-cal-day-num" style="${numColor}">${d}</div>`;
                dayCells.push(`<div class="schedule-cal-day ${hasPlan ? 'has-plan ' + colorClass : ''} ${finalHolidayClass}" style="${bgStyle}" data-date="${dateStr}" onclick="scheduleSelectDay('${dateStr}')">${dayNumWrap}${holidayTag}${colorDotsHtml}${planText}</div>`);
            }
            cal.innerHTML = `<div class="schedule-cal-head">日</div><div class="schedule-cal-head">一</div><div class="schedule-cal-head">二</div><div class="schedule-cal-head">三</div><div class="schedule-cal-head">四</div><div class="schedule-cal-head">五</div><div class="schedule-cal-head">六</div>${pad}${dayCells.join('')}`;
        }

        async function loadDashboardScheduleForMonth() {
            if (!dashboardSelectedYear) {
                dashboardMonthData = [];
                renderDashboardCalendar();
                return;
            }
            try {
                const res = await fetch(`/api/plan-schedule?year=${dashboardCalendarYear}&month=${dashboardCalendarMonth}&t=${Date.now()}`, {
                    credentials: 'include',
                    cache: 'no-store',
                    headers: { 'Cache-Control': 'no-cache' }
                });
                if (!res.ok) {
                    dashboardMonthData = [];
                    renderDashboardCalendar();
                    return;
                }
                const j = await res.json();
                const allData = j.data || [];
                dashboardMonthData = allData.filter(s => {
                    if (!s.start_date && !s.year) return false;
                    if (s.year && String(s.year).trim() === dashboardSelectedYear) return true;
                    if (s.start_date) {
                        const scheduleAdYear = parseInt(s.start_date.slice(0, 4), 10);
                        const scheduleRocYear = scheduleAdYear - 1911;
                        return String(scheduleRocYear).padStart(3, '0') === dashboardSelectedYear;
                    }
                    return false;
                });
                try {
                    const holidayRes = await fetch(`/api/holidays/${dashboardCalendarYear}?t=${Date.now()}`, {
                        credentials: 'include',
                        cache: 'no-store'
                    });
                    if (holidayRes.ok) {
                        const holidayJson = await holidayRes.json();
                        holidayData = {};
                        (holidayJson.data || []).forEach(h => {
                            if (h && h.date && h.isHoliday === true) {
                                const dateStr = String(h.date).slice(0, 10);
                                holidayData[dateStr] = (h.name || '').trim() || '假日';
                            }
                        });
                    }
                } catch (e) {}
                renderDashboardCalendar();
            } catch (e) {
                dashboardMonthData = [];
                renderDashboardCalendar();
            }
        }

        function renderDashboardCalendar() {
            const title = document.getElementById('dashboardScheduleMonthTitle');
            const cal = document.getElementById('dashboardScheduleCalendar');
            if (!title || !cal) return;
            title.textContent = `${dashboardCalendarYear} 年 ${dashboardCalendarMonth} 月`;
            const y = dashboardCalendarYear;
            const m = dashboardCalendarMonth;
            const first = new Date(y, m - 1, 1);
            const last = new Date(y, m, 0);
            const startPad = first.getDay();
            const days = last.getDate();
            const pad = Array(startPad).fill(0).map((_, i) => `<div class="schedule-cal-day schedule-cal-pad"></div>`).join('');
            const dayCells = [];
            for (let d = 1; d <= days; d++) {
                const dateStr = `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                const holidayVal = holidayData[dateStr];
                const isHoliday = !!holidayVal;
                const holidayName = (typeof holidayVal === 'string' && holidayVal) ? holidayVal : '假日';
                const plansForDay = dashboardMonthData.filter(s => {
                    const startStr = (s.start_date || '').slice(0, 10);
                    const endStr = (s.end_date || '').slice(0, 10) || startStr;
                    return startStr && dateStr >= startStr && dateStr <= endStr;
                });
                const hasPlan = plansForDay.length > 0;
                const colorIndices = plansForDay.map(s => schedulePlanColorIndex(s.inspection_type));
                const colorDots = colorIndices.map((idx, i) => {
                    const name = (plansForDay[i].plan_name || '').trim() || '未命名';
                    return `<span class="schedule-cal-color-dot" style="background:${SCHEDULE_PLAN_COLORS[idx]};" title="${name}"></span>`;
                }).join('');
                const colorDotsHtml = colorDots ? `<div class="schedule-cal-dots">${colorDots}</div>` : '';
                const planItems = plansForDay.map((s, i) => {
                    const idx = colorIndices[i];
                    const tc = SCHEDULE_PLAN_TEXT_COLORS[idx] || '#1e3a8a';
                    const name = (s.plan_name || '').trim() || '未命名';
                    const location = (s.location || '').trim() || '';
                    const inspector = (s.inspector || '').trim() || '';
                    const planInfo = [];
                    planInfo.push(`<div class="schedule-cal-plan-name" style="color:${tc}; font-weight:600; font-size:12px; margin-bottom:2px;">${name}</div>`);
                    if (location) planInfo.push(`<div class="schedule-cal-plan-detail" style="color:#64748b; font-size:11px; margin-bottom:1px;">📍 ${location}</div>`);
                    if (inspector) planInfo.push(`<div class="schedule-cal-plan-detail" style="color:#64748b; font-size:11px;">👤 ${inspector}</div>`);
                    return `<div class="schedule-cal-plan-item" style="margin-bottom:3px; padding:2px 0;">${planInfo.join('')}</div>`;
                });
                const planText = planItems.length > 0 ? `<div class="schedule-cal-plan-names">${planItems.join('')}</div>` : '';
                const primaryColorIdx = hasPlan ? colorIndices[0] : 0;
                const colorClass = hasPlan ? `schedule-cal-plan-${primaryColorIdx}` : '';
                const finalHolidayClass = isHoliday ? 'schedule-cal-holiday' : '';
                const bgStyle = isHoliday ? 'background:#fef2f2 !important;' : '';
                const numColor = isHoliday ? 'color:#dc2626;' : '';
                const holidayTag = isHoliday ? `<span class="schedule-cal-holiday-tag" title="${holidayName}">假日</span>` : '';
                const dayNumWrap = hasPlan
                    ? `<div class="schedule-cal-day-num-wrap"><div class="schedule-cal-day-num" style="${numColor}">${d}</div><span class="schedule-cal-day-count">共 ${plansForDay.length} 筆</span></div>`
                    : `<div class="schedule-cal-day-num" style="${numColor}">${d}</div>`;
                const onclickAttr = hasPlan ? ` onclick="dashboardScheduleSelectDay('${dateStr}')"` : '';
                dayCells.push(`<div class="schedule-cal-day ${hasPlan ? 'has-plan ' + colorClass : ''} ${finalHolidayClass}" style="${bgStyle}" data-date="${dateStr}"${onclickAttr}>${dayNumWrap}${holidayTag}${colorDotsHtml}${planText}</div>`);
            }
            cal.innerHTML = `<div class="schedule-cal-head">日</div><div class="schedule-cal-head">一</div><div class="schedule-cal-head">二</div><div class="schedule-cal-head">三</div><div class="schedule-cal-head">四</div><div class="schedule-cal-head">五</div><div class="schedule-cal-head">六</div>${pad}${dayCells.join('')}`;
        }

        function dashboardScheduleSelectDay(dateStr) {
            dashboardRenderDayList(dateStr);
        }

        function dashboardRenderDayList(dateStr) {
            const box = document.getElementById('dashboardScheduleDayListBody');
            if (!box) return;
            if (!dateStr) {
                box.innerHTML = '點選月曆上的日期可查看該日詳細檢查內容';
                return;
            }
            const list = dashboardMonthData.filter(s => {
                const startStr = (s.start_date || '').slice(0, 10);
                const endStr = (s.end_date || '').slice(0, 10) || startStr;
                return dateStr >= startStr && dateStr <= endStr;
            });
            if (list.length === 0) {
                box.innerHTML = '當日尚無排程';
                return;
            }
            box.innerHTML = list.map(s => {
                const startDate = (s.start_date || '').slice(0, 10);
                const endDate = (s.end_date || '').slice(0, 10);
                const range = endDate && endDate !== startDate ? `${startDate} ~ ${endDate}` : startDate;
                const location = (s.location || '').trim() || '';
                const inspector = (s.inspector || '').trim() || '';
                const planNumber = (s.plan_number || '').trim() || '';
                return `<div style="margin-bottom:10px; padding:10px; background:#f1f5f9; border-radius:6px; border-left:3px solid #3b82f6;">
                    <div style="font-weight:600; font-size:14px; margin-bottom:6px; color:#334155;">
                        ${s.plan_name || '-'}
                        ${planNumber ? `<span style="margin-left:8px; font-size:12px; color:#3b82f6; font-weight:500;">[${planNumber}]</span>` : ''}
                    </div>
                    <div style="color:#64748b; font-size:12px; margin-bottom:6px;">📅 ${range}</div>
                    ${location ? `<div style="color:#475569; font-size:12px; margin-bottom:4px;">📍 地點：<span style="font-weight:500;">${location}</span></div>` : ''}
                    ${inspector ? `<div style="color:#475569; font-size:12px;">👤 人員：<span style="font-weight:500;">${inspector}</span></div>` : ''}
                    ${!location && !inspector ? '<div style="color:#94a3b8; font-size:11px; font-style:italic;">無地點及人員資訊</div>' : ''}
                </div>`;
            }).join('');
        }

        function dashboardSchedulePrevMonth() {
            if (dashboardCalendarMonth === 1) {
                dashboardCalendarYear--;
                dashboardCalendarMonth = 12;
            } else {
                dashboardCalendarMonth--;
            }
            const dayListBody = document.getElementById('dashboardScheduleDayListBody');
            if (dayListBody) dayListBody.innerHTML = '點選月曆上的日期可查看該日詳細檢查內容';
            if (dashboardSelectedYear) {
                loadDashboardScheduleForMonth();
            } else {
                renderDashboardCalendar();
            }
        }

        function dashboardScheduleNextMonth() {
            if (dashboardCalendarMonth === 12) {
                dashboardCalendarYear++;
                dashboardCalendarMonth = 1;
            } else {
                dashboardCalendarMonth++;
            }
            const dayListBody = document.getElementById('dashboardScheduleDayListBody');
            if (dayListBody) dayListBody.innerHTML = '點選月曆上的日期可查看該日詳細檢查內容';
            if (dashboardSelectedYear) {
                loadDashboardScheduleForMonth();
            } else {
                renderDashboardCalendar();
            }
        }

        function initDashboardCalendar() {
            const now = new Date();
            dashboardCalendarYear = now.getFullYear();
            dashboardCalendarMonth = now.getMonth() + 1;
            if (dashboardSelectedYear) {
                loadDashboardScheduleForMonth();
            }
        }

        async function printScheduleCalendar() {
            const title = document.getElementById('scheduleMonthTitle');
            if (!title) return;
            
            const y = scheduleCalendarYear;
            const m = scheduleCalendarMonth;
            const first = new Date(y, m - 1, 1);
            const last = new Date(y, m, 0);
            const startPad = first.getDay();
            const days = last.getDate();
            const pad = Array(startPad).fill(0).map(() => '<div class="schedule-cal-day schedule-cal-pad"></div>').join('');
            const dayCells = [];
            
            for (let d = 1; d <= days; d++) {
                const dateStr = `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                const holidayVal = holidayData[dateStr];
                const isHoliday = !!holidayVal;
                const holidayName = (typeof holidayVal === 'string' && holidayVal) ? holidayVal : '假日';
                const plansForDay = scheduleMonthData.filter(s => {
                    const startStr = (s.start_date || '').slice(0, 10);
                    const endStr = (s.end_date || '').slice(0, 10) || startStr;
                    // 確保單日也能顯示（開始日期=結束日期）
                    return startStr && dateStr >= startStr && dateStr <= endStr;
                });
                const hasPlan = plansForDay.length > 0;
                const colorIndices = plansForDay.map(s => schedulePlanColorIndex(s.inspection_type));
                const colorDots = colorIndices.map((idx, i) => {
                    const name = (plansForDay[i].plan_name || '').trim() || '未命名';
                    return `<span class="schedule-cal-color-dot" style="background:${SCHEDULE_PLAN_COLORS[idx]};" title="${name}"></span>`;
                }).join('');
                const colorDotsHtml = colorDots ? `<div class="schedule-cal-dots">${colorDots}</div>` : '';
                // 顯示計畫名稱、地點及人員（列印版本）
                const planItems = plansForDay.map((s, i) => {
                    const idx = colorIndices[i];
                    const tc = SCHEDULE_PLAN_TEXT_COLORS[idx] || '#1e3a8a';
                    const name = (s.plan_name || '').trim() || '未命名';
                    const location = (s.location || '').trim() || '';
                    const inspector = (s.inspector || '').trim() || '';
                    const planInfo = [];
                    planInfo.push(`<div class="schedule-cal-plan-name" style="color:${tc}; font-weight:600; font-size:11px; margin-bottom:2px;">${name}</div>`);
                    if (location) planInfo.push(`<div class="schedule-cal-plan-detail" style="color:#64748b; font-size:10px; margin-bottom:1px;">📍 ${location}</div>`);
                    if (inspector) planInfo.push(`<div class="schedule-cal-plan-detail" style="color:#64748b; font-size:10px;">👤 ${inspector}</div>`);
                    return `<div class="schedule-cal-plan-item" style="margin-bottom:2px; padding:1px 0;">${planInfo.join('')}</div>`;
                });
                const planText = planItems.length > 0 ? `<div class="schedule-cal-plan-names">${planItems.join('')}</div>` : '';
                const primaryColorIdx = hasPlan ? colorIndices[0] : 0;
                const colorClass = hasPlan ? `schedule-cal-plan-${primaryColorIdx}` : '';
                const bgColor = hasPlan ? SCHEDULE_PLAN_COLORS[primaryColorIdx] : '#fff';
                const finalHolidayClass = isHoliday ? 'schedule-cal-holiday' : '';
                const finalBgColor = isHoliday ? '#fef2f2' : bgColor;
                const numColor = isHoliday ? '#dc2626' : '#334155';
                const holidayTag = isHoliday ? `<span class="schedule-cal-holiday-tag" title="${holidayName}">假日</span>` : '';
                dayCells.push(`<div class="schedule-cal-day ${hasPlan ? 'has-plan ' + colorClass : ''} ${finalHolidayClass}" style="background:${finalBgColor};"><div class="schedule-cal-day-num" style="color:${numColor};font-weight:700;">${d}</div>${holidayTag}${colorDotsHtml}${planText}</div>`);
            }
            
            const calendarHtml = `<div class="schedule-cal-head">日</div><div class="schedule-cal-head">一</div><div class="schedule-cal-head">二</div><div class="schedule-cal-head">三</div><div class="schedule-cal-head">四</div><div class="schedule-cal-head">五</div><div class="schedule-cal-head">六</div>${pad}${dayCells.join('')}`;
            const monthTitle = title.textContent;
            
            const printWindow = window.open('', '_blank');
            printWindow.document.write(`
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="UTF-8">
                    <title>${monthTitle} 檢查計畫月曆</title>
                    <style>
                        @page { size: A4 landscape; margin: 18mm 28mm; }
                        * { box-sizing: border-box; margin: 0; padding: 0; }
                        html, body { height: 100%; width: 100%; overflow: hidden; font-family: "Microsoft JhengHei", "微軟正黑體", "Noto Sans TC", Arial, sans-serif; }
                        body { background: #fafafa; padding: 0 0 12px 0; }
                        .print-header {
                            text-align: center; margin-bottom: 14px; padding: 10px 0 10px 0; border-bottom: 3px solid #1e40af;
                            background: linear-gradient(180deg, #eff6ff 0%, #fff 100%);
                        }
                        .print-header h1 { font-size: 20px; font-weight: 700; color: #1e3a8a; letter-spacing: 0.02em; margin: 0; }
                        .print-header .sub { font-size: 11px; color: #64748b; margin-top: 3px; }
                        .schedule-calendar {
                            display: grid; grid-template-columns: repeat(7, 1fr); gap: 3px;
                            background: #64748b; border: 3px solid #475569; border-radius: 8px; overflow: hidden;
                            width: 100%; max-width: 100%; page-break-inside: avoid;
                        }
                        .schedule-cal-head {
                            background: linear-gradient(180deg, #1e40af 0%, #1d4ed8 100%); color: #fff;
                            padding: 10px 6px; text-align: center; font-weight: 700; font-size: 13px;
                            display: flex; align-items: center; justify-content: center; min-height: 34px;
                            border-right: 2px solid rgba(255,255,255,0.3);
                        }
                        .schedule-cal-head:last-child { border-right: none; }
                        .schedule-cal-day {
                            border: 2px solid #94a3b8; padding: 10px 6px; min-height: 100px; background: #fff;
                            display: flex; flex-direction: column; align-items: flex-start; justify-content: flex-start;
                            page-break-inside: avoid; overflow: hidden;
                            border-right: 2px solid #94a3b8;
                            border-bottom: 2px solid #94a3b8;
                        }
                        .schedule-cal-day:nth-child(7n) { border-right: none; }
                        .schedule-cal-day-num { font-weight: 700; font-size: 15px; margin-bottom: 4px; color: #0f172a; }
                        .schedule-cal-holiday { background: #fef2f2 !important; border-color: #fca5a5 !important; }
                        .schedule-cal-holiday .schedule-cal-day-num { color: #b91c1c; }
                        .schedule-cal-holiday-tag { font-size: 10px; color: #b91c1c; font-weight: 600; margin-bottom: 2px; display: block; }
                        .schedule-cal-plan-names { font-size: 11px; line-height: 1.45; margin-top: 4px; width: 100%; word-break: break-word; overflow: hidden; }
                        .schedule-cal-plan-name { font-weight: 600; font-size: 11px; display: block; margin-bottom: 2px; }
                        .schedule-cal-plan-detail { font-size: 10px; color: #475569; display: block; margin-bottom: 1px; }
                        .schedule-cal-plan-item { margin-bottom: 4px; padding-bottom: 2px; border-bottom: 1px solid #e2e8f0; }
                        .schedule-cal-plan-item:last-child { border-bottom: none; }
                        .schedule-cal-pad { background: #f1f5f9; border: 1px solid #e2e8f0; }
                        .schedule-cal-dots { display: flex; gap: 3px; flex-wrap: wrap; margin-bottom: 3px; }
                        .schedule-cal-color-dot { width: 6px; height: 6px; border-radius: 50%; display: inline-block; border: 1px solid rgba(0,0,0,0.15); }
                        .schedule-cal-day.schedule-cal-plan-0 { background: #dbeafe; border-color: #93c5fd; }
                        .schedule-cal-day.schedule-cal-plan-1 { background: #dcfce7; border-color: #86efac; }
                        .schedule-cal-day.schedule-cal-plan-2 { background: #fef9c3; border-color: #fde047; }
                        .schedule-cal-day.schedule-cal-plan-3 { background: #fce7f3; border-color: #f9a8d4; }
                        .schedule-cal-day.schedule-cal-plan-4 { background: #e0e7ff; border-color: #a5b4fc; }
                        .schedule-cal-day.schedule-cal-plan-5 { background: #d1fae5; border-color: #6ee7b7; }
                        .schedule-cal-day.schedule-cal-plan-6 { background: #ffedd5; border-color: #fdba74; }
                        .schedule-cal-day.schedule-cal-plan-7 { background: #ede9fe; border-color: #c4b5fd; }
                        @media print {
                            @page { size: A4 landscape; margin: 18mm 28mm; }
                            body { background: #fff; padding: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                            .print-header { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                            .schedule-cal-head, .schedule-cal-day.schedule-cal-plan-0, .schedule-cal-day.schedule-cal-plan-1,
                            .schedule-cal-day.schedule-cal-plan-2, .schedule-cal-day.schedule-cal-plan-3,
                            .schedule-cal-day.schedule-cal-plan-4, .schedule-cal-day.schedule-cal-plan-5,
                            .schedule-cal-day.schedule-cal-plan-6, .schedule-cal-day.schedule-cal-plan-7,
                            .schedule-cal-holiday { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                            .schedule-cal-day { min-height: 95px; padding: 8px 5px; border: 2px solid #94a3b8 !important; }
                            .schedule-cal-day-num { font-size: 14px; }
                            .schedule-calendar { border: 3px solid #475569 !important; }
                        }
                    </style>
                </head>
                <body>
                    <div class="print-header">
                        <h1>${monthTitle} 檢查計畫月曆</h1>
                        <div class="sub">SMS 開立事項查詢與審查系統 · 列印日期：${new Date().toLocaleDateString('zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit' })}</div>
                    </div>
                    <div class="schedule-calendar">${calendarHtml}</div>
                </body>
                </html>
            `);
            printWindow.document.close();
            setTimeout(() => {
                printWindow.print();
            }, 300);
        }

        async function loadScheduleForMonth() {
            try {
                const res = await fetch(`/api/plan-schedule?year=${scheduleCalendarYear}&month=${scheduleCalendarMonth}&t=${Date.now()}`, { 
                    credentials: 'include',
                    cache: 'no-store',
                    headers: { 'Cache-Control': 'no-cache' }
                });
                if (!res.ok) { 
                    scheduleMonthData = []; 
                    renderScheduleCalendar();
                    return; 
                }
                const j = await res.json();
                scheduleMonthData = j.data || [];
                
                try {
                    const holidayRes = await fetch(`/api/holidays/${scheduleCalendarYear}?t=${Date.now()}`, {
                        credentials: 'include',
                        cache: 'no-store'
                    });
                    if (holidayRes.ok) {
                        const holidayJson = await holidayRes.json();
                        holidayData = {};
                        (holidayJson.data || []).forEach(h => {
                            if (h && h.date && h.isHoliday === true) {
                                const dateStr = String(h.date).slice(0, 10);
                                holidayData[dateStr] = (h.name || '').trim() || '假日';
                            }
                        });
                        renderScheduleCalendar();
                    } else {
                        showToast('無法載入休假日，請稍後再試', 'error');
                        holidayData = {};
                        renderScheduleCalendar();
                    }
                } catch (e) {
                    showToast('無法載入休假日，請稍後再試', 'error');
                    holidayData = {};
                    renderScheduleCalendar();
                }
            } catch (e) {
                scheduleMonthData = [];
            }
            renderScheduleCalendar();
        }

        function scheduleSelectDay(dateStr) {
            const startDateInput = document.getElementById('scheduleStartDate');
            if (startDateInput) {
                startDateInput.value = dateStr;
                scheduleUpdateYearFromStartDate();
            }
            const sel = document.getElementById('scheduleSelectedDate');
            if (sel) sel.value = dateStr;
            scheduleRenderDayList(dateStr);
        }

        function scheduleRenderDayList(dateStr) {
            const box = document.getElementById('scheduleDayListBody');
            if (!box) return;
            if (!dateStr) {
                box.innerHTML = '點選月曆日期後顯示';
                return;
            }
            const list = scheduleMonthData.filter(s => {
                const startStr = (s.start_date || '').slice(0, 10);
                const endStr = (s.end_date || '').slice(0, 10) || startStr;
                return dateStr >= startStr && dateStr <= endStr;
            });
            if (list.length === 0) {
                box.innerHTML = '當日尚無排程';
                return;
            }
            box.innerHTML = list.map(s => {
                const startDate = (s.start_date || '').slice(0, 10);
                const endDate = (s.end_date || '').slice(0, 10);
                const range = endDate && endDate !== startDate ? `${startDate} ~ ${endDate}` : startDate;
                const location = (s.location || '').trim() || '';
                const inspector = (s.inspector || '').trim() || '';
                const planNumber = (s.plan_number || '').trim() || '';
                return `<div style="margin-bottom:10px; padding:10px; background:#f1f5f9; border-radius:6px; border-left:3px solid #3b82f6;">
                    <div style="font-weight:600; font-size:14px; margin-bottom:6px; color:#334155;">
                        ${s.plan_name || '-'}
                        ${planNumber ? `<span style="margin-left:8px; font-size:12px; color:#3b82f6; font-weight:500;">[${planNumber}]</span>` : ''}
                    </div>
                    <div style="color:#64748b; font-size:12px; margin-bottom:6px;">📅 ${range}</div>
                    ${location ? `<div style="color:#475569; font-size:12px; margin-bottom:4px;">📍 地點：<span style="font-weight:500;">${location}</span></div>` : ''}
                    ${inspector ? `<div style="color:#475569; font-size:12px;">👤 人員：<span style="font-weight:500;">${inspector}</span></div>` : ''}
                    ${!location && !inspector ? '<div style="color:#94a3b8; font-size:11px; font-style:italic;">無地點及人員資訊</div>' : ''}
                </div>`;
            }).join('');
        }

        function scheduleClearForm() {
            const startDateInput = document.getElementById('scheduleStartDate');
            const endDateInput = document.getElementById('scheduleEndDate');
            const sel = document.getElementById('scheduleSelectedDate');
            if (startDateInput) {
                startDateInput.value = '';
                // 移除結束日期的限制
                if (endDateInput) {
                    endDateInput.removeAttribute('min');
                    endDateInput.removeAttribute('max');
                }
            }
            if (endDateInput) endDateInput.value = '';
            if (sel) sel.value = '';
            const planSelect = document.getElementById('schedulePlanSelect');
            const location = document.getElementById('scheduleLocation');
            const inspector = document.getElementById('scheduleInspector');
            const planInfoDiv = document.getElementById('schedulePlanInfo');
            if (planSelect) planSelect.value = '';
            if (location) location.value = '';
            if (inspector) inspector.value = '';
            if (planInfoDiv) planInfoDiv.style.display = 'none';
            schedulePlanDetails = {};
            hideSchedulePlanNumber();
            scheduleUpdateYearFromStartDate();
            const dayList = document.getElementById('scheduleDayListBody');
            if (dayList) dayList.textContent = '點選月曆日期後顯示';
        }
        
        let schedulePlanDetails = {};
        
        function onScheduleYearFilterChange() {
            loadSchedulePlanOptions();
        }
        
        async function loadSchedulePlanOptions() {
            try {
                const yearSelect = document.getElementById('scheduleYearFilter');
                const yearParam = yearSelect && yearSelect.value ? `&year=${encodeURIComponent(yearSelect.value)}` : '';
                const res = await fetch(`/api/options/plans?t=${Date.now()}${yearParam}`, { credentials: 'include' });
                if (!res.ok) return;
                const j = await res.json();
                const select = document.getElementById('schedulePlanSelect');
                if (!select) return;
                const currentValue = select.value;
                select.innerHTML = '<option value="">請選擇已建立的檢查計畫</option>';
                if (j.data && Array.isArray(j.data)) {
                    j.data.forEach(p => {
                        const opt = document.createElement('option');
                        opt.value = `${p.name}|||${p.year}`;
                        opt.textContent = `${p.name} (${p.year}年)`;
                        select.appendChild(opt);
                    });
                }
                if (currentValue) select.value = currentValue;
                
                if (yearSelect && (!yearSelect.options || yearSelect.options.length <= 1)) {
                    const allRes = await fetch(`/api/options/plans?t=${Date.now()}`, { credentials: 'include' });
                    if (allRes.ok) {
                        const allJ = await allRes.json();
                        const years = [...new Set((allJ.data || []).map(p => p.year).filter(Boolean))].sort((a, b) => b.localeCompare(a));
                        const firstOpt = yearSelect.options[0] ? yearSelect.options[0].outerHTML : '<option value="">全部年度</option>';
                        yearSelect.innerHTML = firstOpt + years.map(y => `<option value="${y}">${y}年</option>`).join('');
                    }
                }
                
                // 重新設計的選擇計畫處理函數 - 簡化版本
                select.onchange = async function() {
                    const selectedValue = select.value;
                    
                    // 清空之前的資料
                    schedulePlanDetails = {};
                    hideSchedulePlanNumber();
                    const planInfoDiv = document.getElementById('schedulePlanInfo');
                    if (planInfoDiv) planInfoDiv.style.display = 'none';
                    
                    // 如果沒有選擇，直接返回
                    if (!selectedValue) {
                        return;
                    }
                    
                    // 解析計畫名稱和年度
                    const parts = selectedValue.split('|||');
                    if (parts.length !== 2) {
                        showToast('計畫資訊格式錯誤，請重新選擇', 'error');
                        select.value = '';
                        return;
                    }
                    
                    const planName = parts[0].trim();
                    const planYear = parts[1].trim();
                    
                    if (!planName || !planYear) {
                        showToast('計畫資訊不完整，請重新選擇', 'error');
                        select.value = '';
                        return;
                    }
                    
                    // 呼叫 API 取得計畫詳情
                    try {
                        const apiUrl = `/api/plans/by-name?name=${encodeURIComponent(planName)}&year=${encodeURIComponent(planYear)}`;
                        const response = await fetch(apiUrl, {
                            method: 'GET',
                            credentials: 'include',
                            headers: {
                                'Accept': 'application/json'
                            }
                        });
                        
                        // 處理回應
                        if (!response.ok) {
                            let errorMessage = '無法取得計畫資訊';
                            
                            try {
                                const errorData = await response.json();
                                errorMessage = errorData.message || errorData.error || errorMessage;
                            } catch (e) {
                                // 如果無法解析 JSON，使用狀態碼
                                if (response.status === 404) {
                                    errorMessage = '找不到該計畫';
                                } else if (response.status === 500) {
                                    errorMessage = '伺服器錯誤，請稍後再試';
                                } else if (response.status === 503) {
                                    errorMessage = '服務暫時不可用，請稍後再試';
                                }
                            }
                            
                            showToast(errorMessage, 'error');
                            select.value = '';
                            return;
                        }
                        
                        // 解析成功回應
                        const result = await response.json();
                        
                        if (!result.data || !Array.isArray(result.data) || result.data.length === 0) {
                            showToast('無法取得計畫資訊', 'error');
                            select.value = '';
                            return;
                        }
                        
                        const plan = result.data[0];
                        
                        // 處理計畫資料
                        const railway = (plan.railway && plan.railway !== '-') ? String(plan.railway).trim() : '';
                        const inspection_type = (plan.inspection_type && plan.inspection_type !== '-') ? String(plan.inspection_type).trim() : '';
                        
                        // 儲存到變數
                        schedulePlanDetails = {
                            plan_name: planName,
                            year: planYear,
                            railway: railway,
                            inspection_type: inspection_type
                        };
                        
                        // 檢查是否有完整資訊
                        if (!railway || !inspection_type) {
                            if (result.warning) {
                                showToast(result.warning, 'warning');
                            } else {
                                showToast('該計畫缺少必要資訊（鐵路機構、檢查類別），請先在計畫管理中編輯', 'warning');
                            }
                            schedulePlanDetails = {};
                            select.value = '';
                            return;
                        }
                        
                        // 顯示計畫資訊
                        const railwaySpan = document.getElementById('schedulePlanRailway');
                        const inspectionTypeSpan = document.getElementById('schedulePlanInspectionType');
                        
                        if (planInfoDiv && railwaySpan && inspectionTypeSpan) {
                            const railwayNames = { 'T': '臺鐵', 'H': '高鐵', 'A': '林鐵', 'S': '糖鐵' };
                            const inspectionTypeNames = { '1': '年度定期檢查', '2': '特別檢查', '3': '例行性檢查', '4': '臨時檢查', '5': '調查' };
                            
                            railwaySpan.textContent = railwayNames[railway] || railway;
                            inspectionTypeSpan.textContent = inspectionTypeNames[inspection_type] || inspection_type;
                            planInfoDiv.style.display = 'block';
                        }
                        
                        // 如果有開始日期，自動計算取號編號
                        await updateSchedulePlanNumber();
                        
                    } catch (error) {
                        showToast('無法取得計畫資訊，請稍後再試', 'error');
                        select.value = '';
                        schedulePlanDetails = {};
                    }
                };
                
                // 為開始日期添加事件監聽，當日期改變時更新取號編號
                const startDateInput = document.getElementById('scheduleStartDate');
                if (startDateInput) {
                    startDateInput.removeEventListener('change', handleScheduleStartDateChange);
                    startDateInput.addEventListener('change', handleScheduleStartDateChange);
                }
            } catch (e) {
                console.error('載入計畫選項失敗:', e);
            }
        }
        
        async function handleScheduleStartDateChange() {
            await updateSchedulePlanNumber();
        }

        async function scheduleSubmitPlan() {
            const planSelect = document.getElementById('schedulePlanSelect');
            const planValue = planSelect ? planSelect.value : '';
            if (!planValue) {
                showToast('請選擇檢查計畫', 'error');
                return;
            }
            const [planName, planYear] = planValue.split('|||');
            if (!planName || !planYear) {
                showToast('計畫資訊不完整，請重新選擇', 'error');
                return;
            }
            const startDateVal = (document.getElementById('scheduleStartDate') || {}).value;
            const endDateVal = (document.getElementById('scheduleEndDate') || {}).value;
            const locationValue = (document.getElementById('scheduleLocation') || {}).value?.trim();
            const inspectorValue = (document.getElementById('scheduleInspector') || {}).value?.trim();
            
            if (!startDateVal) {
                showToast('請選擇開始日期', 'error');
                return;
            }
            if (!endDateVal) {
                showToast('請選擇結束日期', 'error');
                return;
            }
            if (endDateVal < startDateVal) {
                showToast('結束日期不能早於開始日期', 'error');
                return;
            }
            if (!locationValue) {
                showToast('請填寫地點', 'error');
                return;
            }
            if (!inspectorValue) {
                showToast('請填寫檢查人員', 'error');
                return;
            }
            
            // 如果沒有計畫詳情，嘗試重新載入（簡化版本，不再需要 business）
            if (!schedulePlanDetails.railway || !schedulePlanDetails.inspection_type) {
                try {
                    const apiUrl = `/api/plans/by-name?name=${encodeURIComponent(planName)}&year=${encodeURIComponent(planYear)}`;
                    const planRes = await fetch(apiUrl, {
                        method: 'GET',
                        credentials: 'include',
                        headers: { 'Accept': 'application/json' }
                    });
                    
                    if (!planRes.ok) {
                        const errorData = await planRes.json().catch(() => ({}));
                        showToast(errorData.message || errorData.error || '無法取得計畫資訊', 'error');
                        return;
                    }
                    
                    const planData = await planRes.json();
                    if (!planData.data || !Array.isArray(planData.data) || planData.data.length === 0) {
                        showToast('找不到該計畫，請重新選擇', 'error');
                        return;
                    }
                    
                    const plan = planData.data[0];
                    const railway = (plan.railway && plan.railway !== '-') ? String(plan.railway).trim() : '';
                    const inspection_type = (plan.inspection_type && plan.inspection_type !== '-') ? String(plan.inspection_type).trim() : '';
                    
                    schedulePlanDetails = {
                        plan_name: planName,
                        year: planYear,
                        railway: railway,
                        inspection_type: inspection_type
                    };
                    
                    if (!railway || !inspection_type) {
                        showToast('該計畫缺少必要資訊（鐵路機構、檢查類別），請先在計畫管理中編輯該計畫', 'error');
                        return;
                    }
                } catch (e) {
                    showToast('無法取得計畫資訊，請稍後再試', 'error');
                    return;
                }
            }
            
            // 最終確認計畫詳情是否完整（不再需要 business）
            if (!schedulePlanDetails.railway || !schedulePlanDetails.inspection_type) {
                showToast('該計畫缺少必要資訊（鐵路機構、檢查類別），請先在計畫管理中編輯該計畫', 'error');
                return;
            }
            const adYear = parseInt(startDateVal.slice(0, 4), 10);
            const rocYear = adYear - 1911;
            const yr = String(rocYear).replace(/\D/g, '').slice(-3).padStart(3, '0');
            const planNumberInput = document.getElementById('schedulePlanNumberValue');
            const customPlanNumber = planNumberInput && planNumberInput.value ? String(planNumberInput.value).trim() : '';
            const payload = {
                plan_name: planName,
                start_date: startDateVal,
                end_date: endDateVal,
                year: yr,
                railway: schedulePlanDetails.railway,
                inspection_type: schedulePlanDetails.inspection_type,
                business: null,
                location: locationValue,
                inspector: inspectorValue
            };
            if (customPlanNumber) payload.plan_number = customPlanNumber;
            try {
                const res = await apiFetch('/api/plan-schedule', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const j = await res.json().catch(() => ({}));
                if (!res.ok) {
                    // 根據不同的錯誤狀態碼顯示不同的錯誤訊息
                    let errorMsg = j.error || '儲存失敗';
                    if (res.status === 400) {
                        errorMsg = j.error || '資料格式錯誤，請檢查輸入的資料';
                    } else if (res.status === 403) {
                        errorMsg = '權限不足，請確認您的帳號權限';
                    } else if (res.status === 500) {
                        errorMsg = '伺服器錯誤：' + (j.error || '請稍後再試');
                    }
                    showToast(errorMsg, 'error');
                    return;
                }
                
                // 顯示實際產生的編號
                if (j.planNumber) {
                    showToast(`已上傳，取號：${j.planNumber}`, 'success');
                } else {
                    showToast('已上傳成功', 'success');
                }
                
                // 重新載入月曆資料
                await loadScheduleForMonth();
                scheduleRenderDayList(startDateVal);
                
                // 清空表單欄位，回復預設值
                scheduleClearForm();
                
                loadPlanOptions();
                loadSchedulePlanOptions();
            } catch (e) {
                let errorMsg = '儲存失敗';
                if (e.message) {
                    if (e.message.includes('CSRF')) {
                        errorMsg = '安全驗證失敗，請重新整理頁面後再試';
                    } else if (e.message.includes('fetch')) {
                        errorMsg = '網路連線錯誤，請檢查網路連線後再試';
                    } else {
                        errorMsg = '儲存失敗，請稍後再試';
                    }
                }
                showToast(errorMsg, 'error');
            }
        }

        // 處理開始日期變更，自動設定結束日期為同月最後一天
        function handlePlanStartDateChange() {
            const startDateInput = document.getElementById('planStartDate');
            const endDateInput = document.getElementById('planEndDate');
            if (!startDateInput || !endDateInput) return;
            
            const startDateVal = startDateInput.value;
            if (!startDateVal) {
                endDateInput.value = '';
                return;
            }
            
            // 計算該月的最後一天
            const date = new Date(startDateVal);
            const year = date.getFullYear();
            const month = date.getMonth();
            const lastDay = new Date(year, month + 1, 0).getDate();
            const endDate = `${year}-${String(month + 1).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
            
            // 如果結束日期為空或早於開始日期，自動設定為該月最後一天
            if (!endDateInput.value || endDateInput.value < startDateVal) {
                endDateInput.value = endDate;
            }
        }
        
        function applyScheduleToPlanForm(schedule, basePlan) {
            const scheduleIdInput = document.getElementById('targetScheduleId');
            const nameInput = document.getElementById('planName');
            const yearInput = document.getElementById('planYear');
            const startInput = document.getElementById('planStartDate');
            const endInput = document.getElementById('planEndDate');
            const railwaySelect = document.getElementById('planRailway');
            const inspectionSelect = document.getElementById('planInspectionType');
            const locationInput = document.getElementById('planLocation');
            const inspectorInput = document.getElementById('planInspector');
            const planNumberInput = document.getElementById('planScheduleNumber');
            const planNumberRow = document.getElementById('planNumberRow');
            if (!schedule || !nameInput || !yearInput || !railwaySelect || !inspectionSelect) return;
            
            if (scheduleIdInput) scheduleIdInput.value = schedule.id || '';
            nameInput.value = (basePlan && basePlan.name) ? basePlan.name : (schedule.plan_name || '');
            let rocYear = '';
            if (schedule.start_date) {
                const adYear = parseInt(schedule.start_date.slice(0, 4), 10);
                if (!Number.isNaN(adYear)) rocYear = String(adYear - 1911).padStart(3, '0');
            }
            if (!rocYear && basePlan && basePlan.year) rocYear = String(basePlan.year).padStart(3, '0');
            if (yearInput) yearInput.value = rocYear;
            if (startInput) startInput.value = schedule.start_date ? schedule.start_date.slice(0, 10) : '';
            if (endInput) endInput.value = schedule.end_date ? schedule.end_date.slice(0, 10) : '';
            if (planNumberInput) planNumberInput.value = (schedule.plan_number || '').trim();
            if (planNumberRow) planNumberRow.style.display = 'flex';
            railwaySelect.value = schedule.railway || '';
            inspectionSelect.value = schedule.inspection_type || '';
            if (locationInput) locationInput.value = schedule.location || '';
            if (inspectorInput) inspectorInput.value = schedule.inspector || '';
        }
        
        function selectPlanSchedule(scheduleId) {
            const idNum = Number(scheduleId);
            if (!currentPlanSchedules || currentPlanSchedules.length === 0) return;
            const schedule = currentPlanSchedules.find(s => Number(s.id) === idNum);
            if (!schedule) return;
            const planId = Number((document.getElementById('targetPlanId') || {}).value || '0');
            const basePlan = planList.find(p => Number(p.id) === planId) || {};
            applyScheduleToPlanForm(schedule, basePlan);
            document.querySelectorAll('.plan-schedule-item').forEach(el => el.classList.remove('selected'));
            const el = document.querySelector(`.plan-schedule-item[data-schedule-id="${scheduleId}"]`);
            if (el) el.classList.add('selected');
            var datesSection = document.getElementById('planDatesSection');
            if (datesSection) datesSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        
        async function openPlanModal(mode, id) {
            const m = document.getElementById('planModal');
            const t = document.getElementById('planModalTitle');
            const planDetailsGroup = document.getElementById('planDetailsGroup');
            const planDetailsGroup2 = document.getElementById('planDetailsGroup2');
            const planBusinessGroup = document.getElementById('planBusinessGroup');
            const planPlannedCountGroup = document.getElementById('planPlannedCountGroup');
            const planDatesSection = document.getElementById('planDatesSection');
            const planSchedulesSection = document.getElementById('planSchedulesSection');
            const planStartDateGroup = document.getElementById('planStartDateGroup');
            const planEndDateGroup = document.getElementById('planEndDateGroup');
            const planLocationInspectorRow = document.getElementById('planLocationInspectorRow');
            
            if (mode === 'create') {
                t.innerText = '新增檢查計畫';
                document.getElementById('targetPlanId').value = '';
                document.getElementById('targetScheduleId').value = '';
                document.getElementById('planName').value = '';
                document.getElementById('planYear').value = '';
                document.getElementById('planRailway').value = '';
                document.getElementById('planInspectionType').value = '';
                const planBusinessEl = document.getElementById('planBusiness');
                if (planBusinessEl) planBusinessEl.value = '';
                const planPlannedCountEl = document.getElementById('planPlannedCount');
                if (planPlannedCountEl) planPlannedCountEl.value = '';
                currentPlanSchedules = [];
                const planSchedulesList = document.getElementById('planSchedulesList');
                if (planSchedulesList) planSchedulesList.innerHTML = '<div class="plan-schedules-empty">尚未有檢查行程，請先至「填寫檢查行程」建立。</div>';
                if (planDetailsGroup) planDetailsGroup.style.display = 'block';
                if (planDetailsGroup2) planDetailsGroup2.style.display = 'block';
                if (planBusinessGroup) planBusinessGroup.style.display = 'block';
                if (planPlannedCountGroup) planPlannedCountGroup.style.display = 'block';
                if (planDatesSection) planDatesSection.style.display = 'none';
                if (planSchedulesSection) planSchedulesSection.style.display = 'none';
                if (m) m.classList.add('open');
                return;
            }
            
            const p = planList.find(x => x.id === id) || {};
            t.innerText = '編輯檢查計畫';
            document.getElementById('targetPlanId').value = p.id || '';
            document.getElementById('targetScheduleId').value = '';
            const planPlannedCountEl = document.getElementById('planPlannedCount');
            if (planPlannedCountEl) planPlannedCountEl.value = p.planned_count != null ? String(p.planned_count) : '';
            const planBusinessEl = document.getElementById('planBusiness');
            if (planBusinessEl) planBusinessEl.value = p.business || '';
            if (planDetailsGroup) planDetailsGroup.style.display = 'block';
            if (planDetailsGroup2) planDetailsGroup2.style.display = 'block';
            if (planBusinessGroup) planBusinessGroup.style.display = 'block';
            if (planPlannedCountGroup) planPlannedCountGroup.style.display = 'block';
            if (planDatesSection) planDatesSection.style.display = 'block';
            if (planSchedulesSection) planSchedulesSection.style.display = 'block';
            if (planStartDateGroup) planStartDateGroup.style.display = 'block';
            if (planEndDateGroup) planEndDateGroup.style.display = 'block';
            if (planLocationInspectorRow) planLocationInspectorRow.style.display = 'flex';
            const planStartDateInput = document.getElementById('planStartDate');
            const planEndDateInput = document.getElementById('planEndDate');
            if (planStartDateInput && planEndDateInput) {
                planStartDateInput.removeEventListener('change', handlePlanStartDateChange);
                planStartDateInput.addEventListener('change', handlePlanStartDateChange);
            }
            try {
                const scheduleRes = await fetch(`/api/plans/${p.id}/schedules?t=${Date.now()}`, { credentials: 'include' });
                const planSchedulesList = document.getElementById('planSchedulesList');
                currentPlanSchedules = [];
                if (scheduleRes.ok) {
                    const scheduleData = await scheduleRes.json();
                    const schedules = scheduleData.data || [];
                    currentPlanSchedules = schedules;
                    const validSchedules = schedules.filter(s => s.plan_number && s.plan_number !== '(手動)');
                    if (planSchedulesList) {
                        if (!validSchedules.length) {
                            planSchedulesList.innerHTML = '<div class="plan-schedules-empty">尚未有檢查行程，請先至「填寫檢查行程」或在此建立第一筆行程。</div>';
                        } else {
                            planSchedulesList.innerHTML = validSchedules.map((s, idx) => {
                                const start = (s.start_date || '').slice(0, 10);
                                const end = (s.end_date || '').slice(0, 10);
                                const range = end && end !== start ? `${start} ~ ${end}` : start;
                                const planNumber = (s.plan_number || '').trim();
                                const loc = (s.location || '').trim();
                                const insp = (s.inspector || '').trim();
                                return `<div class="plan-schedule-item" data-schedule-id="${s.id}" onclick="selectPlanSchedule(${s.id})">
                                    <div class="schedule-row-top"><span>第 ${idx + 1} 筆行程</span>${planNumber ? `<span class="plan-number">${planNumber}</span>` : ''}</div>
                                    <div class="schedule-row-detail">📅 ${range || '尚未設定日期'}${loc ? ' · 📍 ' + loc : ''}${insp ? ' · 👤 ' + insp : ''}</div>
                                </div>`;
                            }).join('');
                        }
                    }
                    if (validSchedules.length > 0) {
                        applyScheduleToPlanForm(validSchedules[0], p);
                        const firstEl = planSchedulesList && planSchedulesList.querySelector('.plan-schedule-item');
                        if (firstEl) firstEl.classList.add('selected');
                    } else {
                        document.getElementById('planName').value = p.name || '';
                        document.getElementById('planYear').value = p.year || '';
                        if (document.getElementById('planLocation')) document.getElementById('planLocation').value = '';
                        if (document.getElementById('planInspector')) document.getElementById('planInspector').value = '';
                        var planNumberRow = document.getElementById('planNumberRow');
                        if (planNumberRow) planNumberRow.style.display = 'none';
                        var planScheduleNumber = document.getElementById('planScheduleNumber');
                        if (planScheduleNumber) planScheduleNumber.value = '';
                    }
                } else if (planSchedulesList) {
                    planSchedulesList.innerHTML = '<div class="plan-schedules-empty" style="color:#f97316;">無法載入行程列表，請稍後再試。</div>';
                }
            } catch (e) {
                const planSchedulesList = document.getElementById('planSchedulesList');
                if (planSchedulesList) planSchedulesList.innerHTML = '<div class="plan-schedules-empty" style="color:#f97316;">載入行程列表時發生錯誤。</div>';
            }
            if (m) m.classList.add('open');
        }
        function closePlanModal() {
            const m = document.getElementById('planModal');
            if (m) m.classList.remove('open');
        }
        function openPlanImportModal() {
            const m = document.getElementById('planImportModal');
            if (m) {
                const fileInput = document.getElementById('planImportFile');
                if (fileInput) fileInput.value = '';
                m.classList.add('open');
            }
        }
        function closePlanImportModal() {
            const m = document.getElementById('planImportModal');
            if (m) m.classList.remove('open');
        }
        async function importPlansCSV() {
            const fileInput = document.getElementById('planImportFile');
            if (!fileInput) return showToast('找不到檔案選擇器', 'error');
            const file = fileInput.files[0];
            if (!file) return showToast('請選擇 CSV 檔案', 'error');
            
            const reader = new FileReader();
            reader.onload = async function(e) {
                try {
                    const csv = e.target.result;
                    Papa.parse(csv, {
                        header: true,
                        skipEmptyLines: false, // 改為 false，手動處理空行
                        encoding: "UTF-8",
                        transformHeader: function(header) {
                            // 統一處理欄位名稱，去除空白
                            return header.trim();
                        },
                        transform: function(value) {
                            // 去除值的前後空白
                            return value ? value.trim() : '';
                        },
                        complete: async function(results) {
                            if (results.errors.length && results.data.length === 0) {
                                return showToast('CSV 解析錯誤：' + (results.errors[0]?.message || '未知錯誤'), 'error');
                            }
                            
                            // 顯示解析結果統計
                            
                            // 過濾掉空行，支援多種欄位名稱
                            const validData = [];
                            const invalidRows = [];
                            
                            results.data.forEach((row, index) => {
                                // 檢查是否為完全空行（所有值都為空或只有空白）
                                const isEmptyRow = Object.values(row).every(val => !val || String(val).trim() === '');
                                if (isEmptyRow) {
                                    // 完全空行，跳過
                                    return;
                                }
                                
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
                                        railway = String(row[key] || '').trim();
                                    } else if (cleanKey === '檢查類別' || cleanKey === 'inspection_type' || cleanKey === 'inspectionType') {
                                        inspection_type = String(row[key] || '').trim();
                                    } else if (cleanKey === '業務類別' || cleanKey === 'business') {
                                        business = String(row[key] || '').trim();
                                    }
                                }
                                
                                if (name && start_date && end_date) {
                                    validData.push({ name, year, start_date, end_date, railway, inspection_type, business: null }); // business 改為 null
                                } else {
                                    invalidRows.push({
                                        row: index + 2,
                                        name: name || '(空白)',
                                        start_date: start_date || '(空白)',
                                        end_date: end_date || '(空白)',
                                        rawRow: row
                                    });
                                }
                            });
                            
                            // 已移除調試資訊（只在伺服器 log 中記錄）
                            
                            if (validData.length === 0) {
                                let errorMsg = 'CSV 檔案中沒有有效的資料';
                                if (invalidRows.length > 0) {
                                    errorMsg += `\n發現 ${invalidRows.length} 筆資料缺少必要欄位（計畫名稱或開始日期）`;
                                    console.error('無效行詳情：', invalidRows);
                                }
                                return showToast(errorMsg, 'error');
                            }
                            
                            try {
                                const res = await apiFetch('/api/plans/import', {
                                    method: 'POST',
                                    body: JSON.stringify({ data: validData })
                                });
                                
                                // 先檢查 HTTP 狀態碼
                                if (res.status === 401) {
                                    return showToast('匯入錯誤：請先登入系統', 'error');
                                } else if (res.status === 403) {
                                    return showToast('匯入錯誤：您沒有權限執行此操作', 'error');
                                }
                                
                                // 嘗試解析 JSON
                                let j;
                                let text;
                                try {
                                    text = await res.text();
                                    j = JSON.parse(text);
                                } catch (parseError) {
                                    // 如果解析失敗，檢查狀態碼
                                    if (res.ok) {
                                        // 如果狀態碼是 OK，但解析失敗，可能是格式問題，但實際可能已成功
                                        showToast('匯入可能已完成，但無法解析伺服器回應。請重新整理頁面確認結果。', 'warning');
                                        closePlanImportModal();
                                        await loadPlansPage(1);
                                        await loadPlanOptions();
                                        return;
                                    } else {
                                        return showToast('匯入錯誤：伺服器回應格式錯誤（狀態碼：' + res.status + '）', 'error');
                                    }
                                }
                                
                                // 檢查回應是否成功
                                // 後端回應格式：{ success: true, successCount: 數字, failed: 數字, errors: [], skipped: 數字 }
                                if (res.ok && j.success === true) {
                                    // 取得成功筆數
                                    const successCount = j.successCount || 0;
                                    
                                    let msg = `匯入完成：成功 ${successCount} 筆`;
                                    if (j.skipped > 0) {
                                        msg += `，跳過空行 ${j.skipped} 筆`;
                                    }
                                    if (j.failed > 0) {
                                        msg += `，失敗 ${j.failed} 筆`;
                                        if (j.errors && j.errors.length > 0) {
                                            // 錯誤詳情已在伺服器 log 中記錄
                                            // 顯示前3個錯誤（避免訊息過長）
                                            const errorPreview = j.errors.slice(0, 3).join('；');
                                            if (j.errors.length > 3) {
                                                msg += `\n（前3個錯誤：${errorPreview}...）`;
                                            } else {
                                                msg += `\n（錯誤：${errorPreview}）`;
                                            }
                                        }
                                    }
                                    
                                    // 如果成功筆數少於有效資料筆數，顯示警告
                                    if (successCount < validData.length) {
                                        msg += `\n⚠️ 注意：前端解析到 ${validData.length} 筆有效資料，但只成功匯入 ${successCount} 筆。可能是因為資料庫中已有重複的計畫名稱。`;
                                    }
                                    
                                    showToast(msg, j.failed > 0 ? 'warning' : 'success');
                                    closePlanImportModal();
                                    
                                    // 重新載入計畫列表和選項
                                    await loadPlansPage(1);
                                    await loadPlanOptions();
                                    
                                    // 確保選項已更新
                                    setTimeout(() => {
                                        loadPlanOptions();
                                    }, 500);
                                    return; // 明確返回，避免繼續執行 catch 區塊
                                } else {
                                    // 如果狀態碼不是 OK 或 success 為 false
                                    showToast(j.error || '匯入失敗', 'error');
                                    return; // 明確返回
                                }
                            } catch (e) {
                                // 只有在真正的網路錯誤或無法處理的錯誤時才顯示錯誤
                                // 如果已經在 try 區塊中顯示了成功或錯誤訊息，這裡不應該再顯示
                                // 檢查錯誤類型，避免重複顯示
                                if (e.name === 'TypeError' && (e.message.includes('text') || e.message.includes('already been read'))) {
                                    // 如果已經讀取過 text，可能是重複讀取的問題
                                    // 不顯示錯誤，因為可能已經成功匯入了
                                    return;
                                }
                                // 只有在真正的網路錯誤時才顯示
                                if (e.message.includes('Failed to fetch') || e.message.includes('NetworkError')) {
                                    showToast('匯入錯誤：網路連線失敗', 'error');
                                } else {
                                    // 其他未預期的錯誤，但不要顯示，因為可能已經成功匯入了
                                    console.error('匯入時發生未預期錯誤（可能已成功）：', e);
                                }
                            }
                        }
                    });
                } catch (e) {
                    showToast('讀取檔案錯誤：' + e.message, 'error');
                }
            };
            reader.readAsText(file, 'UTF-8');
        }
        function downloadPlanCSVTemplate() {
            const csv = '計畫名稱,年度,業務類型,規劃檢查次數,開始日期,結束日期,鐵路機構,檢查類別\n上半年定期檢查,113,OP,2,2024-01-01,2024-06-30,T,1\n下半年定期檢查,113,CV,1,2024-07-01,2024-12-31,T,1\n特別檢查,113,,1,2024-03-15,2024-03-20,H,2';
            const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = '檢查計畫匯入範例.csv';
            link.click();
        }
        async function submitPlan() {
            const planId = document.getElementById('targetPlanId').value;
            const scheduleId = document.getElementById('targetScheduleId').value;
            const name = document.getElementById('planName').value.trim();
            const year = document.getElementById('planYear').value.trim();
            const startDate = (document.getElementById('planStartDate') || {}).value;
            const endDate = (document.getElementById('planEndDate') || {}).value;
            const railway = document.getElementById('planRailway').value;
            const inspectionType = document.getElementById('planInspectionType').value;
            const planBusinessEl = document.getElementById('planBusiness');
            const business = planBusinessEl && planBusinessEl.value ? planBusinessEl.value : null;
            const planPlannedCountEl = document.getElementById('planPlannedCount');
            const plannedCount = planPlannedCountEl && planPlannedCountEl.value !== '' ? planPlannedCountEl.value : null;
            const planLocationEl = document.getElementById('planLocation');
            const planInspectorEl = document.getElementById('planInspector');
            const location = planLocationEl ? planLocationEl.value.trim() : '';
            const inspector = planInspectorEl ? planInspectorEl.value.trim() : '';
            
            if (!planId) {
                if (!name) return showToast('請輸入計畫名稱', 'error');
                if (!year) return showToast('請輸入年度', 'error');
                if (!railway) return showToast('請選擇鐵路機構', 'error');
                if (!inspectionType) return showToast('請選擇檢查類別', 'error');
                try {
                    const res = await apiFetch('/api/plans', {
                        method: 'POST',
                        body: JSON.stringify({
                            name,
                            year: year.replace(/\D/g, '').slice(-3).padStart(3, '0'),
                            railway,
                            inspection_type: inspectionType,
                            business,
                            planned_count: plannedCount
                        })
                    });
                    const j = await res.json();
                    if (res.ok) {
                        showToast('新增成功');
                        closePlanModal();
                        loadPlansPage(plansPage || 1);
                        loadPlanOptions();
                        loadSchedulePlanOptions();
                    } else {
                        showToast(j.error || j.message || '新增失敗', 'error');
                    }
                } catch (e) {
                    showToast('操作失敗，請稍後再試', 'error');
                }
                return;
            }
            
            if (scheduleId) {
                if (!name) return showToast('請輸入計畫名稱', 'error');
                if (!startDate) return showToast('請選擇開始日期', 'error');
                if (!endDate) return showToast('請選擇結束日期', 'error');
                if (!railway || !inspectionType) return showToast('請填寫鐵路機構、檢查類別', 'error');
                const adYear = parseInt(startDate.slice(0, 4), 10);
                const rocYear = adYear - 1911;
                const yearStr = String(rocYear).padStart(3, '0');
                const planNumberEl = document.getElementById('planScheduleNumber');
                const customPlanNumber = planNumberEl && planNumberEl.value ? String(planNumberEl.value).trim() : null;
                const payload = {
                    plan_name: name,
                    start_date: startDate,
                    end_date: endDate,
                    year: yearStr,
                    railway,
                    inspection_type: inspectionType,
                    business: null,
                    location: location || null,
                    inspector: inspector || null
                };
                if (customPlanNumber) payload.plan_number = customPlanNumber;
                try {
                    const res = await apiFetch(`/api/plan-schedule/${scheduleId}`, {
                        method: 'PUT',
                        body: JSON.stringify(payload)
                    });
                    const j = await res.json();
                    if (res.ok) {
                        showToast('行程更新成功');
                        closePlanModal();
                        loadPlansPage(plansPage || 1);
                        loadPlanOptions();
                        const scheduleTab = document.getElementById('subtab-plans-schedule');
                        if (scheduleTab && !scheduleTab.classList.contains('hidden')) {
                            scheduleMonthData = [];
                            loadScheduleForMonth();
                        }
                    } else {
                        showToast(j.error || '更新失敗', 'error');
                    }
                } catch (e) {
                    showToast('操作失敗：' + e.message, 'error');
                }
                return;
            }
            
            if (!name) return showToast('請輸入計畫名稱', 'error');
            if (!year) return showToast('請輸入年度', 'error');
            try {
                const res = await apiFetch(`/api/plans/${planId}`, {
                    method: 'PUT',
                    body: JSON.stringify({
                        name,
                        year: year.replace(/\D/g, '').slice(-3).padStart(3, '0'),
                        business,
                        planned_count: plannedCount
                    })
                });
                const j = await res.json();
                if (res.ok) {
                    showToast('計畫更新成功');
                    closePlanModal();
                    loadPlansPage(plansPage || 1);
                    loadPlanOptions();
                    loadSchedulePlanOptions();
                } else {
                    showToast(j.error || '更新失敗', 'error');
                }
            } catch (e) {
                showToast('操作失敗：' + e.message, 'error');
            }
        }
        async function deletePlan(id) {
            const confirmed = await showConfirmModal('確定要刪除這個計畫嗎？\n\n此操作無法復原！', '確定刪除', '取消');
            if (!confirmed) return;
            try {
                const res = await apiFetch(`/api/plans/${id}`, { method: 'DELETE' });
                const j = await res.json();
                if (res.ok) {
                    showToast('刪除成功');
                    loadPlansPage(1);
                    loadPlanOptions();
                    // 如果當前在計畫規劃頁面，強制重新載入月曆（清除快取）
                    const scheduleTab = document.getElementById('subtab-plans-schedule');
                    if (scheduleTab && !scheduleTab.classList.contains('hidden')) {
                        scheduleMonthData = []; // 清除快取資料
                        loadScheduleForMonth();
                    }
                } else {
                    showToast(j.error || '刪除失敗', 'error');
                }
            } catch (e) {
                showToast('刪除失敗', 'error');
            }
        }

        // Profile
        function openProfileModal() { document.getElementById('myProfileName').value = currentUser.name || ''; document.getElementById('myProfilePwd').value = ''; document.getElementById('profileModal').classList.add('open'); }
        async function submitProfile() { 
            const name = document.getElementById('myProfileName').value, 
                pwd = document.getElementById('myProfilePwd').value,
                pwdConfirm = document.getElementById('myProfilePwdConfirm').value; 
            
            try { 
                // 如果有提供密碼，驗證複雜度和確認密碼
                if (pwd) {
                    if (!pwdConfirm) {
                        return showToast('請輸入確認密碼', 'error');
                    }
                    if (pwd !== pwdConfirm) {
                        return showToast('密碼與確認密碼不符', 'error');
                    }
                    const validation = validatePasswordFrontend(pwd);
                    if (!validation.valid) {
                        return showToast(validation.message, 'error');
                    }
                } else if (pwdConfirm) {
                    // 如果只填了確認密碼但沒填密碼
                    return showToast('請輸入新密碼', 'error');
                }
                
                const res = await apiFetch('/api/auth/profile', { 
                    method: 'PUT', 
                    body: JSON.stringify({ name, password: pwd }) 
                }); 
                if (res.ok) { 
                    showToast('更新成功，請重新登入'); 
                    document.getElementById('profileModal').classList.remove('open'); 
                    // 清空密碼欄位
                    document.getElementById('myProfilePwd').value = '';
                    document.getElementById('myProfilePwdConfirm').value = '';
                    logout(); 
                } else { 
                    const j = await res.json(); 
                    showToast(j.error || '更新失敗', 'error'); 
                } 
            } catch (e) { 
                showToast('更新失敗', 'error'); 
            } 
        }

        // 首次登入密碼更新函數
        async function submitChangePassword() {
            const newPwd = document.getElementById('changePwdNew').value;
            const confirmPwd = document.getElementById('changePwdConfirm').value;
            const errorEl = document.getElementById('changePwdError');
            
            // 清除之前的錯誤訊息
            if (errorEl) {
                errorEl.style.display = 'none';
                errorEl.innerText = '';
            }
            
            // 驗證輸入
            if (!newPwd || !confirmPwd) {
                if (errorEl) {
                    errorEl.innerText = '請輸入新密碼和確認密碼';
                    errorEl.style.display = 'block';
                }
                return;
            }
            
            if (newPwd !== confirmPwd) {
                if (errorEl) {
                    errorEl.innerText = '兩次輸入的密碼不一致';
                    errorEl.style.display = 'block';
                }
                return;
            }
            
            // 驗證密碼複雜度
            const validation = validatePasswordFrontend(newPwd);
            if (!validation.valid) {
                if (errorEl) {
                    errorEl.innerText = validation.message;
                    errorEl.style.display = 'block';
                }
                return;
            }
            
            try {
                const res = await apiFetch('/api/auth/change-password', {
                    method: 'POST',
                    body: JSON.stringify({ password: newPwd })
                });
                
                if (res.ok) {
                    showToast('密碼更新成功，請重新登入', 'success');
                    // 關閉模態框
                    const modal = document.getElementById('changePasswordModal');
                    if (modal) {
                        modal.style.display = 'none';
                    }
                    // 清除表單
                    document.getElementById('changePwdNew').value = '';
                    document.getElementById('changePwdConfirm').value = '';
                    // 登出並重新導向到登入頁
                    setTimeout(() => {
                        logout();
                    }, 1000);
                } else {
                    const data = await res.json();
                    if (errorEl) {
                        errorEl.innerText = data.error || '密碼更新失敗';
                        errorEl.style.display = 'block';
                    }
                }
            } catch (e) {
                if (errorEl) {
                    errorEl.innerText = '連線錯誤，請稍後再試';
                    errorEl.style.display = 'block';
                }
            }
        }

        function toggleEditMode(edit) { 
            document.getElementById('viewModeContent').classList.toggle('hidden', edit); 
            document.getElementById('editModeContent').classList.toggle('hidden', !edit); 
            document.getElementById('drawerTitle').innerText = edit ? "審查事項" : "詳細資料"; 
            if (edit) { 
                if (!currentEditItem) return;
                // 清除所有編輯欄位，避免前一個事項的資料殘留
                document.getElementById('editId').value = currentEditItem.id; 
                
                // 編號
                document.getElementById('editHeaderNumber').textContent = currentEditItem.number || '';
                
                // 檢查計畫
                document.getElementById('editHeaderPlanName').textContent = currentEditItem.plan_name || currentEditItem.planName || '(未設定)';
                
                // 檢查種類
                const insName = currentEditItem.inspectionCategoryName || currentEditItem.inspection_category_name || '-';
                document.getElementById('editHeaderInspection').textContent = insName;
                
                // 分組
                const divName = currentEditItem.divisionName || currentEditItem.division_name || '-';
                document.getElementById('editHeaderDivision').textContent = divName;
                
                // 開立日期（發函）
                document.getElementById('editHeaderIssueDate').textContent = currentEditItem.issue_date || currentEditItem.issueDate || '(未設定)';
                
                const st = (currentEditItem.status === 'Open' || !currentEditItem.status) ? '持續列管' : currentEditItem.status; 
                document.getElementById('editStatus').value = st;
                
                // 顯示狀態與類型（缺失、觀察、建議）- 使用統一的字段獲取邏輯
                let k = currentEditItem.item_kind_code || currentEditItem.itemKindCode;
                if (!k) {
                    k = extractKindCodeFromNumber(currentEditItem.number);
                }
                
                let kindLabel = getKindLabel(k);
                let statusBadge = getStatusBadge(st);
                
                // 確保即使只有類型或只有狀態也能顯示
                const statusKindHtml = kindLabel || statusBadge ? `<div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">${kindLabel}${statusBadge}</div>` : '';
                document.getElementById('editHeaderStatusKind').innerHTML = statusKindHtml || ''; 
                
                // 計算應該進行第幾次審查（支持無限次）
                // 邏輯：找到最高的機構辦理情形，檢查是否有對應的審查意見
                // 如果沒有，就應該進行該次的審查；如果有，就進行下一次的審查
                let nextRound = 1;
                let highestHandlingRound = 0;
                
                // 先找到最高的機構辦理情形
                for (let i = 1; i <= 200; i++) {
                    const suffix = i === 1 ? '' : i;
                    const hasHandling = currentEditItem['handling' + suffix] && currentEditItem['handling' + suffix].trim();
                    if (hasHandling) {
                        highestHandlingRound = i;
                    }
                }
                
                // 檢查最高的機構辦理情形是否有對應的審查意見
                if (highestHandlingRound > 0) {
                    const suffix = highestHandlingRound === 1 ? '' : highestHandlingRound;
                    const hasReview = currentEditItem['review' + suffix] && currentEditItem['review' + suffix].trim();
                    if (hasReview) {
                        // 如果有審查意見，就進行下一次審查
                        nextRound = highestHandlingRound + 1;
                    } else {
                        // 如果沒有審查意見，就進行該次審查
                        nextRound = highestHandlingRound;
                    }
                } else {
                    // 如果沒有任何機構辦理情形，就進行第1次審查
                    nextRound = 1;
                }
                // 設置審查次數（隱藏的 input 用於保存）
                document.getElementById('editRound').value = nextRound;
                // 更新顯示文字
                const roundDisplay = document.getElementById('editRoundDisplay');
                if (roundDisplay) {
                    roundDisplay.textContent = `第 ${nextRound} 次`;
                }
                
                document.getElementById('editContentDisplay').innerHTML = stripHtml(currentEditItem.content); 
                // 清除 AI 分析結果
                const aiBox = document.getElementById('aiBox');
                if (aiBox) aiBox.style.display = 'none';
                document.getElementById('aiPreviewText').innerText = '';
                document.getElementById('aiResBadge').innerHTML = '';
                // 清除編輯欄位，loadRoundData 會重新載入正確的資料
                document.getElementById('editReview').value = '';
                document.getElementById('editHandling').value = '';
                loadRoundData();
            }
        }
        function initEditForm() { 
            // 審查次數現在是只讀顯示，不再需要初始化下拉選項
            // 保留此函數以保持代碼兼容性
        }
        
        // 動態添加更多審查次數選項（如果需要超過 100 次，用於隱藏的 select）
        function ensureRoundOption(round) {
            const s = document.getElementById('editRoundSelect');
            if (!s) return;
            const maxRound = Math.max(...Array.from(s.options).map(o => parseInt(o.value) || 0));
            if (round > maxRound) {
                for (let i = maxRound + 1; i <= round + 10; i++) {
                    const o = document.createElement('option');
                    o.value = i;
                    o.text = `第 ${i} 次`;
                    s.add(o);
                }
            }
        }

        function openDetail(id, isEdit) {
            currentEditItem = currentData.find(d => String(d.id) === String(id)); if (!currentEditItem) return;
            
            // 編號
            document.getElementById('dNumber').textContent = currentEditItem.number || '';
            
            // 檢查計畫
            document.getElementById('dPlanName').textContent = currentEditItem.plan_name || currentEditItem.planName || '(未設定)';
            
            // 檢查種類
            const insName = currentEditItem.inspectionCategoryName || currentEditItem.inspection_category_name || '-';
            document.getElementById('dInspection').textContent = insName;
            
            // 分組
            const divName = currentEditItem.divisionName || currentEditItem.division_name || '-';
            document.getElementById('dDivision').textContent = divName;
            
            // 開立日期（發函）
            document.getElementById('dIssueDate').textContent = currentEditItem.issue_date || currentEditItem.issueDate || '(未設定)';
            
            // 事項內容（使用 escapeHtml 防止 XSS）
            document.getElementById('dContent').innerHTML = escapeHtml(currentEditItem.content || '');

            // Status and Kind (狀態與類型) - 使用與dCategoryInfo相同的邏輯
            let k = currentEditItem.item_kind_code || currentEditItem.itemKindCode;
            if (!k) {
                k = extractKindCodeFromNumber(currentEditItem.number);
            }
            
            let kindLabel = getKindLabel(k);
            let statusBadge = getStatusBadge(currentEditItem.status);
            
            // 確保即使只有類型或只有狀態也能顯示
            const statusKindHtml = kindLabel || statusBadge ? `<div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">${kindLabel}${statusBadge}</div>` : '';
            document.getElementById('dStatus').innerHTML = statusKindHtml || '(未設定)';

            let h = '';
            let firstRecord = true;
            // 支持無限次，動態查找（從200開始向下找，實際應該不會超過這個數字）
            // 第N次辦理情形區塊應該包含：第N次機構辦理情形 + 第N次審查意見
            for (let i = 200; i >= 1; i--) {
                const suffix = i === 1 ? '' : i;
                // 第N次機構辦理情形
                const ha = currentEditItem['handling' + suffix];
                // 第N次審查意見（第N次機構辦理情形後，會進行第N次審查）
                const re = currentEditItem['review' + suffix];
                const replyDate = currentEditItem['reply_date_r' + i];
                const responseDate = currentEditItem['response_date_r' + i];

                // 只要有機構辦理情形或審查意見，就顯示該次辦理情形
                if (ha || re) {
                    const latestBadge = firstRecord ? '<span class="badge new" style="margin-left:8px;font-size:11px;">最新進度</span>' : '';

                    let dateInfo = '';
                    if (replyDate || responseDate) {
                        dateInfo = `<div style="margin-bottom:12px;">`;
                        if (replyDate) dateInfo += `<span class="timeline-date-tag">🏢 機構回復: ${replyDate}</span> `;
                        if (responseDate) dateInfo += `<span class="timeline-date-tag">🏛️ 機關函復: ${responseDate}</span>`;
                        dateInfo += `</div>`;
                    }

                    // 第N次辦理情形區塊：先顯示第N次機構辦理情形，再顯示第N+1次審查意見
                    h += `<div class="timeline-item">
                        <div class="timeline-dot"></div>
                        <div class="timeline-title">第 ${i} 次辦理情形 ${latestBadge}</div>
                        ${dateInfo}
                        ${ha ? `<div style="background:#ecfdf5;padding:16px;border-radius:8px;font-size:14px;line-height:1.6;color:#047857;border:1px solid #a7f3d0;margin-bottom:12px;white-space:pre-wrap;"><strong>📝 機構辦理情形：</strong><br>${ha}</div>` : ''}
                        ${re ? `<div style="background:#fff;padding:16px;border-radius:8px;font-size:14px;line-height:1.6;color:#334155;border:1px solid #e2e8f0;border-left:3px solid var(--primary);white-space:pre-wrap;"><strong>👀 審查意見：</strong><br>${re}</div>` : ''}
                    </div>`;
                    firstRecord = false;
                }
            }

            const timelineHtml = `<div class="timeline-line"></div>` + (h || '<div style="color:#999;padding-left:20px;">無歷程紀錄</div>');
            document.getElementById('dTimeline').innerHTML = timelineHtml;

            const canEdit = ['admin', 'manager', 'editor'].includes(currentUser.role); const canDelete = ['admin', 'manager'].includes(currentUser.role); document.getElementById('editBtn').classList.toggle('hidden', !canEdit); document.getElementById('deleteBtnDrawer').classList.toggle('hidden', !canDelete); document.getElementById('drawerBackdrop').classList.add('open'); document.getElementById('detailDrawer').classList.add('open'); toggleEditMode(isEdit);
        }
        function logout() { 
            // 清除視圖狀態
            sessionStorage.removeItem('currentView');
            sessionStorage.removeItem('currentDataTab');
            sessionStorage.removeItem('currentUsersTab');
            fetch('/api/auth/logout', { method: 'POST' }).then(() => window.location.reload()); 
        }
        function closeDrawer() { document.getElementById('drawerBackdrop').classList.remove('open'); document.getElementById('detailDrawer').classList.remove('open'); }
        function initListeners() { document.getElementById('filterKeyword').addEventListener('keyup', (e) => { if (e.key === 'Enter') applyFilters() }); document.getElementById('drawerBackdrop').addEventListener('click', closeDrawer); }
        function onToggleSidebar() { const panel = document.getElementById('filtersPanel'), backdrop = document.getElementById('filterBackdrop'); if (panel.classList.contains('open')) { panel.classList.remove('open'); backdrop.classList.remove('visible'); setTimeout(() => backdrop.style.display = 'none', 300); } else { backdrop.style.display = 'block'; requestAnimationFrame(() => { panel.classList.add('open'); backdrop.classList.add('visible'); }); } }

        function updateChartsData(stats) {
            if (!charts.status || !charts.unit || !charts.trend) return;
            if (!stats) return;
            // 統一顏色方案：使用與主色調一致的顏色
            const colorMap = { 
                '持續列管': '#ef4444',  // 紅色 - 危險/警告
                '解除列管': '#10b981',  // 綠色 - 成功
                '自行列管': '#f59e0b'   // 橙色 - 警告
            };
            const sLabels = stats.status.map(x => x.status).filter(s => s && s !== 'Open');
            const sData = stats.status.filter(x => x.status && x.status !== 'Open').map(x => parseInt(x.count));
            const sColors = sLabels.map(label => colorMap[label] || '#cbd5e1');
            charts.status.data = { labels: sLabels, datasets: [{ data: sData, backgroundColor: sColors }] }; 
            charts.status.update();
            
            // 單位圖表：使用主色調的變體
            const uSorted = stats.unit.sort((a, b) => parseInt(b.count) - parseInt(a.count)); 
            charts.unit.data = { 
                labels: uSorted.map(x => x.unit), 
                datasets: [{ 
                    label: '案件', 
                    data: uSorted.map(x => parseInt(x.count)), 
                    backgroundColor: '#667eea',  // 使用與標題漸變一致的顏色（直接使用顏色值，避免 CSS 變數問題）
                    borderRadius: 8 
                }] 
            }; 
            // 確保更新時保留顏色設定
            if (charts.unit.options && charts.unit.options.scales) {
                charts.unit.options.scales.x.ticks.color = '#64748b';
                charts.unit.options.scales.y.ticks.color = '#64748b';
                charts.unit.options.scales.x.grid.color = '#e2e8f0';
                charts.unit.options.scales.y.grid.color = '#e2e8f0';
            }
            charts.unit.update();
            
            // 趨勢圖表：使用主色調的變體
            const tSorted = stats.year.sort((a, b) => a.year.localeCompare(b.year)); 
            // 清除舊的資料集，避免破圖
            if (charts.trend && charts.trend.data) {
                charts.trend.data.labels = [];
                charts.trend.data.datasets = [];
            }
            charts.trend.data = { 
                labels: tSorted.map(x => x.year), 
                datasets: [{ 
                    label: '開立事項數', 
                    data: tSorted.map(x => parseInt(x.count)), 
                    borderColor: '#667eea',  // 使用與標題漸變一致的顏色（直接使用顏色值）
                    backgroundColor: 'rgba(102, 126, 234, 0.1)', 
                    tension: 0.3, 
                    fill: true,
                    pointRadius: 0, // 不顯示點，避免破圖
                    pointHoverRadius: 4,
                    borderWidth: 2
                }] 
            }; 
            // 確保更新時保留顏色設定
            if (charts.trend.options && charts.trend.options.scales) {
                charts.trend.options.scales.x.ticks.color = '#64748b';
                charts.trend.options.scales.y.ticks.color = '#64748b';
                charts.trend.options.scales.x.grid.color = '#e2e8f0';
                charts.trend.options.scales.y.grid.color = '#e2e8f0';
            }
            if (charts.trend.options && charts.trend.options.plugins && charts.trend.options.plugins.title) {
                charts.trend.options.plugins.title.color = '#64748b';
            }
            // 先清除 canvas，避免破圖
            if (charts.trend && charts.trend.canvas) {
                const ctx = charts.trend.canvas.getContext('2d');
                if (ctx) {
                    ctx.clearRect(0, 0, charts.trend.canvas.width, charts.trend.canvas.height);
                }
            }
            // 使用 updateMode: 'none' 避免動畫導致的破圖，並強制重新渲染
            charts.trend.update('none');
            // 強制重新繪製 canvas 以避免破圖
            if (charts.trend && charts.trend.canvas) {
                setTimeout(() => {
                    charts.trend.resize();
                    charts.trend.draw();
                }, 50);
            }
        }

        function initCharts() {
            try {
                const c1 = document.getElementById('statusChart'), c2 = document.getElementById('unitChart'), c3 = document.getElementById('trendChart');
                if (c1) { charts.status = new Chart(c1, { type: 'doughnut', plugins: [ChartDataLabels], data: { labels: [], datasets: [] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#64748b', font: { size: 12 } } }, datalabels: { formatter: (v, ctx) => { const dataArr = ctx.chart.data.datasets[0].data; if (!dataArr || dataArr.length === 0) return ''; const t = dataArr.reduce((a, b) => a + b, 0); return t > 0 ? ((v / t) * 100).toFixed(1) + '%' : '0%'; }, color: '#64748b', font: { weight: '600', size: 12 } } } } }); }
                if (c2) { charts.unit = new Chart(c2, { type: 'bar', data: { labels: [], datasets: [] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: '#64748b', font: { size: 12 } }, grid: { color: '#e2e8f0' } }, y: { ticks: { color: '#64748b', font: { size: 12 } }, grid: { color: '#e2e8f0' } } } } }); }
                if (c3) { charts.trend = new Chart(c3, { type: 'line', data: { labels: [], datasets: [] }, options: { responsive: true, maintainAspectRatio: false, animation: { duration: 0 }, plugins: { legend: { display: false }, title: { display: true, text: '年度開立事項趨勢', color: '#64748b', font: { size: 14, weight: '600' } } }, scales: { x: { ticks: { color: '#64748b', font: { size: 12 } }, grid: { color: '#e2e8f0' } }, y: { beginAtZero: true, ticks: { stepSize: 1, color: '#64748b', font: { size: 12 } }, grid: { color: '#e2e8f0' } } }, elements: { point: { radius: 0 }, line: { borderWidth: 2 } } } }); }
                if (cachedGlobalStats) updateChartsData(cachedGlobalStats);
            } catch (e) { console.error("Chart Init Error:", e); }
        }
        function loadRoundData() {
            if (!currentEditItem) return;
            const round = parseInt(document.getElementById('editRound').value) || 1;
            const suffix = round === 1 ? '' : round;
            
            // 載入該回合的資料
            // 重要：第N次審查時，應該載入第N次的辦理情形和審查意見
            // 辦理情形應該已經在「資料管理」頁面填寫，這裡只是讀取
            const handling = currentEditItem['handling' + suffix] || '';
            const review = currentEditItem['review' + suffix] || '';
            // 機構回復日期從辦理情形中讀取（不需要在審查頁面編輯）
            const replyDate = currentEditItem['reply_date_r' + round] || '';
            
            // 儲存到隱藏的輸入框（用於儲存時提交）
            // 注意：這裡的 handling 是第N次的辦理情形，review 是第N次的審查意見
            // 在審查頁面，我們只編輯 review，handling 是只讀的（應該已在資料管理頁面填寫）
            // 重要：確保不會把 review 的值錯誤地存到 handling
            document.getElementById('editHandling').value = handling;
            document.getElementById('editReview').value = review;
            // replyDate 從資料中讀取，不需要輸入框
            // responseDate 已移除，不再在審查頁面設定
            
            // 顯示第N次機構辦理情形（只讀，作為參考）
            // 撰寫第N次審查時，右側顯示第N次機構辦理情形
            // 因為第N次機構辦理情形後，會進行第N次審查
            const displayHandlingRound = round;
            const displayHandlingSuffix = displayHandlingRound === 1 ? '' : displayHandlingRound;
            const displayHandling = currentEditItem['handling' + displayHandlingSuffix] || '';
            
            // 更新辦理情形顯示（只讀）
            const currentHandlingDisplay = document.getElementById('currentHandlingDisplay');
            const currentHandlingRoundNum = document.getElementById('currentHandlingRoundNum');
            
            if (currentHandlingDisplay && currentHandlingRoundNum) {
                currentHandlingRoundNum.textContent = displayHandlingRound;
                if (displayHandling && displayHandling.trim()) {
                    currentHandlingDisplay.textContent = displayHandling;
                    currentHandlingDisplay.style.color = '#047857';
                } else {
                    currentHandlingDisplay.textContent = '（尚未有機構辦理情形）';
                    currentHandlingDisplay.style.color = '#94a3b8';
                }
            }
            
            // 顯示上一回合的審查意見（如果有，且不是第1次）
            const prevRound = round - 1;
            if (prevRound >= 1) {
                const prevSuffix = prevRound === 1 ? '' : prevRound;
                const prevReview = currentEditItem['review' + prevSuffix] || '';
                const prevBox = document.getElementById('prevReviewBox');
                const prevText = document.getElementById('prevReviewText');
                const prevRoundNum = document.getElementById('prevRoundNum');
                
                if (prevReview && prevBox && prevText && prevRoundNum) {
                    prevBox.style.display = 'block';
                    prevRoundNum.textContent = prevRound;
                    prevText.textContent = prevReview;
                } else if (prevBox) {
                    prevBox.style.display = 'none';
                }
            } else {
                // 第1次審查，隱藏前次審查意見
                const prevBox = document.getElementById('prevReviewBox');
                if (prevBox) prevBox.style.display = 'none';
            }
            
            // 清除 AI 分析結果（因為回合改變了）
            const aiBox = document.getElementById('aiBox');
            if (aiBox) aiBox.style.display = 'none';
            document.getElementById('aiPreviewText').innerText = '';
            document.getElementById('aiResBadge').innerHTML = '';
            
            // [Added] 初始化查看輪次選擇下拉選單
            initViewRoundSelect();
        }
        
        // [Added] 初始化查看輪次選擇下拉選單
        function initViewRoundSelect() {
            if (!currentEditItem) return;
            
            const select = document.getElementById('viewRoundSelect');
            if (!select) return;
            
            // 找出所有有內容的輪次（有審查意見或辦理情形即可）
            const rounds = [];
            for (let i = 200; i >= 1; i--) {
                const suffix = i === 1 ? '' : i;
                const hasHandling = currentEditItem['handling' + suffix] && currentEditItem['handling' + suffix].trim();
                const hasReview = currentEditItem['review' + suffix] && currentEditItem['review' + suffix].trim();
                // 只要有審查意見或辦理情形就包含
                if (hasHandling || hasReview) {
                    rounds.push(i);
                }
            }
            
            // 生成選項（從最新到最舊）
            select.innerHTML = '<option value="latest">最新進度</option>';
            rounds.forEach(r => {
                select.innerHTML += `<option value="${r}">第 ${r} 次</option>`;
            });
            
            // 預設選擇最新進度
            select.value = 'latest';
            onViewRoundChange();
        }
        
        // [Added] 當查看輪次選擇改變時
        function onViewRoundChange() {
            if (!currentEditItem) return;
            
            const select = document.getElementById('viewRoundSelect');
            if (!select) return;
            
            const selectedValue = select.value;
            
            // 隱藏所有查看區塊
            const viewReviewBox = document.getElementById('viewReviewBox');
            const viewHandlingBox = document.getElementById('viewHandlingBox');
            if (viewReviewBox) viewReviewBox.style.display = 'none';
            if (viewHandlingBox) viewHandlingBox.style.display = 'none';
            
            if (selectedValue === 'latest') {
                // 顯示最新進度 - 優先顯示「同時有審查意見和辦理情形」的最高輪次
                // 如果最高輪次只有其中一個，則顯示次高的完整輪次
                let bestRound = 0;
                let maxRound = 0;
                
                // 先找出最高的完整輪次（同時有審查意見和辦理情形）
                for (let k = 200; k >= 1; k--) {
                    const suffix = k === 1 ? '' : k;
                    const hasHandling = currentEditItem['handling' + suffix] && currentEditItem['handling' + suffix].trim();
                    const hasReview = currentEditItem['review' + suffix] && currentEditItem['review' + suffix].trim();
                    
                    // 記錄最高輪次（有任一內容即可）
                    if ((hasHandling || hasReview) && maxRound === 0) {
                        maxRound = k;
                    }
                    
                    // 優先選擇同時有兩個內容的輪次
                    if (hasHandling && hasReview) {
                        bestRound = k;
                        break;
                    }
                }
                
                // 如果沒有完整的輪次，使用最高輪次
                const displayRound = bestRound > 0 ? bestRound : maxRound;
                
                if (displayRound > 0) {
                    const suffix = displayRound === 1 ? '' : displayRound;
                    const handling = currentEditItem['handling' + suffix] || '';
                    const review = currentEditItem['review' + suffix] || '';
                    
                    // 顯示審查意見
                    if (review && review.trim()) {
                        const viewReviewRoundNum = document.getElementById('viewReviewRoundNum');
                        const viewReviewText = document.getElementById('viewReviewText');
                        const viewReviewDate = document.getElementById('viewReviewDate');
                        if (viewReviewRoundNum) viewReviewRoundNum.textContent = displayRound;
                        if (viewReviewText) viewReviewText.textContent = review;
                        // 顯示審查函復日期
                        const responseDate = currentEditItem['response_date_r' + displayRound] || '';
                        if (viewReviewDate) {
                            viewReviewDate.textContent = responseDate ? `函復日期：${responseDate}` : '';
                        }
                        if (viewReviewBox) viewReviewBox.style.display = 'block';
                    }
                    
                    // 顯示辦理情形
                    if (handling && handling.trim()) {
                        const viewHandlingRoundNum = document.getElementById('viewHandlingRoundNum');
                        const viewHandlingText = document.getElementById('viewHandlingText');
                        const viewHandlingDate = document.getElementById('viewHandlingDate');
                        if (viewHandlingRoundNum) viewHandlingRoundNum.textContent = displayRound;
                        if (viewHandlingText) viewHandlingText.textContent = handling;
                        // 顯示辦理情形回復日期
                        const replyDate = currentEditItem['reply_date_r' + displayRound] || '';
                        if (viewHandlingDate) {
                            viewHandlingDate.textContent = replyDate ? `回復日期：${replyDate}` : '';
                        }
                        if (viewHandlingBox) viewHandlingBox.style.display = 'block';
                    }
                }
            } else {
                // 顯示指定輪次
                const round = parseInt(selectedValue, 10);
                const suffix = round === 1 ? '' : round;
                const handling = currentEditItem['handling' + suffix] || '';
                const review = currentEditItem['review' + suffix] || '';
                
                if (review && review.trim()) {
                    const viewReviewRoundNum = document.getElementById('viewReviewRoundNum');
                    const viewReviewText = document.getElementById('viewReviewText');
                    const viewReviewDate = document.getElementById('viewReviewDate');
                    if (viewReviewRoundNum) viewReviewRoundNum.textContent = round;
                    if (viewReviewText) viewReviewText.textContent = review;
                    // 顯示審查函復日期
                    const responseDate = currentEditItem['response_date_r' + round] || '';
                    if (viewReviewDate) {
                        viewReviewDate.textContent = responseDate ? `函復日期：${responseDate}` : '';
                    }
                    if (viewReviewBox) viewReviewBox.style.display = 'block';
                }
                
                if (handling && handling.trim()) {
                    const viewHandlingRoundNum = document.getElementById('viewHandlingRoundNum');
                    const viewHandlingText = document.getElementById('viewHandlingText');
                    const viewHandlingDate = document.getElementById('viewHandlingDate');
                    if (viewHandlingRoundNum) viewHandlingRoundNum.textContent = round;
                    if (viewHandlingText) viewHandlingText.textContent = handling;
                    // 顯示辦理情形回復日期
                    const replyDate = currentEditItem['reply_date_r' + round] || '';
                    if (viewHandlingDate) {
                        viewHandlingDate.textContent = replyDate ? `回復日期：${replyDate}` : '';
                    }
                    if (viewHandlingBox) viewHandlingBox.style.display = 'block';
                }
            }
        }

        // 從 Drawer 刪除事項
        async function deleteIssueFromDrawer() {
            if (!currentEditItem) {
                showToast('找不到要刪除的事項', 'error');
                return;
            }
            
            const issueId = currentEditItem.id;
            const issueNumber = currentEditItem.number || `ID:${issueId}`;
            
            const confirmed = await showConfirmModal(`確定要刪除事項「${issueNumber}」嗎？\n\n此操作無法復原。`, '確定刪除', '取消');
            if (!confirmed) {
                return;
            }
            
            try {
                const res = await apiFetch(`/api/issues/${issueId}`, {
                    method: 'DELETE'
                });
                
                const data = await res.json();
                if (res.ok) {
                    showToast('刪除成功');
                    closeDrawer();
                    // 重新載入事項列表
                    loadIssuesPage(issuesPage);
                } else {
                    showToast(data.error || '刪除失敗', 'error');
                }
            } catch (e) {
                showToast('刪除失敗: ' + e.message, 'error');
            }
        }
        
        async function saveEdit() {
            if (!currentEditItem) {
                showToast('找不到目前編輯的事項', 'error');
                return;
            }
            
            const id = document.getElementById('editId').value;
            const status = document.getElementById('editStatus').value;
            const round = parseInt(document.getElementById('editRound').value) || 1;
            // 從隱藏欄位讀取辦理情形（僅用於保存，不允許在審查頁面編輯）
            const handling = document.getElementById('editHandling').value.trim() || '';
            const review = document.getElementById('editReview').value.trim();
            // 機構回復日期從資料中讀取（已在辦理情形階段填寫）
            const replyDate = currentEditItem ? (currentEditItem['reply_date_r' + round] || '') : '';
            // responseDate 已移除，不再在審查頁面設定（改為在開立事項建檔頁面批次設定）
            
            if (!id) {
                showToast('找不到事項 ID', 'error');
                return;
            }
            
            // 第N次審查時，必須已有第N次的辦理情形（應該在資料管理頁面先輸入）
            if (!handling) {
                showToast(`第 ${round} 次審查時，必須先有第 ${round} 次機構辦理情形。請至「資料管理」頁面的「年度編輯」功能中新增辦理情形後，再進行審查。`, 'error');
                return;
            }
            
            // 重要：確保 handling 和 review 的對應關係正確
            // handling 應該是第N次的辦理情形（已在資料管理頁面填寫）
            // review 應該是第N次的審查意見（正在審查頁面填寫）
            // 不應該把審查意見存到辦理情形欄位
            
            try {
                const res = await apiFetch(`/api/issues/${id}`, {
                    method: 'PUT',
                    body: JSON.stringify({
                        status,
                        round,
                        handling,
                        review,
                        replyDate: replyDate || null,
                        responseDate: null // 函復日期改為在開立事項建檔頁面批次設定
                    })
                });
                
                if (res.ok) {
                    const json = await res.json();
                    if (json.success) {
                        showToast('儲存成功！');
                        // 關閉 drawer 並返回查詢看板
                        closeDrawer();
                        // 重新載入資料
                        await loadIssuesPage(issuesPage);
                    } else {
                        showToast('儲存失敗', 'error');
                    }
                } else {
                    const json = await res.json();
                    showToast(json.error || '儲存失敗', 'error');
                }
            } catch (e) {
                console.error('Save error:', e);
                showToast('儲存時發生錯誤: ' + e.message, 'error');
            }
        }

        async function runAiInEdit(btn) { 
            btn.disabled = true; 
            btn.innerText = 'AI 分析中...'; 
            // 從隱藏欄位讀取辦理情形
            const handlingTxt = document.getElementById('editHandling').value || ''; 
            const r = [{ handling: handlingTxt, review: '(待審查)' }]; 
            try { 
                if (!currentEditItem || !currentEditItem.content) throw new Error('找不到事項內容'); 
                const cleanContent = stripHtml(currentEditItem.content); 
                const res = await fetch('/api/gemini', { 
                    method: 'POST', 
                    headers: { 'Content-Type': 'application/json' }, 
                    body: JSON.stringify({ content: cleanContent, rounds: r }) 
                }); 
                const j = await res.json(); 
                if (res.ok && j.result) { 
                    document.getElementById('aiBox').style.display = 'block'; 
                    document.getElementById('aiPreviewText').innerText = j.result; 
                    document.getElementById('aiResBadge').innerHTML = j.fulfill && j.fulfill.includes('是') ? `<span class="ai-tag yes">✅ 符合</span>` : `<span class="ai-tag no">⚠️ 需注意</span>`; 
                } else { 
                    showToast('AI 分析失敗', 'error'); 
                } 
            } catch (e) { 
                showToast('AI Error: ' + e.message, 'error'); 
            } finally { 
                btn.disabled = false; 
                btn.innerText = '🤖 AI 智能分析'; 
            } 
        }
        function applyAiSuggestion() { 
            const txt = document.getElementById('aiPreviewText').innerText; 
            if (txt) { 
                document.getElementById('editReview').value = txt; 
                // 移除成功提示，只保留錯誤提示
            } else {
                showToast('沒有可帶入的 AI 建議', 'error');
            }
        }
        
        // --- 事項修正功能 ---
        let yearEditIssue = null; // 儲存當前編輯的事項資料
        let yearEditIssueList = []; // 儲存當前計畫下的事項列表
        
        // 從編號字串中提取數字（用於排序）
        function extractNumberFromString(str) {
            if (!str) return null;
            // 嘗試提取編號最後的數字部分（例如：113ABC-DEF-001 中的 001）
            const matches = str.match(/(\d+)(?!.*\d)/);
            if (matches && matches[1]) {
                return parseInt(matches[1], 10);
            }
            // 如果沒有找到，嘗試提取所有數字
            const allNumbers = str.match(/\d+/g);
            if (allNumbers && allNumbers.length > 0) {
                return parseInt(allNumbers[allNumbers.length - 1], 10);
            }
            return null;
        }
        
        // 載入有開立事項的檢查計畫選項（類似查詢看板的檢查計畫下拉選單）
        async function loadYearEditPlanOptions() {
            const select = document.getElementById('yearEditPlanName');
            if (!select) return;
            
            try {
                select.innerHTML = '<option value="">載入中...</option>';
                
                const res = await fetch('/api/options/plans?withIssues=true&t=' + Date.now(), {
                    cache: 'no-store',
                    headers: {
                        'Cache-Control': 'no-cache'
                    }
                });
                
                if (!res.ok) {
                    throw new Error('載入檢查計畫失敗');
                }
                
                const json = await res.json();
                
                if (!json.data || json.data.length === 0) {
                    select.innerHTML = '<option value="">尚無有開立事項的檢查計畫</option>';
                    return;
                }
                
                // 處理新的資料格式，按年度分組
                const yearGroups = new Map();
                
                json.data.forEach(p => {
                    let planName, planYear, planValue, planDisplay;
                    
                    if (typeof p === 'object' && p !== null) {
                        planName = p.name || '';
                        planYear = p.year || '';
                        planValue = p.value || `${planName}|||${planYear}`;
                        planDisplay = planName;
                    } else {
                        planName = p;
                        planYear = '';
                        planValue = p;
                        planDisplay = p;
                    }
                    
                    if (planName) {
                        const groupKey = planYear || '未分類';
                        if (!yearGroups.has(groupKey)) {
                            yearGroups.set(groupKey, []);
                        }
                        yearGroups.get(groupKey).push({ 
                            value: planValue, 
                            display: planDisplay, 
                            name: planName, 
                            year: planYear 
                        });
                    }
                });
                
                // 建立選項 HTML
                let allOptions = '<option value="">請選擇檢查計畫</option>';
                
                // 將年度分組按年度降序排序（最新的在前）
                const sortedYears = Array.from(yearGroups.keys()).sort((a, b) => {
                    if (a === '未分類') return 1;
                    if (b === '未分類') return -1;
                    const yearA = parseInt(a) || 0;
                    const yearB = parseInt(b) || 0;
                    return yearB - yearA;
                });
                
                sortedYears.forEach(year => {
                    const plans = yearGroups.get(year);
                    // 按計畫名稱排序（同一年度內的計畫按名稱排序）
                    plans.sort((a, b) => {
                        return (a.name || '').localeCompare(b.name || '', 'zh-TW');
                    });
                    
                    // 使用 optgroup 按年度分組
                    const yearLabel = year === '未分類' ? '未分類' : `${year} 年度`;
                    allOptions += `<optgroup label="${yearLabel}">`;
                    plans.forEach(plan => {
                        allOptions += `<option value="${plan.value}">${plan.display}</option>`;
                    });
                    allOptions += `</optgroup>`;
                });
                
                select.innerHTML = allOptions;
            } catch (e) {
                console.error('載入檢查計畫選項失敗:', e);
                select.innerHTML = '<option value="">載入失敗，請重新整理頁面</option>';
                showToast('載入檢查計畫失敗: ' + e.message, 'error');
            }
        }
        
        // 檢查計畫改變時，載入該計畫下的事項列表
        async function onYearEditPlanChange() {
            const planSelect = document.getElementById('yearEditPlanName');
            if (!planSelect) return;
            
            const planValue = planSelect.value;
            
            // 隱藏編輯內容和列表
            hideYearEditIssueContent();
            hideYearEditIssueList();
            
            if (!planValue) {
                document.getElementById('yearEditEmpty').style.display = 'block';
                document.getElementById('yearEditNotFound').style.display = 'none';
                return;
            }
            
            const [planName, planYear] = planValue.split('|||');
            
            try {
                // 載入該計畫下的所有事項（不顯示提示，因為已經確認有開立事項）
                yearEditIssueList = await loadIssuesByPlan(planValue, { showError: true, returnEmpty: true }) || [];
                
                // 對事項列表進行排序：先按類型（缺失N、觀察O、建議R），再按編號（數字小的在前）
                if (yearEditIssueList.length > 0) {
                    yearEditIssueList.sort((a, b) => {
                        // 1. 先按類型排序：缺失(N) -> 觀察(O) -> 建議(R)
                        const kindOrder = { 'N': 1, 'O': 2, 'R': 3 };
                        // 資料庫欄位可能是 item_kind_code 或 itemKindCode，兩種都嘗試
                        const kindCodeA = a.item_kind_code || a.itemKindCode || '';
                        const kindCodeB = b.item_kind_code || b.itemKindCode || '';
                        const kindA = kindOrder[kindCodeA] || 99;
                        const kindB = kindOrder[kindCodeB] || 99;
                        
                        if (kindA !== kindB) {
                            return kindA - kindB;
                        }
                        
                        // 2. 如果類型相同，按編號排序（提取編號中的數字部分）
                        const numA = extractNumberFromString(a.number || '');
                        const numB = extractNumberFromString(b.number || '');
                        
                        if (numA !== null && numB !== null) {
                            return numA - numB;
                        }
                        
                        // 如果無法提取數字，按字串排序
                        return (a.number || '').localeCompare(b.number || '', 'zh-TW');
                    });
                }
                
                if (yearEditIssueList.length === 0) {
                    // 沒有事項
                    document.getElementById('yearEditEmpty').style.display = 'none';
                    document.getElementById('yearEditNotFound').style.display = 'block';
                    document.getElementById('yearEditIssueList').style.display = 'none';
                } else {
                    // 顯示事項列表（不顯示提示，因為已經確認有開立事項）
                    document.getElementById('yearEditEmpty').style.display = 'none';
                    document.getElementById('yearEditNotFound').style.display = 'none';
                    renderYearEditIssueList();
                }
            } catch (e) {
                showToast('載入事項列表失敗: ' + e.message, 'error');
                hideYearEditIssueList();
            }
        }
        
        // 渲染事項列表
        function renderYearEditIssueList() {
            const container = document.getElementById('yearEditIssueListContainer');
            const countEl = document.getElementById('yearEditIssueListCount');
            if (!container) return;
            
            if (countEl) {
                countEl.textContent = yearEditIssueList.length;
            }
            
            if (yearEditIssueList.length === 0) {
                container.innerHTML = '<div style="padding:40px; text-align:center; color:#94a3b8;">尚無事項</div>';
                document.getElementById('yearEditIssueList').style.display = 'none';
                return;
            }
            
            let html = '';
            yearEditIssueList.forEach((issue, index) => {
                const contentPreview = stripHtml(issue.content || '').substring(0, 150);
                
                // 顯示類型（缺失、觀察、建議）
                let k = issue.itemKindCode;
                if (!k) {
                    k = extractKindCodeFromNumber(issue.number);
                }
                
                let kindLabel = getKindLabel(k);
                
                // 顯示狀態徽章
                let badge = '';
                const st = String(issue.status || 'Open');
                if (st !== 'Open' && st) {
                    const stClass = st === '持續列管' ? 'active' : (st === '解除列管' ? 'resolved' : 'self');
                    badge = `<span class="badge ${stClass}">${st}</span>`;
                }
                
                html += `
                    <div class="year-edit-issue-item" 
                         onclick="loadYearEditIssueFromList(${index})"
                         style="padding:16px; border-bottom:1px solid #e2e8f0; cursor:pointer; transition:background 0.2s;"
                         onmouseover="this.style.background='#f8fafc'"
                         onmouseout="this.style.background='#fff'">
                        <div style="display:flex; justify-content:space-between; align-items:start; gap:16px;">
                            <div style="flex:1;">
                                <div style="display:flex; align-items:center; gap:8px; margin-bottom:8px; flex-wrap:wrap;">
                                    <div style="font-weight:700; color:#1e40af; font-size:15px;">
                                        ${issue.number || '未指定編號'}
                                    </div>
                                    <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                                        ${kindLabel}${badge}
                                    </div>
                                </div>
                                <div style="font-size:13px; color:#64748b; line-height:1.6; margin-bottom:8px;">
                                    ${contentPreview}${contentPreview.length >= 150 ? '...' : ''}
                                </div>
                                <div style="display:flex; gap:12px; font-size:12px; color:#94a3b8;">
                                    <span>年度：${issue.year || ''}</span>
                                    <span>機構：${issue.unit || ''}</span>
                                </div>
                            </div>
                            <div style="color:#cbd5e1; font-size:20px; align-self:center;">→</div>
                        </div>
                    </div>
                `;
            });
            
            container.innerHTML = html;
            document.getElementById('yearEditIssueList').style.display = 'block';
        }
        
        // 從列表載入指定事項進入編輯模式
        function loadYearEditIssueFromList(index) {
            if (index < 0 || index >= yearEditIssueList.length) return;
            
            yearEditIssue = yearEditIssueList[index];
            
            // 標準化字段名（確保同時有兩種格式，提高兼容性）
            if (yearEditIssue.division_name && !yearEditIssue.divisionName) {
                yearEditIssue.divisionName = yearEditIssue.division_name;
            }
            if (yearEditIssue.inspection_category_name && !yearEditIssue.inspectionCategoryName) {
                yearEditIssue.inspectionCategoryName = yearEditIssue.inspection_category_name;
            }
            if (yearEditIssue.item_kind_code && !yearEditIssue.itemKindCode) {
                yearEditIssue.itemKindCode = yearEditIssue.item_kind_code;
            }
            if (yearEditIssue.plan_name && !yearEditIssue.planName) {
                yearEditIssue.planName = yearEditIssue.plan_name;
            }
            
            // 隱藏列表，顯示編輯內容
            document.getElementById('yearEditIssueList').style.display = 'none';
            document.getElementById('yearEditEmpty').style.display = 'none';
            document.getElementById('yearEditNotFound').style.display = 'none';
            document.getElementById('yearEditIssueContent').style.display = 'block';
            document.getElementById('yearEditSaveBtn').disabled = false;
            
            renderYearEditIssue();
            
            // 滾動到編輯區域
            document.getElementById('yearEditIssueContent').scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        
        // 隱藏事項列表
        function hideYearEditIssueList() {
            const listEl = document.getElementById('yearEditIssueList');
            if (listEl) listEl.style.display = 'none';
        }
        
        // 隱藏編輯內容
        function hideYearEditIssueContent() {
            const contentEl = document.getElementById('yearEditIssueContent');
            if (contentEl) contentEl.style.display = 'none';
        }
        
        // 返回事項列表
        function backToYearEditIssueList() {
            hideYearEditIssueContent();
            if (yearEditIssueList.length > 0) {
                renderYearEditIssueList();
                document.getElementById('yearEditIssueList').style.display = 'block';
            } else {
                document.getElementById('yearEditEmpty').style.display = 'block';
            }
        }
        
        // 批次設定函復日期（用於開立事項建檔頁面）
        async function batchSetResponseDateForPlan() {
            const roundSelect = document.getElementById('createBatchResponseRound');
            const roundManualInput = document.getElementById('createBatchResponseRoundManual');
            const dateInput = document.getElementById('createBatchResponseDate');
            const planSelect = document.getElementById('createPlanName');
            
            if (!roundSelect || !roundManualInput || !dateInput || !planSelect) return;
            
            // 優先使用下拉選單的值，如果沒有則使用手動輸入
            let round = parseInt(roundSelect.value);
            if (!round || round < 1) {
                round = parseInt(roundManualInput.value);
            }
            
            // 立即從輸入框獲取用戶輸入的日期值並存儲，避免後續被修改
            const userInputResponseDate = dateInput.value.trim();
            const planValue = planSelect.value.trim();
            
            if (!planValue) {
                showToast('請先選擇檢查計畫', 'error');
                return;
            }
            
            if (!round || round < 1) {
                showToast('請選擇或輸入審查輪次', 'error');
                return;
            }
            
            if (round > 200) {
                showToast('審查輪次不能超過200次', 'error');
                return;
            }
            
            if (!userInputResponseDate) {
                showToast('請輸入函復日期', 'error');
                return;
            }
            
            // 驗證日期格式
            if (!validateDateFormat(userInputResponseDate, '日期')) {
                return;
            }
            
            const { name: planName } = parsePlanValue(planValue);
            
            try {
                // 載入該計畫下的所有事項
                const issueList = await loadIssuesByPlan(planValue);
                if (!issueList || issueList.length === 0) {
                    showToast('該檢查計畫下尚無開立事項', 'error');
                    return;
                }
                
                // userInputResponseDate 已經在函數開始時從輸入框獲取並保存
                
                const confirmed = await showConfirmModal(
                    `確定要批次設定第 ${round} 次審查的函復日期為 ${userInputResponseDate} 嗎？\n\n將更新 ${issueList.length} 筆事項。`,
                    '確認設定',
                    '取消'
                );
                
                if (!confirmed) {
                    return;
                }
                
                // 移除批次設定中的提示訊息，只保留錯誤訊息
                
                let successCount = 0;
                let errorCount = 0;
                const errors = [];
                
                // 批次更新所有事項
                for (let i = 0; i < issueList.length; i++) {
                    const issue = issueList[i];
                    const issueId = issue.id;
                    
                    if (!issueId) {
                        errorCount++;
                        errors.push(`${issue.number || '未知編號'}: 缺少事項ID`);
                        continue;
                    }
                    
                    try {
                        // 讀取該輪次的現有資料
                        const suffix = round === 1 ? '' : round;
                        const handling = issue['handling' + suffix] || '';
                        const review = issue['review' + suffix] || '';
                        
                        // 檢查是否有審查內容，沒有審查內容則跳過
                        if (!review || !review.trim()) {
                            errorCount++;
                            errors.push(`${issue.number || '未知編號'}: 第 ${round} 次尚無審查意見，無法設定函復日期`);
                            continue;
                        }
                        
                        // 明確使用用戶輸入的日期，不使用任何從資料庫讀取的日期值
                        // userInputResponseDate 是在函數開始時從輸入框獲取的用戶輸入值，不會被修改
                        // 確保不使用 issue 物件中的任何日期欄位（包括 reply_date_r 和 response_date_r）
                        
                        // 更新該輪次的函復日期
                        // 注意：只更新 responseDate（審查函復日期），不更新 replyDate（回復日期）
                        const updateRes = await apiFetch(`/api/issues/${issueId}`, {
                            method: 'PUT',
                            body: JSON.stringify({
                                status: issue.status || '持續列管',
                                round: round,
                                handling: handling,
                                review: review,
                                // 重要：不發送 replyDate，讓後端保持原有值不變
                                // 只發送 responseDate，使用用戶在輸入框中輸入的日期
                                responseDate: userInputResponseDate  // 明確使用用戶輸入的審查函復日期，不從資料庫讀取
                            })
                        });
                        
                        if (updateRes.ok) {
                            const result = await updateRes.json();
                            if (result.success) {
                                successCount++;
                            } else {
                                errorCount++;
                                errors.push(`${issue.number || '未知編號'}: 更新失敗`);
                            }
                        } else {
                            errorCount++;
                            const errorData = await updateRes.json().catch(() => ({}));
                            errors.push(`${issue.number || '未知編號'}: ${errorData.error || '更新失敗'}`);
                        }
                    } catch (e) {
                        errorCount++;
                        errors.push(`${issue.number || '未知編號'}: ${e.message}`);
                    }
                }
                
                // 顯示資料庫操作結果（成功或警告）
                if (successCount > 0 && errorCount === 0) {
                    // 完全成功時顯示成功訊息
                    showToast(`批次設定完成！成功 ${successCount} 筆`, 'success');
                    // 清空輸入欄位並重置為預設模式
                    roundSelect.value = '';
                    roundManualInput.value = '';
                    dateInput.value = '';
                    // 取消勾選並隱藏設定區塊
                    const toggleCheckbox = document.getElementById('createBatchResponseDateToggle');
                    if (toggleCheckbox) {
                        toggleCheckbox.checked = false;
                        toggleBatchResponseDateSetting();
                    }
                } else if (successCount > 0 && errorCount > 0) {
                    // 部分成功
                    showToast(`批次設定完成，但有 ${errorCount} 筆失敗，成功 ${successCount} 筆`, 'warning');
                    if (errors.length > 0) {
                        console.error('批次設定函復日期錯誤:', errors);
                    }
                } else if (errorCount > 0) {
                    // 全部失敗
                    showToast(`批次設定失敗，所有 ${errorCount} 筆事項都無法更新`, 'error');
                    if (errors.length > 0) {
                        console.error('批次設定函復日期錯誤:', errors);
                        // 顯示第一個錯誤的詳細資訊
                        showToast(`錯誤詳情：${errors[0]}`, 'error');
                    }
                } else {
                    // 沒有處理任何事項（理論上不應該發生）
                    showToast('沒有事項需要更新', 'warning');
                }
            } catch (e) {
                showToast('批次設定失敗: ' + e.message, 'error');
            }
        }
        
        // 批次設定函復日期（用於事項修正頁面，保留向後兼容）
        async function batchSetResponseDate() {
            const roundSelect = document.getElementById('yearEditBatchResponseRound');
            const dateInput = document.getElementById('yearEditBatchResponseDate');
            
            if (!roundSelect || !dateInput) return;
            
            const round = parseInt(roundSelect.value);
            // 確保使用用戶輸入的日期值，存儲在局部變量中避免被修改
            const userInputResponseDate = dateInput.value.trim();
            
            if (!round || round < 1) {
                showToast('請選擇輪次', 'error');
                return;
            }
            
            if (!userInputResponseDate) {
                showToast('請輸入函復日期', 'error');
                return;
            }
            
            // 驗證日期格式
            if (!validateDateFormat(userInputResponseDate, '日期')) {
                return;
            }
            
            if (yearEditIssueList.length === 0) {
                showToast('沒有可設定的事項', 'error');
                return;
            }
            
            const confirmed = await showConfirmModal(`確定要批次設定第 ${round} 次審查的函復日期為 ${responseDate} 嗎？\n\n將更新 ${yearEditIssueList.length} 筆事項。`, '確定設定', '取消');
            if (!confirmed) {
                return;
            }
            
            try {
                showToast('批次設定中，請稍候...', 'info');
                
                let successCount = 0;
                let errorCount = 0;
                const errors = [];
                
                // 批次更新所有事項
                for (let i = 0; i < yearEditIssueList.length; i++) {
                    const issue = yearEditIssueList[i];
                    const issueId = issue.id;
                    
                    if (!issueId) {
                        errorCount++;
                        errors.push(`${issue.number || '未知編號'}: 缺少事項ID`);
                        continue;
                    }
                    
                    try {
                        // 讀取該輪次的現有資料
                        const suffix = round === 1 ? '' : round;
                        const handling = issue['handling' + suffix] || '';
                        const review = issue['review' + suffix] || '';
                        
                        // 檢查是否有審查內容，沒有審查內容則跳過
                        if (!review || !review.trim()) {
                            errorCount++;
                            errors.push(`${issue.number || '未知編號'}: 第 ${round} 次尚無審查意見，無法設定函復日期`);
                            continue;
                        }
                        
                        // 明確使用用戶輸入的日期，不使用任何從資料庫讀取的日期值
                        // userInputResponseDate 是在函數開始時從輸入框獲取的用戶輸入值
                        
                        // 更新該輪次的函復日期
                        // 注意：只更新 responseDate（審查函復日期），不更新 replyDate（回復日期）
                        const res = await apiFetch(`/api/issues/${issueId}`, {
                            method: 'PUT',
                            body: JSON.stringify({
                                status: issue.status || '持續列管',
                                round: round,
                                handling: handling,
                                review: review,
                                // 不發送 replyDate，讓後端保持原有值不變
                                responseDate: userInputResponseDate  // 明確使用用戶輸入的審查函復日期
                            })
                        });
                        
                        if (res.ok) {
                            const result = await res.json();
                            if (result.success) {
                                successCount++;
                                // 更新本地資料
                                issue['response_date_r' + round] = responseDate;
                            } else {
                                errorCount++;
                                errors.push(`${issue.number || '未知編號'}: 更新失敗`);
                            }
                        } else {
                            errorCount++;
                            const errorData = await res.json().catch(() => ({}));
                            errors.push(`${issue.number || '未知編號'}: ${errorData.error || '更新失敗'}`);
                        }
                    } catch (e) {
                        errorCount++;
                        errors.push(`${issue.number || '未知編號'}: ${e.message}`);
                    }
                }
                
                if (successCount > 0) {
                    showToast(`批次設定完成！成功 ${successCount} 筆${errorCount > 0 ? `，失敗 ${errorCount} 筆` : ''}`, errorCount > 0 ? 'warning' : 'success');
                    
                    // 如果有錯誤，顯示詳細資訊
                    if (errorCount > 0 && errors.length > 0) {
                        console.error('批次設定函復日期錯誤:', errors);
                    }
                    
                    // 重新載入事項列表以反映更新
                    const planSelect = document.getElementById('yearEditPlanName');
                    if (planSelect && planSelect.value) {
                        await onYearEditPlanChange();
                    }
                } else {
                    showToast('批次設定失敗，所有事項都無法更新', 'error');
                    if (errors.length > 0) {
                        console.error('批次設定函復日期錯誤:', errors);
                    }
                }
            } catch (e) {
                showToast('批次設定失敗: ' + e.message, 'error');
            }
        }
        
        // 渲染事項詳細內容（包含所有輪次）
        function renderYearEditIssue() {
            const container = document.getElementById('yearEditIssueContainer');
            if (!container || !yearEditIssue) return;
            
            const item = yearEditIssue;
            
            // 收集所有輪次的辦理情形和審查意見
            const rounds = [];
            for (let i = 1; i <= 200; i++) {
                const suffix = i === 1 ? '' : i;
                const handling = item[`handling${suffix}`] || '';
                const review = item[`review${suffix}`] || '';
                const replyDate = item[`reply_date_r${i}`] || '';
                const responseDate = item[`response_date_r${i}`] || '';
                
                if (handling || review || replyDate || responseDate) {
                    rounds.push({
                        round: i,
                        handling: stripHtml(handling),
                        review: stripHtml(review),
                        replyDate: replyDate,
                        responseDate: responseDate
                    });
                }
            }
            
            // 檢查是否有實際的審查和回復紀錄（如果只有開立事項，不顯示此區塊）
            const hasReviewRecords = rounds.length > 0;
            
            // 構建檢查計畫選項（需要從現有的計畫選項中選擇）
            let planOptionsHtml = '<option value="">(未指定)</option>';
            const planSelect = document.getElementById('yearEditPlanName');
            if (planSelect && planSelect.options.length > 1) {
                // 使用現有的計畫選項
                for (let i = 1; i < planSelect.options.length; i++) {
                    const opt = planSelect.options[i];
                    const planValue = opt.value;
                    const { name: planName, year: planYear } = parsePlanValue(planValue);
                    const displayText = planYear ? `${planName} (${planYear})` : planName;
                    const currentPlanName = item.plan_name || item.planName || '';
                    const isSelected = (currentPlanName === planName && (!planYear || item.year === planYear)) || 
                                      (planValue && planValue === `${currentPlanName}|||${item.year}`);
                    planOptionsHtml += `<option value="${planValue}" ${isSelected ? 'selected' : ''}>${displayText}</option>`;
                }
            } else {
                // 如果計畫選項還沒有加載，先添加當前計畫（如果有的話）
                const currentPlanName = item.plan_name || item.planName || '';
                if (currentPlanName) {
                    const currentPlanValue = item.year ? `${currentPlanName}|||${item.year}` : currentPlanName;
                    const displayText = item.year ? `${currentPlanName} (${item.year})` : currentPlanName;
                    planOptionsHtml += `<option value="${currentPlanValue}" selected>${displayText}</option>`;
                }
                // 嘗試加載計畫選項（異步，不阻塞渲染）
                loadPlanOptions().then(() => {
                    // 重新渲染計畫選項
                    const planSelectEl = document.getElementById('yearEditPlanNameSelect');
                    if (planSelectEl && document.getElementById('yearEditPlanName')) {
                        const sourceSelect = document.getElementById('yearEditPlanName');
                        if (sourceSelect && sourceSelect.options.length > 1) {
                            let newOptionsHtml = '<option value="">(未指定)</option>';
                            for (let i = 1; i < sourceSelect.options.length; i++) {
                                const opt = sourceSelect.options[i];
                                const planValue = opt.value;
                                const { name: planName, year: planYear } = parsePlanValue(planValue);
                                const displayText = planYear ? `${planName} (${planYear})` : planName;
                                const isSelected = planSelectEl.value === planValue;
                                newOptionsHtml += `<option value="${planValue}" ${isSelected ? 'selected' : ''}>${displayText}</option>`;
                            }
                            planSelectEl.innerHTML = newOptionsHtml;
                        }
                    }
                }).catch(() => {
                    // 忽略錯誤，使用當前選項
                });
            }
            
            // 確定當前計畫的值（支援兩種字段名格式）
            const currentPlanName = item.plan_name || item.planName || '';
            const currentPlanValue = currentPlanName ? (item.year ? `${currentPlanName}|||${item.year}` : currentPlanName) : '';
            
            let html = `
                <div class="detail-card" style="margin-bottom:20px; border:2px solid #e2e8f0;">
                    <!-- 基本資訊區塊 -->
                    <div style="background:#f8fafc; padding:20px; border-bottom:2px solid #e2e8f0;">
                        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:20px; margin-bottom:16px;">
                            <div>
                                <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">
                                    事項編號 <span style="color:#ef4444;">*</span>
                                </label>
                                <input type="text" id="yearEditNumber" class="filter-input" 
                                    value="${item.number || ''}" 
                                    placeholder="例如: 113-TRA-1-A01-N01" 
                                    style="width:100%; background:white;">
                            </div>
                            <div>
                                <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">
                                    年度 <span style="color:#ef4444;">*</span>
                                </label>
                                <input type="number" id="yearEditYear" class="filter-input" 
                                    value="${item.year || ''}" 
                                    placeholder="例如: 113" 
                                    style="width:100%; background:white;">
                            </div>
                        </div>
                        
                        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:20px; margin-bottom:16px;">
                            <div>
                                <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">
                                    機構 <span style="color:#ef4444;">*</span>
                                </label>
                                <input type="text" id="yearEditUnit" class="filter-input" 
                                    value="${item.unit || ''}" 
                                    placeholder="例如: 臺鐵" 
                                    style="width:100%; background:white;">
                            </div>
                            <div>
                                <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">分組</label>
                                <select id="yearEditDivision" class="filter-select" style="width:100%; background:white;">
                                    <option value="">(未指定)</option>
                                    <option value="運務" ${(item.divisionName || item.division_name) === '運務' ? 'selected' : ''}>運務</option>
                                    <option value="工務" ${(item.divisionName || item.division_name) === '工務' ? 'selected' : ''}>工務</option>
                                    <option value="機務" ${(item.divisionName || item.division_name) === '機務' ? 'selected' : ''}>機務</option>
                                    <option value="電務" ${(item.divisionName || item.division_name) === '電務' ? 'selected' : ''}>電務</option>
                                    <option value="安全" ${(item.divisionName || item.division_name) === '安全' ? 'selected' : ''}>安全</option>
                                    <option value="審核" ${(item.divisionName || item.division_name) === '審核' ? 'selected' : ''}>審核</option>
                                    <option value="災防" ${(item.divisionName || item.division_name) === '災防' ? 'selected' : ''}>災防</option>
                                    <option value="運轉" ${(item.divisionName || item.division_name) === '運轉' ? 'selected' : ''}>運轉</option>
                                    <option value="土木" ${(item.divisionName || item.division_name) === '土木' ? 'selected' : ''}>土木</option>
                                    <option value="機電" ${(item.divisionName || item.division_name) === '機電' ? 'selected' : ''}>機電</option>
                                </select>
                            </div>
                        </div>
                        
                        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:20px; margin-bottom:16px;">
                            <div>
                                <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">檢查種類</label>
                                <select id="yearEditInspection" class="filter-select" style="width:100%; background:white;">
                                    <option value="">(未指定)</option>
                                    <option value="定期檢查" ${(item.inspectionCategoryName || item.inspection_category_name) === '定期檢查' ? 'selected' : ''}>定期檢查</option>
                                    <option value="例行性檢查" ${(item.inspectionCategoryName || item.inspection_category_name) === '例行性檢查' ? 'selected' : ''}>例行性檢查</option>
                                    <option value="特別檢查" ${(item.inspectionCategoryName || item.inspection_category_name) === '特別檢查' ? 'selected' : ''}>特別檢查</option>
                                    <option value="臨時檢查" ${(item.inspectionCategoryName || item.inspection_category_name) === '臨時檢查' ? 'selected' : ''}>臨時檢查</option>
                                </select>
                            </div>
                            <div>
                                <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">開立類型</label>
                                <select id="yearEditKind" class="filter-select" style="width:100%; background:white;">
                                    <option value="">(未指定)</option>
                                    <option value="N" ${(item.item_kind_code || item.itemKindCode) === 'N' || item.category === '缺失事項' ? 'selected' : ''}>缺失事項</option>
                                    <option value="O" ${(item.item_kind_code || item.itemKindCode) === 'O' || item.category === '觀察事項' ? 'selected' : ''}>觀察事項</option>
                                    <option value="R" ${(item.item_kind_code || item.itemKindCode) === 'R' || item.category === '建議事項' ? 'selected' : ''}>建議事項</option>
                                </select>
                            </div>
                        </div>
                        
                        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:20px; margin-bottom:16px;">
                            <div>
                                <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">檢查計畫</label>
                                <select id="yearEditPlanNameSelect" class="filter-select" style="width:100%; background:white;">
                                    ${planOptionsHtml}
                                </select>
                            </div>
                            <div>
                                <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">狀態</label>
                                <select id="yearEditStatus" class="filter-select" style="width:100%; background:white;">
                                    <option value="持續列管" ${item.status === '持續列管' ? 'selected' : ''}>持續列管</option>
                                    <option value="解除列管" ${item.status === '解除列管' ? 'selected' : ''}>解除列管</option>
                                    <option value="自行列管" ${item.status === '自行列管' ? 'selected' : ''}>自行列管</option>
                                </select>
                            </div>
                        </div>
                        
                        <div style="margin-bottom:16px;">
                            <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">開立日期</label>
                            <input type="text" id="yearEditIssueDate" class="filter-input" 
                                value="${item.issue_date || ''}" 
                                placeholder="例如: 1130501" 
                                style="width:100%; background:white;">
                        </div>
                        
                        <div>
                            <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">事項內容</label>
                            <textarea id="yearEditContent" class="filter-input" 
                                style="width:100%; min-height:120px; padding:12px; font-size:14px; line-height:1.6; resize:vertical; background:white;">${stripHtml(item.content || '')}</textarea>
                        </div>
                    </div>
            `;
            
            // 如果有審查和回復紀錄，添加該區塊
            if (hasReviewRecords) {
                html += `
                    <!-- 所有輪次的審查與回復紀錄 -->
                    <div style="padding:20px;">
                        <div style="font-weight:700; font-size:16px; color:#334155; margin-bottom:16px; padding-bottom:12px; border-bottom:2px solid #e2e8f0;">
                            📋 所有審查及回復紀錄（共 ${rounds.length} 輪）
                        </div>
                        
                        <div id="yearEditRoundsContainer">
                `;
                
                // 渲染每個輪次
                rounds.forEach((round, index) => {
                    const isLast = index === rounds.length - 1;
                    html += `
                    <div class="detail-card" style="margin-bottom:16px; border:1px solid #e2e8f0; ${isLast ? 'border-left:4px solid #2563eb;' : ''}">
                        <div style="background:#eff6ff; padding:12px; border-bottom:1px solid #dbeafe; display:flex; justify-content:space-between; align-items:center;">
                            <div style="font-weight:700; color:#1e40af; font-size:15px;">
                                第 ${round.round} 次回復與審查
                            </div>
                            <div style="display:flex; gap:12px; font-size:13px; color:#64748b;">
                                ${round.replyDate ? `<span>鐵路機構回復日期：${round.replyDate}</span>` : ''}
                                ${round.responseDate ? `<span>本次函復日期：${round.responseDate}</span>` : ''}
                            </div>
                        </div>
                        <div style="padding:16px;">
                            <div style="margin-bottom:16px;">
                                <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">
                                    辦理情形 (第 ${round.round} 次回復與審查)
                                </label>
                                <textarea class="filter-input year-edit-round-handling" data-round="${round.round}" 
                                    style="width:100%; min-height:100px; padding:12px; font-size:14px; line-height:1.6; resize:vertical;">${round.handling}</textarea>
                            </div>
                            <div>
                                <label style="display:block; font-weight:600; color:#475569; font-size:14px; margin-bottom:8px;">
                                    審查意見 (第 ${round.round} 次回復與審查)
                                </label>
                                <textarea class="filter-input year-edit-round-review" data-round="${round.round}" 
                                    style="width:100%; min-height:100px; padding:12px; font-size:14px; line-height:1.6; resize:vertical;">${round.review}</textarea>
                            </div>
                            <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px; margin-top:12px;">
                                <div>
                                    <label style="display:block; font-weight:600; color:#475569; font-size:13px; margin-bottom:6px;">鐵路機構回復日期</label>
                                    <input type="text" class="filter-input year-edit-round-reply-date" data-round="${round.round}" 
                                        value="${round.replyDate || ''}" placeholder="例如: 1130601" style="width:100%;">
                                </div>
                                <div>
                                    <label style="display:block; font-weight:600; color:#475569; font-size:13px; margin-bottom:6px;">本次函復日期</label>
                                    <input type="text" class="filter-input year-edit-round-response-date" data-round="${round.round}" 
                                        value="${round.responseDate || ''}" placeholder="例如: 1130615" style="width:100%;">
                                </div>
                            </div>
                        </div>
                    </div>
                `;
                });
                
                html += `
                        </div>
                    </div>
                `;
            }
            // 如果沒有輪次記錄，不顯示任何辦理情形編輯區塊（保持原有邏輯）
            
            html += `
                </div>
            `;
            
            container.innerHTML = html;
        }
        
        // 儲存事項變更
        async function saveYearEditIssue() {
            if (!yearEditIssue) {
                showToast('無事項可儲存', 'error');
                return;
            }
            
            const confirmed = await showConfirmModal('確定要儲存所有變更嗎？', '確定儲存', '取消');
            if (!confirmed) {
                return;
            }
            
            try {
                showToast('儲存中，請稍候...', 'info');
                
                const issueId = yearEditIssue.id;
                // 不進行 trim，保留原始輸入（包括空字串），允許清空欄位
                const number = document.getElementById('yearEditNumber')?.value.trim() || '';
                const year = document.getElementById('yearEditYear')?.value.trim() || '';
                const unit = document.getElementById('yearEditUnit')?.value.trim() || '';
                const divisionName = document.getElementById('yearEditDivision')?.value || '';
                const inspectionCategoryName = document.getElementById('yearEditInspection')?.value || '';
                const itemKindCode = document.getElementById('yearEditKind')?.value || '';
                const planValue = document.getElementById('yearEditPlanNameSelect')?.value || '';
                const { name: planName } = parsePlanValue(planValue);
                const content = document.getElementById('yearEditContent').value;
                const status = document.getElementById('yearEditStatus').value;
                const issueDate = document.getElementById('yearEditIssueDate').value;
                
                // 基本驗證
                if (!number) {
                    showToast('請填寫事項編號', 'error');
                    return;
                }
                if (!year) {
                    showToast('請填寫年度', 'error');
                    return;
                }
                if (!unit) {
                    showToast('請填寫機構', 'error');
                    return;
                }
                
                // 收集所有輪次的資料
                const roundHandlings = document.querySelectorAll('.year-edit-round-handling');
                const roundReviews = document.querySelectorAll('.year-edit-round-review');
                const roundReplyDates = document.querySelectorAll('.year-edit-round-reply-date');
                const roundResponseDates = document.querySelectorAll('.year-edit-round-response-date');
                
                // 找出所有顯示的輪次（不管是否有內容）
                const roundSet = new Set();
                roundHandlings.forEach(el => roundSet.add(parseInt(el.dataset.round)));
                roundReviews.forEach(el => roundSet.add(parseInt(el.dataset.round)));
                roundReplyDates.forEach(el => roundSet.add(parseInt(el.dataset.round)));
                roundResponseDates.forEach(el => roundSet.add(parseInt(el.dataset.round)));
                
                const sortedRounds = Array.from(roundSet).sort((a, b) => a - b);
                
                // 先更新基本資訊（包括所有可編輯欄位）
                // 即使內容為空也要更新（允許清空）
                const updateRes = await apiFetch(`/api/issues/${issueId}`, {
                    method: 'PUT',
                    body: JSON.stringify({
                        status: status,
                        round: 1,
                        handling: '', // 第一輪的辦理情形和審查意見會在後面更新
                        review: '',
                        content: content, // 允許空字串
                        issueDate: issueDate || '', // 允許空字串
                        number: number,
                        year: year,
                        unit: unit,
                        divisionName: divisionName || null,
                        inspectionCategoryName: inspectionCategoryName || null,
                        itemKindCode: itemKindCode || null,
                        category: itemKindCode ? (itemKindCode === 'N' ? '缺失事項' : itemKindCode === 'O' ? '觀察事項' : '建議事項') : null,
                        planName: planName || null,
                        replyDate: '',
                        responseDate: ''
                    })
                });
                
                if (!updateRes.ok) {
                    const errorData = await updateRes.json().catch(() => ({}));
                    throw new Error(errorData.error || '更新基本資訊失敗');
                }
                
                // 更新每個輪次（包括清空的欄位）
                let successCount = 0;
                let errorCount = 0;
                
                // 更新所有顯示的輪次，即使內容為空也要更新（允許清空欄位）
                for (const roundNum of sortedRounds) {
                    const handlingEl = document.querySelector(`.year-edit-round-handling[data-round="${roundNum}"]`);
                    const reviewEl = document.querySelector(`.year-edit-round-review[data-round="${roundNum}"]`);
                    const replyDateEl = document.querySelector(`.year-edit-round-reply-date[data-round="${roundNum}"]`);
                    const responseDateEl = document.querySelector(`.year-edit-round-response-date[data-round="${roundNum}"]`);
                    
                    // 取得值（包括空字串，允許清空）
                    const handling = handlingEl ? handlingEl.value : '';
                    const review = reviewEl ? reviewEl.value : '';
                    const replyDate = replyDateEl ? replyDateEl.value : '';
                    const responseDate = responseDateEl ? responseDateEl.value : '';
                    
                    // 所有顯示的輪次都要更新，即使內容為空（允許清空欄位）
                    try {
                        const updateRes = await apiFetch(`/api/issues/${issueId}`, {
                            method: 'PUT',
                            body: JSON.stringify({
                                status: status, // 保持當前狀態
                                round: roundNum,
                                handling: handling, // 允許空字串
                                review: review, // 允許空字串
                                replyDate: replyDate, // 允許空字串
                                responseDate: responseDate // 允許空字串
                            })
                        });
                        
                        if (updateRes.ok) {
                            successCount++;
                        } else {
                            const errorData = await updateRes.json().catch(() => ({}));
                            console.error(`更新第 ${roundNum} 輪失敗:`, errorData.error || updateRes.statusText);
                            errorCount++;
                        }
                    } catch (e) {
                        console.error(`更新第 ${roundNum} 輪失敗:`, e);
                        errorCount++;
                    }
                }
                
                if (successCount > 0 || errorCount === 0) {
                    showToast(`儲存成功${errorCount > 0 ? `（${errorCount} 個輪次更新失敗）` : ''}`, 
                        errorCount > 0 ? 'warning' : 'success');
                    // 重新載入當前事項的資料（通過編號查詢）
                    try {
                        const currentNumber = document.getElementById('yearEditNumber')?.value.trim() || yearEditIssue?.number;
                        if (currentNumber) {
                            const res = await fetch(`/api/issues?page=1&pageSize=1&q=${encodeURIComponent(currentNumber)}&_t=${Date.now()}`);
                            if (res.ok) {
                                const json = await res.json();
                                if (json.data && json.data.length > 0) {
                                    yearEditIssue = json.data[0];
                                    // 標準化字段名（確保同時有兩種格式，提高兼容性）
                                    if (yearEditIssue.division_name && !yearEditIssue.divisionName) {
                                        yearEditIssue.divisionName = yearEditIssue.division_name;
                                    }
                                    if (yearEditIssue.inspection_category_name && !yearEditIssue.inspectionCategoryName) {
                                        yearEditIssue.inspectionCategoryName = yearEditIssue.inspection_category_name;
                                    }
                                    if (yearEditIssue.item_kind_code && !yearEditIssue.itemKindCode) {
                                        yearEditIssue.itemKindCode = yearEditIssue.item_kind_code;
                                    }
                                    if (yearEditIssue.plan_name && !yearEditIssue.planName) {
                                        yearEditIssue.planName = yearEditIssue.plan_name;
                                    }
                                    // 重新渲染事項內容
                                    renderYearEditIssue();
                                }
                            }
                        }
                    } catch (reloadError) {
                        console.error('重新載入事項資料失敗:', reloadError);
                        // 即使重新載入失敗，也顯示成功訊息（因為已經保存成功）
                        // 嘗試使用當前輸入的值更新 yearEditIssue 並重新渲染
                        if (yearEditIssue) {
                            yearEditIssue.number = document.getElementById('yearEditNumber')?.value.trim() || yearEditIssue.number;
                            yearEditIssue.year = document.getElementById('yearEditYear')?.value.trim() || yearEditIssue.year;
                            yearEditIssue.unit = document.getElementById('yearEditUnit')?.value.trim() || yearEditIssue.unit;
                            // 同時更新兩種格式的字段名（確保兼容性）
                            const divisionValue = document.getElementById('yearEditDivision')?.value || '';
                            yearEditIssue.divisionName = divisionValue;
                            yearEditIssue.division_name = divisionValue;
                            const inspectionValue = document.getElementById('yearEditInspection')?.value || '';
                            yearEditIssue.inspectionCategoryName = inspectionValue;
                            yearEditIssue.inspection_category_name = inspectionValue;
                            const kindValue = document.getElementById('yearEditKind')?.value || '';
                            yearEditIssue.item_kind_code = kindValue;
                            yearEditIssue.itemKindCode = kindValue;
                            const planValue = document.getElementById('yearEditPlanNameSelect')?.value || '';
                            const { name: planName } = parsePlanValue(planValue);
                            if (planName) {
                                yearEditIssue.plan_name = planName;
                                yearEditIssue.planName = planName;
                            }
                            yearEditIssue.status = document.getElementById('yearEditStatus')?.value || yearEditIssue.status;
                            yearEditIssue.issue_date = document.getElementById('yearEditIssueDate')?.value || yearEditIssue.issue_date;
                            yearEditIssue.content = document.getElementById('yearEditContent')?.value || yearEditIssue.content;
                            renderYearEditIssue();
                        }
                    }
                } else {
                    showToast('儲存失敗', 'error');
                }
            } catch (e) {
                showToast('儲存時發生錯誤: ' + e.message, 'error');
            }
        }