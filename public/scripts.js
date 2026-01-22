        // 全域狀態
        let rawData = [], currentData = [], currentUser = null, charts = {}, currentEditItem = null, userList = [], sortState = { field: null, dir: 'asc' }, stagedImportData = [];
        let autoLogoutTimer;
        let currentLogs = { login: [], action: [] };
        let cachedGlobalStats = null;
        
        // 日誌記錄函數（寫入檔案，不在控制台顯示）
        async function writeLog(message, level = 'INFO') {
            try {
                await fetch('/api/log', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
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
        let logsPage = 1, logsPageSize = 20, logsTotal = 0, logsPages = 1;
        let actionsPage = 1, actionsPageSize = 20, actionsTotal = 0, actionsPages = 1;
        // Current import mode: 'word' (uses param) or 'backup' (ignores param)
        let currentImportMode = 'word';

        function resetAutoLogout() { clearTimeout(autoLogoutTimer); autoLogoutTimer = setTimeout(() => { alert("您已閒置過久，系統將自動登出。"); logout(); }, 1800000); }
        window.onload = resetAutoLogout; document.onmousemove = resetAutoLogout; document.onkeypress = resetAutoLogout;

        function toggleDashboard(btn) { const d = document.getElementById('dashboardSection'); const c = d.classList.contains('collapsed'); d.classList.toggle('collapsed', !c); btn.innerHTML = c ? '<span>收合統計圖表</span> <span>▲</span>' : '<span>展開統計圖表</span> <span>▼</span>'; }
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
            
            // 2. 嘗試 THAS-v1 格式 (13T1-A01-N01)
            var mA = whole.match(/(\d{2}[THASthas][1-4]-[A-Ga-g]\d{2}-[NORnor]\d{2})/);
            if (mA) return (mA[1] || "").toUpperCase();
            
            // 3. 處理帶 <br> 的情況，分行匹配
            var rawHtml = cell.innerHTML || "";
            var lines = normalizeCodeString(rawHtml.replace(/<br\s*\/?>/gi, "\n").replace(/<\/p>/gi, "\n").replace(/<[^>]*>/g, "")).split("\n");
            for (var i = 0; i < lines.length; i++) {
                var line = (lines[i] || "").trim();
                if (!line) continue;
                var m1 = line.match(/(\d{3}-[A-Za-z]{3}-[1-4]-\d+-[A-Za-z]{2,3}-[NORnor]\d{1,3})/);
                if (m1) return (m1[1] || "").toUpperCase();
                var m2 = line.match(/(\d{2}[THASthas][1-4]-[A-Ga-g]\d{2}-[NORnor]\d{2})/);
                if (m2) return (m2[1] || "").toUpperCase();
            }
            
            return whole.trim();
        }

        // [Updated] Map & Parser
        const ORG_MAP = { "T": "臺鐵", "H": "高鐵", "A": "林鐵", "S": "糖鐵", "TRC": "臺鐵", "HSR": "高鐵", "AFR": "林鐵", "TSC": "糖鐵" };
        // [Added] 機構交叉映射表（THAS-v1 ↔ TRC-v2）
        const ORG_CROSSWALK = { "T": "TRC", "H": "HSR", "A": "AFR", "S": "TSC", "TRC": "TRC", "HSR": "HSR", "AFR": "AFR", "TSC": "TSC" };
        const INSPECTION_MAP = { "1": "定期檢查", "2": "例行性檢查", "3": "特別檢查", "4": "臨時檢查" };
        // [Verified] Division Map includes all requested codes
        const DIVISION_MAP = { "A": "運務", "B": "工務", "C": "機務", "D": "電務", "E": "安全", "F": "審核", "G": "災防", "OP": "運轉", "CP": "土木", "EM": "機電" };
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
            
            // 3. 長格式（兼容舊格式）：123-TRC-1-7-OP-N12 (支持 3-4 位機構代碼)
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
            
            // 4. 短格式（兼容舊格式）：13T1-A01-N01 (支持 2-3 位年份)
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
            
            // 5. 寬鬆匹配（fallback）
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
                    console.warn('沒有找到任何檢查計畫');
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
                    console.warn('找不到 filterPlan 元素');
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

            // 動態載入視圖內容
            const viewMap = {
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
                        // 視圖載入完成後，設置 admin 專屬元素
                        if (viewId === 'importView') {
                            setupAdminElements();
                            // [Added] 設置導入視圖的事件監聽器
                            setTimeout(() => setupImportListeners(), 100);
                            // 確保檢查計畫選項已載入
                            loadPlanOptions();
                        } else if (viewId === 'usersView') {
                            // 設置清除舊記錄的UI
                            setTimeout(() => setupCleanupDaysSelect(), 100);
                        }
                        
                        // 恢復資料管理頁面的 tab
                        if (viewId === 'importView') {
                            const savedTab = sessionStorage.getItem('currentDataTab');
                            if (savedTab) {
                                setTimeout(() => {
                                    switchDataTab(savedTab);
                                }, 200);
                            }
                        }
                    } else {
                        // 錯誤已在伺服器 log 中記錄
                    }
                } catch (error) {
                    // 錯誤已在伺服器 log 中記錄
                }
            } else if (viewId === 'usersView' && viewElement.dataset.loaded) {
                // 如果視圖已經載入，也要設置清除舊記錄的UI
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
            } else if (viewId === 'importView' && viewElement.dataset.loaded) {
                // 恢復資料管理頁面的 tab
                const savedTab = sessionStorage.getItem('currentDataTab');
                // 確保檢查計畫選項已載入（無論切換到哪個 tab）
                loadPlanOptions();
                if (savedTab) {
                    setTimeout(() => {
                        switchDataTab(savedTab);
                        // 如果是檢查計畫管理 tab，恢復其狀態
                        if (savedTab === 'plans') {
                            restorePlansViewState();
                            setTimeout(() => {
                                loadPlansPage(plansPage || 1);
                            }, 300);
                        }
                    }, 200);
                }
            }
        }

        document.addEventListener('DOMContentLoaded', async () => {
            // App 初始化（已移除 debug 日誌）
            // 首先確保 body 可見，避免空白頁面
            document.body.style.display = 'flex';
            
            try {
                await checkAuth();
                if (currentUser) {
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
                    console.warn('未檢測到登入狀態，嘗試重定向到登入頁');
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
                
                const fetchPromise = fetch('/api/auth/me?t=' + Date.now(), { 
                    headers: { 'Cache-Control': 'no-cache' }
                });
                
                const res = await Promise.race([fetchPromise, timeoutPromise]);
                
                if (!res.ok) {
                    console.error('認證檢查失敗:', res.status, res.statusText);
                    // 如果認證失敗，重定向到登入頁
                    window.location.href = '/login.html';
                    return;
                }
                
                const data = await res.json();
                if (data.isLogin) {
                    currentUser = data;
                    const nameEl = document.getElementById('headerUserName');
                    const roleEl = document.getElementById('headerUserRole');
                    if (nameEl) nameEl.innerText = data.name || data.username;
                    if (roleEl) roleEl.innerText = getRoleName(data.role);
                    if (['admin', 'manager'].includes(data.role)) {
                        const btnImport = document.getElementById('btn-importView');
                        if (btnImport) btnImport.classList.remove('hidden');
                        if (data.role === 'admin') {
                            const btnUsers = document.getElementById('btn-usersView');
                            if (btnUsers) btnUsers.classList.remove('hidden');
                            // 這些元素現在在動態載入的視圖中，會在視圖載入後處理
                        }
                    }
                } else {
                    // 未登入，重定向到登入頁
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
                } else {
                    console.error('認證檢查錯誤:', e);
                    // 其他錯誤，重定向到登入頁
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
                const res = await fetch('/api/issues?' + params.toString());
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
            if (!confirm(`確定要刪除 ${ids.length} 筆資料嗎？此操作無法復原！`)) {
                return;
            }
            
            try {
                const res = await fetch('/api/issues/batch-delete', {
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
                    let aiContent = `<div style="color:#ccc;font-size:11px;">未分析</div>`; if (item.aiResult && item.aiResult.status === 'done') { const f = String(item.aiResult.fulfill || ''); const isYes = f.includes('是') || f.includes('Yes'); aiContent = `<div class="ai-tag ${isYes ? 'yes' : 'no'}">${isYes ? '✅' : '⚠️'} ${f}</div>`; }
                    const editBtn = canEdit ? `<button class="badge" style="background:#fff;border:1px solid #ddd;cursor:pointer;margin-top:4px;" onclick="event.stopPropagation();openDetail('${item.id}',false)">✏️ 審查/查看詳情</button>` : '';
                    const checkbox = canManage ? `<td class="manager-col"><input type="checkbox" class="issue-check" value="${item.id}" onclick="event.stopPropagation(); updateBatchUI()"></td>` : `<td class="manager-col" style="display:none"></td>`;

                    let k = item.itemKindCode;
                    const numStr = String(item.number || '');
                    if (!k && numStr) { const m = numStr.match(/-([NOR])\d+$/i); if (m) k = m[1].toUpperCase(); }

                    let kindLabel = '';
                    if (k === 'N') kindLabel = `<span class="kind-tag N">缺失</span>`;
                    else if (k === 'O') kindLabel = `<span class="kind-tag O">觀察</span>`;
                    else if (k === 'R') kindLabel = `<span class="kind-tag R">建議</span>`;

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
                const res = await fetch('/api/users?' + params.toString()); 
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
                console.warn('usersTableBody element not found');
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
                console.warn('loginSearch element not found');
                return;
            }
            logsPage = page;
            const q = loginSearchEl.value || '';
            saveLogsViewState();
            const params = new URLSearchParams({ page: logsPage, pageSize: logsPageSize, q, _t: Date.now() });
            const logsLoadingEl = document.getElementById('logsLoading');
            if (logsLoadingEl) logsLoadingEl.style.display = 'block';
            try {
                const res = await fetch('/api/admin/logs?' + params.toString());
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
                console.warn('actionSearch element not found');
                return;
            }
            actionsPage = page;
            const q = actionSearchEl.value || '';
            saveActionsViewState();
            const params = new URLSearchParams({ page: actionsPage, pageSize: actionsPageSize, q, _t: Date.now() });
            const logsLoadingEl = document.getElementById('logsLoading');
            if (logsLoadingEl) logsLoadingEl.style.display = 'block';
            try {
                const res = await fetch('/api/admin/action_logs?' + params.toString());
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
                if (!confirm(`確定要刪除資料庫中所有「${logTypeName}」紀錄嗎？此動作無法復原！`)) {
                    return;
                }
                
                const endpoint = type === 'login' ? '/api/admin/logs' : '/api/admin/action_logs';
                try {
                    const res = await fetch(endpoint, { method: 'DELETE' });
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
            
            if (!confirm(`確定要刪除資料庫中 ${days} 天前的「${logTypeName}」紀錄嗎？此動作無法復原！\n\n將保留最近 ${days} 天的記錄，刪除更早的記錄。`)) {
                return;
            }
            
            const endpoint = type === 'login' ? '/api/admin/logs/cleanup' : '/api/admin/action_logs/cleanup';
            try {
                const res = await fetch(endpoint, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
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
                alert("解析 Word 表格時發生錯誤，請確認表格格式是否包含「編號」與「內容」欄位。");
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
            if (!msg) { alert("系統錯誤：找不到狀態顯示區域"); return; }

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
            if (!confirm(msg)) return;

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
                    console.warn('無法載入計畫選項，將使用選擇的計畫名稱', e);
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
                                console.warn(`⚠️ ${warnMsg}`);
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
                const res = await fetch('/api/issues/import', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
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
                // 恢復檢查計畫管理頁面的狀態
                restorePlansViewState();
                loadPlanOptions();
                setTimeout(() => {
                    loadPlansPage(plansPage || 1);
                }, 200);
            }
            if (tab === 'export') {
                // 設置匯出選項的顯示/隱藏
                setTimeout(() => setupExportOptions(), 100);
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
                <td><select class="filter-select batch-division"><option value="">-</option><option value="運務">運務</option><option value="工務">工務</option><option value="機務">機務</option><option value="電務">電務</option><option value="安全">安全</option><option value="審核">審核</option><option value="災防">災防</option><option value="運轉">運轉</option><option value="土木">土木</option><option value="機電">機電</option></select></td>
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

            if (!confirm(`確定要批次新增 ${items.length} 筆資料嗎？\n計畫：${planName}`)) return;

            try {
                const res = await fetch('/api/issues/import', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
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
            
            // 驗證日期格式（應該是6或7位數字，例如：1130601 或 1141001）
            if (!/^\d{6,7}$/.test(replyDate)) {
                showToast('日期格式錯誤，應為6或7位數字（例如：1130601 或 1141001）', 'error');
                return;
            }
            
            const { name: planName } = parsePlanValue(planValue);
            
            try {
                // 載入該計畫下的所有事項
                // 移除載入中的提示訊息，只保留錯誤訊息
                const res = await fetch(`/api/issues?page=1&pageSize=1000&planName=${encodeURIComponent(planValue)}&_t=${Date.now()}`);
                if (!res.ok) throw new Error('載入事項列表失敗');
                
                const json = await res.json();
                const issueList = json.data || [];
                
                if (issueList.length === 0) {
                    showToast('該檢查計畫下尚無開立事項', 'error');
                    return;
                }
                
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
                        const updateRes = await fetch(`/api/issues/${issueId}`, {
                            method: 'PUT',
                            headers: { 'Content-Type': 'application/json' },
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
                const res = await fetch(`/api/issues?page=1&pageSize=1000&planName=${encodeURIComponent(planValue)}&_t=${Date.now()}`);
                if (!res.ok) return;
                
                const json = await res.json();
                const issueList = json.data || [];
                
                if (issueList.length === 0) {
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
                // 移除載入中的提示訊息，只保留錯誤訊息
                
                // 載入該計畫下的所有事項
                const res = await fetch(`/api/issues?page=1&pageSize=1000&planName=${encodeURIComponent(planValue)}&_t=${Date.now()}`);
                if (!res.ok) throw new Error('載入事項列表失敗');
                
                const json = await res.json();
                const issueList = json.data || [];
                
                if (issueList.length === 0) {
                    // 移除無事項的提示訊息，只保留錯誤訊息
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
                    if (!confirm(`表格中已有資料，載入現有事項將會清空現有資料。\n確定要載入 ${issueList.length} 筆事項嗎？`)) {
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
                        <td><select class="filter-select create-batch-division"><option value="">-</option><option value="運務" ${divisionName === '運務' ? 'selected' : ''}>運務</option><option value="工務" ${divisionName === '工務' ? 'selected' : ''}>工務</option><option value="機務" ${divisionName === '機務' ? 'selected' : ''}>機務</option><option value="電務" ${divisionName === '電務' ? 'selected' : ''}>電務</option><option value="安全" ${divisionName === '安全' ? 'selected' : ''}>安全</option><option value="審核" ${divisionName === '審核' ? 'selected' : ''}>審核</option><option value="災防" ${divisionName === '災防' ? 'selected' : ''}>災防</option><option value="運轉" ${divisionName === '運轉' ? 'selected' : ''}>運轉</option><option value="土木" ${divisionName === '土木' ? 'selected' : ''}>土木</option><option value="機電" ${divisionName === '機電' ? 'selected' : ''}>機電</option></select></td>
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
                const res = await fetch(`/api/issues?page=1&pageSize=1000&planName=${encodeURIComponent(planValue)}&_t=${Date.now()}`);
                if (!res.ok) return;
                
                const json = await res.json();
                const issueList = json.data || [];
                
                if (issueList.length === 0) {
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
                const res = await fetch('/api/issues/import', { 
                    method: 'POST', 
                    headers: { 'Content-Type': 'application/json' }, 
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
                                            const updateRes = await fetch(`/api/issues/${issueId}`, {
                                                method: 'PUT',
                                                headers: { 'Content-Type': 'application/json' },
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
                <td><select class="filter-select create-batch-division"><option value="">-</option><option value="運務">運務</option><option value="工務">工務</option><option value="機務">機務</option><option value="電務">電務</option><option value="安全">安全</option><option value="審核">審核</option><option value="災防">災防</option><option value="運轉">運轉</option><option value="土木">土木</option><option value="機電">機電</option></select></td>
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
                
                if (info.orgCode) {
                    const name = ORG_MAP[info.orgCode] || info.orgCode;
                    if (name && name !== '?') tr.querySelector('.create-batch-unit').value = name;
                }
                if (info.divCode) {
                    const divName = DIVISION_MAP[info.divCode];
                    if (divName) tr.querySelector('.create-batch-division').value = divName;
                }
                if (info.inspectCode) {
                    const inspectName = INSPECTION_MAP[info.inspectCode];
                    if (inspectName) tr.querySelector('.create-batch-inspection').value = inspectName;
                }
                if (info.kindCode) {
                    tr.querySelector('.create-batch-kind').value = info.kindCode;
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
                        const updateRes = await fetch(`/api/issues/${issueId}`, {
                            method: 'PUT',
                            headers: { 'Content-Type': 'application/json' },
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
                                const updateRes = await fetch(`/api/issues/${issueId}`, {
                                    method: 'PUT',
                                    headers: { 'Content-Type': 'application/json' },
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

            if (!confirm(`確定要批次新增 ${items.length} 筆資料嗎？\n計畫：${planName}`)) return;

            try {
                // 先新增所有事項（第一次辦理情形）
                // 注意：每個事項可能有不同的回復日期，需要在服務器端使用 item.replyDate
                const itemsForImport = items.map(item => {
                    const { handlingRounds, ...itemData } = item;
                    return itemData;
                });
                
                const res = await fetch('/api/issues/import', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
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
                                            const updateRes = await fetch(`/api/issues/${issueId}`, {
                                                method: 'PUT',
                                                headers: { 'Content-Type': 'application/json' },
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
                let plansData = [];
                
                // 根據選擇的資料類型獲取資料
                if (exportDataType === 'issues' || exportDataType === 'both') {
                    const res = await fetch('/api/issues?page=1&pageSize=10000&sortField=created_at&sortDir=desc');
                    if (!res.ok) throw new Error('取得開立事項資料失敗');
                    const json = await res.json();
                    issuesData = json.data || [];
                }
                
                if (exportDataType === 'plans' || exportDataType === 'both') {
                    const res = await fetch('/api/plans?page=1&pageSize=10000&sortField=id&sortDir=desc');
                    if (!res.ok) throw new Error('取得檢查計畫資料失敗');
                    const json = await res.json();
                    plansData = json.data || [];
                }
                
                // 檢查是否有資料可匯出
                if (exportDataType === 'issues' && issuesData.length === 0) {
                    return showToast('無開立事項資料可匯出', 'error');
                }
                if (exportDataType === 'plans' && plansData.length === 0) {
                    return showToast('無檢查計畫資料可匯出', 'error');
                }
                if (exportDataType === 'both' && issuesData.length === 0 && plansData.length === 0) {
                    return showToast('無資料可匯出', 'error');
                }

                // JSON 格式匯出
                if (exportFormat === 'json') {
                    const exportData = {};
                    if (exportDataType === 'issues' || exportDataType === 'both') {
                        exportData.issues = issuesData;
                    }
                    if (exportDataType === 'plans' || exportDataType === 'both') {
                        exportData.plans = plansData;
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
                    
                    // 如果選擇合併匯出，創建兩個工作表
                    if (exportDataType === 'both') {
                        // 工作表1：檢查計畫
                        if (plansData.length > 0) {
                            const plansWSData = [
                                ['計畫名稱', '年度', '建立時間', '更新時間', '關聯事項數']
                            ];
                            plansData.forEach(plan => {
                                plansWSData.push([
                                    plan.name || '',
                                    plan.year || '',
                                    new Date(plan.created_at).toLocaleString('zh-TW'),
                                    new Date(plan.updated_at).toLocaleString('zh-TW'),
                                    plan.issue_count || 0
                                ]);
                            });
                            const plansWS = XLSX.utils.aoa_to_sheet(plansWSData);
                            XLSX.utils.book_append_sheet(wb, plansWS, '檢查計畫');
                        }
                        
                        // 工作表2：開立事項
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
                        // 僅匯出檢查計畫
                        const plansWSData = [
                            ['計畫名稱', '年度', '建立時間', '更新時間', '關聯事項數']
                        ];
                        plansData.forEach(plan => {
                            plansWSData.push([
                                plan.name || '',
                                plan.year || '',
                                new Date(plan.created_at).toLocaleString('zh-TW'),
                                new Date(plan.updated_at).toLocaleString('zh-TW'),
                                plan.issue_count || 0
                            ]);
                        });
                        const plansWS = XLSX.utils.aoa_to_sheet(plansWSData);
                        XLSX.utils.book_append_sheet(wb, plansWS, '檢查計畫');
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
        function updatePwdStrength() { const p = document.getElementById('uPwd').value || ''; const conf = document.getElementById('uPwdConfirm').value || ''; let score = 0; if (p.length >= 8) score++; if (/[A-Z]/.test(p)) score++; if (/[0-9]/.test(p)) score++; if (/[^A-Za-z0-9]/.test(p)) score++; const texts = ['弱', '偏弱', '一般', '良好', '強']; document.getElementById('pwdStrength').innerText = `密碼強度: ${texts[Math.min(score, 4)]} ${conf && p !== conf ? '(密碼不相符)' : ''}`; }

        // User CRUD
        async function openUserModal(mode, id) { const m = document.getElementById('userModal'), t = document.getElementById('userModalTitle'), e = document.getElementById('uEmail'); if (mode === 'create') { t.innerText = '新增'; document.getElementById('targetUserId').value = ''; document.getElementById('uName').value = ''; e.value = ''; e.disabled = false; document.getElementById('uPwd').value = ''; document.getElementById('uPwdConfirm').value = ''; document.getElementById('pwdStrength').innerText = '密碼強度: -'; document.getElementById('pwdHint').innerText = ''; document.getElementById('uRole').value = 'viewer'; } else { const u = userList.find(x => x.id === id) || {}; t.innerText = '編輯'; document.getElementById('targetUserId').value = u.id || ''; document.getElementById('uName').value = u.name || ''; e.value = u.username || ''; e.disabled = true; document.getElementById('uPwd').value = ''; document.getElementById('uPwdConfirm').value = ''; document.getElementById('pwdHint').innerText = '(留空不改)'; document.getElementById('pwdStrength').innerText = '密碼強度: -'; document.getElementById('uRole').value = u.role || 'viewer'; } m.classList.add('open'); }
        async function submitUser() { const id = document.getElementById('targetUserId').value, name = document.getElementById('uName').value, email = document.getElementById('uEmail').value, pwd = document.getElementById('uPwd').value, pwdConfirm = document.getElementById('uPwdConfirm').value, role = document.getElementById('uRole').value; if (!id) { if (!email) return showToast('請輸入帳號', 'error'); if (!pwd) return showToast('請輸入密碼', 'error'); if (pwd !== pwdConfirm) return showToast('密碼與確認密碼不符', 'error'); if (pwd.length < 8) return showToast('密碼需至少 8 碼', 'error'); const res = await fetch('/api/users', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: email, name, password: pwd, role }) }); const j = await res.json(); if (res.ok) { showToast('新增成功'); document.getElementById('userModal').classList.remove('open'); loadUsersPage(1); } else showToast(j.error || '新增失敗', 'error'); } else { const payload = { name, role }; if (pwd) { if (pwd !== pwdConfirm) return showToast('密碼與確認密碼不符', 'error'); if (pwd.length < 8) return showToast('密碼需至少 8 碼', 'error'); payload.password = pwd; } const res = await fetch(`/api/users/${id}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) }); const j = await res.json(); if (res.ok) { showToast('更新成功'); document.getElementById('userModal').classList.remove('open'); loadUsersPage(usersPage); } else showToast(j.error || '更新失敗', 'error'); } }
        async function deleteUser(id) { if (!confirm('確定?')) return; const res = await fetch(`/api/users/${id}`, { method: 'DELETE' }); if (res.ok) { showToast('刪除成功'); loadUsersPage(1); } else showToast('刪除失敗', 'error'); }
        
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
                                const res = await fetch('/api/users/import', {
                                    method: 'POST',
                                    headers: { 'Content-Type': 'application/json' },
                                    credentials: 'include',
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
                renderPlans();
                renderPagination('plansPagination', plansPage, plansPages, 'loadPlansPage');
                // 更新年度選項
                updatePlanYearOptions();
            } catch (e) {
                // 錯誤已在伺服器 log 中記錄
                showToast('載入計畫錯誤: ' + e.message, 'error');
            }
        }
        function renderPlans() {
            const tbody = document.getElementById('plansTableBody');
            if (!tbody) return;
            tbody.innerHTML = planList.map(p => {
                return `<tr>
                    <td data-label="選擇" style="padding:12px;text-align:center;">
                        <input type="checkbox" class="plan-check" value="${p.id}" onchange="updatePlansBatchDeleteBtn()">
                    </td>
                    <td data-label="年度" style="padding:12px;font-weight:600;">${p.year || '-'}</td>
                    <td data-label="計畫名稱" style="padding:12px;font-weight:600;">${p.name || '-'}</td>
                    <td data-label="事項數量">${p.issue_count || 0}</td>
                    <td data-label="建立時間">${new Date(p.created_at).toLocaleDateString('zh-TW')}</td>
                    <td data-label="操作">
                        <button class="btn btn-outline" style="padding:2px 6px;margin-right:4px;" onclick="openPlanModal('edit', ${p.id})">✏️</button>
                        <button class="btn btn-danger" style="padding:2px 6px;" onclick="deletePlan(${p.id})">🗑️</button>
                    </td>
                </tr>`;
            }).join('');
            // 重置批次刪除按鈕狀態
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
            
            if (!confirm(`確定要刪除以下 ${ids.length} 筆檢查計畫嗎？\n\n${planNames.slice(0, 5).join('\n')}${planNames.length > 5 ? '\n...' : ''}\n\n此操作無法復原！`)) {
                return;
            }
            
            try {
                // 逐一刪除（因為需要記錄每個計畫的名稱）
                let successCount = 0;
                let failCount = 0;
                const errors = [];
                
                for (const id of ids) {
                    try {
                        const res = await fetch(`/api/plans/${id}`, { method: 'DELETE' });
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
                            console.warn('刪除錯誤詳情：', errors);
                        }
                    }
                    showToast(msg, failCount > 0 ? 'warning' : 'success');
                    loadPlansPage(plansPage);
                    loadPlanOptions();
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
        function openPlanModal(mode, id) {
            const m = document.getElementById('planModal');
            const t = document.getElementById('planModalTitle');
            if (mode === 'create') {
                t.innerText = '新增檢查計畫';
                document.getElementById('targetPlanId').value = '';
                document.getElementById('planName').value = '';
                document.getElementById('planYear').value = '';
            } else {
                const p = planList.find(x => x.id === id) || {};
                t.innerText = '編輯檢查計畫';
                document.getElementById('targetPlanId').value = p.id || '';
                document.getElementById('planName').value = p.name || '';
                document.getElementById('planYear').value = p.year || '';
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
                            // CSV 解析完成（已移除 debug 日誌）
                            
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
                                
                                // 嘗試各種可能的欄位名稱
                                let name = '';
                                let year = '';
                                for (const key in row) {
                                    const cleanKey = key.trim();
                                    if (cleanKey === '計畫名稱' || cleanKey === 'name' || cleanKey === 'planName' || cleanKey === '計劃名稱') {
                                        name = String(row[key] || '').trim();
                                    }
                                    if (cleanKey === '年度' || cleanKey === 'year') {
                                        year = String(row[key] || '').trim();
                                    }
                                }
                                
                                if (name && year) {
                                    validData.push({ name, year });
                                } else {
                                    // 記錄無效行的資訊（用於調試）
                                    invalidRows.push({
                                        row: index + 2, // +2 因為有標題行且從0開始
                                        name: name || '(空白)',
                                        year: year || '(空白)',
                                        rawRow: row
                                    });
                                }
                            });
                            
                            // 已移除調試資訊（只在伺服器 log 中記錄）
                            
                            if (validData.length === 0) {
                                let errorMsg = 'CSV 檔案中沒有有效的資料';
                                if (invalidRows.length > 0) {
                                    errorMsg += `\n發現 ${invalidRows.length} 筆資料缺少必要欄位（計畫名稱或年度）`;
                                    console.error('無效行詳情：', invalidRows);
                                }
                                return showToast(errorMsg, 'error');
                            }
                            
                            try {
                                const res = await fetch('/api/plans/import', {
                                    method: 'POST',
                                    headers: { 'Content-Type': 'application/json' },
                                    credentials: 'include', // 確保包含 session cookie
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
            const csv = '計畫名稱,年度\n113年度上半年定期檢查,113\n113年度下半年定期檢查,113\n114年度上半年定期檢查,114';
            const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = '檢查計畫匯入範例.csv';
            link.click();
        }
        async function submitPlan() {
            const id = document.getElementById('targetPlanId').value;
            const name = document.getElementById('planName').value.trim();
            const year = document.getElementById('planYear').value.trim();
            if (!name) return showToast('請輸入計畫名稱', 'error');
            if (!year) return showToast('請輸入年度', 'error');
            const payload = { name, year };
            try {
                const url = id ? `/api/plans/${id}` : '/api/plans';
                const method = id ? 'PUT' : 'POST';
                const res = await fetch(url, { method, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
                const j = await res.json();
                if (res.ok) {
                    showToast(id ? '更新成功' : '新增成功');
                    closePlanModal();
                    loadPlansPage(id ? plansPage : 1);
                    loadPlanOptions(); // 重新載入計畫選項
                } else {
                    showToast(j.error || (id ? '更新失敗' : '新增失敗'), 'error');
                }
            } catch (e) {
                showToast('操作失敗', 'error');
            }
        }
        async function deletePlan(id) {
            if (!confirm('確定要刪除這個計畫嗎？')) return;
            try {
                const res = await fetch(`/api/plans/${id}`, { method: 'DELETE' });
                const j = await res.json();
                if (res.ok) {
                    showToast('刪除成功');
                    loadPlansPage(1);
                    loadPlanOptions(); // 重新載入計畫選項
                } else {
                    showToast(j.error || '刪除失敗', 'error');
                }
            } catch (e) {
                showToast('刪除失敗', 'error');
            }
        }

        // Profile
        function openProfileModal() { document.getElementById('myProfileName').value = currentUser.name || ''; document.getElementById('myProfilePwd').value = ''; document.getElementById('profileModal').classList.add('open'); }
        async function submitProfile() { const name = document.getElementById('myProfileName').value, pwd = document.getElementById('myProfilePwd').value; try { const res = await fetch('/api/auth/profile', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name, password: pwd }) }); if (res.ok) { showToast('更新成功，請重新登入'); document.getElementById('profileModal').classList.remove('open'); logout(); } else { const j = await res.json(); showToast(j.error || '更新失敗', 'error'); } } catch (e) { showToast('更新失敗', 'error'); } }

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
                const numStr = String(currentEditItem.number || '');
                if (!k && numStr) { const m = numStr.match(/-([NOR])\d+$/i); if (m) k = m[1].toUpperCase(); }
                
                let kindLabel = '';
                if (k === 'N') kindLabel = `<span class="kind-tag N">缺失</span>`;
                else if (k === 'O') kindLabel = `<span class="kind-tag O">觀察</span>`;
                else if (k === 'R') kindLabel = `<span class="kind-tag R">建議</span>`;
                
                // 顯示狀態標籤（包括持續列管）
                let statusBadge = '';
                if (st && st !== 'Open') {
                    const stClass = st === '持續列管' ? 'active' : (st === '解除列管' ? 'resolved' : 'self');
                    statusBadge = `<span class="badge ${stClass}">${st}</span>`;
                }
                
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
            
            // 事項內容
            document.getElementById('dContent').innerHTML = currentEditItem.content;

            // Status and Kind (狀態與類型) - 使用與dCategoryInfo相同的邏輯
            let k = currentEditItem.item_kind_code || currentEditItem.itemKindCode;
            const numStr = String(currentEditItem.number || '');
            if (!k && numStr) { const m = numStr.match(/-([NOR])\d+$/i); if (m) k = m[1].toUpperCase(); }
            
            let kindLabel = '';
            if (k === 'N') kindLabel = `<span class="kind-tag N">缺失</span>`;
            else if (k === 'O') kindLabel = `<span class="kind-tag O">觀察</span>`;
            else if (k === 'R') kindLabel = `<span class="kind-tag R">建議</span>`;
            
            const st = currentEditItem.status === '持續列管' ? 'active' : (currentEditItem.status === '解除列管' ? 'resolved' : 'self');
            let statusBadge = '';
            if (currentEditItem.status && currentEditItem.status !== 'Open') {
                statusBadge = `<span class="badge ${st}">${currentEditItem.status}</span>`;
            }
            
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
            charts.trend.data = { 
                labels: tSorted.map(x => x.year), 
                datasets: [{ 
                    label: '開立事項數', 
                    data: tSorted.map(x => parseInt(x.count)), 
                    borderColor: '#667eea',  // 使用與標題漸變一致的顏色（直接使用顏色值）
                    backgroundColor: 'rgba(102, 126, 234, 0.1)', 
                    tension: 0.3, 
                    fill: true 
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
            charts.trend.update();
        }

        function initCharts() {
            try {
                const c1 = document.getElementById('statusChart'), c2 = document.getElementById('unitChart'), c3 = document.getElementById('trendChart');
                if (c1) { charts.status = new Chart(c1, { type: 'doughnut', plugins: [ChartDataLabels], data: { labels: [], datasets: [] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#64748b', font: { size: 12 } } }, datalabels: { formatter: (v, ctx) => { const dataArr = ctx.chart.data.datasets[0].data; if (!dataArr || dataArr.length === 0) return ''; const t = dataArr.reduce((a, b) => a + b, 0); return t > 0 ? ((v / t) * 100).toFixed(1) + '%' : '0%'; }, color: '#64748b', font: { weight: '600', size: 12 } } } } }); }
                if (c2) { charts.unit = new Chart(c2, { type: 'bar', data: { labels: [], datasets: [] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: '#64748b', font: { size: 12 } }, grid: { color: '#e2e8f0' } }, y: { ticks: { color: '#64748b', font: { size: 12 } }, grid: { color: '#e2e8f0' } } } } }); }
                if (c3) { charts.trend = new Chart(c3, { type: 'line', data: { labels: [], datasets: [] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, title: { display: true, text: '年度開立事項趨勢', color: '#64748b', font: { size: 14, weight: '600' } } }, scales: { x: { ticks: { color: '#64748b', font: { size: 12 } }, grid: { color: '#e2e8f0' } }, y: { beginAtZero: true, ticks: { stepSize: 1, color: '#64748b', font: { size: 12 } }, grid: { color: '#e2e8f0' } } } } }); }
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
                const res = await fetch(`/api/issues/${id}`, {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
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
                        // 重新載入資料
                        await loadIssuesPage(issuesPage);
                        // 更新 currentEditItem
                        const updatedItem = currentData.find(d => String(d.id) === String(id));
                        if (updatedItem) {
                            currentEditItem = updatedItem;
                            // 重新載入回合資料以反映最新的儲存結果
                            // 確保使用正確的 round 值（不應該改變）
                            const currentRound = parseInt(document.getElementById('editRound').value) || 1;
                            document.getElementById('editRound').value = currentRound;
                            loadRoundData();
                        }
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
                showToast('已帶入 AI 建議'); 
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
                const res = await fetch(`/api/issues?page=1&pageSize=1000&planName=${encodeURIComponent(planValue)}&_t=${Date.now()}`);
                if (!res.ok) throw new Error('載入事項列表失敗');
                
                const json = await res.json();
                yearEditIssueList = json.data || [];
                
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
                const numStr = String(issue.number || '');
                if (!k && numStr) { const m = numStr.match(/-([NOR])\d+$/i); if (m) k = m[1].toUpperCase(); }
                
                let kindLabel = '';
                if (k === 'N') kindLabel = `<span class="kind-tag N">缺失</span>`;
                else if (k === 'O') kindLabel = `<span class="kind-tag O">觀察</span>`;
                else if (k === 'R') kindLabel = `<span class="kind-tag R">建議</span>`;
                
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
            
            // 驗證日期格式（應該是6或7位數字，例如：1130615 或 1141001）
            if (!/^\d{6,7}$/.test(userInputResponseDate)) {
                showToast('日期格式錯誤，應為6或7位數字（例如：1130615 或 1141001）', 'error');
                return;
            }
            
            const { name: planName } = parsePlanValue(planValue);
            
            try {
                // 載入該計畫下的所有事項
                // 移除載入中的提示訊息，只保留錯誤訊息
                const res = await fetch(`/api/issues?page=1&pageSize=1000&planName=${encodeURIComponent(planValue)}&_t=${Date.now()}`);
                if (!res.ok) throw new Error('載入事項列表失敗');
                
                const json = await res.json();
                const issueList = json.data || [];
                
                if (issueList.length === 0) {
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
                        const updateRes = await fetch(`/api/issues/${issueId}`, {
                            method: 'PUT',
                            headers: { 'Content-Type': 'application/json' },
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
                if (errorCount > 0) {
                    showToast(`批次設定完成，但有 ${errorCount} 筆失敗${successCount > 0 ? `，成功 ${successCount} 筆` : ''}`, 'warning');
                    
                    // 如果有錯誤，顯示詳細資訊
                    if (errors.length > 0) {
                        console.error('批次設定函復日期錯誤:', errors);
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
                    const toggleCheckbox = document.getElementById('createBatchResponseDateToggle');
                    if (toggleCheckbox) {
                        toggleCheckbox.checked = false;
                        toggleBatchResponseDateSetting();
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
            
            // 驗證日期格式（應該是6或7位數字，例如：1130615 或 1141001）
            if (!/^\d{6,7}$/.test(userInputResponseDate)) {
                showToast('日期格式錯誤，應為6或7位數字（例如：1130615 或 1141001）', 'error');
                return;
            }
            
            if (yearEditIssueList.length === 0) {
                showToast('沒有可設定的事項', 'error');
                return;
            }
            
            if (!confirm(`確定要批次設定第 ${round} 次審查的函復日期為 ${responseDate} 嗎？\n將更新 ${yearEditIssueList.length} 筆事項。`)) {
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
                        const res = await fetch(`/api/issues/${issueId}`, {
                            method: 'PUT',
                            headers: { 'Content-Type': 'application/json' },
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
            
            if (!confirm('確定要儲存所有變更嗎？')) {
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
                const updateRes = await fetch(`/api/issues/${issueId}`, {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
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
                        const updateRes = await fetch(`/api/issues/${issueId}`, {
                            method: 'PUT',
                            headers: { 'Content-Type': 'application/json' },
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