        // 全域狀態
        let rawData = [], currentData = [], currentUser = null, charts = {}, currentEditItem = null, userList = [], sortState = { field: null, dir: 'asc' }, stagedImportData = [];
        let autoLogoutTimer;
        let currentLogs = { login: [], action: [] };
        let cachedGlobalStats = null;
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
                    return;
                }
                
                // 載入計畫選項（已移除 debug 日誌）
                
                // 更新所有計畫選擇下拉選單
                const selectIds = ['filterPlan', 'importPlanName', 'batchPlanName', 'manualPlanName'];
                selectIds.forEach(selectId => {
                    const select = document.getElementById(selectId);
                    if (select) {
                        const currentValue = select.value;
                        // 保留第一個選項（通常是「全部計畫」或「請選擇計畫」）
                        const firstOption = select.options[0] ? select.options[0].outerHTML : '';
                            // 建立完整的選項列表（避免重複）
                            const existingValues = new Set();
                            if (firstOption) {
                                // 從第一個選項中提取值
                                const tempDiv = document.createElement('div');
                                tempDiv.innerHTML = firstOption;
                                const firstOpt = tempDiv.querySelector('option');
                                if (firstOpt && firstOpt.value) {
                                    existingValues.add(firstOpt.value);
                                }
                            }
                            
                            // 處理新的資料格式（可能包含 {name, year, display, value} 或舊的字串格式）
                            const allOptions = json.data.map(p => {
                                let planValue, planDisplay;
                                
                                // 檢查是否為新格式（物件）
                                if (typeof p === 'object' && p !== null) {
                                    planValue = p.value || `${p.name}|||${p.year || ''}`;
                                    planDisplay = p.display || `${p.name}${p.year ? ` (${p.year})` : ''}`;
                                } else {
                                    // 舊格式（字串），向後兼容
                                    planValue = p;
                                    planDisplay = p;
                                }
                                
                                if (!existingValues.has(planValue)) {
                                    existingValues.add(planValue);
                                    return `<option value="${planValue}">${planDisplay}</option>`;
                                }
                                return '';
                            }).filter(opt => opt).join('');
                            
                            // 完全重建選項列表，確保所有計畫都顯示
                            select.innerHTML = firstOption + allOptions;
                        
                        // 恢復之前選擇的值
                        if (currentValue && Array.from(select.options).some(opt => opt.value === currentValue)) {
                            select.value = currentValue;
                        }
                    }
                });
            } catch (e) {
                console.error("Load plans failed", e);
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
            try {
                await checkAuth();
                if (currentUser) {
                    // 確保 body 可見
                    document.body.style.display = 'flex';
                    
                    // 嘗試恢復上次的視圖
                    const savedView = sessionStorage.getItem('currentView');
                    let targetView = savedView || 'searchView';
                    
                    // 確保視圖存在
                    const viewElement = document.getElementById(targetView);
                    if (!viewElement) {
                        targetView = 'searchView';
                    }
                    
                    // 切換到目標視圖
                    await switchView(targetView);
                    
                    initListeners();
                    initEditForm();
                    initCharts();
                    loadPlanOptions();
                    initImportRoundOptions();
                    
                    // 如果目標視圖是 searchView，載入資料
                    if (targetView === 'searchView') {
                        await loadIssuesPage(1);
                    }
                    // Preload users if needed
                    if(currentUser.role === 'admin' && targetView === 'usersView') {
                        loadUsersPage(1);
                    }
                }
            } catch (error) {
                console.error('初始化錯誤:', error);
                // 即使出錯也嘗試顯示頁面
                document.body.style.display = 'flex';
            }
        });

        async function checkAuth() {
            try {
                const res = await fetch('/api/auth/me?t=' + Date.now(), { headers: { 'Cache-Control': 'no-cache' } });
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
                } else window.location.href = '/login.html';
            } catch (e) { window.location.href = '/login.html'; }
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
            // 從計畫選項值中提取計畫名稱（用於查詢）
            const planName = planValue ? parsePlanValue(planValue).name : '';

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
                if (document.getElementById('filterKeyword')) document.getElementById('filterKeyword').value = state.keyword || '';
                if (document.getElementById('filterYear')) document.getElementById('filterYear').value = state.year || '';
                if (document.getElementById('filterPlan')) document.getElementById('filterPlan').value = state.plan || '';
                if (document.getElementById('filterUnit')) document.getElementById('filterUnit').value = state.unit || '';
                if (document.getElementById('filterStatus')) document.getElementById('filterStatus').value = state.status || '';
                if (document.getElementById('filterKind')) document.getElementById('filterKind').value = state.kind || '';
                if (document.getElementById('filterDivision')) document.getElementById('filterDivision').value = state.division || '';
                if (document.getElementById('filterInspection')) document.getElementById('filterInspection').value = state.inspection || '';
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
                    showToast(`成功刪除 ${ids.length} 筆資料`);
                    loadIssuesPage(issuesPage);
                } else {
                    const json = await res.json();
                    showToast(json.error || '刪除失敗', 'error');
                }
            } catch (e) {
                showToast('刪除時發生錯誤: ' + e.message, 'error');
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
                    const editBtn = canEdit ? `<button class="badge" style="background:#fff;border:1px solid #ddd;cursor:pointer;margin-top:4px;" onclick="event.stopPropagation();openDetail('${item.id}',false)">✏️ 編輯</button>` : '';
                    const delBtn = canManage ? `<button class="badge" style="background:#fee2e2;color:#ef4444;border:1px solid #fca5a5;cursor:pointer;margin-top:4px;margin-left:2px;" onclick="event.stopPropagation();deleteIssue('${item.id}')">🗑️</button>` : '';
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

                    html += `<tr onclick="openDetail('${item.id}',false)"> ${checkbox} <td data-label="年度">${item.year}</td><td data-label="編號" style="font-weight:600;color:var(--primary);">${item.number}</td><td data-label="機構">${item.unit}</td><td data-label="狀態與類型">${statusHtml}</td><td data-label="事項內容"><div class="text-content">${snippet}${(stripHtml(item.content || '').length > 180 ? ` <a href='javascript:void(0)' onclick="event.stopPropagation();showPreview(${JSON.stringify(fullHtml)}, '編號 ${item.number} 內容')">...更多</a>` : '')}</div></td><td data-label="最新辦理/審查情形"><div class="text-content">${stripHtml(updateTxt)}</div></td><td data-label="管理功能"><div>${aiContent}${editBtn}${delBtn}</div></td></tr>`;
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
            // 從計畫選項值中提取計畫名稱
            const planName = isBackup ? '' : parsePlanValue(planValue).name;

            let cleanData = stagedImportData.map(({ _importStatus, ...item }) => {
                if (currentImportMode === 'word') {
                    if (!item.planName && planName) item.planName = planName;
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
            document.getElementById('tab-data-import').classList.toggle('hidden', tab !== 'import'); 
            document.getElementById('tab-data-manual').classList.toggle('hidden', tab !== 'manual'); 
            document.getElementById('tab-data-export').classList.toggle('hidden', tab !== 'export'); 
            document.getElementById('tab-data-batch').classList.toggle('hidden', tab !== 'batch'); 
            const plansTab = document.getElementById('tab-data-plans');
            if (plansTab) plansTab.classList.toggle('hidden', tab !== 'plans');
            if (tab === 'batch' && document.querySelectorAll('#batchGridBody tr').length === 0) initBatchGrid();
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
                            exportIssuesOptions.style.display = 'flex';
                        }
                    });
                });
                
                // 初始化顯示狀態
                const checked = document.querySelector('input[name="exportDataType"]:checked');
                if (checked && checked.value === 'plans') {
                    exportIssuesOptions.style.display = 'none';
                } else {
                    exportIssuesOptions.style.display = 'flex';
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

        function autoFillFromNumber() {
            const val = document.getElementById('manualNumber').value;
            const info = parseItemNumber(val);
            if (info) {
                if (info.yearRoc) {
                    const yearDisplay = document.getElementById('manualYearDisplay');
                    if (yearDisplay) yearDisplay.value = info.yearRoc;
                }
                if (info.orgCode) {
                    const name = ORG_MAP[info.orgCode] || info.orgCode;
                    if (name && name !== '?') document.getElementById('manualUnit').value = name;
                }
                if (info.divCode) {
                    const divName = DIVISION_MAP[info.divCode];
                    if (divName) document.getElementById('manualDivision').value = divName;
                }
                if (info.inspectCode) {
                    const inspectName = INSPECTION_MAP[info.inspectCode];
                    if (inspectName) document.getElementById('manualInspection').value = inspectName;
                }
                if (info.kindCode) {
                    document.getElementById('manualKind').value = info.kindCode;
                }
            }
        }

        async function submitManualIssue() {
            const number = document.getElementById('manualNumber').value.trim();
            const yearDisplay = document.getElementById('manualYearDisplay');
            const year = yearDisplay ? yearDisplay.value.trim() : '';
            const unit = document.getElementById('manualUnit').value.trim();
            const division = document.getElementById('manualDivision').value;
            const inspection = document.getElementById('manualInspection').value;
            const kind = document.getElementById('manualKind').value;

            const planValue = document.getElementById('manualPlanName').value.trim();
            const issueDate = document.getElementById('manualIssueDate').value.trim();
            const continuousMode = document.getElementById('manualContinuousMode').checked;

            const status = document.getElementById('manualStatus').value;
            const content = document.getElementById('manualContent').value.trim();
            if (!number || !year || !unit || !content) return showToast('請填寫所有必填欄位', 'error');
            // 從計畫選項值中提取計畫名稱
            const planName = parsePlanValue(planValue).name;

            const payload = {
                data: [{
                    number, year, unit, content, status,
                    itemKindCode: kind,
                    divisionName: division,
                    inspectionCategoryName: inspection,
                    planName: planName,
                    issueDate: issueDate,
                    scheme: 'MANUAL'
                }],
                round: 1, reviewDate: '', replyDate: ''
            };

            try {
                const res = await fetch('/api/issues/import', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
                if (res.ok) {
                    showToast('新增成功');

                    if (continuousMode) {
                        document.getElementById('manualNumber').value = '';
                        document.getElementById('manualKind').value = '';
                        document.getElementById('manualContent').value = '';
                        document.getElementById('manualNumber').focus();
                    } else {
                        document.getElementById('manualNumber').value = '';
                        const yearDisplay = document.getElementById('manualYearDisplay');
                        if (yearDisplay) yearDisplay.value = '';
                        document.getElementById('manualUnit').value = '';
                        document.getElementById('manualDivision').value = '';
                        document.getElementById('manualInspection').value = '';
                        document.getElementById('manualKind').value = '';
                        document.getElementById('manualContent').value = '';
                        document.getElementById('manualPlanName').value = '';
                        document.getElementById('manualIssueDate').value = '';
                    }

                    loadIssuesPage(1);
                    loadPlanOptions();
                } else { showToast('新增失敗', 'error'); }
            } catch (e) { showToast('Error: ' + e.message, 'error'); }
        }

        // 保留舊函數名稱以向後兼容
        async function exportAllIssues() {
            return exportAllData();
        }

        async function exportAllData() {
            try {
                const exportDataType = document.querySelector('input[name="exportDataType"]:checked')?.value || 'issues';
                const exportScope = document.querySelector('input[name="exportScope"]:checked')?.value || 'latest';
                const exportFormat = document.querySelector('input[name="exportFormat"]:checked').value;
                showToast('準備匯出中，請稍候...', 'info');
                
                const clean = (t) => `"${String(t || '').replace(/"/g, '""').replace(/<[^>]*>/g, '').trim()}"`;
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

                // CSV 格式匯出
                let csvContent = '\uFEFF';
                
                if (exportDataType === 'both') {
                    // 合併匯出：先匯出檢查計畫，再匯出開立事項
                    // 檢查計畫
                    csvContent += "=== 檢查計畫 ===\n";
                    csvContent += "計畫名稱,年度,建立時間,更新時間,關聯事項數\n";
                    plansData.forEach(plan => {
                        csvContent += `${clean(plan.name)},${clean(plan.year)},${clean(new Date(plan.created_at).toLocaleString('zh-TW'))},${clean(new Date(plan.updated_at).toLocaleString('zh-TW'))},${clean(plan.issue_count || 0)}\n`;
                    });
                    
                    csvContent += "\n=== 開立事項 ===\n";
                }
                
                // 開立事項
                if (exportDataType === 'issues' || exportDataType === 'both') {
                    const baseHeader = "編號,年度,機構,分組,檢查種類,類型,狀態,事項內容";
                    
                    if (exportScope === 'latest') {
                        csvContent += baseHeader + ",最新辦理情形,最新審查意見\n";
                        issuesData.forEach(item => {
                            let latestH = '', latestR = '';
                            for (let i = 200; i >= 1; i--) { 
                                const suffix = i === 1 ? '' : i;
                                if (!latestH && (item[`handling${suffix}`])) latestH = item[`handling${suffix}`]; 
                                if (!latestR && (item[`review${suffix}`])) latestR = item[`review${suffix}`]; 
                            }
                            csvContent += `${clean(item.number)},${clean(item.year)},${clean(item.unit)},${clean(item.divisionName)},${clean(item.inspectionCategoryName)},${clean(item.category)},${clean(item.status)},${clean(item.content)},${clean(latestH)},${clean(latestR)}\n`;
                        });
                    } else {
                        csvContent += baseHeader + ",完整辦理情形歷程,完整審查意見歷程\n";
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
                            csvContent += `${clean(item.number)},${clean(item.year)},${clean(item.unit)},${clean(item.divisionName)},${clean(item.inspectionCategoryName)},${clean(item.category)},${clean(item.status)},${clean(item.content)},${clean(joinedH)},${clean(joinedR)}\n`;
                        });
                    }
                }
                
                // 僅匯出檢查計畫
                if (exportDataType === 'plans') {
                    csvContent += "計畫名稱,年度,建立時間,更新時間,關聯事項數\n";
                    plansData.forEach(plan => {
                        csvContent += `${clean(plan.name)},${clean(plan.year)},${clean(new Date(plan.created_at).toLocaleString('zh-TW'))},${clean(new Date(plan.updated_at).toLocaleString('zh-TW'))},${clean(plan.issue_count || 0)}\n`;
                    });
                }
                
                const link = document.createElement("a");
                link.setAttribute("href", URL.createObjectURL(new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })));
                let fileName = '';
                if (exportDataType === 'issues') {
                    const typeLabel = exportScope === 'latest' ? 'Latest' : 'FullHistory';
                    fileName = `SMS_Issues_${typeLabel}_${new Date().toISOString().slice(0, 10)}.csv`;
                } else if (exportDataType === 'plans') {
                    fileName = `SMS_Plans_${new Date().toISOString().slice(0, 10)}.csv`;
                } else {
                    fileName = `SMS_AllData_${new Date().toISOString().slice(0, 10)}.csv`;
                }
                link.setAttribute("download", fileName);
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                showToast('CSV 匯出完成', 'success');
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
                
                // 詢問匯出格式
                const format = confirm('選擇匯出格式：\n確定 = CSV\n取消 = JSON') ? 'csv' : 'json';
                
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
                    // CSV 格式（不包含密碼）
                    let csvContent = '\uFEFF';
                    csvContent += "姓名,帳號,權限,建立時間\n";
                    users.forEach(user => {
                        const clean = (t) => `"${String(t || '').replace(/"/g, '""').trim()}"`;
                        csvContent += `${clean(user.name)},${clean(user.username)},${clean(getRoleName(user.role))},${clean(new Date(user.created_at).toLocaleString('zh-TW'))}\n`;
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
            const csv = '姓名,帳號,權限,密碼\n張三,zhang@example.com,editor,password123\n李四,li@example.com,manager,password123';
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
                                
                                // 驗證權限值
                                const validRoles = ['admin', 'manager', 'editor', 'viewer'];
                                if (!validRoles.includes(role.toLowerCase())) {
                                    invalidRows.push({
                                        row: index + 2,
                                        error: `無效的權限值：${role}（應為：admin, manager, editor, viewer）`
                                    });
                                    return;
                                }
                                
                                validData.push({ name, username, role: role.toLowerCase(), password });
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
                                    closeUserImportModal();
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
            document.getElementById('drawerTitle').innerText = edit ? "編輯事項" : "詳細資料"; 
            if (edit) { 
                if (!currentEditItem) return;
                // 清除所有編輯欄位，避免前一個事項的資料殘留
                document.getElementById('editId').value = currentEditItem.id; 
                document.getElementById('editHeaderNumber').innerText = currentEditItem.number; 
                document.getElementById('editHeaderYear').innerText = currentEditItem.year + '年'; 
                document.getElementById('editHeaderUnit').innerText = currentEditItem.unit; 
                const st = (currentEditItem.status === 'Open' || !currentEditItem.status) ? '持續列管' : currentEditItem.status; 
                document.getElementById('editStatus').value = st; 
                
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
                document.getElementById('editReplyDate').value = '';
                document.getElementById('editResponseDate').value = '';
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
            document.getElementById('dNumber').textContent = currentEditItem.number; document.getElementById('dUnit').textContent = currentEditItem.unit; document.getElementById('dContent').innerHTML = currentEditItem.content;

            // Plan Info
            document.getElementById('dPlanName').textContent = currentEditItem.plan_name || currentEditItem.planName || '(未設定)';
            document.getElementById('dIssueDate').textContent = currentEditItem.issue_date || currentEditItem.issueDate || '(未設定)';

            // Category Info
            const divName = currentEditItem.divisionName || currentEditItem.division_name || '-';
            const insName = currentEditItem.inspectionCategoryName || currentEditItem.inspection_category_name || '-';
            const kindName = currentEditItem.category || '-';
            document.getElementById('dCategoryInfo').textContent = `${divName} / ${insName} / ${kindName}`;

            // Status
            const st = currentEditItem.status === '持續列管' ? 'active' : (currentEditItem.status === '解除列管' ? 'resolved' : 'self');
            document.getElementById('dStatus').innerHTML = currentEditItem.status && currentEditItem.status !== 'Open' ? `<span class="badge ${st}">${currentEditItem.status}</span>` : '(未設定)';

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
        function logout() { fetch('/api/auth/logout', { method: 'POST' }).then(() => window.location.reload()); }
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
            const handling = currentEditItem['handling' + suffix] || '';
            const review = currentEditItem['review' + suffix] || '';
            const replyDate = currentEditItem['reply_date_r' + round] || '';
            const responseDate = currentEditItem['response_date_r' + round] || '';
            
            // 儲存到隱藏的輸入框（用於儲存時提交）
            document.getElementById('editHandling').value = handling;
            document.getElementById('editReview').value = review;
            document.getElementById('editReplyDate').value = replyDate;
            document.getElementById('editResponseDate').value = responseDate;
            
            // 顯示第N次機構辦理情形（只讀，作為參考）
            // 撰寫第N次審查時，右側顯示第N次機構辦理情形
            // 因為第N次機構辦理情形後，會進行第N次審查
            const displayHandlingRound = round;
            const displayHandlingSuffix = displayHandlingRound === 1 ? '' : displayHandlingRound;
            const displayHandling = currentEditItem['handling' + displayHandlingSuffix] || '';
            
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
            
            // 找出所有同時有審查意見和辦理情形的輪次（完整內容）
            const rounds = [];
            for (let i = 200; i >= 1; i--) {
                const suffix = i === 1 ? '' : i;
                const hasHandling = currentEditItem['handling' + suffix] && currentEditItem['handling' + suffix].trim();
                const hasReview = currentEditItem['review' + suffix] && currentEditItem['review' + suffix].trim();
                // 只包含同時有兩個內容的輪次
                if (hasHandling && hasReview) {
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
                // 顯示最新進度 - 找出最高輪次，且該輪次必須同時有審查意見和辦理情形
                let maxRound = 0;
                
                // 找出最高的完整輪次（同時有審查意見和辦理情形）
                for (let k = 200; k >= 1; k--) {
                    const suffix = k === 1 ? '' : k;
                    const hasHandling = currentEditItem['handling' + suffix] && currentEditItem['handling' + suffix].trim();
                    const hasReview = currentEditItem['review' + suffix] && currentEditItem['review' + suffix].trim();
                    // 只選擇同時有兩個內容的輪次
                    if (hasHandling && hasReview) {
                        maxRound = k;
                        break;
                    }
                }
                
                if (maxRound > 0) {
                    const suffix = maxRound === 1 ? '' : maxRound;
                    const handling = currentEditItem['handling' + suffix] || '';
                    const review = currentEditItem['review' + suffix] || '';
                    
                    // 顯示審查意見
                    if (review && review.trim()) {
                        const viewReviewRoundNum = document.getElementById('viewReviewRoundNum');
                        const viewReviewText = document.getElementById('viewReviewText');
                        if (viewReviewRoundNum) viewReviewRoundNum.textContent = maxRound;
                        if (viewReviewText) viewReviewText.textContent = review;
                        if (viewReviewBox) viewReviewBox.style.display = 'block';
                    }
                    
                    // 顯示辦理情形
                    if (handling && handling.trim()) {
                        const viewHandlingRoundNum = document.getElementById('viewHandlingRoundNum');
                        const viewHandlingText = document.getElementById('viewHandlingText');
                        if (viewHandlingRoundNum) viewHandlingRoundNum.textContent = maxRound;
                        if (viewHandlingText) viewHandlingText.textContent = handling;
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
                    if (viewReviewRoundNum) viewReviewRoundNum.textContent = round;
                    if (viewReviewText) viewReviewText.textContent = review;
                    if (viewReviewBox) viewReviewBox.style.display = 'block';
                }
                
                if (handling && handling.trim()) {
                    const viewHandlingRoundNum = document.getElementById('viewHandlingRoundNum');
                    const viewHandlingText = document.getElementById('viewHandlingText');
                    if (viewHandlingRoundNum) viewHandlingRoundNum.textContent = round;
                    if (viewHandlingText) viewHandlingText.textContent = handling;
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
            const handling = document.getElementById('editHandling').value.trim();
            const review = document.getElementById('editReview').value.trim();
            const replyDate = document.getElementById('editReplyDate').value.trim();
            const responseDate = document.getElementById('editResponseDate').value.trim();
            
            if (!id) {
                showToast('找不到事項 ID', 'error');
                return;
            }
            
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
                        responseDate: responseDate || null
                    })
                });
                
                if (res.ok) {
                    const json = await res.json();
                    if (json.success) {
                        showToast('儲存成功！');
                        // 重新載入資料
                        await loadIssuesPage(issuesPage);
                        // 更新 currentEditItem
                        currentEditItem = currentData.find(d => String(d.id) === String(id));
                        if (currentEditItem) {
                            // 重新載入回合資料以反映最新的儲存結果
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
            const handlingTxt = document.getElementById('editHandling').value; 
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