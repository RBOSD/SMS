        // 全域狀態
        let rawData = [], currentData = [], currentUser = null, charts = {}, currentEditItem = null, userList = [], sortState = { field: null, dir: 'asc' }, stagedImportData = [];
        let autoLogoutTimer;
        let currentLogs = { login: [], action: [] };
        let cachedGlobalStats = null;
        let issuesPage = 1, issuesPageSize = 20, issuesTotal = 0, issuesPages = 1;
        let usersPage = 1, usersPageSize = 20, usersTotal = 0, usersPages = 1, usersSortField = 'id', usersSortDir = 'asc';
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
        function getRoleName(r) { const map = { 'admin': '系統管理員', 'manager': '資料管理者', 'editor': '審查人員', 'viewer': '檢視人員' }; return map[r] || r; }
        function extractNumberFromCell(cell) { if (!cell) return ""; var whole = normalizeCodeString(cell.innerText || cell.textContent || ""); return whole.trim(); }

        // [Updated] Map & Parser
        const ORG_MAP = { "T": "臺鐵", "H": "高鐵", "A": "林鐵", "S": "糖鐵", "TRC": "臺鐵", "HSR": "高鐵", "AFR": "林鐵", "TSC": "糖鐵" };
        const INSPECTION_MAP = { "1": "定期檢查", "2": "例行性檢查", "3": "特別檢查", "4": "臨時檢查" };
        // [Verified] Division Map includes all requested codes
        const DIVISION_MAP = { "A": "運務", "B": "工務", "C": "機務", "D": "電務", "E": "安全", "F": "審核", "G": "災防", "OP": "運轉", "CP": "土木", "EM": "機電" };
        const KIND_MAP = { "N": "缺失事項", "O": "觀察事項", "R": "建議事項" };
        const FILLED_MARKS = ["■", "☑", "☒", "✔", "✅", "●", "◉", "✓"]; var EMPTY_MARKS = ["□", "☐", "◻", "○", "◯", "◇", "△"];

        function parseItemNumber(numberStr) {
            var raw = normalizeCodeString(numberStr || "");
            if (!raw) return null;
            var cleanRaw = raw.replace(/[^a-zA-Z0-9\-]/g, "");
            var mLong = cleanRaw.match(/^(\d{3})-([A-Z]{3,4})-([0-9])-(\d+)-([A-Z]{2,4})-([NOR])(\d+)$/i);
            if (mLong) {
                return { raw: mLong[0], yearRoc: parseInt(mLong[1], 10), orgCode: mLong[2].toUpperCase(), inspectCode: mLong[3], divCode: mLong[5].toUpperCase(), kindCode: mLong[6].toUpperCase() };
            }
            var mShort = cleanRaw.match(/^(\d{2,3})([A-Z])([0-9])-([A-Z])(\d{2})-([NOR])(\d{2})$/i);
            if (mShort) {
                var yy = parseInt(mShort[1], 10);
                var rocYear = (yy < 1000) ? (yy + (yy < 100 ? 100 : 0)) : (yy - 1911);
                return { raw: mShort[0], yearRoc: rocYear, orgCode: mShort[2].toUpperCase(), inspectCode: mShort[3], divCode: mShort[4].toUpperCase(), kindCode: mShort[6].toUpperCase() };
            }
            var mLoose = cleanRaw.match(/(\d{2,3}).*([NOR])\d+/i);
            if (mLoose) {
                return { raw: mLoose[0], yearRoc: parseInt(mLoose[1], 10), orgCode: "?", inspectCode: "?", divCode: "?", kindCode: mLoose[2].toUpperCase() };
            }
            return { raw: cleanRaw, yearRoc: "", orgCode: "", inspectCode: "", divCode: "", kindCode: "" };
        }

        function normalizeMultiline(s) { s = String(s || ""); return s.replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/\u00A0/g, " ").replace(/\u3000/g, " ").replace(/\r/g, "").replace(/[ \t]+\n/g, "\n").replace(/\n[ \t]+/g, "\n").trim(); }

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
            const icon = type === 'success' ? '✅' : '⚠️';
            const title = type === 'success' ? '成功' : '錯誤';
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
                const res = await fetch('/api/options/plans?t=' + Date.now());
                const json = await res.json();
                const list = document.getElementById('planOptionsList');
                if (list && json.data) {
                    list.innerHTML = json.data.map(p => `<option value="${p}">`).join('');
                }
            } catch (e) {
                console.error("Load plans failed", e);
            }
        }

        function initImportRoundOptions() {
            const s = document.getElementById('importRoundSelect');
            if (!s) return;
            s.innerHTML = '';
            // 支援無限次審查，先建立前 100 次選項
            for (let i = 1; i <= 100; i++) {
                s.innerHTML += `<option value="${i}">第 ${i} 次審查</option>`;
            }
        }

        async function switchView(viewId) {
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
                    } else {
                        console.error('Failed to load view:', viewId);
                    }
                } catch (error) {
                    console.error('Error loading view:', viewId, error);
                }
            }

            if(viewId === 'searchView') {
                loadIssuesPage(1);
            } else if (viewId === 'usersView') {
                loadUsersPage(1);
            }
        }

        document.addEventListener('DOMContentLoaded', async () => {
            console.log("App init...");
            await checkAuth();
            if (currentUser) {
                document.body.style.display = 'flex';
                initListeners();
                initEditForm();
                initCharts();
                loadPlanOptions();
                initImportRoundOptions();
                await loadIssuesPage(1);
                // Preload users if needed
                if(currentUser.role === 'admin') {
                    loadUsersPage(1);
                }
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
                        document.getElementById('btn-importView').classList.remove('hidden');
                        if (data.role === 'admin') {
                            document.getElementById('btn-usersView').classList.remove('hidden');
                            document.getElementById('uploadCardBackup').classList.remove('hidden');
                            document.getElementById('exportJsonOption').style.display = 'flex';
                        }
                    }
                } else window.location.href = '/login.html';
            } catch (e) { window.location.href = '/login.html'; }
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
            issuesPage = page; document.getElementById('issuesPageSizeTop').value = issuesPageSize; document.getElementById('issuesPageSizeBottom').value = issuesPageSize;
            const q = document.getElementById('filterKeyword').value || '', year = document.getElementById('filterYear').value || '', unit = document.getElementById('filterUnit').value || '', status = document.getElementById('filterStatus').value || '', kind = document.getElementById('filterKind').value || '';
            const division = document.getElementById('filterDivision') ? document.getElementById('filterDivision').value : '';
            const inspection = document.getElementById('filterInspection') ? document.getElementById('filterInspection').value : '';
            const planName = document.getElementById('filterPlan') ? document.getElementById('filterPlan').value : '';

            // 預設以年度最新排序（降序）
            let sortField = 'year', sortDir = 'desc';
            if (sortState.field) { 
                if (sortState.field === 'number') sortField = 'title'; 
                else if (sortState.field === 'year') sortField = 'year'; 
                else if (sortState.field === 'unit') sortField = 'unit'; 
                else if (sortState.field === 'status') sortField = 'status'; 
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

        function applyFilters() { issuesPage = 1; loadIssuesPage(1); }
        function resetFilters() { document.querySelectorAll('.filter-input,.filter-select').forEach(e => e.value = ''); applyFilters(); }
        function sortData(field) { if (sortState.field === field) sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc'; else { sortState.field = field; sortState.dir = 'asc'; } loadIssuesPage(1); updateSortUI(); }
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
                    let updateTxt = '-'; const lr = getLatest(item, 'review'), lh = getLatest(item, 'handling'); if (lr) updateTxt = `[審] ${stripHtml(lr).slice(0, 80)}`; else if (lh) updateTxt = `[回] ${stripHtml(lh).slice(0, 80)}`;
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

        function onIssuesPageSizeChange(val) { issuesPageSize = parseInt(val, 10); loadIssuesPage(1); }
        async function loadUsersPage(page = 1) { usersPage = page; usersPageSize = parseInt(document.getElementById('usersPageSize').value, 10); const q = document.getElementById('userSearch').value || ''; const params = new URLSearchParams({ page: usersPage, pageSize: usersPageSize, q, sortField: usersSortField, sortDir: usersSortDir, _t: Date.now() }); try { const res = await fetch('/api/users?' + params.toString()); if (!res.ok) { showToast('載入使用者失敗', 'error'); return; } const j = await res.json(); userList = j.data || []; usersTotal = j.total || 0; usersPages = j.pages || 1; renderUsers(); renderPagination('usersPagination', usersPage, usersPages, 'loadUsersPage'); } catch (e) { console.error(e); showToast('載入使用者錯誤', 'error'); } }
        function renderUsers() { document.getElementById('usersTableBody').innerHTML = userList.map(u => `<tr><td data-label="姓名" style="padding:12px;">${u.name || '-'}</td><td data-label="帳號">${u.username}</td><td data-label="權限">${getRoleName(u.role)}</td><td data-label="註冊時間">${new Date(u.created_at).toLocaleDateString()}</td><td data-label="操作">${u.id !== currentUser.userId ? `<button class="btn btn-outline" style="padding:2px 6px;margin-right:4px;" onclick="openUserModal('edit', ${u.id})">✏️</button><button class="btn btn-danger" style="padding:2px 6px;" onclick="deleteUser(${u.id})">🗑️</button>` : '-'}</td></tr>`).join(''); }
        function usersSortBy(field) { if (usersSortField === field) usersSortDir = usersSortDir === 'asc' ? 'desc' : 'asc'; else { usersSortField = field; usersSortDir = 'asc'; } loadUsersPage(1); }

        async function loadLogsPage(page = 1) { logsPage = page; const q = document.getElementById('loginSearch').value || ''; const params = new URLSearchParams({ page: logsPage, pageSize: logsPageSize, q, _t: Date.now() }); document.getElementById('logsLoading').style.display = 'block'; try { const res = await fetch('/api/admin/logs?' + params.toString()); if (!res.ok) { showToast('載入登入紀錄失敗', 'error'); return; } const j = await res.json(); currentLogs.login = j.data || []; logsTotal = j.total || 0; logsPages = j.pages || 1; document.getElementById('logsTableBody').innerHTML = currentLogs.login.map(l => `<tr><td data-label="時間" style="padding:12px;">${new Date(l.login_time).toLocaleString('zh-TW')}</td><td data-label="帳號">${l.username}</td><td data-label="IP">${l.ip_address || '-'}</td></tr>`).join(''); renderPagination('logsPagination', logsPage, logsPages, 'loadLogsPage'); } catch (e) { console.error(e); showToast('載入登入紀錄錯誤', 'error'); } finally { document.getElementById('logsLoading').style.display = 'none'; } }
        async function loadActionsPage(page = 1) { actionsPage = page; const q = document.getElementById('actionSearch').value || ''; const params = new URLSearchParams({ page: actionsPage, pageSize: actionsPageSize, q, _t: Date.now() }); document.getElementById('logsLoading').style.display = 'block'; try { const res = await fetch('/api/admin/action_logs?' + params.toString()); if (!res.ok) { showToast('載入操作紀錄失敗', 'error'); return; } const j = await res.json(); currentLogs.action = j.data || []; actionsTotal = j.total || 0; actionsPages = j.pages || 1; document.getElementById('actionsTableBody').innerHTML = currentLogs.action.map(l => `<tr><td data-label="時間" style="padding:12px;white-space:nowrap;">${new Date(l.created_at).toLocaleString('zh-TW')}</td><td data-label="帳號">${l.username}</td><td data-label="動作"><span class="badge new">${l.action}</span></td><td data-label="詳細內容"><div style="font-size:12px;color:#666;">${l.details}</div></td></tr>`).join(''); renderPagination('actionsPagination', actionsPage, actionsPages, 'loadActionsPage'); } catch (e) { console.error(e); showToast('載入操作紀錄錯誤', 'error'); } finally { document.getElementById('logsLoading').style.display = 'none'; } }

        function exportLogs(type) { const data = type === 'login' ? currentLogs.login : currentLogs.action; if (!data || data.length === 0) return showToast('無資料可匯出', 'error'); let csvContent = '\uFEFF'; if (type === 'login') { csvContent += "時間,帳號,IP位址\n"; data.forEach(row => { csvContent += `"${new Date(row.login_time).toLocaleString('zh-TW')}","${row.username}","${row.ip_address}"\n`; }); } else { csvContent += "時間,帳號,動作,詳細內容\n"; data.forEach(row => { csvContent += `"${new Date(row.created_at).toLocaleString('zh-TW')}","${row.username}","${row.action}","${(row.details || '').replace(/"/g, '""')}"\n`; }); } const link = document.createElement("a"); link.setAttribute("href", URL.createObjectURL(new Blob([csvContent], { type: 'text/csv;charset=utf-8;' }))); link.setAttribute("download", `${type}_logs_${new Date().toISOString().slice(0, 10)}.csv`); document.body.appendChild(link); link.click(); document.body.removeChild(link); }
        async function clearLogs(type) { if (!confirm(`確定要清空所有「${type === 'login' ? '登入' : '操作'}」紀錄嗎？此動作無法復原！`)) return; const endpoint = type === 'login' ? '/api/admin/logs' : '/api/admin/action_logs'; try { const res = await fetch(endpoint, { method: 'DELETE' }); if (res.ok) { showToast('紀錄已清空'); if (type === 'login') loadLogsPage(1); else loadActionsPage(1); } else showToast('清除失敗', 'error'); } catch (e) { showToast('Error: ' + e.message, 'error'); } }

        function switchAdminTab(tab) { document.querySelectorAll('.admin-tab-btn').forEach(b => b.classList.remove('active')); event.target.classList.add('active'); document.getElementById('tab-users').classList.toggle('hidden', tab !== 'users'); document.getElementById('tab-logs').classList.toggle('hidden', tab !== 'logs'); document.getElementById('tab-actions').classList.toggle('hidden', tab !== 'actions'); if (tab === 'logs') loadLogsPage(1); if (tab === 'actions') loadActionsPage(1); }

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

                        var rawNumText = cells[col.number] ? (cells[col.number].innerText || cells[col.number].textContent || "") : "";
                        var info = parseItemNumber(rawNumText);
                        if (!info || !info.raw) continue;

                        var unitName = ORG_MAP[info.orgCode] || info.orgCode || "";
                        var inspectName = INSPECTION_MAP[info.inspectCode] || info.inspectCode || "";
                        var divName = DIVISION_MAP[info.divCode] || info.divCode || "";
                        var kindName = KIND_MAP[info.kindCode] || "其他";

                        var item = {
                            number: info.raw.toUpperCase(),
                            year: String(info.yearRoc),
                            unit: unitName,
                            itemKindCode: info.kindCode,
                            category: kindName,
                            inspectionCategoryName: inspectName,
                            divisionName: divName,
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
            const f = document.getElementById('wordInput').files[0];
            if (currentImportMode === 'backup') return;

            const stageRadio = document.querySelector('input[name="importStage"]:checked');
            if (!stageRadio) return;

            const stage = stageRadio.value;
            let valid = false;

            if (stage === 'initial') {
                const d = document.getElementById('importIssueDate').value.trim();
                valid = (d.length > 0);
            } else {
                valid = true;
            }

            document.getElementById('wordInput').disabled = !valid;
            document.getElementById('btnParseWord').disabled = !valid || !f;
        }
        document.getElementById('wordInput').addEventListener('change', checkImportReady);

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

            const planName = isBackup ? '' : document.getElementById('importPlanName').value;

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
                if (res.ok) { showToast('匯入成功！'); cancelImport(); loadIssuesPage(1); loadPlanOptions(); } else { showToast('匯入失敗', 'error'); }
            } catch (e) { showToast('Error: ' + e.message, 'error'); }
        }

        function switchDataTab(tab) { document.querySelectorAll('#importView .admin-tab-btn').forEach(b => b.classList.remove('active')); event.target.classList.add('active'); document.getElementById('tab-data-import').classList.toggle('hidden', tab !== 'import'); document.getElementById('tab-data-manual').classList.toggle('hidden', tab !== 'manual'); document.getElementById('tab-data-export').classList.toggle('hidden', tab !== 'export'); document.getElementById('tab-data-batch').classList.toggle('hidden', tab !== 'batch'); if (tab === 'batch' && document.querySelectorAll('#batchGridBody tr').length === 0) initBatchGrid(); }

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
            const planName = document.getElementById('batchPlanName').value.trim();
            const issueDate = document.getElementById('batchIssueDate').value.trim();

            if (!planName) return showToast('請填寫檢查計畫名稱', 'error');
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
                if (info.yearRoc) document.getElementById('manualYear').value = info.yearRoc;
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
            const year = document.getElementById('manualYear').value.trim();
            const unit = document.getElementById('manualUnit').value.trim();
            const division = document.getElementById('manualDivision').value;
            const inspection = document.getElementById('manualInspection').value;
            const kind = document.getElementById('manualKind').value;

            const planName = document.getElementById('manualPlanName').value.trim();
            const issueDate = document.getElementById('manualIssueDate').value.trim();
            const continuousMode = document.getElementById('manualContinuousMode').checked;

            const status = document.getElementById('manualStatus').value;
            const content = document.getElementById('manualContent').value.trim();
            if (!number || !year || !unit || !content) return showToast('請填寫所有必填欄位', 'error');

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
                        document.getElementById('manualYear').value = '';
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

        async function exportAllIssues() {
            try {
                const exportScope = document.querySelector('input[name="exportScope"]:checked').value;
                const exportFormat = document.querySelector('input[name="exportFormat"]:checked').value;
                showToast('準備匯出中，請稍候...');
                const res = await fetch('/api/issues?page=1&pageSize=10000&sortField=created_at&sortDir=desc');
                if (!res.ok) throw new Error('取得資料失敗');
                const json = await res.json();
                const data = json.data || [];
                if (data.length === 0) return showToast('無資料可匯出', 'error');

                if (exportFormat === 'json') {
                    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
                    const link = document.createElement("a"); link.href = URL.createObjectURL(blob); link.download = `SMS_Backup_${new Date().toISOString().slice(0, 10)}.json`; document.body.appendChild(link); link.click(); document.body.removeChild(link); showToast('JSON 匯出完成'); return;
                }

                let csvContent = '\uFEFF'; const clean = (t) => `"${String(t || '').replace(/"/g, '""').replace(/<[^>]*>/g, '').trim()}"`;

                const baseHeader = "編號,年度,機構,分組,檢查種類,類型,狀態,事項內容";

                if (exportScope === 'latest') {
                    csvContent += baseHeader + ",最新辦理情形,最新審查意見\n";
                    data.forEach(item => {
                        let latestH = '', latestR = '';
                        // 支持無限次，動態查找最新資料（從200向下找）
                        for (let i = 200; i >= 1; i--) { 
                            const suffix = i === 1 ? '' : i;
                            if (!latestH && (item[`handling${suffix}`])) latestH = item[`handling${suffix}`]; 
                            if (!latestR && (item[`review${suffix}`])) latestR = item[`review${suffix}`]; 
                        }

                        csvContent += `${clean(item.number)},${clean(item.year)},${clean(item.unit)},${clean(item.divisionName)},${clean(item.inspectionCategoryName)},${clean(item.category)},${clean(item.status)},${clean(item.content)},${clean(latestH)},${clean(latestR)}\n`;
                    });
                } else {
                    csvContent += baseHeader + ",完整辦理情形歷程,完整審查意見歷程\n";
                    data.forEach(item => {
                        let fullH = [], fullR = [];
                        // 支持無限次，動態查找（從1到200，實際應該不會超過這個數字）
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
                const link = document.createElement("a"); link.setAttribute("href", URL.createObjectURL(new Blob([csvContent], { type: 'text/csv;charset=utf-8;' }))); const typeLabel = exportScope === 'latest' ? 'Latest' : 'FullHistory'; link.setAttribute("download", `SMS_Issues_${typeLabel}_${new Date().toISOString().slice(0, 10)}.csv`); document.body.appendChild(link); link.click(); document.body.removeChild(link); showToast('CSV 匯出完成');
            } catch (e) { showToast('匯出失敗: ' + e.message, 'error'); }
        }

        // --- User modal submit & password strength ---
        document.getElementById('uPwd')?.addEventListener('input', updatePwdStrength); document.getElementById('uPwdConfirm')?.addEventListener('input', updatePwdStrength);
        function updatePwdStrength() { const p = document.getElementById('uPwd').value || ''; const conf = document.getElementById('uPwdConfirm').value || ''; let score = 0; if (p.length >= 8) score++; if (/[A-Z]/.test(p)) score++; if (/[0-9]/.test(p)) score++; if (/[^A-Za-z0-9]/.test(p)) score++; const texts = ['弱', '偏弱', '一般', '良好', '強']; document.getElementById('pwdStrength').innerText = `密碼強度: ${texts[Math.min(score, 4)]} ${conf && p !== conf ? '(密碼不相符)' : ''}`; }

        // User CRUD
        async function openUserModal(mode, id) { const m = document.getElementById('userModal'), t = document.getElementById('userModalTitle'), e = document.getElementById('uEmail'); if (mode === 'create') { t.innerText = '新增'; document.getElementById('targetUserId').value = ''; document.getElementById('uName').value = ''; e.value = ''; e.disabled = false; document.getElementById('uPwd').value = ''; document.getElementById('uPwdConfirm').value = ''; document.getElementById('pwdStrength').innerText = '密碼強度: -'; document.getElementById('pwdHint').innerText = ''; document.getElementById('uRole').value = 'viewer'; } else { const u = userList.find(x => x.id === id) || {}; t.innerText = '編輯'; document.getElementById('targetUserId').value = u.id || ''; document.getElementById('uName').value = u.name || ''; e.value = u.username || ''; e.disabled = true; document.getElementById('uPwd').value = ''; document.getElementById('uPwdConfirm').value = ''; document.getElementById('pwdHint').innerText = '(留空不改)'; document.getElementById('pwdStrength').innerText = '密碼強度: -'; document.getElementById('uRole').value = u.role || 'viewer'; } m.classList.add('open'); }
        async function submitUser() { const id = document.getElementById('targetUserId').value, name = document.getElementById('uName').value, email = document.getElementById('uEmail').value, pwd = document.getElementById('uPwd').value, pwdConfirm = document.getElementById('uPwdConfirm').value, role = document.getElementById('uRole').value; if (!id) { if (!email) return showToast('請輸入帳號', 'error'); if (!pwd) return showToast('請輸入密碼', 'error'); if (pwd !== pwdConfirm) return showToast('密碼與確認密碼不符', 'error'); if (pwd.length < 8) return showToast('密碼需至少 8 碼', 'error'); const res = await fetch('/api/users', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: email, name, password: pwd, role }) }); const j = await res.json(); if (res.ok) { showToast('新增成功'); document.getElementById('userModal').classList.remove('open'); loadUsersPage(1); } else showToast(j.error || '新增失敗', 'error'); } else { const payload = { name, role }; if (pwd) { if (pwd !== pwdConfirm) return showToast('密碼與確認密碼不符', 'error'); if (pwd.length < 8) return showToast('密碼需至少 8 碼', 'error'); payload.password = pwd; } const res = await fetch(`/api/users/${id}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) }); const j = await res.json(); if (res.ok) { showToast('更新成功'); document.getElementById('userModal').classList.remove('open'); loadUsersPage(usersPage); } else showToast(j.error || '更新失敗', 'error'); } }
        async function deleteUser(id) { if (!confirm('確定?')) return; const res = await fetch(`/api/users/${id}`, { method: 'DELETE' }); if (res.ok) { showToast('刪除成功'); loadUsersPage(1); } else showToast('刪除失敗', 'error'); }

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
                
                // 計算已經審查的最高次數，然後自動跳到下一次（支持無限次）
                let highestRound = 0;
                // 動態查找所有可能的回合（檢查所有欄位）
                for (let i = 1; i <= 200; i++) {
                    const suffix = i === 1 ? '' : i;
                    const hasHandling = currentEditItem['handling' + suffix] && currentEditItem['handling' + suffix].trim();
                    const hasReview = currentEditItem['review' + suffix] && currentEditItem['review' + suffix].trim();
                    if (hasHandling || hasReview) {
                        highestRound = i;
                    }
                }
                // 跳到下一次審查（如果最高是第2次，就跳到第3次）
                const nextRound = highestRound + 1;
                // 確保選項存在
                ensureRoundOption(nextRound);
                document.getElementById('editRound').value = nextRound;
                
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
            const s = document.getElementById('editRound'); 
            if (!s || s.options.length > 0) return; 
            // 支援無限次審查，先建立前 100 次選項（如果需要更多，可以動態添加）
            for (let i = 1; i <= 100; i++) { 
                const o = document.createElement('option'); 
                o.value = i; 
                o.text = `第 ${i} 次`; 
                s.add(o); 
            } 
        }
        
        // 動態添加更多審查次數選項（如果需要超過 100 次）
        function ensureRoundOption(round) {
            const s = document.getElementById('editRound');
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
            for (let i = 200; i >= 1; i--) {
                const suffix = i === 1 ? '' : i;
                const ha = currentEditItem['handling' + suffix], re = currentEditItem['review' + suffix];
                const replyDate = currentEditItem['reply_date_r' + i];
                const responseDate = currentEditItem['response_date_r' + i];

                if (ha || re) {
                    const latestBadge = firstRecord ? '<span class="badge new" style="margin-left:8px;font-size:11px;">最新進度</span>' : '';

                    let dateInfo = '';
                    if (replyDate || responseDate) {
                        dateInfo = `<div style="margin-bottom:12px;">`;
                        if (replyDate) dateInfo += `<span class="timeline-date-tag">🏢 機構回復: ${replyDate}</span> `;
                        if (responseDate) dateInfo += `<span class="timeline-date-tag">🏛️ 機關函復: ${responseDate}</span>`;
                        dateInfo += `</div>`;
                    }

                    h += `<div class="timeline-item">
                        <div class="timeline-dot"></div>
                        <div class="timeline-title">第 ${i} 次辦理情形 ${latestBadge}</div>
                        ${dateInfo}
                        ${ha ? `<div style="background:#f8fafc;padding:16px;border-radius:8px;font-size:14px;line-height:1.6;color:#334155;border:1px solid #e2e8f0;margin-bottom:12px;">${ha}</div>` : ''}
                        ${re ? `<div style="background:#fff;padding:16px;border-radius:8px;font-size:14px;line-height:1.6;color:#334155;border:1px solid #e2e8f0;border-left:3px solid var(--primary);"><strong>審查意見：</strong><br>${re}</div>` : ''}
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
            
            document.getElementById('editHandling').value = handling;
            document.getElementById('editReview').value = review;
            document.getElementById('editReplyDate').value = replyDate;
            document.getElementById('editResponseDate').value = responseDate;
            
            // 顯示上一回合的審查意見（如果有）
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
                const prevBox = document.getElementById('prevReviewBox');
                if (prevBox) prevBox.style.display = 'none';
            }
            
            // 清除 AI 分析結果（因為回合改變了）
            const aiBox = document.getElementById('aiBox');
            if (aiBox) aiBox.style.display = 'none';
            document.getElementById('aiPreviewText').innerText = '';
            document.getElementById('aiResBadge').innerHTML = '';
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