    // tsm-se.js — TSM SE Dashboard Module v5.0.0
    // SP List mode: load / upload TSM SE accounts (~22k) from SharePoint list + Excel
    (function () {
        'use strict';

        var TSM_SE_MODULE_VERSION = '5.0.0';
        var TSM_SE_CANDIDATE_LISTS = [
            'Service Manager Request SE',
            'TSM SE Request',
            'TSM_SE_Request',
            'TSM SE Accounts',
            'TSM_SE_Accounts',
            'Undedicated Accounts'
        ];
        var TSM_SE_SP_CONCURRENCY = 8;
        var TSM_SE_MIN_LIST_ROWS = 5000;

        var tsmSeState = {
            listName: null,
            entityType: null,
            listFields: {},
            allRows: [],
            filteredRows: [],
            uploadRows: [],
            loaded: false,
            loading: false
        };

        window.TSM_SE_LIST = window.TSM_SE_LIST || null;
        window.tsmSeAllData = window.tsmSeAllData || [];
        window.tsmSeRows = window.tsmSeRows || [];
        window.tsmSeGridApi = window.tsmSeGridApi || null;

        function tsmSeSpUrl() {
            return (typeof SP_URL !== 'undefined' && SP_URL) ? SP_URL : '';
        }

        function tsmSeEsc(s) {
            return String(s == null ? '' : s)
                .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;');
        }

        function tsmSeOdata(s) {
            return String(s == null ? '' : s).replace(/'/g, "''");
        }

        function tsmSeNormHeader(h) {
            return String(h || '').replace(/\u00a0/g, ' ').trim().toLowerCase().replace(/[_\s\-]+/g, ' ');
        }

        function tsmSePlain(val) {
            if (val == null || val === '') return '';
            if (typeof val === 'object') {
                if (val.Title) return String(val.Title).trim();
                if (val.results && val.results.length) return String(val.results[0]).trim();
                return '';
            }
            return String(val).trim();
        }

        function tsmSeToast(msg, type) {
            if (typeof window.showToast === 'function') window.showToast(msg, type || 'info');
            else alert(msg);
        }

        function tsmSeCanUpload() {
            var role = (window.USER_CONTEXT && USER_CONTEXT.role) || '';
            return !!(window.USER_CONTEXT && (window.USER_CONTEXT.isAdmin || role === 'Admin' || role === 'Service Director'));
        }

        async function tsmSeGetDigest() {
            var res = await fetch(tsmSeSpUrl() + '/_api/contextinfo', {
                method: 'POST',
                headers: { 'Accept': 'application/json;odata=verbose' },
                credentials: 'include'
            });
            if (!res.ok) throw new Error('Digest failed');
            var data = await res.json();
            return data.d.GetContextWebInformation.FormDigestValue;
        }

        async function tsmSeDiscoverListName() {
            if (tsmSeState.listName) return tsmSeState.listName;
            var named = [window.TSM_SE_LIST, window.TSM_SE_SP_LIST, window.TSMSE_LIST, window.TSM_SE_SP_LIST_NAME];
            for (var i = 0; i < named.length; i++) {
                if (named[i]) {
                    tsmSeState.listName = named[i];
                    window.TSM_SE_LIST = named[i];
                    return named[i];
                }
            }
            try {
                var listsRes = await fetch(tsmSeSpUrl() + "/_api/web/lists?$select=Title,ItemCount&$filter=Hidden eq false&$top=500", {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'include'
                });
                if (listsRes.ok) {
                    var lists = (((await listsRes.json()).d || {}).results) || [];
                    var scored = lists.filter(function (l) {
                        return l.Title !== 'Service Manager Request' && (l.ItemCount || 0) >= TSM_SE_MIN_LIST_ROWS;
                    }).sort(function (a, b) {
                        var sa = /TSM|SE|Undedicated/i.test(a.Title) ? 1 : 0;
                        var sb = /TSM|SE|Undedicated/i.test(b.Title) ? 1 : 0;
                        if (sb !== sa) return sb - sa;
                        return (b.ItemCount || 0) - (a.ItemCount || 0);
                    });
                    if (scored.length) {
                        tsmSeState.listName = scored[0].Title;
                        window.TSM_SE_LIST = scored[0].Title;
                        console.log('[TSM SE] SP List mode:', scored[0].Title, scored[0].ItemCount);
                        return tsmSeState.listName;
                    }
                }
            } catch (e) {
                console.warn('[TSM SE] list discovery failed', e);
            }
            for (var c = 0; c < TSM_SE_CANDIDATE_LISTS.length; c++) {
                try {
                    var cRes = await fetch(tsmSeSpUrl() + "/_api/web/lists/getbytitle('" + tsmSeOdata(TSM_SE_CANDIDATE_LISTS[c]) + "')?$select=Title,ItemCount", {
                        headers: { 'Accept': 'application/json;odata=verbose' },
                        credentials: 'include'
                    });
                    if (cRes.ok) {
                        tsmSeState.listName = TSM_SE_CANDIDATE_LISTS[c];
                        window.TSM_SE_LIST = TSM_SE_CANDIDATE_LISTS[c];
                        return tsmSeState.listName;
                    }
                } catch (e) {}
            }
            throw new Error('TSM SE SharePoint list not found');
        }

        async function tsmSeLoadListFields(listName) {
            var url = tsmSeSpUrl() + "/_api/web/lists/getbytitle('" + tsmSeOdata(listName) + "')/fields?$select=Title,InternalName,TypeAsString&$filter=Hidden eq false and ReadOnlyField eq false&$top=500";
            var res = await fetch(url, { headers: { 'Accept': 'application/json;odata/verbose' }, credentials: 'include' });
            if (!res.ok) return;
            var rows = (((await res.json()).d || {}).results) || [];
            tsmSeState.listFields = {};
            rows.forEach(function (f) {
                tsmSeState.listFields[f.InternalName] = f.TypeAsString || 'Text';
            });
        }

        async function tsmSeGetEntityType(listName) {
            var res = await fetch(tsmSeSpUrl() + "/_api/web/lists/getbytitle('" + tsmSeOdata(listName) + "')?$select=ListItemEntityTypeFullName", {
                headers: { 'Accept': 'application/json;odata=verbose' },
                credentials: 'include'
            });
            if (!res.ok) throw new Error('Could not read list entity type');
            var data = await res.json();
            tsmSeState.entityType = data.d.ListItemEntityTypeFullName;
            return tsmSeState.entityType;
        }

        function tsmSeMapSpItem(item) {
            var code = tsmSePlain(item.Title);
            var parent = tsmSePlain(item.Parent_x0020_Code) || code;
            var isGroup = !item.Parent_x0020_Code || parent === code;
            return {
                spItemId: item.ID || item.Id || null,
                code: code,
                parent: parent,
                customer: tsmSePlain(item.Customer_x0020_Name),
                team: tsmSePlain(item.Team) || 'TSM_SE',
                lm: tsmSePlain(item.Line_x0020_Manager),
                sm: tsmSePlain(item.Service_x0020_Manager),
                secondarySm: tsmSePlain(item.Secondary_x0020_Service_x0020_Ma),
                am: tsmSePlain(item.Account_x0020_Manager),
                ad: tsmSePlain(item.Account_x0020_Director),
                pocName: tsmSePlain(item.POC_x0020_Name),
                pocEmail: tsmSePlain(item.POC_x0020_Email_x0020_ID),
                pocPhone: tsmSePlain(item.POC_x0020_Contact_x0020_No),
                type: isGroup ? 'Group' : 'Child',
                avg: parseFloat(item.Jan_x002d_26 || item.Feb_x002d_26 || 0) || 0,
                _raw: item
            };
        }

        async function tsmSeFetchAllRows() {
            var listName = await tsmSeDiscoverListName();
            await tsmSeLoadListFields(listName);
            var rows = [];
            var url = tsmSeSpUrl() + "/_api/web/lists/getbytitle('" + tsmSeOdata(listName) + "')/items?$top=5000";
            while (url) {
                var res = await fetch(url, { headers: { 'Accept': 'application/json;odata=verbose' }, credentials: 'include' });
                if (!res.ok) throw new Error('Failed to load TSM SE list: ' + res.status);
                var data = await res.json();
                var batch = (data.d && data.d.results) || [];
                batch.forEach(function (item) {
                    if (item.Title) rows.push(tsmSeMapSpItem(item));
                });
                url = data.d && data.d.__next ? data.d.__next : null;
            }
            tsmSeState.allRows = rows;
            window.tsmSeAllData = rows;
            window.tsmSeRows = rows;
            window.TSM_SE_ROWS = rows;
            tsmSeState.loaded = true;
            console.log('[TSM SE] Loaded', rows.length, 'accounts from', listName);
            return rows;
        }

        var TSM_SE_HEADER_MAP = {
            Title: ['account code', 'account', 'title', 'code', 'account no', 'account number', 'acct code'],
            Parent_x0020_Code: ['parent code', 'parent', 'parent account', 'parent account code'],
            Customer_x0020_Name: ['customer name', 'customer', 'company', 'company name'],
            Team: ['team'],
            Line_x0020_Manager: ['line manager', 'lm', 'line mgr'],
            Service_x0020_Manager: ['service manager', 'sm', 'service mgr', 'service manager name'],
            Secondary_x0020_Service_x0020_Ma: ['secondary sm', 'secondary service manager', 'secondary service mgr'],
            Account_x0020_Manager: ['account manager', 'am', 'account mgr'],
            Account_x0020_Director: ['account director', 'ad', 'account dir'],
            POC_x0020_Name: ['poc name', 'primary poc', 'poc'],
            POC_x0020_Email_x0020_ID: ['poc email', 'poc email id', 'email'],
            POC_x0020_Contact_x0020_No: ['poc contact', 'poc phone', 'poc mobile', 'contact']
        };

        function tsmSeMonthFieldFromHeader(h) {
            var m = String(h || '').trim().match(/^([A-Za-z]{3})[\s\-\/]?(\d{2,4})$/);
            if (!m) return null;
            var mon = m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase();
            var yr = m[2].length === 4 ? m[2].slice(-2) : m[2];
            return mon + '_x002d_' + yr;
        }

        function tsmSeBuildColumnIndex(headers) {
            var index = {};
            headers.forEach(function (h, i) {
                var norm = tsmSeNormHeader(h);
                if (!norm) return;
                Object.keys(TSM_SE_HEADER_MAP).forEach(function (field) {
                    if (index[field] != null) return;
                    if (TSM_SE_HEADER_MAP[field].indexOf(norm) >= 0) index[field] = i;
                });
                if (index.Title == null && (norm === 'title' || norm.indexOf('account') >= 0)) index.Title = i;
                var monthField = tsmSeMonthFieldFromHeader(h);
                if (monthField) index[monthField] = i;
            });
            return index;
        }

        function tsmSeParseExcelRows(workbook) {
            if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) return [];
            var sheet = workbook.Sheets[workbook.SheetNames[0]];
            var matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            if (!matrix || matrix.length < 2) return [];

            var headerRowIdx = 0;
            for (var r = 0; r < Math.min(matrix.length, 15); r++) {
                var joined = matrix[r].map(function (c) { return tsmSeNormHeader(c); }).join(' ');
                if (/account|title|customer|service manager|parent/.test(joined)) {
                    headerRowIdx = r;
                    break;
                }
            }

            var colIndex = tsmSeBuildColumnIndex(matrix[headerRowIdx]);
            if (colIndex.Title == null) {
                throw new Error('Could not find Account Code column in Excel (row ' + (headerRowIdx + 1) + ')');
            }

            var out = [];
            for (var i = headerRowIdx + 1; i < matrix.length; i++) {
                var row = matrix[i];
                if (!row || !row.length) continue;
                var title = String(row[colIndex.Title] || '').trim();
                if (!title) continue;

                var fields = { Title: title };
                Object.keys(colIndex).forEach(function (field) {
                    if (field === 'Title') return;
                    var val = row[colIndex[field]];
                    if (val == null || val === '') return;
                    fields[field] = String(val).trim();
                });

                if (!fields.Parent_x0020_Code) fields.Parent_x0020_Code = title;
                if (!fields.Team) fields.Team = 'TSM_SE';

                out.push(fields);
            }
            return out;
        }

        function tsmSeFilterPayload(fields) {
            var out = {};
            Object.keys(fields).forEach(function (k) {
                if (k === 'Title' || fields[k] == null || fields[k] === '') return;
                if (tsmSeState.listFields[k] || k.indexOf('_x002d_') >= 0 || k.indexOf('_x0020_') >= 0) {
                    out[k] = fields[k];
                }
            });
            return out;
        }

        async function tsmSeCreateItem(fields, digest, entityType) {
            var payload = tsmSeFilterPayload(fields);
            payload.Title = fields.Title;
            var body = {
                __metadata: { type: entityType },
                Title: payload.Title
            };
            Object.keys(payload).forEach(function (k) {
                if (k !== 'Title') body[k] = payload[k];
            });
            var listName = tsmSeState.listName;
            var res = await fetch(tsmSeSpUrl() + "/_api/web/lists/getbytitle('" + tsmSeOdata(listName) + "')/items", {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose',
                    'X-RequestDigest': digest
                },
                credentials: 'include',
                body: JSON.stringify(body)
            });
            if (!res.ok) {
                var err = await res.text();
                throw new Error('Create failed for ' + fields.Title + ': ' + err.slice(0, 180));
            }
            return res.json();
        }

        async function tsmSeDeleteItem(id, digest) {
            var listName = tsmSeState.listName;
            var res = await fetch(tsmSeSpUrl() + "/_api/web/lists/getbytitle('" + tsmSeOdata(listName) + "')/items(" + id + ")", {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'X-RequestDigest': digest,
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE'
                },
                credentials: 'include'
            });
            if (!res.ok) throw new Error('Delete failed for ID ' + id);
        }

        async function tsmSeRunPool(items, worker, options) {
            options = options || {};
            var concurrency = options.concurrency || TSM_SE_SP_CONCURRENCY;
            var onProgress = options.onProgress;
            var getDigest = options.getDigest;
            var setDigest = options.setDigest;
            var total = items.length;
            var idx = 0, ok = 0, fail = 0;
            if (!total) return { ok: 0, fail: 0 };

            async function runWorker() {
                while (true) {
                    var i = idx++;
                    if (i >= total) return;
                    if (getDigest && setDigest && i > 0 && i % 100 === 0) {
                        try { setDigest(await tsmSeGetDigest()); } catch (e) {}
                    }
                    try {
                        await worker(items[i], i);
                        ok++;
                    } catch (e) {
                        fail++;
                        console.error('[TSM SE upload]', e);
                    }
                    if (onProgress) onProgress(i + 1, total, ok, fail);
                }
            }
            var workers = [];
            for (var w = 0; w < Math.min(concurrency, total); w++) workers.push(runWorker());
            await Promise.all(workers);
            return { ok: ok, fail: fail };
        }

        function tsmSeSetProgress(msg) {
            var el = document.getElementById('tsmSeUploadProgress');
            if (el) el.textContent = msg;
        }

        window.tsmSeConfirmUpload = async function (smartReplace) {
            var rows = tsmSeState.uploadRows || [];
            if (!rows.length) {
                tsmSeToast('No rows to upload. Choose an Excel file first.', 'warn');
                return;
            }
            if (!tsmSeCanUpload()) {
                tsmSeToast('Only Admin can upload TSM SE accounts', 'warn');
                return;
            }

            var modeLabel = smartReplace ? 'Smart Upload (replace all accounts in list)' : 'Upload (add/update)';
            if (!confirm(modeLabel + '\n\nRows to upload: ' + rows.length + '\n\nContinue?')) return;

            var digest, entityType;
            try {
                await tsmSeDiscoverListName();
                await tsmSeLoadListFields(tsmSeState.listName);
                digest = await tsmSeGetDigest();
                entityType = await tsmSeGetEntityType(tsmSeState.listName);
            } catch (e) {
                tsmSeToast('SharePoint digest/list error: ' + e.message, 'error');
                return;
            }

            if (smartReplace) {
                tsmSeSetProgress('Loading existing list items for replace...');
                if (!tsmSeState.loaded) await tsmSeFetchAllRows();
                var existingIds = (tsmSeState.allRows || []).map(function (r) { return r.spItemId; }).filter(Boolean);
                if (existingIds.length) {
                    if (!confirm('Smart Upload will DELETE ' + existingIds.length + ' existing accounts, then upload ' + rows.length + ' from Excel.\n\nProceed?')) return;
                    tsmSeSetProgress('Deleting ' + existingIds.length + ' existing accounts...');
                    var delResult = await tsmSeRunPool(existingIds, function (id) {
                        return tsmSeDeleteItem(id, digest);
                    }, {
                        concurrency: TSM_SE_SP_CONCURRENCY,
                        getDigest: function () { return digest; },
                        setDigest: function (d) { digest = d; },
                        onProgress: function (done, total) {
                            tsmSeSetProgress('Deleting old accounts: ' + done + ' / ' + total);
                        }
                    });
                    console.log('[TSM SE] Deleted', delResult.ok, 'existing rows');
                }
            }

            tsmSeSetProgress('Uploading 0 / ' + rows.length);
            var uploadResult = await tsmSeRunPool(rows, function (fields) {
                return tsmSeCreateItem(fields, digest, entityType);
            }, {
                concurrency: TSM_SE_SP_CONCURRENCY,
                getDigest: function () { return digest; },
                setDigest: function (d) { digest = d; },
                onProgress: function (done, total, ok, fail) {
                    tsmSeSetProgress('Uploading: ' + ok + ' ok, ' + fail + ' failed (' + done + ' / ' + total + ')');
                }
            });

            tsmSeSetProgress('');
            tsmSeToast('Uploaded ' + uploadResult.ok + ' accounts' + (uploadResult.fail ? ' (' + uploadResult.fail + ' failed — check console)' : ''), uploadResult.fail ? 'warn' : 'success');
            tsmSeState.loaded = false;
            await tsmSeFetchAllRows();
            window.tsmSeRenderTable();
        };

        window.tsmSeSmartUpload = function () {
            window.tsmSeConfirmUpload(true);
        };

        window.tsmSeParseFile = function (ev) {
            var file = ev.target.files && ev.target.files[0];
            var preview = document.getElementById('tsmSeUploadPreview');
            if (!file) return;
            if (typeof XLSX === 'undefined') {
                tsmSeToast('XLSX library not loaded', 'error');
                return;
            }
            var reader = new FileReader();
            reader.onload = function (e) {
                try {
                    var wb = XLSX.read(e.target.result, { type: 'array' });
                    var rows = tsmSeParseExcelRows(wb);
                    tsmSeState.uploadRows = rows;
                    if (preview) {
                        preview.innerHTML = '<div style="font-size:.82rem;color:var(--t1);font-weight:700;">Parsed ' + rows.length + ' account rows from <b>' + tsmSeEsc(file.name) + '</b></div>' +
                            (rows.length ? '<div style="margin-top:.65rem;display:flex;gap:.5rem;flex-wrap:wrap;">' +
                                '<button type="button" class="export-btn" onclick="tsmSeConfirmUpload(false)">Upload ' + rows.length + ' Accounts</button>' +
                                '<button type="button" class="export-btn" onclick="tsmSeSmartUpload()" style="background:linear-gradient(135deg,#2563eb,#1d4ed8);color:#fff;border:none;">Smart Upload (Replace All)</button>' +
                            '</div>' : '<div style="color:#ef4444;margin-top:.5rem;">No valid account rows found — check Excel headers.</div>');
                    }
                    if (!rows.length) tsmSeToast('0 accounts parsed — check Excel column headers', 'warn');
                    else tsmSeToast('Parsed ' + rows.length + ' accounts — click Upload or Smart Upload', 'success');
                } catch (err) {
                    tsmSeState.uploadRows = [];
                    if (preview) preview.innerHTML = '<div style="color:#ef4444;font-weight:700;">' + tsmSeEsc(err.message) + '</div>';
                    tsmSeToast(err.message, 'error');
                }
                ev.target.value = '';
            };
            reader.readAsArrayBuffer(file);
        };

        function tsmSeGetFilterValues() {
            function fromMs(key) {
                if (typeof smGetMsFilterValues === 'function') return smGetMsFilterValues(key) || [];
                return [];
            }
            function fromSelect(id) {
                var el = document.getElementById(id);
                if (!el || !el.value) return [];
                return [el.value];
            }
            var lm = fromMs('LM').length ? fromMs('LM') : fromSelect('filterLM');
            var sm = fromMs('SM').length ? fromMs('SM') : fromSelect('filterSM');
            return { lm: lm, sm: sm };
        }

        function tsmSeApplyFilters(rows) {
            var f = tsmSeGetFilterValues();
            return (rows || []).filter(function (r) {
                if (f.lm.length && f.lm.indexOf(r.lm) < 0) return false;
                if (f.sm.length && f.sm.indexOf(r.sm) < 0) return false;
                return true;
            });
        }

        function tsmSeEnsureToolbar() {
            var gridDiv = document.getElementById('myGrid');
            if (!gridDiv) return null;
            var host = gridDiv.parentElement;
            if (!host) return null;
            var bar = document.getElementById('tsmSeToolbar');
            if (!bar) {
                bar = document.createElement('div');
                bar.id = 'tsmSeToolbar';
                bar.style.cssText = 'margin:0 0 12px;padding:14px;border:1px solid var(--border);border-radius:12px;background:var(--bg-card);';
                host.insertBefore(bar, gridDiv);
            }
            bar.innerHTML = '<div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;">' +
                '<div><div style="font-size:.92rem;font-weight:800;color:var(--t1);">TSM SE Accounts</div>' +
                '<div style="font-size:.76rem;color:var(--t3);">SP List mode · v' + TSM_SE_MODULE_VERSION + (tsmSeState.listName ? ' · ' + tsmSeEsc(tsmSeState.listName) : '') + '</div></div>' +
                (tsmSeCanUpload() ?
                    '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">' +
                        '<label class="export-btn" style="cursor:pointer;margin:0;">' +
                            '<input type="file" accept=".xlsx,.xls,.csv" style="display:none;" onchange="tsmSeParseFile(event)">Choose Excel</label>' +
                        '<button type="button" class="export-btn" onclick="tsmSeSmartUpload()" style="background:linear-gradient(135deg,#2563eb,#1d4ed8);color:#fff;border:none;">Smart Upload</button>' +
                    '</div>' : '') +
                '</div>' +
                '<div id="tsmSeUploadPreview" style="margin-top:.75rem;"></div>' +
                '<div id="tsmSeUploadProgress" style="margin-top:.5rem;font-size:.78rem;color:var(--t2);font-weight:600;"></div>';
            return bar;
        }

        function tsmSeDestroyGrid() {
            if (window.tsmSeGridApi) {
                try { window.tsmSeGridApi.destroy(); } catch (e) {}
                window.tsmSeGridApi = null;
            }
            var gridDiv = document.getElementById('myGrid');
            if (gridDiv) gridDiv.innerHTML = '';
        }

        function tsmSeColumnDefs() {
            return [
                { field: 'code', headerName: 'Account Code', pinned: 'left', width: 160 },
                { field: 'parent', headerName: 'Parent Code', width: 140 },
                { field: 'customer', headerName: 'Customer Name', width: 220 },
                { field: 'type', headerName: 'Type', width: 90 },
                { field: 'team', headerName: 'Team', width: 100 },
                { field: 'lm', headerName: 'Line Manager', width: 150 },
                { field: 'sm', headerName: 'Service Manager', width: 150 },
                { field: 'secondarySm', headerName: 'Secondary SM', width: 140 },
                { field: 'am', headerName: 'Account Manager', width: 150 },
                { field: 'ad', headerName: 'Account Director', width: 150 }
            ];
        }

        window.tsmSeRenderTable = async function () {
            if (!window.TSM_SE_ACTIVE) return;
            var gridDiv = document.getElementById('myGrid');
            if (!gridDiv) return;

            tsmSeEnsureToolbar();
            tsmSeDestroyGrid();

            if (!tsmSeState.loaded && !tsmSeState.loading) {
                tsmSeState.loading = true;
                gridDiv.innerHTML = '<div style="padding:40px;text-align:center;color:var(--t2);font-weight:600;">Loading TSM SE accounts from SharePoint...</div>';
                try {
                    await tsmSeFetchAllRows();
                } catch (e) {
                    gridDiv.innerHTML = '<div style="padding:40px;text-align:center;color:#ef4444;font-weight:700;">' + tsmSeEsc(e.message) + '</div>';
                    tsmSeState.loading = false;
                    return;
                }
                tsmSeState.loading = false;
            }

            var rows = tsmSeApplyFilters(tsmSeState.allRows || []);
            tsmSeState.filteredRows = rows;
            window.CURRENT_TABLE_DATA = rows;

            if (typeof agGrid === 'undefined') {
                gridDiv.innerHTML = '<div style="padding:40px;color:#ef4444;">AG Grid not loaded</div>';
                return;
            }

            var colDefs = tsmSeColumnDefs();
            if (typeof SmSetColumnFilter !== 'undefined') {
                colDefs = colDefs.map(function (c) {
                    c.filter = SmSetColumnFilter;
                    return c;
                });
            }

            var gridOptions = {
                columnDefs: colDefs,
                rowData: rows,
                defaultColDef: { sortable: true, resizable: true, filter: true, minWidth: 90 },
                animateRows: true,
                suppressCellFocus: true
            };

            if (agGrid.createGrid) {
                window.tsmSeGridApi = agGrid.createGrid(gridDiv, gridOptions);
            } else {
                new agGrid.Grid(gridDiv, gridOptions);
                window.tsmSeGridApi = gridOptions.api;
            }

            if (typeof lucide !== 'undefined') lucide.createIcons();
            console.log('[TSM SE] Grid rendered with', rows.length, 'rows');
        };

        window.tsmSeFindAccount = function (accountCode) {
            var want = String(accountCode || '').trim().toUpperCase();
            if (!want) return null;
            var rows = tsmSeState.allRows.length ? tsmSeState.allRows : (window.tsmSeAllData || []);
            return rows.find(function (r) {
                return String(r.code || '').trim().toUpperCase() === want;
            }) || null;
        };

        window.tsmSeFindAccountForSearch = async function (accountCode) {
            if (!tsmSeState.loaded) {
                try { await tsmSeFetchAllRows(); } catch (e) { return null; }
            }
            var hit = window.tsmSeFindAccount(accountCode);
            if (!hit || !hit._raw) return hit;
            return hit._raw;
        };

        window.tsmSeInit = function () {
            console.log('[TSM SE] Module loaded v' + TSM_SE_MODULE_VERSION);
            tsmSeDiscoverListName().catch(function (e) {
                console.warn('[TSM SE] Init list discovery:', e.message);
            });
            return true;
        };

        window.tsmSeDestroy = tsmSeDestroyGrid;
        window.tsmSeReset = function () { tsmSeState.loaded = false; tsmSeState.allRows = []; window.tsmSeAllData = []; };
        window.tsmSeClear = window.tsmSeReset;
        window.tsmSeRestoreMainGrid = function () {
            var bar = document.getElementById('tsmSeToolbar');
            if (bar) bar.remove();
            tsmSeDestroyGrid();
        };
        window.tsmSeHide = window.tsmSeRestoreMainGrid;
        window.tsmSeGetAllRows = function () { return tsmSeState.allRows.slice(); };

        window.TSM_SE_MODULE_VERSION = TSM_SE_MODULE_VERSION;
        window.tsmSeInit();
    })();
