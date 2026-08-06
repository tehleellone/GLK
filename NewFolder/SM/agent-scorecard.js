/* ═══════════════════════════════════════════════════════════════
   AGENT SCORE CARD — UI + Query Layer v2.0
   All numbers computed from window.ASC_DATA (agent-scorecard-data.js)
   ═══════════════════════════════════════════════════════════════ */
'use strict';

var ASC_CLR = {
    disaster: '#e74c3c', passive: '#f39c12', promo: '#27ae60',
    within: '#27ae60', outside: '#e74c3c', acc: '#a855f7', blue: '#4c6fff',
    smColors: ['#4c6fff', '#a855f7', '#c724b1', '#27ae60', '#e74c3c', '#f39c12', '#3498db', '#9b59b6', '#1abc9c', '#e67e22']
};

var ASC = {
    team: 'TSM_SE',
    view: 'landing',
    drillAgentId: null,
    drillAccountId: null,
    drillTab: 'overview',
    selectedCallId: null,
    period: 'month',
    filterAgentIds: [],
    filterLocations: [],
    filterScoreBand: '',
    filterKpi: null,
    filterSla: null,
    filterCallDir: 'all',
    filterEmailStatus: null,
    search: '',
    chartsVisible: false,
    charts: {},
    detailGridApi: null,
    filteredAgents: [],
    filteredCalls: [],
    filteredEmails: [],
    filteredPending: []
};

function ascData() {
    if (!window.ASC_DATA) throw new Error('ASC_DATA not loaded — include agent-scorecard-data.js before agent-scorecard.js');
    return window.ASC_DATA;
}
function ascCfg() { return ascData().config; }
function ascTargets() { return ascCfg().targets; }

function ascFmtDate(iso) {
    if (!iso) return '—';
    return new Date(iso).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}
function ascFmtPct(v) { return v == null || isNaN(v) ? '—' : Math.round(v) + '%'; }
function ascIcons() { if (typeof lucide !== 'undefined') lucide.createIcons(); }

function ascToast(msg) {
    var el = document.getElementById('asc-toast');
    if (!el) return;
    el.textContent = msg;
    el.classList.add('show');
    setTimeout(function() { el.classList.remove('show'); }, 2800);
}

function ascPeriodStart(period) {
    var now = new Date(), d = new Date(now);
    if (period === 'today') { d.setHours(0, 0, 0, 0); return d; }
    if (period === 'week') { d.setDate(d.getDate() - 7); return d; }
    if (period === 'month') { d.setMonth(d.getMonth() - 1); return d; }
    if (period === 'quarter') { d.setMonth(d.getMonth() - 3); return d; }
    if (period === 'ytd') return new Date(now.getFullYear(), 0, 1);
    d.setFullYear(d.getFullYear() - 1);
    return d;
}
function ascInPeriod(iso, period) {
    if (!iso) return false;
    return new Date(iso) >= ascPeriodStart(period);
}

function ascAgentById(id) {
    return ascData().agents.find(function(a) { return a.id === id; });
}
function ascAccountById(id) {
    return ascData().accounts.find(function(a) { return a.id === id; });
}
function ascLmName(lmId) {
    var lm = ascData().lineManagers.find(function(l) { return l.id === lmId; });
    return lm ? lm.name : '—';
}

function ascScoreBand(score) {
    var bands = ascCfg().scoreBands;
    if (score >= bands[0].min) return 'high';
    if (score >= bands[1].min) return 'mid';
    return 'low';
}

/* ── QUERY LAYER (all metrics from dataset) ─────────────────────── */
function ascQCalls(opts) {
    opts = opts || {};
    return ascData().calls.filter(function(c) {
        if (opts.team && ascAgentById(c.agentId).team !== opts.team) return false;
        if (opts.agentId && c.agentId !== opts.agentId) return false;
        if (opts.accountId && c.accountId !== opts.accountId) return false;
        if (opts.period && !ascInPeriod(c.datetime, opts.period)) return false;
        if (opts.direction && opts.direction !== 'all' && c.direction !== opts.direction) return false;
        if (opts.agentIds && opts.agentIds.indexOf(c.agentId) === -1) return false;
        return true;
    });
}

function ascQEmails(opts) {
    opts = opts || {};
    return ascData().emails.filter(function(e) {
        if (opts.team && ascAgentById(e.agentId).team !== opts.team) return false;
        if (opts.agentId && e.agentId !== opts.agentId) return false;
        if (opts.accountId && e.accountId !== opts.accountId) return false;
        if (opts.period && !ascInPeriod(e.receivedAt, opts.period)) return false;
        if (opts.status && e.status !== opts.status) return false;
        if (opts.slaStatus && e.slaStatus !== opts.slaStatus) return false;
        if (opts.openOnly && e.status === 'Closed') return false;
        if (opts.agentIds && opts.agentIds.indexOf(e.agentId) === -1) return false;
        return true;
    });
}

function ascQPending(opts) {
    opts = opts || {};
    return ascData().pendingItems.filter(function(p) {
        if (opts.team && ascAgentById(p.agentId).team !== opts.team) return false;
        if (opts.agentId && p.agentId !== opts.agentId) return false;
        if (opts.accountId && p.accountId !== opts.accountId) return false;
        if (opts.period && !ascInPeriod(p.createdAt, opts.period)) return false;
        if (opts.recordType && p.recordType !== opts.recordType) return false;
        if (opts.status && p.status !== opts.status) return false;
        if (opts.agentIds && opts.agentIds.indexOf(p.agentId) === -1) return false;
        return true;
    });
}

function ascQAttendance(agentId) {
    return ascData().attendanceWeeks.filter(function(a) { return a.agentId === agentId; });
}

function ascQQa(agentId, period) {
    return ascData().qaReviews.filter(function(q) {
        return q.agentId === agentId && (!period || ascInPeriod(q.reviewDate, period));
    });
}

function ascQAccounts(agentId) {
    return ascData().accounts.filter(function(a) { return a.agentId === agentId; });
}

function ascQAccountMetrics(accountId, period) {
    var d = ascData();
    var inP = function(iso) { return !period || ascInPeriod(iso, period); };
    var tt = d.troubleTickets.filter(function(t) { return t.accountId === accountId && inP(t.openedAt); });
    var ev = d.events.filter(function(e) { return e.accountId === accountId && inP(e.eventDate); });
    var hc = d.healthChecks.filter(function(h) { return h.accountId === accountId && inP(h.checkDate); });
    var qsr = d.qsrRecords.filter(function(q) { return q.accountId === accountId; });
    var rnps = d.rnpsSurveys.filter(function(r) { return r.accountId === accountId && inP(r.surveyDate); });
    var latestHc = hc.sort(function(a, b) { return new Date(b.checkDate) - new Date(a.checkDate); })[0];
    var qsrRow = qsr[0];
    var avgRnps = rnps.length ? rnps.reduce(function(s, r) { return s + r.score; }, 0) / rnps.length : null;
    var activity = tt.length * 3 + ev.length * 2 + (rnps.length * 5);
    var engagement = activity >= 25 ? 'High' : activity >= 12 ? 'Medium' : 'Low';
    return {
        ttCount: tt.length,
        openTT: tt.filter(function(t) { return t.status !== 'Closed'; }).length,
        eventCount: ev.length,
        healthCheckStatus: latestHc ? latestHc.status : '—',
        healthCheckScore: latestHc ? latestHc.score : null,
        qsrStatus: qsrRow ? qsrRow.status : '—',
        qsrScore: qsrRow ? qsrRow.score : null,
        avgRnps: avgRnps,
        rnpsCount: rnps.length,
        activityScore: activity,
        engagement: engagement,
        tt: tt,
        events: ev,
        healthChecks: hc,
        rnps: rnps,
        qsr: qsr
    };
}

function ascComputeTnps(calls) {
    var fb = calls.filter(function(c) { return c.feedback.submitted && c.feedback.type; });
    if (!fb.length) return null;
    var promo = fb.filter(function(c) { return c.feedback.type === 'Promotion'; }).length;
    return Math.round((promo / fb.length) * 100);
}

function ascComputeAgentMetrics(agentId, period) {
    var agent = ascAgentById(agentId);
    if (!agent) return null;
    var targets = ascTargets();
    var weights = ascCfg().scoreWeights[agent.team];

    var calls = ascQCalls({ agentId: agentId, period: period });
    var emails = ascQEmails({ agentId: agentId, period: period });
    var pending = ascQPending({ agentId: agentId, period: period });
    var att = ascQAttendance(agentId);
    var qaRows = ascQQa(agentId, period);

    var fbCalls = calls.filter(function(c) { return c.feedback.submitted; });
    var tnps = ascComputeTnps(calls);
    var releaseRate = calls.length ? Math.round((fbCalls.length / calls.length) * 100) : 0;

    var openEmails = emails.filter(function(e) { return e.status !== 'Closed'; });
    var slaWithin = openEmails.filter(function(e) { return e.slaStatus === 'Within SLA'; }).length;
    var slaCompliance = openEmails.length ? Math.round((slaWithin / openEmails.length) * 100) : 100;

    var attRate = att.length ? Math.round(att.reduce(function(s, a) { return s + a.attendanceRate; }, 0) / att.length) : 0;
    var qaScore = qaRows.length ? Math.round(qaRows.reduce(function(s, q) { return s + q.score; }, 0) / qaRows.length) : null;
    var qaPassRate = qaRows.length ? Math.round(qaRows.filter(function(q) { return q.passed; }).length / qaRows.length * 100) : null;

    var accounts = ascQAccounts(agentId);
    var accountMetrics = accounts.map(function(acct) {
        return Object.assign({ account: acct }, ascQAccountMetrics(acct.id, period));
    });
    var avgRnps = accountMetrics.length && accountMetrics.some(function(a) { return a.avgRnps != null; })
        ? Math.round(accountMetrics.filter(function(a) { return a.avgRnps != null; }).reduce(function(s, a) { return s + a.avgRnps; }, 0) / accountMetrics.filter(function(a) { return a.avgRnps != null; }).length)
        : null;
    var hcPassRate = accountMetrics.length
        ? Math.round(accountMetrics.filter(function(a) { return a.healthCheckStatus === 'Pass'; }).length / accountMetrics.length * 100)
        : null;

    var parts = {};
    var composite = 0;
    if (agent.team === 'TSM_SE') {
        parts = {
            tnps: { value: tnps, weight: weights.tnps, target: targets.tnps },
            emailSla: { value: slaCompliance, weight: weights.emailSla, target: targets.emailSla },
            attendance: { value: attRate, weight: weights.attendance, target: targets.attendance },
            qa: { value: qaScore, weight: weights.qa, target: targets.qa },
            callRelease: { value: releaseRate, weight: weights.callRelease, target: targets.callRelease }
        };
        composite = Math.round(
            (tnps != null ? tnps : 0) * weights.tnps +
            slaCompliance * weights.emailSla +
            attRate * weights.attendance +
            (qaScore != null ? qaScore : 0) * weights.qa +
            releaseRate * weights.callRelease
        );
    } else {
        parts = {
            tnps: { value: tnps, weight: weights.tnps, target: targets.tnps },
            rnps: { value: avgRnps, weight: weights.rnps, target: targets.rnps },
            emailSla: { value: slaCompliance, weight: weights.emailSla, target: targets.emailSla },
            healthCheck: { value: hcPassRate, weight: weights.healthCheck, target: targets.healthCheck },
            qa: { value: qaScore, weight: weights.qa, target: targets.qa },
            attendance: { value: attRate, weight: weights.attendance, target: targets.attendance }
        };
        composite = Math.round(
            (tnps != null ? tnps : 0) * weights.tnps +
            (avgRnps != null ? avgRnps : 0) * weights.rnps +
            slaCompliance * weights.emailSla +
            (hcPassRate != null ? hcPassRate : 0) * weights.healthCheck +
            (qaScore != null ? qaScore : 0) * weights.qa +
            attRate * weights.attendance
        );
    }

    var qsrRows = agent.team === 'TSM_ME' ? ascData().qsrRecords.filter(function(q) { return q.agentId === agentId; }) : [];
    var qsrSubmitted = qsrRows.filter(function(q) { return q.status === 'Submitted'; }).length;
    var qsrAchieved = qsrRows.length ? Math.round((qsrSubmitted / qsrRows.length) * 100) : null;

    var ttRows = agent.team === 'TSM_ME' ? ascData().troubleTickets.filter(function(t) {
        return t.agentId === agentId && (!period || ascInPeriod(t.openedAt, period));
    }) : [];
    var ttClosed = ttRows.filter(function(t) { return t.status === 'Closed'; }).length;
    var ttAchieved = ttRows.length ? Math.round((ttClosed / ttRows.length) * 100) : null;

    var evRows = agent.team === 'TSM_ME' ? ascData().events.filter(function(e) {
        return e.agentId === agentId && (!period || ascInPeriod(e.eventDate, period));
    }) : [];

    var hcRows = agent.team === 'TSM_ME' ? ascData().healthChecks.filter(function(h) {
        return h.agentId === agentId && (!period || ascInPeriod(h.checkDate, period));
    }) : [];
    var hcPass = hcRows.filter(function(h) { return h.status === 'Pass'; }).length;
    var hcAchieved = hcRows.length ? Math.round((hcPass / hcRows.length) * 100) : hcPassRate;

    var rnpsRows = agent.team === 'TSM_ME' ? ascData().rnpsSurveys.filter(function(r) {
        return r.agentId === agentId && (!period || ascInPeriod(r.surveyDate, period));
    }) : [];
    var avgActivity = accountMetrics.length
        ? Math.round(accountMetrics.reduce(function(s, a) { return s + a.activityScore; }, 0) / accountMetrics.length)
        : null;
    var activityAchieved = avgActivity != null ? Math.min(100, Math.round(avgActivity * 2.5)) : null;

    var callsAchieved = tnps != null
        ? Math.round(Math.min(100, (tnps / targets.tnps) * 60 + (releaseRate / targets.callRelease) * 40))
        : Math.round(releaseRate / targets.callRelease * 100);

    var sections = {
        calls: {
            id: 'calls', label: 'Calls / TNPS', achieved: callsAchieved, target: targets.tnps,
            subs: { tnps: tnps, releaseRate: releaseRate, total: calls.length, inbound: calls.filter(function(c) { return c.direction === 'Inbound'; }).length, outbound: calls.filter(function(c) { return c.direction === 'Outbound'; }).length }
        },
        emails: {
            id: 'emails', label: 'Emails', achieved: slaCompliance, target: targets.emailSla,
            subs: { open: openEmails.length, closed: emails.filter(function(e) { return e.status === 'Closed'; }).length, slaWithin: slaWithin, slaOutside: openEmails.length - slaWithin, pending: pending.filter(function(p) { return p.recordType === 'Pending' && p.status === 'In Progress'; }).length, escalations: pending.filter(function(p) { return p.recordType === 'Escalation' && p.status === 'In Progress'; }).length }
        },
        qa: {
            id: 'qa', label: 'Quality Assurance', achieved: qaScore, target: targets.qa,
            subs: { reviews: qaRows.length, passRate: qaPassRate }
        },
        attendance: {
            id: 'attendance', label: 'Attendance', achieved: attRate, target: targets.attendance,
            subs: { weeks: att.length }
        }
    };

    if (agent.team === 'TSM_ME') {
        sections.rnps = { id: 'rnps', label: 'RNPS Survey', achieved: avgRnps, target: targets.rnps, subs: { surveys: rnpsRows.length, promoters: rnpsRows.filter(function(r) { return r.responseType === 'Promoter'; }).length } };
        sections.qsr = { id: 'qsr', label: 'QSR', achieved: qsrAchieved, target: 80, subs: { submitted: qsrSubmitted, pending: qsrRows.length - qsrSubmitted, total: qsrRows.length } };
        sections.tt = { id: 'tt', label: 'Trouble Tickets', achieved: ttAchieved, target: 85, subs: { total: ttRows.length, open: ttRows.filter(function(t) { return t.status !== 'Closed'; }).length, closed: ttClosed } };
        sections.events = { id: 'events', label: 'Events', achieved: evRows.length ? Math.min(100, evRows.length * 8) : 0, target: 70, subs: { total: evRows.length } };
        sections.healthcheck = { id: 'healthcheck', label: 'Health Check', achieved: hcAchieved, target: targets.healthCheck, subs: { total: hcRows.length, pass: hcPass, fail: hcRows.length - hcPass } };
        sections.activities = { id: 'activities', label: 'Activities', achieved: activityAchieved, target: 75, subs: { avgScore: avgActivity, accounts: accountMetrics.length } };
    }

    return {
        composite: composite,
        sections: sections,
        tnps: tnps,
        releaseRate: releaseRate,
        feedbackCount: fbCalls.length,
        totalCalls: calls.length,
        inbound: calls.filter(function(c) { return c.direction === 'Inbound'; }).length,
        outbound: calls.filter(function(c) { return c.direction === 'Outbound'; }).length,
        disaster: fbCalls.filter(function(c) { return c.feedback.type === 'Disaster'; }).length,
        passive: fbCalls.filter(function(c) { return c.feedback.type === 'Passive'; }).length,
        promotion: fbCalls.filter(function(c) { return c.feedback.type === 'Promotion'; }).length,
        totalEmails: emails.length,
        openEmails: openEmails.length,
        closedEmails: emails.filter(function(e) { return e.status === 'Closed'; }).length,
        slaCompliance: slaCompliance,
        slaWithin: slaWithin,
        slaOutside: openEmails.length - slaWithin,
        pendingOpen: pending.filter(function(p) { return p.recordType === 'Pending' && p.status === 'In Progress'; }).length,
        escalationsOpen: pending.filter(function(p) { return p.recordType === 'Escalation' && p.status === 'In Progress'; }).length,
        attRate: attRate,
        qaScore: qaScore,
        qaPassRate: qaPassRate,
        qaReviewCount: qaRows.length,
        avgRnps: avgRnps,
        accountCount: accounts.length,
        totalTT: accountMetrics.reduce(function(s, a) { return s + a.ttCount; }, 0),
        totalEvents: accountMetrics.reduce(function(s, a) { return s + a.eventCount; }, 0),
        hcPassRate: hcPassRate,
        scoreParts: parts,
        accountMetrics: accountMetrics
    };
}

function ascEnrichCall(row) {
    var acct = ascAccountById(row.accountId);
    var agent = ascAgentById(row.agentId);
    return Object.assign({}, row, {
        agentName: agent ? agent.name : '—',
        accountCode: acct ? acct.code : '—',
        customerName: acct ? acct.customerName : '—',
        feedbackType: row.feedback.type || '—',
        feedbackPoints: row.feedback.points != null ? row.feedback.points : '—'
    });
}

function ascEnrichEmail(row) {
    var acct = ascAccountById(row.accountId);
    var agent = ascAgentById(row.agentId);
    return Object.assign({}, row, {
        agentName: agent ? agent.name : '—',
        accountCode: acct ? acct.code : '—',
        customerName: acct ? acct.customerName : '—'
    });
}

function ascEnrichPending(row) {
    var acct = row.accountId ? ascAccountById(row.accountId) : null;
    var agent = ascAgentById(row.agentId);
    return Object.assign({}, row, {
        agentName: agent ? agent.name : '—',
        accountCode: acct ? acct.code : '—'
    });
}

/* ── FILTER PIPELINE ───────────────────────────────────────────── */
function ascApplyFilters() {
    var team = ASC.team;
    var period = ASC.period;
    var agents = ascData().agents.filter(function(a) { return a.team === team && a.active; });

    if (ASC.filterAgentIds.length) agents = agents.filter(function(a) { return ASC.filterAgentIds.indexOf(a.id) !== -1; });
    if (ASC.filterLocations.length) agents = agents.filter(function(a) { return ASC.filterLocations.indexOf(a.location) !== -1; });

    agents = agents.map(function(a) {
        return Object.assign({}, a, {
            lineManager: ascLmName(a.lineManagerId),
            metrics: ascComputeAgentMetrics(a.id, period)
        });
    });

    if (ASC.filterScoreBand === 'high') agents = agents.filter(function(a) { return a.metrics.composite >= 80; });
    else if (ASC.filterScoreBand === 'mid') agents = agents.filter(function(a) { return a.metrics.composite >= 70 && a.metrics.composite < 80; });
    else if (ASC.filterScoreBand === 'low') agents = agents.filter(function(a) { return a.metrics.composite < 70; });

    if (ASC.search) {
        var q = ASC.search.toLowerCase();
        agents = agents.filter(function(a) {
            return a.name.toLowerCase().indexOf(q) !== -1 || a.citrixId.toLowerCase().indexOf(q) !== -1;
        });
    }

    var tnpsTarget = ascTargets().tnps;
    if (ASC.filterKpi === 'below_target') agents = agents.filter(function(a) { return (a.metrics.tnps || 0) < tnpsTarget; });
    else if (ASC.filterKpi === 'sla_risk') agents = agents.filter(function(a) { return a.metrics.slaCompliance < ascTargets().emailSla; });
    else if (ASC.filterKpi === 'escalations') agents = agents.filter(function(a) { return a.metrics.escalationsOpen > 0; });

    agents.sort(function(a, b) { return b.metrics.composite - a.metrics.composite; });
    ASC.filteredAgents = agents;

    var agentIds = agents.map(function(a) { return a.id; });
    ASC.filteredCalls = ascQCalls({ team: team, period: period, agentIds: agentIds, direction: ASC.filterCallDir === 'all' ? null : ASC.filterCallDir });
    ASC.filteredEmails = ascQEmails({ team: team, period: period, agentIds: agentIds, status: ASC.filterEmailStatus || null, slaStatus: ASC.filterSla ? (ASC.filterSla === 'within' ? 'Within SLA' : 'Outside SLA') : null, openOnly: !!ASC.filterSla });
    ASC.filteredPending = ascQPending({ team: team, period: period, agentIds: agentIds });
}

function ascTeamKPIs() {
    var agents = ASC.filteredAgents;
    if (!agents.length) return { agentCount: 0 };
    var tnpsVals = agents.filter(function(a) { return a.metrics.tnps != null; });
    return {
        avgScore: Math.round(agents.reduce(function(s, a) { return s + a.metrics.composite; }, 0) / agents.length),
        avgTnps: tnpsVals.length ? Math.round(tnpsVals.reduce(function(s, a) { return s + a.metrics.tnps; }, 0) / tnpsVals.length) : null,
        belowTarget: agents.filter(function(a) { return (a.metrics.tnps || 0) < ascTargets().tnps; }).length,
        totalCalls: ASC.filteredCalls.length,
        totalEmails: ASC.filteredEmails.length,
        openEmails: ASC.filteredEmails.filter(function(e) { return e.status !== 'Closed'; }).length,
        escalations: ASC.filteredPending.filter(function(p) { return p.recordType === 'Escalation' && p.status === 'In Progress'; }).length,
        agentCount: agents.length
    };
}

function ascBuildInsights() {
    var k = ascTeamKPIs();
    var cards = [];
    var targets = ascTargets();
    if (k.belowTarget > 0) cards.push({ color: '#e74c3c', title: k.belowTarget + ' agents below TNPS target', sub: 'Target is ' + targets.tnps + '% — computed from call feedback records' });
    var top = ASC.filteredAgents[0];
    if (top) cards.push({ color: '#27ae60', title: top.name + ' leads with ' + top.metrics.composite + '%', sub: 'Highest composite from weighted KPI matrix' });
    if (k.escalations > 0) cards.push({ color: '#f39c12', title: k.escalations + ' open escalations', sub: 'From pendingItems where recordType=Escalation' });
    cards.push({ color: '#4c6fff', title: k.totalCalls.toLocaleString() + ' calls · ' + k.totalEmails.toLocaleString() + ' emails', sub: 'Period: ' + ASC.period + ' · ' + ascData().meta.counts.calls.toLocaleString() + ' total in dataset' });
    return cards.slice(0, 4);
}

/* ── RENDER SHELL ──────────────────────────────────────────────── */
function ascInjectShell() {
    var meta = ascData().meta;
    document.getElementById('asc-root').innerHTML =
        '<div id="asc-landing">' +
        '<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:.75rem;margin-bottom:1rem;">' +
        '<div><h2 class="section-title" style="margin:0!important;"><i data-lucide="layout-grid" style="width:20px;height:20px;"></i>Agent Score Card</h2>' +
        '<div style="font-size:.75rem;color:var(--t3);">Dataset v' + meta.version + ' · seed ' + meta.seed + ' · ' + meta.counts.calls.toLocaleString() + ' calls · ' + meta.counts.emails.toLocaleString() + ' emails · all metrics from records</div></div>' +
        '<div style="display:flex;gap:8px;flex-wrap:wrap;">' +
        '<button class="export-btn" onclick="AgentScoreCard.exportCSV()"><i data-lucide="download" style="width:14px;height:14px;"></i>Export CSV</button>' +
        '<button class="export-btn primary" onclick="AgentScoreCard.exportPDF()"><i data-lucide="file-down" style="width:14px;height:14px;"></i>Export PDF</button>' +
        '</div></div>' +
        '<div class="team-tabs" id="asc-team-tabs"></div>' +
        '<div id="asc-insights" class="insight-row"></div>' +
        '<div id="asc-filters"></div>' +
        '<div class="top-stats" id="asc-kpis"></div>' +
        '<div class="table-section"><div class="table-header"><h3 class="table-title"><i data-lucide="users" style="width:18px;height:18px;"></i>Agent Tiles</h3>' +
        '<div class="table-actions"><input type="text" class="search-box" id="asc-search" placeholder="Search agent..." oninput="AgentScoreCard.setSearch(this.value)">' +
        '<span class="record-count" id="asc-agent-count"></span></div></div>' +
        '<div class="agent-grid" id="asc-agent-grid"></div></div>' +
        '<div class="table-section"><div class="table-header"><h3 class="table-title"><i data-lucide="trophy" style="width:18px;height:18px;"></i>Leaderboard</h3></div>' +
        '<div id="asc-leaderboard"></div></div>' +
        '<div style="text-align:center;margin-bottom:1rem;"><button class="export-btn" id="asc-toggle-charts" onclick="AgentScoreCard.toggleCharts()">' +
        '<i data-lucide="bar-chart-2" style="width:14px;height:14px;"></i>Show Analytics Charts</button></div>' +
        '<div id="asc-charts-section" class="hidden"></div></div>' +
        '<div id="asc-drill" class="hidden"></div>';
}

function ascRenderFilters() {
    var teamAgents = ascData().agents.filter(function(a) { return a.team === ASC.team; });
    var names = teamAgents.map(function(a) { return ({ id: a.id, label: a.name }); }).sort(function(x, y) { return x.label.localeCompare(y.label); });
    var locs = [...new Set(teamAgents.map(function(a) { return a.location; }))].sort();
    var periods = [
        { v: 'today', l: 'Today' }, { v: 'week', l: 'This Week' }, { v: 'month', l: 'This Month' },
        { v: 'quarter', l: 'Quarter' }, { v: 'ytd', l: 'YTD' }, { v: 'year', l: 'Last 12 Months' }
    ];

    function dd(id, label, opts, selectedIds, fn) {
        var cnt = selectedIds.length;
        var btnLabel = cnt === 0 ? 'All ' + label : cnt + ' selected';
        return '<div style="flex:1;min-width:120px;"><div style="font-size:.62rem;font-weight:700;text-transform:uppercase;color:var(--t3);margin-bottom:4px;">' + label + '</div>' +
            '<div style="position:relative;"><div id="' + id + '-trigger" onclick="AgentScoreCard.toggleDD(\'' + id + '\')" style="display:flex;align-items:center;justify-content:space-between;padding:7px 10px;border:1px solid var(--border);border-radius:8px;background:var(--bg-input);cursor:pointer;font-size:12px;">' +
            '<span>' + btnLabel + '</span><i data-lucide="chevron-down" style="width:12px;height:12px;"></i></div>' +
            '<div id="' + id + '-dd" style="display:none;position:absolute;top:calc(100% + 4px);left:0;min-width:200px;z-index:9999;background:var(--bg-card);border:1px solid var(--border);border-radius:10px;box-shadow:0 8px 24px rgba(0,0,0,.18);max-height:240px;overflow-y:auto;">' +
            '<div style="padding:6px 10px;border-bottom:1px solid var(--border);"><label style="display:flex;align-items:center;gap:7px;font-size:12px;cursor:pointer;font-weight:700;">' +
            '<input type="checkbox" ' + (cnt === 0 ? 'checked' : '') + ' onchange="' + fn + '(\'__all__\',this)">All</label></div>' +
            opts.map(function(o) {
                var val = o.id || o, lbl = o.label || o;
                return '<div style="padding:5px 10px;"><label style="display:flex;align-items:center;gap:7px;font-size:12px;cursor:pointer;">' +
                    '<input type="checkbox" value="' + val + '" ' + (selectedIds.indexOf(val) !== -1 ? 'checked' : '') + ' onchange="' + fn + '(\'' + val + '\',this)">' + lbl + '</label></div>';
            }).join('') + '</div></div></div>';
    }

    document.getElementById('asc-filters').innerHTML =
        '<div style="display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap;padding:12px 16px;background:var(--bg-card);border:1px solid var(--border);border-radius:12px;box-shadow:var(--cs);margin-bottom:1rem;">' +
        '<div style="flex:1;min-width:120px;"><div style="font-size:.62rem;font-weight:700;text-transform:uppercase;color:var(--t3);margin-bottom:4px;">Period</div>' +
        '<select onchange="AgentScoreCard.setPeriod(this.value)" style="width:100%;padding:7px 10px;border:1px solid var(--border);border-radius:8px;background:var(--bg-input);font-size:12px;cursor:pointer;">' +
        periods.map(function(p) { return '<option value="' + p.v + '" ' + (ASC.period === p.v ? 'selected' : '') + '>' + p.l + '</option>'; }).join('') +
        '</select></div>' +
        dd('asc-f-agent', 'Agent', names, ASC.filterAgentIds, 'AgentScoreCard.filterAgent') +
        dd('asc-f-loc', 'Location', locs, ASC.filterLocations, 'AgentScoreCard.filterLoc') +
        '<div style="flex:1;min-width:120px;"><div style="font-size:.62rem;font-weight:700;text-transform:uppercase;color:var(--t3);margin-bottom:4px;">Score Band</div>' +
        '<select onchange="AgentScoreCard.setScoreBand(this.value)" style="width:100%;padding:7px 10px;border:1px solid var(--border);border-radius:8px;background:var(--bg-input);font-size:12px;cursor:pointer;">' +
        '<option value="">All Scores</option>' +
        ascCfg().scoreBands.map(function(b) {
            var v = b.id, lbl = b.label + (b.min ? ' (' + b.min + '+' : '') + (b.max ? ' ≤' + b.max : '') + (b.min || b.max ? ')' : '');
            return '<option value="' + v + '" ' + (ASC.filterScoreBand === v ? 'selected' : '') + '>' + lbl + '</option>';
        }).join('') +
        '</select></div>' +
        '<button class="reset-btn" onclick="AgentScoreCard.clearFilters()" style="align-self:flex-end;"><i data-lucide="rotate-ccw" style="width:11px;height:11px;"></i>Clear All</button></div>';
}

function ascRenderLanding() {
    ascRenderTeamTabs();
    document.getElementById('asc-insights').innerHTML = ascBuildInsights().map(function(c) {
        return '<div class="insight-card" style="border-left-color:' + c.color + ';"><div class="insight-title">' + c.title + '</div><div class="insight-sub">' + c.sub + '</div></div>';
    }).join('');
    ascRenderFilters();
    ascRenderKPIs();
    ascRenderAgentGrid();
    ascRenderLeaderboard();
    var btn = document.getElementById('asc-toggle-charts');
    var sec = document.getElementById('asc-charts-section');
    if (btn) btn.innerHTML = ASC.chartsVisible ? '<i data-lucide="bar-chart-2" style="width:14px;height:14px;"></i>Hide Analytics Charts' : '<i data-lucide="bar-chart-2" style="width:14px;height:14px;"></i>Show Analytics Charts';
    if (sec) { sec.classList.toggle('hidden', !ASC.chartsVisible); if (ASC.chartsVisible) ascRenderLandingCharts(); }
}

function ascRenderTeamTabs() {
    var se = ascData().agents.filter(function(a) { return a.team === 'TSM_SE'; }).length;
    var me = ascData().agents.filter(function(a) { return a.team === 'TSM_ME'; }).length;
    document.getElementById('asc-team-tabs').innerHTML =
        '<div class="team-tab ' + (ASC.team === 'TSM_SE' ? 'active' : '') + '" onclick="AgentScoreCard.setTeam(\'TSM_SE\')">TSM_SE <span class="tab-count">Undedicated · ' + se + ' agents</span></div>' +
        '<div class="team-tab ' + (ASC.team === 'TSM_ME' ? 'active' : '') + '" onclick="AgentScoreCard.setTeam(\'TSM_ME\')">TSM_ME <span class="tab-count">Dedicated · ' + me + ' agents</span></div>';
}

function ascRenderKPIs() {
    var k = ascTeamKPIs();
    var fk = ASC.filterKpi;
    var t = ascTargets();
    var tiles = [
        { key: 'score', label: 'Team Score', val: k.avgScore + '%', sub: 'Avg composite from records', color: 'var(--acc)', kpi: null },
        { key: 'tnps', label: 'TNPS', val: ascFmtPct(k.avgTnps), sub: 'Target ' + t.tnps + '%', color: '#4c6fff', kpi: 'below_target' },
        { key: 'calls', label: 'Calls', val: (k.totalCalls || 0).toLocaleString(), sub: 'Filtered call records', color: '#3498db', kpi: null },
        { key: 'emails', label: 'Open Emails', val: (k.openEmails || 0).toLocaleString(), sub: 'From email table', color: '#f39c12', kpi: 'sla_risk' },
        { key: 'esc', label: 'Escalations', val: k.escalations || 0, sub: 'Open escalation records', color: '#e74c3c', kpi: 'escalations' }
    ];
    document.getElementById('asc-kpis').innerHTML = tiles.map(function(tl) {
        return '<div class="stat-card ' + (tl.kpi ? 'clickable' : '') + ' ' + (fk === tl.kpi ? 'active-kpi' : '') + '" ' +
            (tl.kpi ? 'onclick="AgentScoreCard.toggleKpi(\'' + tl.kpi + '\')"' : '') + ' style="' + (fk === tl.kpi ? 'outline:2px solid ' + tl.color : '') + '">' +
            '<div class="stat-label">' + tl.label + '</div><div class="stat-value" style="color:' + tl.color + ';">' + tl.val + '</div>' +
            '<div class="stat-subtitle">' + tl.sub + '</div></div>';
    }).join('');
}

function ascSectionValClass(achieved, target) {
    if (achieved == null) return '';
    if (achieved >= target) return 'ok';
    if (achieved >= target - 10) return 'warn';
    return 'bad';
}

function ascSectionList(team) {
    var base = ['calls', 'emails', 'qa', 'attendance'];
    if (team === 'TSM_ME') return base.concat(['rnps', 'qsr', 'tt', 'events', 'healthcheck', 'activities']);
    return base;
}

function ascRenderAgentGrid() {
    var grid = document.getElementById('asc-agent-grid');
    document.getElementById('asc-agent-count').textContent = ASC.filteredAgents.length + ' agents';
    if (!grid) return;

    grid.innerHTML = ASC.filteredAgents.map(function(a) {
        var m = a.metrics, band = ascScoreBand(m.composite);
        var secIds = ascSectionList(a.team);
        var sectionRows = secIds.map(function(sid) {
            var sec = m.sections[sid];
            if (!sec) return '';
            var cls = ascSectionValClass(sec.achieved, sec.target);
            return '<div class="tile-section-row" onclick="event.stopPropagation(); AgentScoreCard.openDrill(\'' + a.id + '\',\'' + sid + '\')">' +
                '<span class="sec-name">' + sec.label + '</span>' +
                '<span class="sec-val ' + cls + '">' + (sec.achieved != null ? sec.achieved + '%' : '—') + '</span></div>';
        }).join('');

        return '<div class="agent-tile score-' + band + '" onclick="AgentScoreCard.openDrill(\'' + a.id + '\',\'overview\')">' +
            '<div class="agent-tile-head"><div><div class="agent-name">' + a.name + '</div><div class="agent-meta">' + a.location + ' · ' + a.citrixId + '</div></div></div>' +
            '<div class="score-achieved-label">Score Achieved</div>' +
            '<div class="score-pill ' + band + '" style="font-size:22px;display:inline-block;margin-bottom:4px;">' + m.composite + '%</div>' +
            '<div class="score-bar"><span style="width:' + m.composite + '%;background:' + (band === 'high' ? '#27ae60' : band === 'mid' ? '#f39c12' : '#e74c3c') + ';"></span></div>' +
            '<div class="tile-sections">' + sectionRows + '</div></div>';
    }).join('') || '<div style="padding:20px;color:var(--t3);">No agents match filters.</div>';
}

function ascRenderLeaderboard() {
    var cols = ASC.team === 'TSM_SE'
        ? ['Rank', 'Agent', 'Score', 'TNPS', 'Calls', 'Release %', 'Email SLA', 'QA', 'Attendance']
        : ['Rank', 'Agent', 'Score', 'TNPS', 'RNPS', 'Accounts', 'TT', 'Events', 'QSR Pending'];
    document.getElementById('asc-leaderboard').innerHTML = '<div class="table-wrapper"><table><thead><tr>' + cols.map(function(c) { return '<th>' + c + '</th>'; }).join('') + '</tr></thead><tbody>' +
        ASC.filteredAgents.map(function(a, i) {
            var m = a.metrics;
            var qsrPending = m.accountMetrics ? m.accountMetrics.filter(function(x) { return x.qsrStatus === 'Pending'; }).length : 0;
            var cells = ASC.team === 'TSM_SE'
                ? [i + 1, a.name, m.composite + '%', ascFmtPct(m.tnps), m.totalCalls, m.releaseRate + '%', m.slaCompliance + '%', ascFmtPct(m.qaScore), m.attRate + '%']
                : [i + 1, a.name, m.composite + '%', ascFmtPct(m.tnps), ascFmtPct(m.avgRnps), m.accountCount, m.totalTT, m.totalEvents, qsrPending];
            return '<tr style="cursor:pointer;" onclick="AgentScoreCard.openDrill(\'' + a.id + '\')">' + cells.map(function(c) { return '<td>' + c + '</td>'; }).join('') + '</tr>';
        }).join('') + '</tbody></table></div>';
}

/* ── DRILL DOWN ────────────────────────────────────────────────── */
function ascDrillTabs(agent) {
    return ['overview'].concat(ascSectionList(agent.team));
}
var ASC_TAB_LABELS = {
    overview: 'Overview',
    calls: 'Calls / TNPS',
    emails: 'Emails',
    qa: 'Quality Assurance',
    attendance: 'Attendance',
    rnps: 'RNPS Survey',
    qsr: 'QSR',
    tt: 'Trouble Tickets',
    events: 'Events',
    healthcheck: 'Health Check',
    activities: 'Activities'
};

function ascRenderDrill() {
    var agent = ascAgentById(ASC.drillAgentId);
    if (!agent) return;
    var m = ascComputeAgentMetrics(agent.id, ASC.period);
    var acct = ASC.drillAccountId ? ascAccountById(ASC.drillAccountId) : null;

    document.getElementById('asc-landing').classList.add('hidden');
    document.getElementById('asc-drill').classList.remove('hidden');

    var tabs = ascDrillTabs(agent);
    document.getElementById('asc-drill').innerHTML =
        '<div class="drill-header"><div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">' +
        '<button class="back-btn" onclick="AgentScoreCard.closeDrill()"><i data-lucide="arrow-left" style="width:14px;height:14px;"></i>Back</button>' +
        '<div><div style="font-size:18px;font-weight:900;">' + agent.name + '</div>' +
        '<div style="font-size:12px;color:var(--t3);">' + agent.team + ' · ' + agent.location + ' · ' + agent.citrixId + ' · LM: ' + ascLmName(agent.lineManagerId) +
        (acct ? ' · Account: <strong>' + acct.code + '</strong> <button class="reset-btn" style="padding:2px 8px;font-size:10px;margin-left:6px;" onclick="AgentScoreCard.clearAccountDrill()">Clear account filter</button>' : '') +
        '</div></div></div>' +
        '<div><div class="score-achieved-label">Score Achieved</div><div class="score-pill ' + ascScoreBand(m.composite) + '" style="font-size:24px;">' + m.composite + '%</div></div></div>' +
        '<div class="drill-tabs">' + tabs.map(function(t) {
            return '<div class="drill-tab ' + (ASC.drillTab === t ? 'active' : '') + '" onclick="AgentScoreCard.setDrillTab(\'' + t + '\')">' + ASC_TAB_LABELS[t] + '</div>';
        }).join('') + '</div><div id="asc-drill-content"></div>';

    ascRenderDrillContent(agent, m);
    ascIcons();
}

function ascRenderDrillContent(agent, m) {
    var el = document.getElementById('asc-drill-content');
    if (!el) return;
    var tab = ASC.drillTab;
    var period = ASC.period;
    var acctFilter = ASC.drillAccountId;

    if (tab === 'overview') {
        var secIds = ascSectionList(agent.team);
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;"><div class="table-title" style="margin-bottom:4px;">Score Achieved: <span style="color:var(--acc);font-size:1.4rem;">' + m.composite + '%</span></div>' +
            '<div style="font-size:12px;color:var(--t3);">Weighted composite from all sections below · click any section to drill down</div></div>' +
            '<div class="section-card-grid">' +
            secIds.map(function(sid) {
                var sec = m.sections[sid];
                var cls = ascSectionValClass(sec.achieved, sec.target);
                var subTxt = Object.keys(sec.subs || {}).slice(0, 3).map(function(k) { return k + ': ' + sec.subs[k]; }).join(' · ');
                return '<div class="section-card" onclick="AgentScoreCard.setDrillTab(\'' + sid + '\')">' +
                    '<div class="section-card-head"><div class="section-card-title">' + sec.label + '</div>' +
                    '<div class="section-card-achieved ' + cls + '" style="color:' + (cls === 'ok' ? '#27ae60' : cls === 'warn' ? '#f39c12' : '#e74c3c') + ';">' + (sec.achieved != null ? sec.achieved + '%' : '—') + '</div></div>' +
                    '<div class="section-card-sub">Target ' + sec.target + '% · ' + subTxt + '</div></div>';
            }).join('') +
            '</div>' +
            '<div class="table-section"><div class="table-title" style="margin-bottom:12px;">Score Weight Breakdown</div>' +
            Object.keys(m.scoreParts).map(function(key) {
                var p = m.scoreParts[key];
                var v = p.value != null ? Math.round(p.value) : 0;
                var ok = v >= p.target;
                return '<div style="margin-bottom:10px;"><div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px;"><span style="font-weight:700;">' + key + ' (' + Math.round(p.weight * 100) + '%)</span><span style="color:' + (ok ? '#27ae60' : '#e74c3c') + ';">' + (p.value != null ? v + '%' : '—') + ' · target ' + p.target + '%</span></div>' +
                    '<div style="height:8px;background:var(--bg-secondary);border-radius:99px;overflow:hidden;"><div style="width:' + Math.min(v, 100) + '%;height:100%;background:' + (ok ? '#27ae60' : '#e74c3c') + ';"></div></div></div>';
            }).join('') + '</div>';
    }

    else if (tab === 'calls') {
        var calls = ascQCalls({ agentId: agent.id, accountId: acctFilter, period: period, direction: ASC.filterCallDir === 'all' ? null : ASC.filterCallDir }).map(ascEnrichCall);
        var sec = m.sections.calls;
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;padding:12px 14px;"><div style="display:flex;justify-content:space-between;align-items:center;"><div><div class="score-achieved-label">Section Achieved</div><div style="font-size:1.6rem;font-weight:900;color:var(--acc);">' + sec.achieved + '%</div><div style="font-size:11px;color:var(--t3);">Calls / TNPS · target TNPS ' + ascTargets().tnps + '%</div></div></div></div>' +
            '<div class="sub-section-label">Call Volume</div>' +
            '<div class="top-stats" style="grid-template-columns:repeat(3,1fr);margin-bottom:1rem;">' +
            '<div class="stat-card clickable" onclick="AgentScoreCard.setCallDir(\'all\')"><div class="stat-label">Total Calls</div><div class="stat-value">' + m.totalCalls + '</div></div>' +
            '<div class="stat-card clickable" onclick="AgentScoreCard.setCallDir(\'Inbound\')"><div class="stat-label">Inbound</div><div class="stat-value" style="color:#4c6fff;">' + m.inbound + '</div></div>' +
            '<div class="stat-card clickable" onclick="AgentScoreCard.setCallDir(\'Outbound\')"><div class="stat-label">Outbound</div><div class="stat-value" style="color:#9b59b6;">' + m.outbound + '</div></div></div>' +
            '<div class="sub-section-label">TNPS & Call Release</div>' +
            '<div class="top-stats" style="grid-template-columns:repeat(4,1fr);margin-bottom:1rem;">' +
            '<div class="stat-card"><div class="stat-label">TNPS Achieved</div><div class="stat-value" style="color:' + ((m.tnps || 0) >= ascTargets().tnps ? '#27ae60' : '#e74c3c') + ';">' + ascFmtPct(m.tnps) + '</div><div class="stat-subtitle">Target ' + ascTargets().tnps + '%</div></div>' +
            '<div class="stat-card"><div class="stat-label">Call Release</div><div class="stat-value">' + m.releaseRate + '%</div><div class="stat-subtitle">' + m.feedbackCount + ' / ' + m.totalCalls + ' got feedback</div></div>' +
            '<div class="stat-card"><div class="stat-label">Disaster · 0</div><div class="stat-value" style="color:#e74c3c;">' + m.disaster + '</div></div>' +
            '<div class="stat-card"><div class="stat-label">Passive · 5 / Promo · 10</div><div class="stat-value" style="font-size:1rem;">' + m.passive + ' / ' + m.promotion + '</div></div></div>' +
            '<div class="sub-section-label">Trend & Record Drill-down</div>' +
            '<div class="chart-card"><div class="chart-title">Daily call volume</div><div style="height:220px;"><canvas id="asc-chart-calls"></canvas></div></div>' +
            '<div id="asc-call-detail" style="display:none;margin-bottom:1rem;" class="table-section"></div>' +
            '<div class="table-section"><div class="table-header"><h3 class="table-title">All call records</h3><span class="record-count">' + calls.length + ' rows · click row for detail</span></div>' +
            '<div id="asc-call-grid" class="ag-theme-alpine" style="height:420px;"></div></div>';
        ascRenderCallGrid(calls);
        ascRenderCallChart(calls);
    }

    else if (tab === 'emails') {
        var emailOpts = { agentId: agent.id, accountId: acctFilter, period: period };
        if (ASC.filterEmailStatus) emailOpts.status = ASC.filterEmailStatus;
        if (ASC.filterSla) { emailOpts.openOnly = true; emailOpts.slaStatus = ASC.filterSla === 'within' ? 'Within SLA' : 'Outside SLA'; }
        var emails = ascQEmails(emailOpts).map(ascEnrichEmail);
        var openRows = ascQEmails({ agentId: agent.id, accountId: acctFilter, period: period, openOnly: true });
        var within = openRows.filter(function(e) { return e.slaStatus === 'Within SLA'; }).length;
        var outside = openRows.length - within;
        var pendingItems = ascQPending({ agentId: agent.id, accountId: acctFilter, period: period }).map(ascEnrichPending);
        var secEm = m.sections.emails;
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;padding:12px 14px;"><div class="score-achieved-label">Section Achieved</div><div style="font-size:1.6rem;font-weight:900;color:var(--acc);">' + secEm.achieved + '%</div><div style="font-size:11px;color:var(--t3);">Emails SLA · target ' + ascTargets().emailSla + '%</div></div>' +
            '<div class="sub-section-label">Open / Closed & SLA</div>' +
            '<div class="sla-grid" style="margin-bottom:1rem;">' +
            '<div class="sla-tile" onclick="AgentScoreCard.toggleSla(\'within\')"><div style="font-size:.68rem;font-weight:700;color:var(--t3);">WITHIN SLA</div><div class="sla-tile-val" style="color:' + ASC_CLR.within + ';">' + within + '</div></div>' +
            '<div class="sla-tile" onclick="AgentScoreCard.toggleSla(\'outside\')"><div style="font-size:.68rem;font-weight:700;color:var(--t3);">OUTSIDE SLA</div><div class="sla-tile-val" style="color:' + ASC_CLR.outside + ';">' + outside + '</div></div>' +
            '<div class="sla-tile"><div style="font-size:.68rem;font-weight:700;color:var(--t3);">COMPLIANCE</div><div class="sla-tile-val">' + m.slaCompliance + '%</div></div>' +
            '<div class="sla-tile" onclick="AgentScoreCard.setEmailStatus(\'Open\')"><div style="font-size:.68rem;font-weight:700;color:var(--t3);">OPEN</div><div class="sla-tile-val">' + m.openEmails + '</div></div>' +
            '<div class="sla-tile" onclick="AgentScoreCard.setEmailStatus(\'Closed\')"><div style="font-size:.68rem;font-weight:700;color:var(--t3);">CLOSED</div><div class="sla-tile-val" style="color:#27ae60;">' + m.closedEmails + '</div></div></div>' +
            '<div class="sub-section-label">Categories & Email Drill-down</div>' +
            '<div class="chart-card"><div class="chart-title">Email categories</div><div style="height:200px;"><canvas id="asc-chart-email-cat"></canvas></div></div>' +
            '<div class="table-section"><div class="table-header"><h3 class="table-title">Email records</h3><span class="record-count">' + emails.length + ' rows</span></div><div id="asc-email-grid" class="ag-theme-alpine" style="height:340px;"></div></div>' +
            '<div class="sub-section-label">Pending Tracker & Escalation Tracker</div>' +
            '<div class="top-stats" style="grid-template-columns:repeat(4,1fr);margin-bottom:1rem;">' +
            '<div class="stat-card"><div class="stat-label">Pending Open</div><div class="stat-value" style="color:#f39c12;">' + secEm.subs.pending + '</div></div>' +
            '<div class="stat-card"><div class="stat-label">Escalations</div><div class="stat-value" style="color:#e74c3c;">' + secEm.subs.escalations + '</div></div>' +
            '<div class="stat-card"><div class="stat-label">Resolved</div><div class="stat-value" style="color:#27ae60;">' + pendingItems.filter(function(p) { return p.status === 'Resolved'; }).length + '</div></div>' +
            '<div class="stat-card"><div class="stat-label">High Priority</div><div class="stat-value">' + pendingItems.filter(function(p) { return p.priority === 'High'; }).length + '</div></div></div>' +
            '<div class="table-section"><div id="asc-pending-grid" class="ag-theme-alpine" style="height:300px;"></div></div>';
        ascRenderEmailGrid(emails, agent.team === 'TSM_ME');
        ascRenderEmailCatChart(emails);
        setTimeout(function() { ascRenderPendingGrid(pendingItems); }, 50);
    }

    else if (tab === 'attendance') {
        var att = ascQAttendance(agent.id).sort(function(a, b) { return a.weekKey.localeCompare(b.weekKey); });
        var secAtt = m.sections.attendance;
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;padding:12px 14px;"><div class="score-achieved-label">Section Achieved</div><div style="font-size:1.6rem;font-weight:900;color:var(--acc);">' + secAtt.achieved + '%</div><div style="font-size:11px;color:var(--t3);">Attendance vs Schedule · target ' + ascTargets().attendance + '%</div></div>' +
            '<div class="sub-section-label">Weekly Drill-down</div>' +
            '<div class="chart-card"><div class="chart-title">Scheduled vs attended</div><div style="height:260px;"><canvas id="asc-chart-att"></canvas></div></div>' +
            '<div class="table-section"><div class="table-wrapper"><table><thead><tr><th>Week</th><th>Scheduled</th><th>Attended</th><th>Rate</th></tr></thead><tbody>' +
            att.slice(-16).map(function(a) {
                return '<tr><td>' + a.weekKey + '</td><td>' + a.scheduledDays + '</td><td>' + a.attendedDays + '</td><td style="font-weight:700;color:' + (a.attendanceRate >= 90 ? '#27ae60' : '#e74c3c') + ';">' + a.attendanceRate + '%</td></tr>';
            }).join('') + '</tbody></table></div></div>';
        ascRenderAttChart(att.slice(-16));
    }

    else if (tab === 'qa') {
        var qaRows = ascQQa(agent.id, period);
        var secQa = m.sections.qa;
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;padding:12px 14px;"><div class="score-achieved-label">Section Achieved</div><div style="font-size:1.6rem;font-weight:900;color:var(--acc);">' + ascFmtPct(secQa.achieved) + '</div><div style="font-size:11px;color:var(--t3);">Quality Assurance · target ' + ascTargets().qa + '%</div></div>' +
            '<div class="sub-section-label">Review Drill-down</div>' +
            '<div class="table-section"><div class="table-wrapper"><table><thead><tr><th>Date</th><th>Score</th><th>Pass</th><th>Rubric</th><th>Reviewer</th></tr></thead><tbody>' +
            qaRows.map(function(q) {
                return '<tr><td>' + ascFmtDate(q.reviewDate) + '</td><td style="font-weight:700;">' + q.score + '%</td><td>' + (q.passed ? 'Yes' : 'No') + '</td><td>' + q.rubric + '</td><td>' + q.reviewer + '</td></tr>';
            }).join('') + '</tbody></table></div></div>';
    }

    else if (tab === 'rnps') {
        var rnpsRows = ascData().rnpsSurveys.filter(function(r) { return r.agentId === agent.id && ascInPeriod(r.surveyDate, period); }).map(function(r) {
            var acct = ascAccountById(r.accountId);
            return Object.assign({}, r, { accountCode: acct ? acct.code : '—', customerName: acct ? acct.customerName : '—' });
        });
        var sec = m.sections.rnps;
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;padding:12px 14px;"><div class="score-achieved-label">Section Achieved · RNPS Survey</div><div style="font-size:1.6rem;font-weight:900;color:var(--acc);">' + ascFmtPct(sec.achieved) + '</div><div style="font-size:11px;color:var(--t3);">Target ' + ascTargets().rnps + '% · ' + rnpsRows.length + ' survey records</div></div>' +
            '<div class="sub-section-label">Survey Drill-down by Account</div>' +
            '<div class="table-section"><div class="table-wrapper"><table><thead><tr><th>Date</th><th>Account</th><th>Customer</th><th>Score</th><th>Response</th></tr></thead><tbody>' +
            rnpsRows.map(function(r) {
                return '<tr style="cursor:pointer;" onclick="AgentScoreCard.drillAccount(\'' + r.accountId + '\')"><td>' + ascFmtDate(r.surveyDate) + '</td><td>' + r.accountCode + '</td><td>' + r.customerName + '</td><td style="font-weight:700;">' + r.score + '%</td><td>' + r.responseType + '</td></tr>';
            }).join('') + '</tbody></table></div></div>';
    }

    else if (tab === 'qsr') {
        var qsrRows = ascData().qsrRecords.filter(function(q) { return q.agentId === agent.id; }).map(function(q) {
            var acct = ascAccountById(q.accountId);
            return Object.assign({}, q, { accountCode: acct ? acct.code : '—', customerName: acct ? acct.customerName : '—' });
        });
        var secQsr = m.sections.qsr;
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;padding:12px 14px;"><div class="score-achieved-label">Section Achieved · QSR</div><div style="font-size:1.6rem;font-weight:900;color:var(--acc);">' + ascFmtPct(secQsr.achieved) + '</div></div>' +
            '<div class="sub-section-label">QSR by Account</div>' +
            '<div class="table-section"><div class="table-wrapper"><table><thead><tr><th>Account</th><th>Quarter</th><th>Status</th><th>Score</th></tr></thead><tbody>' +
            qsrRows.map(function(q) {
                return '<tr style="cursor:pointer;" onclick="AgentScoreCard.drillAccount(\'' + q.accountId + '\')"><td>' + q.accountCode + '</td><td>' + q.quarter + '</td><td>' + q.status + '</td><td>' + (q.score || '—') + '</td></tr>';
            }).join('') + '</tbody></table></div></div>';
    }

    else if (tab === 'tt') {
        var ttRows = ascData().troubleTickets.filter(function(t) { return t.agentId === agent.id && ascInPeriod(t.openedAt, period); }).map(function(t) {
            var acct = ascAccountById(t.accountId);
            return Object.assign({}, t, { accountCode: acct ? acct.code : '—' });
        });
        var secTt = m.sections.tt;
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;padding:12px 14px;"><div class="score-achieved-label">Section Achieved · TT</div><div style="font-size:1.6rem;font-weight:900;color:var(--acc);">' + ascFmtPct(secTt.achieved) + '</div><div style="font-size:11px;color:var(--t3);">' + secTt.subs.closed + ' closed / ' + secTt.subs.total + ' tickets</div></div>' +
            '<div class="sub-section-label">TT Drill-down</div>' +
            '<div class="table-section"><div class="table-wrapper"><table><thead><tr><th>Opened</th><th>Account</th><th>Category</th><th>Status</th><th>Title</th></tr></thead><tbody>' +
            ttRows.map(function(t) {
                return '<tr style="cursor:pointer;" onclick="AgentScoreCard.drillAccount(\'' + t.accountId + '\')"><td>' + ascFmtDate(t.openedAt) + '</td><td>' + t.accountCode + '</td><td>' + t.category + '</td><td>' + t.status + '</td><td>' + t.title + '</td></tr>';
            }).join('') + '</tbody></table></div></div>';
    }

    else if (tab === 'events') {
        var evRows = ascData().events.filter(function(e) { return e.agentId === agent.id && ascInPeriod(e.eventDate, period); }).map(function(e) {
            var acct = ascAccountById(e.accountId);
            return Object.assign({}, e, { accountCode: acct ? acct.code : '—' });
        });
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;padding:12px 14px;"><div class="score-achieved-label">Events on Accounts</div><div style="font-size:1.6rem;font-weight:900;color:var(--acc);">' + evRows.length + '</div></div>' +
            '<div class="sub-section-label">Events Drill-down</div>' +
            '<div class="table-section"><div class="table-wrapper"><table><thead><tr><th>Date</th><th>Account</th><th>Type</th><th>Category</th></tr></thead><tbody>' +
            evRows.map(function(e) {
                return '<tr style="cursor:pointer;" onclick="AgentScoreCard.drillAccount(\'' + e.accountId + '\')"><td>' + ascFmtDate(e.eventDate) + '</td><td>' + e.accountCode + '</td><td>' + e.eventType + '</td><td>' + e.category + '</td></tr>';
            }).join('') + '</tbody></table></div></div>';
    }

    else if (tab === 'healthcheck') {
        var hcRows = ascData().healthChecks.filter(function(h) { return h.agentId === agent.id && ascInPeriod(h.checkDate, period); }).map(function(h) {
            var acct = ascAccountById(h.accountId);
            return Object.assign({}, h, { accountCode: acct ? acct.code : '—' });
        });
        var secHc = m.sections.healthcheck;
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;padding:12px 14px;"><div class="score-achieved-label">Section Achieved · Health Check</div><div style="font-size:1.6rem;font-weight:900;color:var(--acc);">' + ascFmtPct(secHc.achieved) + '</div></div>' +
            '<div class="sub-section-label">Health Check Drill-down</div>' +
            '<div class="table-section"><div class="table-wrapper"><table><thead><tr><th>Date</th><th>Account</th><th>Status</th><th>Score</th></tr></thead><tbody>' +
            hcRows.map(function(h) {
                return '<tr style="cursor:pointer;" onclick="AgentScoreCard.drillAccount(\'' + h.accountId + '\')"><td>' + ascFmtDate(h.checkDate) + '</td><td>' + h.accountCode + '</td><td style="color:' + (h.status === 'Pass' ? '#27ae60' : '#e74c3c') + ';">' + h.status + '</td><td>' + h.score + '</td></tr>';
            }).join('') + '</tbody></table></div></div>';
    }

    else if (tab === 'activities') {
        var actRows = m.accountMetrics.map(function(am) {
            return { accountId: am.account.id, code: am.account.code, customerName: am.account.customerName, activityScore: am.activityScore, engagement: am.engagement, ttCount: am.ttCount, eventCount: am.eventCount };
        });
        var secAct = m.sections.activities;
        el.innerHTML =
            '<div class="table-section" style="margin-bottom:1rem;padding:12px 14px;"><div class="score-achieved-label">Section Achieved · Activities</div><div style="font-size:1.6rem;font-weight:900;color:var(--acc);">' + ascFmtPct(secAct.achieved) + '</div></div>' +
            '<div class="chart-card"><div class="chart-title">Activity by account</div><div style="height:240px;"><canvas id="asc-chart-acct"></canvas></div></div>' +
            '<div id="asc-acct-detail" style="display:none;margin-bottom:1rem;" class="table-section"></div>' +
            '<div class="table-section"><div id="asc-acct-grid" class="ag-theme-alpine" style="height:400px;"></div></div>';
        ascRenderAcctGrid(actRows);
        ascRenderAcctChart(actRows);
    }

    ascIcons();
}

/* ── GRIDS ─────────────────────────────────────────────────────── */
function ascDestroyGrid() {
    if (ASC.detailGridApi) { ASC.detailGridApi.destroy(); ASC.detailGridApi = null; }
}

function ascGridOpts(cols, rows, id, onRowClick) {
    var el = document.getElementById(id);
    if (!el) return;
    ascDestroyGrid();
    ASC.detailGridApi = agGrid.createGrid(el, {
        columnDefs: cols,
        rowData: rows,
        defaultColDef: { sortable: true, filter: true, resizable: true, flex: 1, minWidth: 90 },
        pagination: true,
        paginationPageSize: 25,
        paginationPageSizeSelector: [25, 50, 100],
        onRowClicked: onRowClick || null
    });
}

function ascRenderCallGrid(rows) {
    ascGridOpts([
        { field: 'id', headerName: 'ID', width: 100 },
        { field: 'datetime', headerName: 'Date', valueFormatter: function(p) { return ascFmtDate(p.value); } },
        { field: 'direction', headerName: 'Direction' },
        { field: 'category', headerName: 'Category' },
        { field: 'accountCode', headerName: 'Account' },
        { field: 'durationMinutes', headerName: 'Min', width: 70 },
        { field: 'feedbackType', headerName: 'Feedback' },
        { field: 'feedbackPoints', headerName: 'Pts', width: 60 },
        { field: 'disconnectedBy', headerName: 'Cut By' }
    ], rows, 'asc-call-grid', function(e) {
        ASC.selectedCallId = e.data.id;
        var c = e.data;
        var panel = document.getElementById('asc-call-detail');
        if (!panel) return;
        panel.style.display = 'block';
        panel.innerHTML = '<div class="table-title" style="margin-bottom:8px;">Call detail · ' + c.id + '</div>' +
            '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;font-size:12px;">' +
            ['Customer: ' + c.customerName, 'Phone: ' + c.customerPhone, 'Notes: ' + c.notes, 'Agent: ' + c.agentName, 'Account: ' + c.accountCode, 'Feedback: ' + c.feedbackType + ' (' + c.feedbackPoints + ' pts)'].map(function(x) {
                return '<div style="padding:8px;background:var(--bg-input);border-radius:8px;">' + x + '</div>';
            }).join('') + '</div>';
    });
}

function ascRenderEmailGrid(rows, showAcct) {
    var cols = [
        { field: 'id', headerName: 'ID', width: 100 },
        { field: 'receivedAt', headerName: 'Received', valueFormatter: function(p) { return ascFmtDate(p.value); } },
        { field: 'subject', headerName: 'Subject', minWidth: 180 },
        { field: 'status', headerName: 'Status' },
        { field: 'category', headerName: 'Category' },
        { field: 'slaStatus', headerName: 'SLA' },
        { field: 'senderType', headerName: 'Sender' },
        { field: 'ageDays', headerName: 'Age', width: 70 }
    ];
    if (showAcct) cols.push({ field: 'accountCode', headerName: 'Account' });
    ascGridOpts(cols, rows, 'asc-email-grid');
}

function ascRenderPendingGrid(rows) {
    ascGridOpts([
        { field: 'id', headerName: 'ID', width: 100 },
        { field: 'createdAt', headerName: 'Date', valueFormatter: function(p) { return ascFmtDate(p.value); } },
        { field: 'recordType', headerName: 'Type' },
        { field: 'issueType', headerName: 'Issue' },
        { field: 'status', headerName: 'Status' },
        { field: 'pendingWith', headerName: 'Pending With' },
        { field: 'priority', headerName: 'Priority' },
        { field: 'accountCode', headerName: 'Account' },
        { field: 'summary', headerName: 'Summary', minWidth: 160 }
    ], rows, 'asc-pending-grid');
}

function ascRenderAcctGrid(rows) {
    ascGridOpts([
        { field: 'code', headerName: 'Account' },
        { field: 'customerName', headerName: 'Customer' },
        { field: 'ttCount', headerName: 'TT' },
        { field: 'openTT', headerName: 'Open TT' },
        { field: 'eventCount', headerName: 'Events' },
        { field: 'healthCheckStatus', headerName: 'Health Check' },
        { field: 'qsrStatus', headerName: 'QSR' },
        { field: 'engagement', headerName: 'Engagement' },
        { field: 'avgRnps', headerName: 'RNPS', valueFormatter: function(p) { return p.value != null ? Math.round(p.value) + '%' : '—'; } },
        { field: 'activityScore', headerName: 'Activity' }
    ], rows, 'asc-acct-grid', function(e) {
        ASC.drillAccountId = e.data.accountId;
        var am = ascQAccountMetrics(e.data.accountId, ASC.period);
        var panel = document.getElementById('asc-acct-detail');
        if (panel) {
            panel.style.display = 'block';
            panel.innerHTML = '<div class="table-title" style="margin-bottom:8px;">Account drill · ' + e.data.code + ' · ' + e.data.customerName + '</div>' +
                '<div style="font-size:12px;color:var(--t2);margin-bottom:10px;">Filtering calls/emails to this account. TT: ' + am.tt.length + ' records · Events: ' + am.events.length + ' · RNPS surveys: ' + am.rnps.length + '</div>' +
                '<button class="export-btn" onclick="AgentScoreCard.setDrillTab(\'emails\')">View emails for this account</button> ' +
                '<button class="export-btn" onclick="AgentScoreCard.setDrillTab(\'calls\')">View calls for this account</button>';
        }
        ascToast('Account filter set: ' + e.data.code);
    });
}

/* ── CHARTS (all from filtered rows) ───────────────────────────── */
function ascDestroyChart(key) {
    if (ASC.charts[key]) { ASC.charts[key].destroy(); delete ASC.charts[key]; }
}

function ascRenderCallChart(rows) {
    ascDestroyChart('calls');
    var canvas = document.getElementById('asc-chart-calls');
    if (!canvas) return;
    var buckets = {};
    rows.forEach(function(c) {
        var day = c.datetime.substring(0, 10);
        buckets[day] = (buckets[day] || 0) + 1;
    });
    var labels = Object.keys(buckets).sort().slice(-14);
    ASC.charts.calls = new Chart(canvas, {
        type: 'bar',
        data: { labels: labels.map(function(l) { return ascFmtDate(l); }), datasets: [{ label: 'Calls', data: labels.map(function(l) { return buckets[l]; }), backgroundColor: ASC_CLR.acc + '99' }] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });
}

function ascRenderEmailCatChart(rows) {
    ascDestroyChart('emailCat');
    var canvas = document.getElementById('asc-chart-email-cat');
    if (!canvas) return;
    var cats = {};
    rows.forEach(function(e) { cats[e.category] = (cats[e.category] || 0) + 1; });
    ASC.charts.emailCat = new Chart(canvas, {
        type: 'doughnut',
        data: { labels: Object.keys(cats), datasets: [{ data: Object.values(cats), backgroundColor: ASC_CLR.smColors }] },
        options: { responsive: true, maintainAspectRatio: false }
    });
}

function ascRenderAttChart(rows) {
    ascDestroyChart('att');
    var canvas = document.getElementById('asc-chart-att');
    if (!canvas) return;
    ASC.charts.att = new Chart(canvas, {
        type: 'line',
        data: {
            labels: rows.map(function(a) { return a.weekKey; }),
            datasets: [
                { label: 'Scheduled', data: rows.map(function(a) { return a.scheduledDays; }), borderColor: ASC_CLR.blue, tension: 0.3 },
                { label: 'Attended', data: rows.map(function(a) { return a.attendedDays; }), borderColor: ASC_CLR.promo, tension: 0.3 }
            ]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });
}

function ascRenderAcctChart(rows) {
    ascDestroyChart('acct');
    var canvas = document.getElementById('asc-chart-acct');
    if (!canvas) return;
    ASC.charts.acct = new Chart(canvas, {
        type: 'bar',
        data: {
            labels: rows.map(function(a) { return a.code; }),
            datasets: [
                { label: 'TT', data: rows.map(function(a) { return a.ttCount; }), backgroundColor: ASC_CLR.disaster + '99' },
                { label: 'Events', data: rows.map(function(a) { return a.eventCount; }), backgroundColor: ASC_CLR.passive + '99' },
                { label: 'Activity', data: rows.map(function(a) { return a.activityScore; }), backgroundColor: ASC_CLR.acc + '99' }
            ]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });
}

function ascRenderLandingCharts() {
    var section = document.getElementById('asc-charts-section');
    if (!section) return;
    section.innerHTML =
        '<div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1rem;">' +
        '<div class="chart-card"><div class="chart-title">TNPS distribution (from agent metrics)</div><div style="height:240px;"><canvas id="asc-chart-tnps-dist"></canvas></div></div>' +
        '<div class="chart-card"><div class="chart-title">TNPS vs Composite (each dot = 1 agent)</div><div style="height:240px;"><canvas id="asc-chart-scatter"></canvas></div></div></div>' +
        '<div class="chart-card"><div class="chart-title">Call volume by agent (top 10 from filtered calls)</div><div style="height:260px;"><canvas id="asc-chart-team-calls"></canvas></div></div>';

    ascDestroyChart('tnpsDist');
    var tnpsCanvas = document.getElementById('asc-chart-tnps-dist');
    if (tnpsCanvas) {
        var buckets = { 'Below 60': 0, '60–69': 0, '70–79': 0, '80+': 0 };
        ASC.filteredAgents.forEach(function(a) {
            var t = a.metrics.tnps || 0;
            if (t < 60) buckets['Below 60']++;
            else if (t < 70) buckets['60–69']++;
            else if (t < 80) buckets['70–79']++;
            else buckets['80+']++;
        });
        ASC.charts.tnpsDist = new Chart(tnpsCanvas, {
            type: 'bar',
            data: { labels: Object.keys(buckets), datasets: [{ data: Object.values(buckets), backgroundColor: [ASC_CLR.disaster, ASC_CLR.passive, ASC_CLR.blue, ASC_CLR.promo] }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });
    }

    ascDestroyChart('scatter');
    var scCanvas = document.getElementById('asc-chart-scatter');
    if (scCanvas) {
        ASC.charts.scatter = new Chart(scCanvas, {
            type: 'scatter',
            data: { datasets: [{ label: 'Agents', data: ASC.filteredAgents.map(function(a) { return { x: a.metrics.tnps || 0, y: a.metrics.composite }; }), backgroundColor: ASC_CLR.acc + 'aa' }] },
            options: { responsive: true, maintainAspectRatio: false, scales: { x: { title: { display: true, text: 'TNPS %' } }, y: { title: { display: true, text: 'Composite %' } } } }
        });
    }

    ascDestroyChart('teamCalls');
    var tcCanvas = document.getElementById('asc-chart-team-calls');
    if (tcCanvas) {
        var byAgent = {};
        ASC.filteredCalls.forEach(function(c) {
            var name = ascAgentById(c.agentId).name;
            byAgent[name] = (byAgent[name] || 0) + 1;
        });
        var sorted = Object.entries(byAgent).sort(function(a, b) { return b[1] - a[1]; }).slice(0, 10);
        ASC.charts.teamCalls = new Chart(tcCanvas, {
            type: 'bar',
            data: { labels: sorted.map(function(s) { return s[0].split(' ')[0]; }), datasets: [{ label: 'Calls', data: sorted.map(function(s) { return s[1]; }), backgroundColor: ASC_CLR.blue + '99' }] },
            options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        });
    }
}

function ascRenderAll() {
    ascApplyFilters();
    if (ASC.view === 'landing') {
        document.getElementById('asc-landing').classList.remove('hidden');
        document.getElementById('asc-drill').classList.add('hidden');
        ascRenderLanding();
    } else {
        ascRenderDrill();
    }
    ascIcons();
}

/* ── PUBLIC API ────────────────────────────────────────────────── */
var AgentScoreCard = {
    init: function() {
        if (!window.ASC_DATA) {
            document.getElementById('asc-root').innerHTML = '<div style="padding:20px;color:#e74c3c;font-weight:700;">Missing agent-scorecard-data.js — load it before agent-scorecard.js</div>';
            return;
        }
        ascInjectShell();
        document.addEventListener('click', function(e) {
            if (!e.target.closest('[id$="-trigger"]') && !e.target.closest('[id$="-dd"]'))
                document.querySelectorAll('[id$="-dd"]').forEach(function(el) { el.style.display = 'none'; });
        });
        ascApplyFilters();
        ascRenderAll();
        console.log('[AgentScoreCard] Dataset loaded:', ascData().meta.counts);
    },

    setTeam: function(team) { ASC.team = team; ASC.filterAgentIds = []; ASC.filterKpi = null; ASC.view = 'landing'; ascRenderAll(); },
    setPeriod: function(p) { ASC.period = p; ascRenderAll(); },
    setSearch: function(v) { ASC.search = v; ascRenderAll(); },
    setScoreBand: function(v) { ASC.filterScoreBand = v; ascRenderAll(); },

    toggleDD: function(id) {
        var dd = document.getElementById(id + '-dd');
        if (!dd) return;
        document.querySelectorAll('[id$="-dd"]').forEach(function(el) { if (el.id !== id + '-dd') el.style.display = 'none'; });
        dd.style.display = dd.style.display === 'none' ? 'block' : 'none';
    },

    _setArr: function(arr, val, cb) {
        if (val === '__all__') arr.length = 0;
        else if (cb.checked) { if (arr.indexOf(val) === -1) arr.push(val); }
        else { var i = arr.indexOf(val); if (i > -1) arr.splice(i, 1); }
        ascRenderAll();
    },
    filterAgent: function(v, cb) { AgentScoreCard._setArr(ASC.filterAgentIds, v, cb); },
    filterLoc: function(v, cb) { AgentScoreCard._setArr(ASC.filterLocations, v, cb); },

    clearFilters: function() {
        ASC.filterAgentIds = []; ASC.filterLocations = []; ASC.filterScoreBand = '';
        ASC.filterKpi = null; ASC.filterSla = null; ASC.filterCallDir = 'all';
        ASC.filterEmailStatus = null; ASC.search = '';
        var s = document.getElementById('asc-search'); if (s) s.value = '';
        ascRenderAll();
    },

    toggleKpi: function(k) { ASC.filterKpi = ASC.filterKpi === k ? null : k; ascRenderAll(); },
    toggleCharts: function() { ASC.chartsVisible = !ASC.chartsVisible; ascRenderAll(); },

    openDrill: function(id, tab) { ASC.drillAgentId = id; ASC.drillAccountId = null; ASC.drillTab = tab || 'overview'; ASC.view = 'drill'; ascDestroyGrid(); ascRenderAll(); },
    drillAccount: function(accountId) { ASC.drillAccountId = accountId; ascToast('Account filter: ' + (ascAccountById(accountId) || {}).code); ascRenderDrill(); },
    closeDrill: function() { ASC.view = 'landing'; ASC.drillAgentId = null; ASC.drillAccountId = null; ascDestroyGrid(); ascRenderAll(); },
    setDrillTab: function(t) { ASC.drillTab = t; ascDestroyGrid(); ascRenderDrill(); },
    clearAccountDrill: function() { ASC.drillAccountId = null; ascRenderDrill(); },

    setCallDir: function(d) { ASC.filterCallDir = d; ascRenderDrill(); },
    toggleSla: function(s) { ASC.filterSla = ASC.filterSla === s ? null : s; ascRenderDrill(); },
    setEmailStatus: function(st) { ASC.filterEmailStatus = ASC.filterEmailStatus === st ? null : st; ascRenderDrill(); },

    exportCSV: function() {
        var rows = ASC.filteredAgents.map(function(a) {
            return { Agent: a.name, Team: a.team, Score: a.metrics.composite, TNPS: a.metrics.tnps, Calls: a.metrics.totalCalls, Emails: a.metrics.totalEmails, EmailSLA: a.metrics.slaCompliance };
        });
        var csv = ['Agent,Team,Score,TNPS,Calls,Emails,EmailSLA'].concat(rows.map(function(r) { return [r.Agent, r.Team, r.Score, r.TNPS, r.Calls, r.Emails, r.EmailSLA].join(','); })).join('\n');
        var blob = new Blob([csv], { type: 'text/csv' });
        var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'agent-scorecard.csv'; a.click();
        ascToast('Exported ' + rows.length + ' agents from computed metrics');
    },
    exportPDF: function() { ascToast('PDF export placeholder — data source: ASC_DATA'); }
};

window.AgentScoreCard = AgentScoreCard;
