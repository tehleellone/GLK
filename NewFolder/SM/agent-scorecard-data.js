/* ═══════════════════════════════════════════════════════════════
   AGENT SCORE CARD — Relational Dummy Dataset v2.0
   Every KPI, tile, grid row, and chart series is derived from these
   tables via foreign keys (agentId, accountId). Nothing is display-only.
   ═══════════════════════════════════════════════════════════════ */
'use strict';

(function buildAscDataset() {
    function mulberry32(a) {
        return function() {
            a |= 0; a = a + 0x6D2B79F5 | 0;
            var t = Math.imul(a ^ a >>> 15, 1 | a);
            t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
            return ((t ^ t >>> 14) >>> 0) / 4294967296;
        };
    }

    var SEED = 20260806;
    var rng = mulberry32(SEED);
    function rand(min, max) { return Math.floor(rng() * (max - min + 1)) + min; }
    function pick(arr) { return arr[rand(0, arr.length - 1)]; }
    function pickW(arr, weights) {
        var t = weights.reduce(function(a, b) { return a + b; }, 0);
        var r = rng() * t, s = 0;
        for (var i = 0; i < arr.length; i++) { s += weights[i]; if (r <= s) return arr[i]; }
        return arr[arr.length - 1];
    }
    function pad(n, w) { return String(n).padStart(w || 2, '0'); }
    function isoDaysAgo(days, hour) {
        var d = new Date();
        d.setDate(d.getDate() - days);
        d.setHours(hour != null ? hour : rand(8, 18), rand(0, 59), 0, 0);
        return d.toISOString();
    }
    function weekKey(daysAgo) {
        var d = new Date();
        d.setDate(d.getDate() - daysAgo);
        var one = new Date(d.getFullYear(), 0, 1);
        var wk = Math.ceil(((d - one) / 86400000 + one.getDay() + 1) / 7);
        return d.getFullYear() + '-W' + pad(wk);
    }

    var CONFIG = {
        targets: {
            tnps: 70,
            rnps: 75,
            emailSla: 90,
            attendance: 95,
            qa: 85,
            callRelease: 60,
            healthCheck: 80
        },
        scoreWeights: {
            TSM_SE: { tnps: 0.35, emailSla: 0.25, attendance: 0.15, qa: 0.15, callRelease: 0.10 },
            TSM_ME: { tnps: 0.25, rnps: 0.25, emailSla: 0.20, healthCheck: 0.10, qa: 0.10, attendance: 0.10 }
        },
        scoreBands: [
            { id: 'high', label: 'Excellent', min: 80 },
            { id: 'mid', label: 'Good', min: 70, max: 79 },
            { id: 'low', label: 'Needs Attention', max: 69 }
        ],
        feedbackScores: {
            Disaster: 0,
            Passive: 5,
            Promotion: 10
        },
        callCategories: ['Billing Inquiry', 'Technical Support', 'Service Request', 'Complaint', 'Provisioning', 'Outage', 'General', 'Escalation'],
        emailCategories: ['Billing', 'Technical', 'Service', 'Documentation', 'Escalation', 'Provisioning', 'Complaint'],
        issueTypes: ['Billing', 'Technical', 'Service', 'Escalation', 'Documentation', 'Other'],
        eventTypes: ['Outage', 'Maintenance', 'Review Meeting', 'QBR', 'Health Review', 'Escalation Review'],
        ttCategories: ['Connectivity', 'Billing', 'Hardware', 'Software', 'Configuration', 'Performance'],
        locations: ['Dubai', 'Abu Dhabi', 'Remote', 'Sharjah'],
        senderTypes: ['Customer', 'Internal', 'Vendor', 'AM', 'LM'],
        pendingWith: ['Agent', 'SM', 'LM', 'AM', 'Customer', 'PSD', 'Other'],
        priorities: ['Low', 'Medium', 'High'],
        engagementLevels: ['High', 'Medium', 'Low'],
        disconnectedBy: ['Customer', 'Agent', 'System', 'Unknown']
    };

    var lineManagers = [
        { id: 'LM-01', name: 'James Wilson' },
        { id: 'LM-02', name: 'Priya Sharma' },
        { id: 'LM-03', name: 'Michael Chen' }
    ];

    var seAgentDefs = [
        'Sarah Al-Mansoori', 'Ahmed Hassan', 'Fatima Khan', 'Omar Khalid', 'Layla Ibrahim',
        'Youssef Nasser', 'Mariam Said', 'Khalid Rahman', 'Noor Al-Zahra', 'Hassan Farouk',
        'Aisha Malik', 'Tariq Jamil', 'Dina Sorour', 'Rashid Al-Otaibi', 'Hana Youssef',
        'Samir Khoury', 'Leila Mansour', 'Karim Boutros', 'Nadia Farah', 'Zain Abidi'
    ];
    var meAgentDefs = [
        'Fatima Khan', 'Omar Khalid', 'Rashid Al-Otaibi', 'Samir Khoury', 'Leila Mansour',
        'Karim Boutros', 'Nadia Farah', 'Zain Abidi', 'Amira Haddad', 'Fadi Nassar',
        'Rana Sleiman', 'Bassel Hamdan', 'Maya Khoury', 'Wael Sabbagh', 'Tarek Mansour'
    ];

    var customerNames = [
        'ACME Corporation', 'Beta Industries', 'Gamma Telecom', 'Delta Holdings', 'Epsilon Group',
        'Zeta Solutions', 'Eta Systems', 'Theta Networks', 'Iota Services', 'Kappa Digital',
        'Lambda Corp', 'Mu Enterprises', 'Nu Partners', 'Xi Global', 'Omicron Ltd', 'Pi Ventures',
        'Rho Analytics', 'Sigma Retail', 'Tau Logistics', 'Upsilon Finance'
    ];

    var agents = [];
    seAgentDefs.forEach(function(name, i) {
        agents.push({
            id: 'AG-SE-' + pad(i + 1, 3),
            name: name,
            team: 'TSM_SE',
            location: pick(CONFIG.locations),
            citrixId: 'CNFX' + (1000 + i),
            lineManagerId: pick(lineManagers).id,
            active: true
        });
    });
    meAgentDefs.forEach(function(name, i) {
        agents.push({
            id: 'AG-ME-' + pad(i + 1, 3),
            name: name,
            team: 'TSM_ME',
            location: pick(CONFIG.locations),
            citrixId: 'CNFX' + (2000 + i),
            lineManagerId: pick(lineManagers).id,
            active: true
        });
    });

    var accounts = [];
    var acctSeq = 1;
    agents.filter(function(a) { return a.team === 'TSM_ME'; }).forEach(function(agent) {
        var count = rand(8, 14);
        for (var n = 0; n < count; n++) {
            var cust = pick(customerNames);
            accounts.push({
                id: 'ACC-' + pad(acctSeq++, 4),
                code: cust.substring(0, 4).toUpperCase() + '-' + rand(100, 999),
                customerName: cust,
                agentId: agent.id,
                team: 'TSM_ME',
                segment: pick(['Enterprise', 'Mid-Market', 'SMB']),
                region: pick(['UAE-N', 'UAE-S', 'GCC'])
            });
        }
    });

    var poolAccounts = [];
    for (var p = 0; p < 120; p++) {
        var pcust = pick(customerNames);
        poolAccounts.push({
            id: 'ACC-POOL-' + pad(p + 1, 4),
            code: 'POOL-' + rand(10000, 99999),
            customerName: pcust,
            agentId: null,
            team: 'TSM_SE',
            segment: 'Pool',
            region: pick(['UAE-N', 'UAE-S', 'GCC'])
        });
    }
    accounts = accounts.concat(poolAccounts);

    function accountForAgent(agent) {
        if (agent.team === 'TSM_ME') {
            var mine = accounts.filter(function(a) { return a.agentId === agent.id; });
            return pick(mine);
        }
        return pick(poolAccounts);
    }

    var calls = [];
    var emails = [];
    var pendingItems = [];
    var attendanceWeeks = [];
    var qaReviews = [];
    var troubleTickets = [];
    var events = [];
    var healthChecks = [];
    var qsrRecords = [];
    var rnpsSurveys = [];

    var callSeq = 1, emailSeq = 1, pendSeq = 1, ttSeq = 1, evSeq = 1, hcSeq = 1, qsrSeq = 1, rnpsSeq = 1, qaSeq = 1, attSeq = 1;

    agents.forEach(function(agent) {
        var callTarget = agent.team === 'TSM_SE' ? rand(200, 450) : rand(100, 240);
        for (var c = 0; c < callTarget; c++) {
            var daysAgo = rand(0, 364);
            var acct = accountForAgent(agent);
            var direction = rng() < 0.72 ? 'Inbound' : 'Outbound';
            var hasFeedback = rng() < 0.58;
            var fbType = hasFeedback ? pickW(['Disaster', 'Passive', 'Promotion'], [0.12, 0.38, 0.50]) : null;
            calls.push({
                id: 'CALL-' + pad(callSeq++, 5),
                agentId: agent.id,
                accountId: acct.id,
                datetime: isoDaysAgo(daysAgo),
                direction: direction,
                category: pick(CONFIG.callCategories),
                durationMinutes: rand(2, 45),
                customerPhone: '+9715' + rand(1000000, 9999999),
                feedback: {
                    submitted: hasFeedback,
                    type: fbType,
                    points: fbType ? CONFIG.feedbackScores[fbType] : null
                },
                disconnectedBy: pick(CONFIG.disconnectedBy),
                notes: direction + ' call regarding ' + pick(CONFIG.callCategories).toLowerCase()
            });
        }

        var emailTarget = agent.team === 'TSM_SE' ? rand(140, 320) : rand(70, 180);
        for (var e = 0; e < emailTarget; e++) {
            var eDays = rand(0, 364);
            var recvIso = isoDaysAgo(eDays, rand(7, 20));
            var recv = new Date(recvIso);
            var isClosed = rng() < 0.78;
            var slaHours = pick([24, 48, 72]);
            var resolveDays = isClosed ? rand(1, 10) : null;
            var closedIso = isClosed ? new Date(recv.getTime() + resolveDays * 86400000).toISOString() : null;
            var ageDays = isClosed ? resolveDays : Math.ceil((Date.now() - recv.getTime()) / 86400000);
            var withinSla = isClosed ? resolveDays <= slaHours / 24 : ageDays <= slaHours / 24;
            var eAcct = accountForAgent(agent);
            emails.push({
                id: 'EMAIL-' + pad(emailSeq++, 5),
                agentId: agent.id,
                accountId: eAcct.id,
                subject: pick(CONFIG.emailCategories) + ' request — ' + eAcct.code,
                receivedAt: recvIso,
                closedAt: closedIso,
                status: isClosed ? 'Closed' : (rng() < 0.12 ? 'Unassigned' : 'Open'),
                category: pick(CONFIG.emailCategories),
                senderType: pick(CONFIG.senderTypes),
                slaHours: slaHours,
                slaStatus: withinSla ? 'Within SLA' : 'Outside SLA',
                ageDays: ageDays,
                threadId: 'TH-' + rand(1000, 9999)
            });
        }

        var pendTarget = rand(10, 40);
        for (var pi = 0; pi < pendTarget; pi++) {
            var pAcct = agent.team === 'TSM_ME' ? accountForAgent(agent) : (rng() < 0.4 ? pick(poolAccounts) : null);
            var isEsc = rng() < 0.24;
            pendingItems.push({
                id: 'PEND-' + pad(pendSeq++, 5),
                agentId: agent.id,
                accountId: pAcct ? pAcct.id : null,
                createdAt: isoDaysAgo(rand(0, 120)),
                recordType: isEsc ? 'Escalation' : 'Pending',
                issueType: pick(CONFIG.issueTypes),
                status: pickW(['In Progress', 'Resolved'], [0.55, 0.45]),
                pendingWith: pick(CONFIG.pendingWith),
                priority: pickW(CONFIG.priorities, [0.45, 0.35, 0.20]),
                summary: pick(CONFIG.issueTypes) + ' follow-up for ' + (pAcct ? pAcct.code : 'shared queue')
            });
        }

        for (var w = 0; w < 52; w++) {
            var scheduled = rand(4, 5);
            var missed = rand(0, 2);
            var attended = Math.max(0, scheduled - missed);
            attendanceWeeks.push({
                id: 'ATT-' + pad(attSeq++, 5),
                agentId: agent.id,
                weekKey: weekKey(w * 7 + rand(0, 6)),
                scheduledDays: scheduled,
                attendedDays: attended,
                attendanceRate: Math.round((attended / scheduled) * 100)
            });
        }

        var reviewCount = rand(5, 14);
        for (var q = 0; q < reviewCount; q++) {
            var score = rand(62, 99);
            qaReviews.push({
                id: 'QA-' + pad(qaSeq++, 5),
                agentId: agent.id,
                reviewDate: isoDaysAgo(rand(0, 180)),
                score: score,
                passed: score >= 75,
                reviewer: pick(lineManagers).name,
                rubric: pick(['Call Handling', 'Email Quality', 'Documentation', 'Customer Empathy']),
                notes: 'Review session ' + (q + 1)
            });
        }

        if (agent.team === 'TSM_ME') {
            var agentAccounts = accounts.filter(function(a) { return a.agentId === agent.id; });
            agentAccounts.forEach(function(acct) {
                var ttCount = rand(0, 9);
                for (var t = 0; t < ttCount; t++) {
                    var opened = isoDaysAgo(rand(0, 200));
                    var closed = rng() < 0.7 ? new Date(new Date(opened).getTime() + rand(1, 14) * 86400000).toISOString() : null;
                    troubleTickets.push({
                        id: 'TT-' + pad(ttSeq++, 5),
                        accountId: acct.id,
                        agentId: agent.id,
                        openedAt: opened,
                        closedAt: closed,
                        category: pick(CONFIG.ttCategories),
                        status: closed ? 'Closed' : pick(['Open', 'In Progress']),
                        priority: pick(CONFIG.priorities),
                        title: pick(CONFIG.ttCategories) + ' issue on ' + acct.code
                    });
                }

                var evCount = rand(0, 6);
                for (var ev = 0; ev < evCount; ev++) {
                    events.push({
                        id: 'EVT-' + pad(evSeq++, 5),
                        accountId: acct.id,
                        agentId: agent.id,
                        eventDate: isoDaysAgo(rand(0, 180)),
                        eventType: pick(CONFIG.eventTypes),
                        category: pick(CONFIG.issueTypes),
                        description: pick(CONFIG.eventTypes) + ' with ' + acct.customerName
                    });
                }

                for (var h = 0; h < rand(2, 5); h++) {
                    healthChecks.push({
                        id: 'HC-' + pad(hcSeq++, 5),
                        accountId: acct.id,
                        agentId: agent.id,
                        checkDate: isoDaysAgo(rand(0, 90)),
                        status: rng() < 0.78 ? 'Pass' : 'Fail',
                        score: rand(55, 100),
                        notes: 'Scheduled health check'
                    });
                }

                qsrRecords.push({
                    id: 'QSR-' + pad(qsrSeq++, 5),
                    accountId: acct.id,
                    agentId: agent.id,
                    quarter: pick(['2025-Q1', '2025-Q2', '2025-Q3', '2025-Q4']),
                    status: rng() < 0.62 ? 'Submitted' : 'Pending',
                    submittedAt: rng() < 0.62 ? isoDaysAgo(rand(0, 120)) : null,
                    score: rand(70, 98)
                });

                var surveyCount = rand(1, 4);
                for (var s = 0; s < surveyCount; s++) {
                    rnpsSurveys.push({
                        id: 'RNPS-' + pad(rnpsSeq++, 5),
                        accountId: acct.id,
                        agentId: agent.id,
                        surveyDate: isoDaysAgo(rand(0, 365)),
                        score: rand(40, 98),
                        responseType: pickW(['Promoter', 'Passive', 'Detractor'], [0.48, 0.32, 0.20])
                    });
                }
            });
        }
    });

    window.ASC_DATA = {
        meta: {
            version: '2.0',
            seed: SEED,
            generatedAt: new Date().toISOString(),
            counts: {
                agents: agents.length,
                accounts: accounts.length,
                calls: calls.length,
                emails: emails.length,
                pendingItems: pendingItems.length,
                attendanceWeeks: attendanceWeeks.length,
                qaReviews: qaReviews.length,
                troubleTickets: troubleTickets.length,
                events: events.length,
                healthChecks: healthChecks.length,
                qsrRecords: qsrRecords.length,
                rnpsSurveys: rnpsSurveys.length
            }
        },
        config: CONFIG,
        lineManagers: lineManagers,
        agents: agents,
        accounts: accounts,
        calls: calls,
        emails: emails,
        pendingItems: pendingItems,
        attendanceWeeks: attendanceWeeks,
        qaReviews: qaReviews,
        troubleTickets: troubleTickets,
        events: events,
        healthChecks: healthChecks,
        qsrRecords: qsrRecords,
        rnpsSurveys: rnpsSurveys
    };
})();
