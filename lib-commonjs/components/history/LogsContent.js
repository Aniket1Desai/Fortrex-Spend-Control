"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
/**
 * LogsContent — accordion Approval + Audit log sections with no Panel wrapper.
 * Used by RequestLogsPanel (wraps in a Panel) and MyRequests detail view (inline).
 */
var React = tslib_1.__importStar(require("react"));
var Spinner_1 = require("@fluentui/react/lib/Spinner");
var MessageBar_1 = require("@fluentui/react/lib/MessageBar");
var Icon_1 = require("@fluentui/react/lib/Icon");
var DataverseService_1 = require("../../services/DataverseService");
var dashboardUtils_1 = require("../dashboard/dashboardUtils");
// ── Choice maps ───────────────────────────────────────────────────────────────
var APPROVAL_STATUS = {
    1: { label: 'Approved', color: '#107C10', bg: '#DFF6DD' },
    2: { label: 'Rejected', color: '#A4262C', bg: '#FDE7E9' },
    3: { label: 'Return for Info', color: '#905400', bg: '#FFF4CE' },
};
// sct_tiername Edm.Int32 confirmed from Dataverse — Accepted Values: 1=Tier 1, 2=Tier 2, 3=Tier 3, 4=Manager
var APPROVAL_TIER = { 1: 'Tier 1', 2: 'Tier 2', 3: 'Tier 3', 4: 'Manager' };
var AUDIT_ACTION = {
    1: { label: 'Submitted', color: '#004578', bg: '#DBEAFE' },
    2: { label: 'Edited', color: '#00897B', bg: '#D1FAE5' },
    3: { label: 'Draft', color: '#605E5C', bg: '#F3F2F1' },
    4: { label: 'Returned', color: '#905400', bg: '#FFF4CE' },
    5: { label: 'Rejected', color: '#A4262C', bg: '#FDE7E9' },
};
// ── Helpers ───────────────────────────────────────────────────────────────────
function fmtDateTime(record, fallback) {
    var _a, _b, _c;
    if (!record && !fallback)
        return '—';
    // CRITICAL: Prioritize raw 'createdon' (or 'sct_decisionon') EDM timestamp.
    // Never read @OData.Community.Display.V1.FormattedValue for dates because Dataverse
    // formats it on the server in UTC and in US MM/dd/yyyy format.
    var rawIso = record
        ? ((_c = (_b = (_a = record['createdon']) !== null && _a !== void 0 ? _a : record['sct_decisionon']) !== null && _b !== void 0 ? _b : fallback) !== null && _c !== void 0 ? _c : '')
        : (fallback !== null && fallback !== void 0 ? fallback : '');
    return (0, dashboardUtils_1.formatDateTime)(String(rawIso));
}
function resolveChoice(item, field, map) {
    var _a;
    var raw = item[field];
    if (typeof raw === 'number' && map[raw])
        return map[raw];
    if (typeof raw === 'string' && /^\d+$/.test(raw) && map[Number(raw)])
        return map[Number(raw)];
    var fmt = (0, DataverseService_1.strVal)(item, field);
    if (fmt)
        return fmt;
    if (typeof raw === 'number')
        return (_a = map[raw]) !== null && _a !== void 0 ? _a : String(raw);
    return '—';
}
// ── Micro-components ──────────────────────────────────────────────────────────
var Badge = function (_a) {
    var label = _a.label, color = _a.color, bg = _a.bg;
    return (React.createElement("span", { style: {
            display: 'inline-block', padding: '2px 10px', borderRadius: 10,
            fontSize: 11, fontWeight: 700, background: bg,
            color: color,
            whiteSpace: 'nowrap',
        } }, label));
};
var DetailRow = function (_a) {
    var label = _a.label, value = _a.value;
    return (React.createElement("div", { style: { display: 'flex', gap: 8, marginBottom: 6, alignItems: 'baseline', flexWrap: 'wrap' } },
        React.createElement("span", { style: {
                fontSize: 10, fontWeight: 700, color: '#A19F9D',
                textTransform: 'uppercase', letterSpacing: '0.6px', flexShrink: 0, minWidth: 88,
            } }, label),
        React.createElement("span", { style: { fontSize: 13, color: '#201F1E' } }, value || '—')));
};
var RowAccordion = function (_a) {
    var header = _a.header, children = _a.children, isOpen = _a.isOpen, onToggle = _a.onToggle, accentColor = _a.accentColor;
    return (React.createElement("div", { style: {
            border: '1px solid',
            borderColor: isOpen ? accentColor : '#EDEBE9',
            borderRadius: 8, marginBottom: 6, overflow: 'hidden',
            boxShadow: isOpen ? '0 2px 8px rgba(0,0,0,0.07)' : '0 1px 2px rgba(0,0,0,0.04)',
        } },
        React.createElement("button", { onClick: onToggle, style: {
                width: '100%', display: 'flex', alignItems: 'center', gap: 10,
                padding: '9px 12px', background: isOpen ? '#F8FFFE' : '#FAFAFA',
                border: 'none', cursor: 'pointer', textAlign: 'left', fontFamily: 'inherit',
            } },
            React.createElement(Icon_1.Icon, { iconName: isOpen ? 'ChevronDown' : 'ChevronRight', style: { fontSize: 11, color: accentColor, flexShrink: 0 } }),
            React.createElement("div", { style: { flex: 1, minWidth: 0 } }, header)),
        isOpen && (React.createElement("div", { style: { padding: '10px 14px 12px 36px', borderTop: '1px solid #EDEBE9', background: '#FFFFFF', borderLeft: "3px solid ".concat(accentColor) } }, children))));
};
var SectionHeader = function (_a) {
    var title = _a.title, count = _a.count, loading = _a.loading, isOpen = _a.isOpen, onToggle = _a.onToggle, color = _a.color, bg = _a.bg, iconName = _a.iconName;
    return (React.createElement("button", { onClick: onToggle, style: {
            width: '100%', display: 'flex', alignItems: 'center', gap: 12,
            padding: '12px 16px', background: isOpen ? bg : '#F3F2F1',
            border: '1px solid', borderColor: isOpen ? color : '#EDEBE9',
            borderRadius: isOpen ? '8px 8px 0 0' : 8,
            cursor: 'pointer', textAlign: 'left', fontFamily: 'inherit',
            transition: 'background 0.15s, border-color 0.15s',
            boxShadow: isOpen ? '0 2px 6px rgba(0,0,0,0.06)' : 'none',
        } },
        React.createElement(Icon_1.Icon, { iconName: iconName, style: { fontSize: 18, color: isOpen ? color : '#605E5C', flexShrink: 0 } }),
        React.createElement("span", { style: { flex: 1, fontSize: 14, fontWeight: 700, color: isOpen ? color : '#201F1E' } }, title),
        loading && React.createElement(Spinner_1.Spinner, { size: Spinner_1.SpinnerSize.xSmall }),
        !loading && count !== undefined && (React.createElement("span", { style: {
                background: isOpen ? color : '#C8C6C4', color: '#fff',
                fontSize: 11, fontWeight: 700, padding: '2px 8px', borderRadius: 10,
                minWidth: 20, textAlign: 'center', display: 'inline-block',
            } }, count)),
        React.createElement(Icon_1.Icon, { iconName: isOpen ? 'ChevronDown' : 'ChevronRight', style: { fontSize: 13, color: isOpen ? color : '#A19F9D', flexShrink: 0 } })));
};
// ── Loading / empty ───────────────────────────────────────────────────────────
var Loader = function () { return (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10, padding: '36px 16px' } },
    React.createElement(Spinner_1.Spinner, { size: Spinner_1.SpinnerSize.medium }),
    React.createElement("span", { style: { fontSize: 13, color: '#605E5C' } }, "Loading\u2026"))); };
var Empty = function (_a) {
    var message = _a.message;
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8, padding: '40px 0', textAlign: 'center' } },
        React.createElement(Icon_1.Icon, { iconName: "EntitlementPolicy", style: { fontSize: 36, color: '#C8C6C4' } }),
        React.createElement("span", { style: { fontSize: 13, color: '#605E5C' } }, message)));
};
var ApprovalSection = function (_a) {
    var rows = _a.rows, loading = _a.loading, error = _a.error, isOpen = _a.isOpen, onToggleSection = _a.onToggleSection;
    var _b = React.useState({}), expanded = _b[0], setExpanded = _b[1];
    var toggle = function (i) {
        return setExpanded(function (prev) {
            var next = {};
            Object.keys(prev).forEach(function (k) { next[Number(k)] = prev[Number(k)]; });
            next[i] = !prev[i];
            return next;
        });
    };
    var body = function () {
        if (loading)
            return React.createElement(Loader, null);
        if (error)
            return React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.error, styles: { root: { borderRadius: 6 } } }, error);
        if (!rows.length)
            return React.createElement(Empty, { message: "No approval log entries for this request." });
        return (React.createElement("div", null, rows.map(function (row, i) {
            var _a;
            var logName = (0, DataverseService_1.strVal)(row, 'sct_name') || '—';
            var statusRaw = row['sct_approvalstatus'];
            var statusDef = (_a = APPROVAL_STATUS[statusRaw]) !== null && _a !== void 0 ? _a : { label: (0, DataverseService_1.strVal)(row, 'sct_approvalstatus') || 'Unknown', color: '#605E5C', bg: '#F3F2F1' };
            var tier = resolveChoice(row, 'sct_tiername', APPROVAL_TIER);
            var approver = (0, DataverseService_1.strVal)(row, '_sct_approverid_value') || '—';
            var date = fmtDateTime(row);
            var comments = (0, DataverseService_1.strVal)(row, 'sct_comments');
            var header = (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' } },
                React.createElement("span", { style: { width: 8, height: 8, borderRadius: '50%', background: statusDef.color, flexShrink: 0 } }),
                React.createElement("span", { style: { fontSize: 11, fontFamily: 'monospace', fontWeight: 700, color: '#0078D4', flexShrink: 0 } }, logName),
                React.createElement("span", { style: { fontSize: 13, fontWeight: 600, color: '#201F1E', flex: 1, minWidth: 0 } }, approver),
                React.createElement(Badge, { label: tier, color: "#0078D4", bg: "#EBF3FB" }),
                React.createElement("span", { style: { fontSize: 11, color: '#A19F9D', whiteSpace: 'nowrap' } }, date)));
            return (React.createElement(RowAccordion, { key: i, isOpen: !!expanded[i], onToggle: function () { return toggle(i); }, accentColor: statusDef.color, header: header },
                React.createElement(DetailRow, { label: "Log No.", value: React.createElement("span", { style: { fontFamily: 'monospace', fontWeight: 700, color: '#0078D4' } }, logName) }),
                React.createElement(DetailRow, { label: "Tier", value: React.createElement("span", { style: { fontWeight: 600, color: '#0078D4' } }, tier) }),
                React.createElement(DetailRow, { label: "Approver", value: approver }),
                React.createElement(DetailRow, { label: "Status", value: React.createElement(Badge, { label: statusDef.label, color: statusDef.color, bg: statusDef.bg }) }),
                React.createElement(DetailRow, { label: "Date", value: date }),
                comments ? (React.createElement("div", { style: { marginTop: 8 } },
                    React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', marginBottom: 5 } }, "Comments"),
                    React.createElement("div", { style: { fontSize: 12, color: '#201F1E', lineHeight: '1.6', background: '#F8F8F8', border: '1px solid #EDEBE9', borderRadius: 6, padding: '8px 12px' } }, comments))) : (React.createElement(DetailRow, { label: "Comments", value: "\u2014" }))));
        })));
    };
    return (React.createElement("div", { style: { marginBottom: 12 } },
        React.createElement(SectionHeader, { title: "Approval Logs", count: loading ? undefined : rows.length, loading: loading, isOpen: isOpen, onToggle: onToggleSection, color: "#107C10", bg: "#DFF6DD", iconName: "Checklist" }),
        isOpen && (React.createElement("div", { style: { border: '1px solid #107C10', borderTop: 'none', borderRadius: '0 0 8px 8px', padding: 14, background: '#FAFFFE' } }, body()))));
};
var AuditSection = function (_a) {
    var rows = _a.rows, loading = _a.loading, error = _a.error, isOpen = _a.isOpen, onToggleSection = _a.onToggleSection;
    var _b = React.useState({}), expanded = _b[0], setExpanded = _b[1];
    var toggle = function (i) {
        return setExpanded(function (prev) {
            var next = {};
            Object.keys(prev).forEach(function (k) { next[Number(k)] = prev[Number(k)]; });
            next[i] = !prev[i];
            return next;
        });
    };
    var body = function () {
        if (loading)
            return React.createElement(Loader, null);
        if (error)
            return React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.error, styles: { root: { borderRadius: 6 } } }, error);
        if (!rows.length)
            return React.createElement(Empty, { message: "No audit log entries for this request." });
        return (React.createElement("div", null, rows.map(function (row, i) {
            var _a;
            var logName = (0, DataverseService_1.strVal)(row, 'sct_name') || '—';
            var actionRaw = row['sct_action'];
            var desc = (0, DataverseService_1.strVal)(row, 'sct_description');
            var actionDef = (_a = AUDIT_ACTION[actionRaw]) !== null && _a !== void 0 ? _a : { label: (0, DataverseService_1.strVal)(row, 'sct_action') || 'Audit', color: '#605E5C', bg: '#F3F2F1' };
            if (desc && desc.toLowerCase().indexOf('approved') !== -1) {
                actionDef = { label: 'Approved', color: '#107C10', bg: '#DFF6DD' };
            }
            else if (desc && desc.toLowerCase().indexOf('submitted') !== -1) {
                actionDef = { label: 'Submitted', color: '#004578', bg: '#DBEAFE' };
            }
            var actionBy = (0, DataverseService_1.strVal)(row, '_sct_actionby_value') || '—';
            var date = fmtDateTime(row);
            var oldVal = (0, DataverseService_1.strVal)(row, 'sct_oldvalue');
            var newVal = (0, DataverseService_1.strVal)(row, 'sct_newvalue');
            var header = (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' } },
                React.createElement(Badge, { label: actionDef.label, color: actionDef.color, bg: actionDef.bg }),
                React.createElement("span", { style: { fontSize: 11, fontFamily: 'monospace', fontWeight: 700, color: '#8764B8', flexShrink: 0 } }, logName),
                React.createElement("span", { style: { fontSize: 13, fontWeight: 600, color: '#201F1E', flex: 1, minWidth: 0 } }, actionBy),
                React.createElement("span", { style: { fontSize: 11, color: '#A19F9D', whiteSpace: 'nowrap' } }, date)));
            return (React.createElement(RowAccordion, { key: i, isOpen: !!expanded[i], onToggle: function () { return toggle(i); }, accentColor: actionDef.color, header: header },
                React.createElement(DetailRow, { label: "Audit No.", value: React.createElement("span", { style: { fontFamily: 'monospace', fontWeight: 700, color: '#8764B8' } }, logName) }),
                React.createElement(DetailRow, { label: "Action", value: React.createElement(Badge, { label: actionDef.label, color: actionDef.color, bg: actionDef.bg }) }),
                React.createElement(DetailRow, { label: "Action By", value: actionBy }),
                React.createElement(DetailRow, { label: "Date", value: date }),
                React.createElement(DetailRow, { label: "Description", value: desc }),
                (oldVal || newVal) && (React.createElement("div", { style: { marginTop: 10 } },
                    React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.6px', marginBottom: 6 } }, "Changes"),
                    React.createElement("div", { style: { display: 'flex', borderRadius: 6, border: '1px solid #EDEBE9', overflow: 'hidden', fontSize: 12 } },
                        React.createElement("div", { style: { flex: 1, padding: '8px 12px', background: '#FEF3F2', borderRight: '1px solid #EDEBE9' } },
                            React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#A4262C', textTransform: 'uppercase', marginBottom: 4 } }, "Before"),
                            React.createElement("div", { style: { color: '#605E5C', wordBreak: 'break-word', lineHeight: '1.5' } }, oldVal || '(empty)')),
                        React.createElement("div", { style: { flex: 1, padding: '8px 12px', background: '#F0FDF4' } },
                            React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#107C10', textTransform: 'uppercase', marginBottom: 4 } }, "After"),
                            React.createElement("div", { style: { color: '#201F1E', fontWeight: 500, wordBreak: 'break-word', lineHeight: '1.5' } }, newVal || '(empty)')))))));
        })));
    };
    return (React.createElement("div", { style: { marginBottom: 12 } },
        React.createElement(SectionHeader, { title: "Audit Logs", count: loading ? undefined : rows.length, loading: loading, isOpen: isOpen, onToggle: onToggleSection, color: "#8764B8", bg: "#F4F0FF", iconName: "History" }),
        isOpen && (React.createElement("div", { style: { border: '1px solid #8764B8', borderTop: 'none', borderRadius: '0 0 8px 8px', padding: 14, background: '#FDFCFF' } }, body()))));
};
// ── Workflow Stepper ──────────────────────────────────────────────────────────
var WorkflowStepper = function (_a) {
    var rows = _a.rows;
    var isApproved = function (tier, nameMatch) {
        return rows.some(function (r) {
            var _a;
            return (r['sct_tiername'] === tier || String((_a = r['sct_name']) !== null && _a !== void 0 ? _a : '').toLowerCase().indexOf(nameMatch) !== -1) &&
                (r['sct_approvalstatus'] === 1 || r['sct_approvalstatus'] === 854000000);
        });
    };
    var isRejected = rows.some(function (r) { return r['sct_approvalstatus'] === 2 || r['sct_approvalstatus'] === 854000001; });
    var isReturned = rows.some(function (r) { return r['sct_approvalstatus'] === 3 || r['sct_approvalstatus'] === 854000002; });
    var mgrDone = isApproved(4, 'manager');
    var t1Done = isApproved(1, 'tier 1');
    var t2Done = isApproved(2, 'tier 2');
    var t3Done = isApproved(3, 'tier 3');
    var activeStage = 'Manager';
    if (mgrDone && !t1Done)
        activeStage = 'Tier 1';
    else if (t1Done && !t2Done)
        activeStage = 'Tier 2';
    else if (t2Done && !t3Done)
        activeStage = 'Tier 3';
    else if (t3Done)
        activeStage = 'PO Issued';
    var stages = [
        { label: 'Submitted', done: true, active: false },
        { label: 'Manager', done: mgrDone, active: activeStage === 'Manager' && !isRejected && !isReturned },
        { label: 'Tier 1', done: t1Done, active: activeStage === 'Tier 1' && !isRejected && !isReturned },
        { label: 'Tier 2', done: t2Done, active: activeStage === 'Tier 2' && !isRejected && !isReturned },
        { label: 'Tier 3', done: t3Done, active: activeStage === 'Tier 3' && !isRejected && !isReturned },
        { label: 'PO Issued', done: t3Done, active: activeStage === 'PO Issued' && !isRejected && !isReturned },
    ];
    return (React.createElement("div", { style: {
            background: '#ffffff',
            border: '1px solid #E2E8F0',
            borderRadius: 10,
            padding: '12px 14px',
            marginBottom: 14,
            boxShadow: '0 1px 4px rgba(15, 23, 42, 0.04)',
        } },
        React.createElement("div", { style: { fontSize: 11, fontWeight: 700, color: '#475569', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 10 } }, "Workflow Pipeline Progress"),
        React.createElement("div", { style: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 4 } }, stages.map(function (st, i) {
            var isCurrentActive = st.active;
            var isDone = st.done;
            return (React.createElement(React.Fragment, { key: i },
                React.createElement("div", { style: { display: 'flex', flexDirection: 'column', alignItems: 'center', flex: 1, minWidth: 0 } },
                    React.createElement("div", { style: {
                            width: 24,
                            height: 24,
                            borderRadius: '50%',
                            background: isDone ? '#107C10' : isCurrentActive ? '#0078D4' : '#F1F5F9',
                            color: isDone || isCurrentActive ? '#ffffff' : '#94A3B8',
                            border: isCurrentActive ? '2px solid #60A5FA' : '1px solid #CBD5E1',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            fontSize: 10,
                            fontWeight: 700,
                            boxShadow: isCurrentActive ? '0 0 0 3px rgba(0, 120, 212, 0.2)' : 'none',
                        } }, isDone ? React.createElement(Icon_1.Icon, { iconName: "CheckMark", style: { fontSize: 10 } }) : i + 1),
                    React.createElement("div", { style: {
                            fontSize: 10.5,
                            fontWeight: isCurrentActive ? 700 : isDone ? 600 : 500,
                            color: isDone ? '#107C10' : isCurrentActive ? '#0078D4' : '#64748B',
                            marginTop: 4,
                            textAlign: 'center',
                            overflow: 'hidden',
                            textOverflow: 'ellipsis',
                            whiteSpace: 'nowrap',
                        } }, st.label)),
                i < stages.length - 1 && (React.createElement("div", { style: {
                        flex: 1,
                        height: 2,
                        minWidth: 8,
                        background: st.done ? '#107C10' : '#E2E8F0',
                        marginBottom: 16,
                    } }))));
        }))));
};
// ── Root: fetches data + renders both sections ─────────────────────────────────
var LogsContent = function (_a) {
    var requestId = _a.requestId, service = _a.service;
    var _b = React.useState([]), approvalRows = _b[0], setApprovalRows = _b[1];
    var _c = React.useState([]), auditRows = _c[0], setAuditRows = _c[1];
    var _d = React.useState(true), loadingApproval = _d[0], setLoadingApproval = _d[1];
    var _e = React.useState(true), loadingAudit = _e[0], setLoadingAudit = _e[1];
    var _f = React.useState(null), approvalError = _f[0], setApprovalError = _f[1];
    var _g = React.useState(null), auditError = _g[0], setAuditError = _g[1];
    var _h = React.useState(true), approvalOpen = _h[0], setApprovalOpen = _h[1];
    var _j = React.useState(true), auditOpen = _j[0], setAuditOpen = _j[1];
    React.useEffect(function () {
        var cancelled = false;
        var parseEpoch = function (rec) {
            var _a, _b;
            var val = String((_b = (_a = rec['createdon']) !== null && _a !== void 0 ? _a : rec['sct_decisionon']) !== null && _b !== void 0 ? _b : '').trim();
            if (!val)
                return 0;
            if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(val)) {
                val += 'Z';
            }
            var t = new Date(val).getTime();
            return isNaN(t) ? 0 : t;
        };
        var sortByNewest = function (records) {
            return tslib_1.__spreadArray([], records, true).sort(function (a, b) {
                var _a, _b;
                var timeA = parseEpoch(a);
                var timeB = parseEpoch(b);
                if (timeB !== timeA)
                    return timeB - timeA;
                return String((_a = b['sct_name']) !== null && _a !== void 0 ? _a : '').localeCompare(String((_b = a['sct_name']) !== null && _b !== void 0 ? _b : ''));
            });
        };
        var fetchApproval = function () { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
            var rows, err_1;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        setLoadingApproval(true);
                        setApprovalError(null);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, 4, 5]);
                        return [4 /*yield*/, service.getApprovalLogs(requestId)];
                    case 2:
                        rows = _a.sent();
                        if (!cancelled)
                            setApprovalRows(sortByNewest(Array.isArray(rows) ? rows : []));
                        return [3 /*break*/, 5];
                    case 3:
                        err_1 = _a.sent();
                        if (!cancelled)
                            setApprovalError(err_1 instanceof Error ? err_1.message : 'Failed to load approval logs.');
                        return [3 /*break*/, 5];
                    case 4:
                        if (!cancelled)
                            setLoadingApproval(false);
                        return [7 /*endfinally*/];
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        var fetchAudit = function () { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
            var rows, err_2;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        setLoadingAudit(true);
                        setAuditError(null);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, 4, 5]);
                        return [4 /*yield*/, service.getAuditLogs(requestId)];
                    case 2:
                        rows = _a.sent();
                        if (!cancelled)
                            setAuditRows(sortByNewest(Array.isArray(rows) ? rows : []));
                        return [3 /*break*/, 5];
                    case 3:
                        err_2 = _a.sent();
                        if (!cancelled)
                            setAuditError(err_2 instanceof Error ? err_2.message : 'Failed to load audit logs.');
                        return [3 /*break*/, 5];
                    case 4:
                        if (!cancelled)
                            setLoadingAudit(false);
                        return [7 /*endfinally*/];
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        void fetchApproval();
        void fetchAudit();
        return function () { cancelled = true; };
    }, [requestId]);
    var totalLoaded = !loadingApproval && !loadingAudit;
    return (React.createElement("div", null,
        totalLoaded && (React.createElement(React.Fragment, null,
            React.createElement(WorkflowStepper, { rows: approvalRows }),
            React.createElement("div", { style: {
                    display: 'flex', gap: 20, padding: '10px 14px',
                    background: '#F8F9FA', borderRadius: 8, marginBottom: 14,
                    border: '1px solid #EDEBE9', flexWrap: 'wrap', alignItems: 'center',
                } },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 7 } },
                    React.createElement("span", { style: { width: 10, height: 10, borderRadius: '50%', background: '#107C10', flexShrink: 0, display: 'inline-block' } }),
                    React.createElement("span", { style: { fontSize: 12, color: '#323130', fontWeight: 600 } },
                        approvalRows.length,
                        " Approval ",
                        approvalRows.length === 1 ? 'Entry' : 'Entries')),
                React.createElement("div", { style: { width: 1, alignSelf: 'stretch', background: '#EDEBE9', flexShrink: 0 } }),
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 7 } },
                    React.createElement("span", { style: { width: 10, height: 10, borderRadius: '50%', background: '#8764B8', flexShrink: 0, display: 'inline-block' } }),
                    React.createElement("span", { style: { fontSize: 12, color: '#323130', fontWeight: 600 } },
                        auditRows.length,
                        " Audit ",
                        auditRows.length === 1 ? 'Entry' : 'Entries'))))),
        React.createElement(ApprovalSection, { rows: approvalRows, loading: loadingApproval, error: approvalError, isOpen: approvalOpen, onToggleSection: function () { return setApprovalOpen(function (o) { return !o; }); } }),
        React.createElement(AuditSection, { rows: auditRows, loading: loadingAudit, error: auditError, isOpen: auditOpen, onToggleSection: function () { return setAuditOpen(function (o) { return !o; }); } })));
};
exports.default = LogsContent;
//# sourceMappingURL=LogsContent.js.map