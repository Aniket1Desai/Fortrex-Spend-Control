var _a, _b, _c;
// Partial<Record<>> because the Dashboard uses OData formatted values for display —
// these maps are only needed when rendering enum integers directly from the model.
export var REQUEST_STATUS_LABELS = (_a = {},
    _a[854000001 /* SCTRequestStatus.Draft */] = 'Draft',
    _a[854000002 /* SCTRequestStatus.Submitted */] = 'Submitted',
    _a[854000003 /* SCTRequestStatus.InReview */] = 'In Review',
    _a[854000004 /* SCTRequestStatus.POIssued */] = 'PO Issued',
    _a[854000005 /* SCTRequestStatus.Rejected */] = 'Rejected',
    _a[854000006 /* SCTRequestStatus.Returned */] = 'Returned',
    _a[854000007 /* SCTRequestStatus.Closed */] = 'Closed',
    _a);
export var REQUEST_STATUS_COLORS = (_b = {},
    _b[854000001 /* SCTRequestStatus.Draft */] = '#605E5C',
    _b[854000002 /* SCTRequestStatus.Submitted */] = '#0078D4',
    _b[854000003 /* SCTRequestStatus.InReview */] = '#FF8C00',
    _b[854000004 /* SCTRequestStatus.POIssued */] = '#00B294',
    _b[854000005 /* SCTRequestStatus.Rejected */] = '#A4262C',
    _b[854000006 /* SCTRequestStatus.Returned */] = '#8764B8',
    _b[854000007 /* SCTRequestStatus.Closed */] = '#107C10',
    _b);
export var REQUEST_TYPE_LABELS = (_c = {},
    _c[100000000 /* SCTRequestType.Goods */] = 'Goods',
    _c[100000001 /* SCTRequestType.Services */] = 'Services',
    _c[100000002 /* SCTRequestType.Software */] = 'Software',
    _c[100000003 /* SCTRequestType.CapitalExpenditure */] = 'Capital Expenditure',
    _c[100000004 /* SCTRequestType.Works */] = 'Works',
    _c);
/** All status labels in workflow order — confirmed from Power Apps sct_Request_Status option set. */
export var ALL_REQUEST_STATUS_VALUES = [
    'Draft',
    'Submitted',
    'In Review',
    'PO Issued',
    'Rejected',
    'Returned',
    'Closed',
];
// ── Choice Field Integer → Human-readable Label Maps ────────────────────────
export var BUSINESS_UNIT_LABELS = {
    854000000: 'PSSI',
    854000001: 'PCI',
    854000002: 'SAFE',
    854000003: 'CCCS',
};
export var PURCHASE_TYPE_LABELS = {
    854000000: 'Goods',
    854000001: 'Repair',
    854000002: 'Rental',
    854000003: 'Service',
    854000004: 'Software',
    854000005: 'Temporary Labor',
    854000006: 'Other',
};
export var URGENCY_LABELS = {
    854000000: 'Standard',
    854000001: 'Urgent',
    854000002: 'Emergency',
};
export var PRICING_BASIS_LABELS = {
    854000000: 'Fixed Price',
    854000001: 'Rate Card',
    854000002: 'Time and Materials with Cap',
    854000003: 'Usage-Based Estimate',
};
export var COST_IS_FOR_LABELS = {
    854000000: 'This Site',
    854000001: 'A Customer Job',
    854000002: 'Several Sites',
    854000003: 'Company-Wide',
};
/** Resolves the human-readable text label for any choice or lookup column */
export function getChoiceLabel(record, field, map) {
    if (!record)
        return '';
    // 1. Prefer OData formatted value annotation
    var formatted = record["".concat(field, "@OData.Community.Display.V1.FormattedValue")];
    if (formatted !== undefined && formatted !== null && String(formatted).trim() !== '') {
        return String(formatted).trim();
    }
    // 2. Read raw value
    var raw = record[field];
    if (raw === undefined || raw === null)
        return '';
    var rawStr = String(raw).trim();
    if (!rawStr)
        return '';
    // 3. Map numeric codes to text labels
    if (map) {
        var mapped = map[rawStr] || map[Number(rawStr)];
        if (mapped)
            return mapped;
    }
    return rawStr;
}
export function formatCurrency(amount) {
    return new Intl.NumberFormat('en-AU', {
        style: 'currency',
        currency: 'AUD',
        minimumFractionDigits: 2,
    }).format(amount);
}
export function formatDate(iso) {
    try {
        return new Intl.DateTimeFormat('en-AU', {
            day: '2-digit', month: 'short', year: 'numeric',
        }).format(new Date(iso));
    }
    catch (_a) {
        return iso;
    }
}
/**
 * Formats a date/timestamp to strict 'dd/MM/yyyy HH:mm tt' format in user's local timezone.
 * Example: 18/11/2026 09:31 AM
 *
 * Automatically handles:
 * - Dataverse UTC ISO-8601 strings (e.g. '2026-11-18T04:01:23Z') converted to local time
 * - Raw ISO strings missing timezone offsets (appends 'Z' so they parse as UTC)
 * - Date objects
 */
export function formatDateTime(val) {
    if (!val)
        return '—';
    try {
        var iso = val instanceof Date ? val.toISOString() : String(val).trim();
        if (!iso)
            return '—';
        // If raw ISO has no timezone designator, append 'Z' so it is parsed as UTC
        if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(iso)) {
            iso += 'Z';
        }
        var d = new Date(iso);
        if (isNaN(d.getTime()))
            return typeof val === 'string' ? val : '—';
        var pad2 = function (n) { return (n < 10 ? '0' + n : String(n)); };
        var day = pad2(d.getDate());
        var month = pad2(d.getMonth() + 1);
        var year = d.getFullYear();
        var h = d.getHours();
        var tt = h >= 12 ? 'PM' : 'AM';
        var h12 = h % 12 || 12;
        var hh = pad2(h12);
        var mm = pad2(d.getMinutes());
        return "".concat(day, "/").concat(month, "/").concat(year, " ").concat(hh, ":").concat(mm, " ").concat(tt);
    }
    catch (_a) {
        return typeof val === 'string' ? val : '—';
    }
}
/**
 * Canonical status label resolver for an SCT request record.
 * Accurately maps numeric choice integers, formatted values, and string columns
 * to the exact canonical status label: 'Draft' | 'Submitted' | 'In Review' | 'PO Issued' | 'Rejected' | 'Returned' | 'Closed'.
 *
 * Strict Workflow Rules:
 * 1. Explicit terminal rejections, returns, closures, or drafts always take priority.
 * 2. Active Pending Tier: If a record has an active pending tier (Manager, Tier 1, Tier 2, Tier 3),
 *    it is in-flight. It can NEVER be 'PO Issued'.
 *    - Pending Manager: 'Submitted'
 *    - Pending Tier 1, 2, or 3: 'In Review'
 * 3. Terminal 'PO Issued': Only valid when there is NO active pending tier AND status is PO Issued / Approved.
 */
export function getRecordStatusLabel(record) {
    var _a, _b;
    if (!record)
        return '';
    var raw = record['sct_request_status'];
    var rawInt = typeof raw === 'number' ? raw : parseInt(String(raw !== null && raw !== void 0 ? raw : ''), 10);
    var formatted = record['sct_request_status@OData.Community.Display.V1.FormattedValue'];
    var fStr = formatted ? String(formatted).trim() : '';
    var reqStatusStr = String((_a = record['sct_requeststatus']) !== null && _a !== void 0 ? _a : '').trim();
    var reqStatusLower = reqStatusStr.toLowerCase();
    var str = String((_b = record['sct_request_status']) !== null && _b !== void 0 ? _b : '').trim();
    var sLower = str.toLowerCase();
    var fLower = fStr.toLowerCase();
    // 1. Explicit terminal / non-review states take absolute priority
    if (rawInt === 854000005 ||
        rawInt === 854000004 ||
        reqStatusLower === 'rejected' ||
        sLower === 'rejected' ||
        fLower === 'rejected') {
        return 'Rejected';
    }
    if (rawInt === 854000006 ||
        rawInt === 854000005 ||
        reqStatusLower === 'returned' ||
        sLower === 'returned' ||
        fLower === 'returned') {
        return 'Returned';
    }
    if (rawInt === 854000007 ||
        rawInt === 854000006 ||
        reqStatusLower === 'closed' ||
        sLower === 'closed' ||
        fLower === 'closed') {
        return 'Closed';
    }
    if (rawInt === 854000001 ||
        rawInt === 854000000 ||
        reqStatusLower === 'draft' ||
        sLower === 'draft' ||
        fLower === 'draft') {
        return 'Draft';
    }
    // 2. Active Pending Tier detection
    var rawPt = String(record['cr2a4_pendingtier'] ||
        record['cr2a4_PendingTier'] ||
        record['sct_pendingtier'] ||
        record['pendingtier'] ||
        '').trim();
    var ptNorm = rawPt.toLowerCase().replace(/[^a-z0-9]/g, '');
    var hasActivePendingTier = ptNorm === 'manager' || ptNorm.indexOf('tier') !== -1 || ptNorm === '1' || ptNorm === '2' || ptNorm === '3';
    if (hasActivePendingTier) {
        if (ptNorm === 'manager') {
            // Pending with Line Manager is canonical 'Submitted'
            return 'Submitted';
        }
        // Pending with Tier 1, 2, or 3 is canonical 'In Review'
        return 'In Review';
    }
    // 3. Terminal 'PO Issued' (ONLY when NO active pending tier)
    if (rawInt === 854000004 ||
        rawInt === 854000003 ||
        reqStatusLower === 'po issued' ||
        reqStatusLower === 'approved' ||
        sLower === 'po issued' ||
        sLower === 'approved' ||
        fLower === 'po issued' ||
        fLower === 'approved') {
        return 'PO Issued';
    }
    // 4. Fallbacks when no pending tier column is present on record
    if (rawInt === 854000002 || rawInt === 854000001 || reqStatusLower === 'submitted' || sLower === 'submitted' || fLower === 'submitted') {
        return 'Submitted';
    }
    if (rawInt === 854000003 || rawInt === 854000002 || reqStatusLower.indexOf('review') !== -1 || sLower.indexOf('review') !== -1 || fLower.indexOf('review') !== -1) {
        return 'In Review';
    }
    if (fStr)
        return fStr;
    if (reqStatusStr)
        return reqStatusStr;
    return '';
}
//# sourceMappingURL=dashboardUtils.js.map