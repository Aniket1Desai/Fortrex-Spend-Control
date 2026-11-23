"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var DetailsList_1 = require("@fluentui/react/lib/DetailsList");
var Icon_1 = require("@fluentui/react/lib/Icon");
var Text_1 = require("@fluentui/react/lib/Text");
var Button_1 = require("@fluentui/react/lib/Button");
var Panel_1 = require("@fluentui/react/lib/Panel");
var Dialog_1 = require("@fluentui/react/lib/Dialog");
var TextField_1 = require("@fluentui/react/lib/TextField");
var Toggle_1 = require("@fluentui/react/lib/Toggle");
var MessageBar_1 = require("@fluentui/react/lib/MessageBar");
var DatePicker_1 = require("@fluentui/react/lib/DatePicker");
var Calendar_1 = require("@fluentui/react/lib/Calendar");
var Spinner_1 = require("@fluentui/react/lib/Spinner");
var Pivot_1 = require("@fluentui/react/lib/Pivot");
var GraphMailService_1 = require("../../services/GraphMailService");
var DataverseService_1 = require("../../services/DataverseService");
var RequestLogsPanel_1 = tslib_1.__importDefault(require("../history/RequestLogsPanel"));
var auditLogger_1 = require("../../utils/auditLogger");
var ApprovalRoutingService_1 = require("../../services/ApprovalRoutingService");
var common_1 = require("../common");
var AppShell_module_scss_1 = tslib_1.__importDefault(require("../shell/AppShell.module.scss"));
var dashboardUtils_1 = require("./dashboardUtils");
// ── Constants ─────────────────────────────────────────────────────────────────
var PAGE_SIZE_OPTIONS = [10, 15, 25, 50];
var ALWAYS_SKIP = new Set([
    '@odata.etag', '@odata.context', 'versionnumber',
    'timezoneruleversionnumber', 'utcconversiontimezonecode',
    'importsequencenumber', 'overriddencreatedon',
]);
// Fields to show as table columns — confirmed from actual Dataverse sct_request record.
// Boolean fields excluded per user request (they are stored but not shown in the table).
// Lookup fields use the _xxx_value OData pattern and show display text via FormattedValue annotation.
var PRIORITY_FIELDS = [
    // Reference & identity
    'sct_sct_reference',
    // Business status
    'sct_request_status',
    // Site lookup (_sct_site_value@associatednavigationproperty = "sct_Site")
    '_sct_site_value',
    // Requester lookup
    '_sct_requester_value',
    // Purchase classification
    'sct_purchase_type',
    'sct_urgency',
    'sct_business_unit',
    'sct_category',
    // Financials
    'sct_estimated_total_value',
    // Dates
    'sct_date_needed_by',
    'sct_submitted_date',
    'sct_term_start_date',
    'sct_term_end_date',
    // Supplier
    '_sct_suggested_supplier_value',
    // Description & function
    'sct_description',
    'sct_function',
    // Logistics
    'sct_cost_is_for',
    'sct_delivery_address',
    'sct_cost_center',
    'sct_customer_name',
    'sct_equipment_asset_tag',
    'sct_urgency_reason',
    // PO tracking
    'sct_po_number',
    'sct_po_issue_date',
    // Contact
    'sct_receiving_contact_email',
    '_sct_receiving_contact_value',
];
// ── Dashboard table columns — the 9 specific fields shown in the list view ────
// Ordered by business relevance; shown regardless of detection order.
var DASHBOARD_VISIBLE_COLUMNS = [
    { field: 'sct_sct_reference', minW: 110, maxW: 140 },
    { field: 'sct_request_status', minW: 130, maxW: 160 },
    { field: '_sct_requester_value', minW: 120, maxW: 160 },
    { field: 'sct_business_unit', minW: 90, maxW: 120 },
    { field: '_sct_site_value', minW: 120, maxW: 180 },
    { field: 'sct_purchase_type', minW: 100, maxW: 130 },
    { field: 'sct_estimated_total_value', minW: 110, maxW: 140 },
    { field: 'sct_urgency', minW: 90, maxW: 110 },
    { field: 'sct_submitted_date', minW: 110, maxW: 140 },
];
// ── Edit panel — fields editable by a requester ───────────────────────────────
// Read-only fields (reference, status, dates, choice ints, lookups, procurement)
// are shown in the View panel only — NOT as inputs in the Edit panel.
var EDIT_TEXT_FIELDS = new Set([
    'sct_description', 'sct_function', 'sct_urgency_reason',
    'sct_delivery_address', 'sct_equipment_asset_tag',
    'sct_customer_name', 'sct_receiving_contact_email',
    'sct_emergency_declared_by_email',
]);
var EDIT_NUMBER_FIELDS = new Set(['sct_estimated_total_value']);
var EDIT_DATE_FIELDS = new Set(['sct_date_needed_by', 'sct_term_start_date', 'sct_term_end_date']);
// Boolean fields are already defined in BOOLEAN_FIELDS — all are editable via Toggle.
// Fields NOT in any of the above sets are read-only: shown in View panel, not Edit panel.
// ── Helpers ───────────────────────────────────────────────────────────────────
/** True if the field only contains a bare GUID with no display-text annotation.
 *  Dataverse lookup fields (e.g. _sct_siteid_value) store the FK GUID as their raw value
 *  BUT also carry a FormattedValue annotation with the human-readable display text.
 *  If that annotation exists, strVal() will return the text — so we must NOT skip the field. */
function isGuidField(records, field) {
    for (var _i = 0, records_1 = records; _i < records_1.length; _i++) {
        var r = records_1[_i];
        // If a formatted-value annotation exists, strVal() returns display text, not the GUID.
        var fmtKey = "".concat(field, "@OData.Community.Display.V1.FormattedValue");
        if (r[fmtKey] != null && String(r[fmtKey]).trim() !== '')
            return false;
        var v = r[field];
        if (v == null)
            continue;
        return typeof v === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v);
    }
    return false;
}
/** All usable fields (not @annotations, not system noise) */
function allUsableFields(records) {
    if (!records.length)
        return [];
    return Object.keys(records[0]).filter(function (k) {
        if (ALWAYS_SKIP.has(k))
            return false;
        if (k.startsWith('@'))
            return false;
        if (k.endsWith('@OData.Community.Display.V1.FormattedValue'))
            return false;
        return true;
    });
}
// Boolean fields confirmed from Dataverse record — excluded from table columns per spec.
var BOOLEAN_FIELDS = new Set([
    'sct_is_new_supplier', 'sct_cost_recharged_to_customer', 'sct_customer_agreement_attached',
    'sct_tech_or_data_involved', 'sct_supplier_on_premises', 'sct_possible_capital_purchase',
    'sct_is_confidential', 'sct_is_retrospective', 'sct_existing_agreement_or_po',
    'sct_declaration_no_prior_commitment', 'sct_declaration_no_conflict_of_interest',
]);
/** Fields to show as table columns — sct_ fields first, skip bare GUIDs and boolean fields */
function detectDisplayFields(records) {
    var usable = allUsableFields(records);
    var nonGuid = usable.filter(function (k) { return !isGuidField(records, k); });
    // Separate custom (sct_) from system fields; exclude boolean columns
    var customFields = nonGuid.filter(function (k) { return k.startsWith('sct_') && !BOOLEAN_FIELDS.has(k); });
    var systemDateFields = ['createdon', 'modifiedon'].filter(function (k) { return nonGuid.indexOf(k) !== -1; });
    // Within custom fields: priority ones first, then the rest alpha
    var sorted = [];
    PRIORITY_FIELDS.forEach(function (p) { if (customFields.indexOf(p) !== -1)
        sorted.push(p); });
    customFields.forEach(function (k) { if (sorted.indexOf(k) === -1)
        sorted.push(k); });
    // Append key system dates at the end
    systemDateFields.forEach(function (k) { if (sorted.indexOf(k) === -1)
        sorted.push(k); });
    return sorted;
}
/** Safely convert any truthy string/value to a strict JavaScript boolean for Edm.Boolean fields.
 *  Dataverse rejects "Yes", "No", "true", "false" strings — must send actual true/false. */
function toBoolean(value) {
    if (typeof value === 'boolean')
        return value;
    if (typeof value === 'string') {
        var lower = value.trim().toLowerCase();
        return lower === 'true' || lower === 'yes' || lower === '1';
    }
    return Boolean(value);
}
/**
 * Convert any date string to strict Edm.Date format (YYYY-MM-DD) required by Dataverse.
 * Handles: already-correct YYYY-MM-DD, localized M/D/YYYY, and JS Date strings.
 * Returns null when the input cannot be parsed — caller should skip the field.
 * Note: avoids padStart — not available in this project's TS lib target.
 */
function formatToEdmDate(dateInput) {
    if (!dateInput)
        return null;
    var str = String(dateInput).trim();
    if (!str)
        return null;
    // Already in strict Edm.Date format
    if (/^\d{4}-\d{2}-\d{2}$/.test(str))
        return str;
    // Try JS Date parsing (handles ISO strings, long-form locale strings, etc.)
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
        var y = d.getFullYear();
        var mo = d.getMonth() + 1;
        var da = d.getDate();
        var mm = mo < 10 ? ('0' + mo) : String(mo);
        var dd = da < 10 ? ('0' + da) : String(da);
        return y + '-' + mm + '-' + dd;
    }
    // Fallback: split on / or - for M/D/YYYY or M-D-YYYY patterns
    var parts = str.split(/[\/\-]/);
    if (parts.length === 3) {
        var rawMonth = parts[0].trim();
        var rawDay = parts[1].trim();
        var rawYear = parts[2].trim();
        if (rawYear.length === 2)
            rawYear = '20' + rawYear;
        var mm = rawMonth.length < 2 ? ('0' + rawMonth) : rawMonth;
        var dd = rawDay.length < 2 ? ('0' + rawDay) : rawDay;
        if (rawYear.length === 4)
            return rawYear + '-' + mm + '-' + dd;
    }
    return null;
}
/** Human-readable column header — field names confirmed from actual Dataverse record */
function fieldLabel(logicalName) {
    var KNOWN_LABELS = {
        // Reference
        'sct_sct_reference': 'Reference',
        // Status
        'sct_request_status': 'Status',
        // Site lookup: _sct_site_value@associatednavigationproperty = "sct_Site"
        '_sct_site_value': 'Site',
        // Requester / contacts
        '_sct_requester_value': 'Requester',
        '_sct_receiving_contact_value': 'Receiving Contact',
        '_sct_emergency_declared_by_value': 'Emergency Declared By',
        // Supplier: _sct_suggested_supplier_value@associatednavigationproperty = "sct_Suggested_Supplier"
        '_sct_suggested_supplier_value': 'Suggested Supplier',
        // Purchase details
        'sct_purchase_type': 'Purchase Type',
        'sct_urgency': 'Urgency',
        'sct_business_unit': 'Business Unit',
        'sct_category': 'Category',
        'sct_pricing_basis': 'Pricing Basis',
        'sct_cost_is_for': 'Cost Is For',
        // Financials
        'sct_estimated_total_value': 'Estimated Value',
        // Dates
        'sct_date_needed_by': 'Date Needed By',
        'sct_submitted_date': 'Submitted Date',
        'sct_term_start_date': 'Term Start',
        'sct_term_end_date': 'Term End',
        'sct_po_issue_date': 'PO Issue Date',
        // Text
        'sct_description': 'Description',
        'sct_function': 'Function',
        'sct_delivery_address': 'Delivery Address',
        'sct_cost_center': 'Cost Centre',
        'sct_customer_name': 'Customer Name',
        'sct_equipment_asset_tag': 'Asset Tag',
        'sct_urgency_reason': 'Urgency Reason',
        'sct_po_number': 'PO Number',
        'sct_receiving_contact_email': 'Receiving Contact Email',
        'sct_emergency_declared_by_email': 'Emergency Email',
        'sct_sct_notes': 'Notes',
        // System lookups
        '_ownerid_value': 'Owner',
        '_createdby_value': 'Created By',
        '_modifiedby_value': 'Modified By',
    };
    if (KNOWN_LABELS[logicalName])
        return KNOWN_LABELS[logicalName];
    return logicalName
        .replace(/^_/, '')
        .replace(/_value$/, '')
        .replace(/^sct_/, '')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, function (c) { return c.toUpperCase(); });
}
/**
 * Detect live status values from records using canonical status label resolution.
 */
function detectStatusValues(records) {
    var vals = new Set();
    records.forEach(function (r) {
        var label = (0, dashboardUtils_1.getRecordStatusLabel)(r);
        if (label && label.trim() !== '')
            vals.add(label.trim());
    });
    return Array.from(vals).sort();
}
/** Find the primary key GUID field for a set of records */
function findPrimaryKey(records) {
    var _a, _b;
    if (!records.length)
        return null;
    var guidFields = Object.keys(records[0]).filter(function (k) {
        return !k.startsWith('@') &&
            !k.endsWith('@OData.Community.Display.V1.FormattedValue') &&
            isGuidField(records, k) &&
            !k.startsWith('_');
    });
    // prefer sct_requestid
    return (_b = (_a = guidFields.find(function (k) { return k.endsWith('requestid'); })) !== null && _a !== void 0 ? _a : guidFields[0]) !== null && _b !== void 0 ? _b : null;
}
// ── Status pill ───────────────────────────────────────────────────────────────
var STATUS_COLOURS = {
    // Core workflow states — mint-cyan palette
    draft: { bg: '#F3F2F1', fg: '#605E5C' },
    submitted: { bg: '#E0F5F2', fg: '#00897B' },
    'in review': { bg: '#E3F2FB', fg: '#0277BD' },
    'under review': { bg: '#E3F2FB', fg: '#0277BD' },
    pending: { bg: '#FFF8E1', fg: '#F57F17' },
    'awaiting': { bg: '#FFF8E1', fg: '#F57F17' },
    approved: { bg: '#E0F5F2', fg: '#00695C' },
    rejected: { bg: '#FDE7E9', fg: '#A4262C' },
    returned: { bg: '#FFF4CE', fg: '#7A4F01' },
    cancelled: { bg: '#F3F2F1', fg: '#A19F9D' },
    completed: { bg: '#E0F5F2', fg: '#00695C' },
    closed: { bg: '#E0F5F2', fg: '#00695C' },
    // PO-specific
    'po draft': { bg: '#FFF8E1', fg: '#F57F17' },
    'po issued': { bg: '#E0F5F2', fg: '#00695C' },
    'po raised': { bg: '#E0F5F2', fg: '#00695C' },
    'po pending': { bg: '#FFF8E1', fg: '#F57F17' },
    // Triage / sourcing
    triage: { bg: '#F3E5F5', fg: '#6A1B9A' },
    sourcing: { bg: '#F3E5F5', fg: '#6A1B9A' },
    'fp&a': { bg: '#F3E5F5', fg: '#6A1B9A' },
    'cfo': { bg: '#FDE7E9', fg: '#A4262C' },
    receiving: { bg: '#E0F5F2', fg: '#00695C' },
    // System state
    active: { bg: '#E0F5F2', fg: '#00695C' },
    inactive: { bg: '#F3F2F1', fg: '#605E5C' },
    // Choice / boolean
    'fixed price': { bg: '#E0F5F2', fg: '#00897B' },
    'time and material': { bg: '#FFF8E1', fg: '#F57F17' },
    yes: { bg: '#E0F5F2', fg: '#00695C' },
    no: { bg: '#F3F2F1', fg: '#605E5C' },
};
var INT_TO_STATUS = {
    '854000001': 'Draft',
    '854000002': 'Submitted',
    '854000003': 'In Review',
    '854000004': 'PO Issued',
    '854000005': 'Rejected',
    '854000006': 'Returned',
    '854000007': 'Closed',
    '854000000': 'Draft',
};
var StatusPill = function (_a) {
    var _b;
    var value = _a.value;
    var displayVal = (_b = INT_TO_STATUS[value]) !== null && _b !== void 0 ? _b : value;
    var key = Object.keys(STATUS_COLOURS).find(function (k) { return displayVal.toLowerCase().indexOf(k) !== -1; });
    var c = key ? STATUS_COLOURS[key] : { bg: '#F3F2F1', fg: '#323130' };
    return (React.createElement("span", { style: {
            background: c.bg, color: c.fg, fontSize: 12, fontWeight: 600,
            padding: '3px 10px', borderRadius: 12, whiteSpace: 'nowrap',
        } }, displayVal || '—'));
};
// ── Main component ────────────────────────────────────────────────────────────
// ── Confirmed status integers (Power Apps > sct_Request_Status option set: 854000001 to 854000007) ────
var DSTATUS = {
    DRAFT: 854000001,
    SUBMITTED: 854000002,
    IN_REVIEW: 854000003,
    PO_ISSUED: 854000004,
    REJECTED: 854000005,
    RETURNED: 854000006,
    CLOSED: 854000007,
};
// CFO/FP&A gate threshold
var D_CFO_THRESHOLD = 10000;
// Date helper — no padStart (not in this project's TS lib target)
function _dateToEdm(d) {
    var y = d.getFullYear();
    var mo = d.getMonth() + 1;
    var da = d.getDate();
    return y + '-' + (mo < 10 ? '0' + mo : String(mo)) + '-' + (da < 10 ? '0' + da : String(da));
}
var Dashboard = /** @class */ (function (_super) {
    tslib_1.__extends(Dashboard, _super);
    function Dashboard(props) {
        var _this = _super.call(this, props) || this;
        _this._service = null;
        _this._primaryKey = null;
        _this._router = null;
        _this._currentSystemUserId = null;
        _this._onSortColumn = function (columnKey) {
            _this.setState(function (prevState) { return ({
                sortField: columnKey,
                sortDesc: prevState.sortField === columnKey ? !prevState.sortDesc : false,
                currentPage: 1,
            }); });
        };
        _this.state = {
            records: [], loading: false, saving: false,
            error: null, saveError: null, saveSuccess: null,
            viewRecord: null, editRecord: null, editValues: {}, editInitialValues: {},
            poDate: undefined, resolvedApprovers: null, loadingApprovers: false,
            deleteRecord: null, historyRecord: null,
            showFieldInspector: false,
            detectedFields: [], allFields: [],
            searchText: '', statusFilter: 'all',
            currentPage: 1, pageSize: 10,
            roleFilteredIds: null, roleFilterLoading: false,
            requesterTab: 'pending',
            managerApprovedIds: new Set(),
            masterRequestData: [],
            approverApprovedIds: new Set(),
            directReportIds: new Set(),
            takeActionRecord: null, takeActionComments: '', takeActionSaving: false, takeActionError: null,
            sortField: 'sct_submitted_date',
            sortDesc: true,
            deepLinkProcessed: false,
            sodHiddenCount: 0,
        };
        return _this;
    }
    Dashboard.prototype.componentDidMount = function () {
        this._service = this.props.dataverseService;
        this._router = new ApprovalRoutingService_1.ApprovalRoutingService(this.props.dataverseService);
        void this._load();
        void this._resolveCurrentUser();
        // Apply initial role filter after data loads — handled inside _load() callback
    };
    Dashboard.prototype.componentDidUpdate = function (prevProps, prevState) {
        var _this = this;
        if (prevProps.activeRole !== this.props.activeRole) {
            // Immediately clear grid data and state to prevent data bleed across roles
            this.setState({
                records: [],
                masterRequestData: [],
                loading: true,
                requesterTab: 'pending',
                currentPage: 1,
                managerApprovedIds: new Set(),
                approverApprovedIds: new Set(),
                directReportIds: new Set(),
                roleFilteredIds: null,
            }, function () {
                var role = _this.props.activeRole || 'admin';
                if (role === 'manager') {
                    void _this._loadManagerData();
                }
                else if (role === 'approver_tier1' || role === 'approver_tier2' || role === 'approver_tier3') {
                    void _this._loadApproverData(role);
                    void _this._load();
                }
                else {
                    void _this._applyRoleFilter(role);
                    void _this._load();
                }
            });
        }
        // Check deep-link parameter once data is populated
        if (this.props.deepLinkRequestId && !this.state.deepLinkProcessed && (this.state.records.length > 0 || this.state.masterRequestData.length > 0)) {
            this._processDeepLink();
        }
        // useEffect([selectedRole, activeTab]) equivalent: reset to page 1 when tab changes
        if (prevState.requesterTab !== this.state.requesterTab) {
            this.setState({ currentPage: 1 });
        }
    };
    /**
     * Builds `roleFilteredIds` based on the selected role.
     *
     * - 'requester'      → Set of IDs where _sct_requester_value === currentSystemUserId
     * - 'approver_tier1' → Set of In Review IDs pending at Tier 1 (async cross-reference)
     * - 'approver_tier2' → Set of In Review IDs pending at Tier 2 (async cross-reference)
     * - 'admin' / other  → null (no restriction — show all records)
     */
    Dashboard.prototype._applyRoleFilter = function (role) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var userId_1, ids_1, tier, pendingIds, _a, ids_2;
            var _this = this;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        // Admin sees everything — clear any restriction immediately
                        if (!role || role === 'admin') {
                            this.setState({ roleFilteredIds: null, roleFilterLoading: false, currentPage: 1 });
                            return [2 /*return*/];
                        }
                        if (role === 'requester') {
                            userId_1 = this._currentSystemUserId;
                            if (!userId_1) {
                                // User not yet resolved — wait for _resolveCurrentUser to finish, then re-filter
                                this.setState({ roleFilteredIds: new Set(), currentPage: 1 });
                                return [2 /*return*/];
                            }
                            ids_1 = new Set();
                            this.state.records.forEach(function (r) {
                                var _a;
                                var requesterId = String((_a = r['_sct_requester_value']) !== null && _a !== void 0 ? _a : '');
                                if (requesterId === userId_1) {
                                    var pk = _this._getRecordId(r);
                                    if (pk)
                                        ids_1.add(pk);
                                }
                            });
                            this.setState({ roleFilteredIds: ids_1, roleFilterLoading: false, currentPage: 1 });
                            return [2 /*return*/];
                        }
                        if (role === 'manager') {
                            // Manager tab filtering is handled entirely in _filteredRecords() per tab.
                            // Set null so all records pass the role-filter step, then _filteredRecords() splits by tab.
                            this.setState({ roleFilteredIds: null, roleFilterLoading: false, currentPage: 1 });
                            return [2 /*return*/];
                        }
                        if (!(role === 'approver_tier1' || role === 'approver_tier2' || role === 'approver_tier3')) return [3 /*break*/, 5];
                        if (!this._service)
                            return [2 /*return*/];
                        this.setState({ roleFilterLoading: true });
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        tier = role === 'approver_tier1' ? 1 : role === 'approver_tier2' ? 2 : 3;
                        return [4 /*yield*/, this._service.getTierPendingRequestIds(tier)];
                    case 2:
                        pendingIds = _b.sent();
                        this.setState({
                            roleFilteredIds: new Set(pendingIds),
                            roleFilterLoading: false,
                            currentPage: 1,
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _b.sent();
                        ids_2 = new Set();
                        this.state.records.forEach(function (r) {
                            var _a;
                            var statusLabel = String((_a = r['sct_request_status@OData.Community.Display.V1.FormattedValue']) !== null && _a !== void 0 ? _a : '');
                            if (statusLabel === 'In Review') {
                                var pk = _this._getRecordId(r);
                                if (pk)
                                    ids_2.add(pk);
                            }
                        });
                        this.setState({ roleFilteredIds: ids_2, roleFilterLoading: false, currentPage: 1 });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                    case 5:
                        // Unknown role key — show all
                        this.setState({ roleFilteredIds: null, roleFilterLoading: false, currentPage: 1 });
                        return [2 /*return*/];
                }
            });
        });
    };
    Dashboard.prototype._resolveCurrentUser = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var ctx, aadId, email, user, role, _a;
            var _b, _c, _d, _e, _f;
            return tslib_1.__generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        if (!this._service)
                            return [2 /*return*/];
                        _g.label = 1;
                    case 1:
                        _g.trys.push([1, 6, , 7]);
                        ctx = this.props.context;
                        aadId = (_d = (_c = (_b = ctx === null || ctx === void 0 ? void 0 : ctx.pageContext) === null || _b === void 0 ? void 0 : _b.aadInfo) === null || _c === void 0 ? void 0 : _c.userId) === null || _d === void 0 ? void 0 : _d.toString();
                        email = (_f = (_e = ctx === null || ctx === void 0 ? void 0 : ctx.pageContext) === null || _e === void 0 ? void 0 : _e.user) === null || _f === void 0 ? void 0 : _f.email;
                        user = null;
                        if (!aadId) return [3 /*break*/, 3];
                        return [4 /*yield*/, this._service.getCurrentSystemUser(aadId)];
                    case 2:
                        user = _g.sent();
                        _g.label = 3;
                    case 3:
                        if (!(!user && email)) return [3 /*break*/, 5];
                        return [4 /*yield*/, this._service.getSystemUserByEmail(email)];
                    case 4:
                        user = _g.sent();
                        _g.label = 5;
                    case 5:
                        if (user) {
                            this._currentSystemUserId = user.systemuserid;
                            role = this.props.activeRole;
                            if (role === 'manager') {
                                void this._loadManagerData();
                            }
                            else if (role === 'approver_tier1' || role === 'approver_tier2' || role === 'approver_tier3') {
                                void this._loadApproverData(role);
                            }
                            else {
                                this.setState({ currentPage: 1 });
                            }
                        }
                        return [3 /*break*/, 7];
                    case 6:
                        _a = _g.sent();
                        return [3 /*break*/, 7];
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    Dashboard.prototype._loadManagerApprovedIds = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var ids, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!this._service || !this._currentSystemUserId)
                            return [2 /*return*/];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._service.getMyApprovedRequestIds(this._currentSystemUserId)];
                    case 2:
                        ids = _b.sent();
                        this.setState({ managerApprovedIds: new Set(ids) });
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _b.sent();
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Three-source data load for the Manager role.
     * Fetches authored requests, pending queue, and approval-log history,
     * merges and deduplicates into masterRequestData.
     * Also refreshes managerApprovedIds in the same call so both are in sync.
     */
    Dashboard.prototype._loadManagerData = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a, masterData, approvedIds, reports, _b;
            return tslib_1.__generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        if (!this._service || !this._currentSystemUserId)
                            return [2 /*return*/];
                        this.setState({ loading: true });
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, Promise.all([
                                this._service.getManagerMasterRequests(this._currentSystemUserId),
                                this._service.getMyApprovedRequestIds(this._currentSystemUserId),
                                this._service.getDirectReportUserIds(this._currentSystemUserId).catch(function () { return []; }),
                            ])];
                    case 2:
                        _a = _c.sent(), masterData = _a[0], approvedIds = _a[1], reports = _a[2];
                        // Update primary key detection from merged records if not yet set
                        if (!this._primaryKey && masterData.length > 0) {
                            this._primaryKey = findPrimaryKey(masterData);
                        }
                        this.setState({
                            masterRequestData: masterData,
                            records: masterData,
                            managerApprovedIds: new Set(approvedIds),
                            directReportIds: new Set(reports.map(function (id) { return id.toLowerCase().replace(/[{}]/g, ''); })),
                            // Ensure column detection covers master data fields
                            detectedFields: detectDisplayFields(masterData.length > 0 ? masterData : this.state.records),
                            allFields: allUsableFields(masterData.length > 0 ? masterData : this.state.records),
                            loading: false,
                            currentPage: 1,
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        _b = _c.sent();
                        this.setState({ loading: false });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Loads both pending queue IDs and already-approved IDs for a tier approver role.
     * Tier approver 'My Requests' tab shows strictly what they approved at their tier.
     */
    Dashboard.prototype._loadApproverData = function (role) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var userId, email, user, _a, tierIndex, approvedIds, _b;
            return tslib_1.__generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        if (!this._service)
                            return [2 /*return*/];
                        this.setState({ roleFilterLoading: true });
                        userId = this._currentSystemUserId;
                        if (!!userId) return [3 /*break*/, 5];
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 4, , 5]);
                        email = this.props.context.pageContext.user.email;
                        if (!email) return [3 /*break*/, 3];
                        return [4 /*yield*/, this._service.getSystemUserByEmail(email)];
                    case 2:
                        user = _c.sent();
                        if (user) {
                            userId = user.systemuserid;
                            this._currentSystemUserId = userId;
                        }
                        _c.label = 3;
                    case 3: return [3 /*break*/, 5];
                    case 4:
                        _a = _c.sent();
                        return [3 /*break*/, 5];
                    case 5:
                        if (!userId) {
                            this.setState({ roleFilterLoading: false });
                            return [2 /*return*/];
                        }
                        tierIndex = role === 'approver_tier1' ? 1 : role === 'approver_tier2' ? 2 : 3;
                        _c.label = 6;
                    case 6:
                        _c.trys.push([6, 8, , 9]);
                        return [4 /*yield*/, this._service.getMyTierApprovedRequestIds(userId, tierIndex)];
                    case 7:
                        approvedIds = _c.sent();
                        this.setState({
                            approverApprovedIds: new Set(approvedIds),
                            roleFilteredIds: null, // not needed; Review Queue uses cr2a4_pendingtier filter
                            roleFilterLoading: false,
                            currentPage: 1,
                        });
                        return [3 /*break*/, 9];
                    case 8:
                        _b = _c.sent();
                        this.setState({ roleFilterLoading: false });
                        return [3 /*break*/, 9];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    // ── Data ──────────────────────────────────────────────────────────────────
    Dashboard.prototype._load = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var records, err_1;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this._service)
                            return [2 /*return*/];
                        this.setState({ loading: true, error: null, saveSuccess: null });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._service.getRawRequests()];
                    case 2:
                        records = _a.sent();
                        this._primaryKey = findPrimaryKey(records);
                        this.setState({
                            records: records,
                            masterRequestData: records,
                            detectedFields: detectDisplayFields(records),
                            allFields: allUsableFields(records),
                            loading: false,
                            currentPage: 1,
                        }, function () {
                            if (_this.props.deepLinkRequestId && !_this.state.deepLinkProcessed) {
                                _this._processDeepLink();
                            }
                            var role = _this.props.activeRole;
                            if (role === 'manager') {
                                if (_this._currentSystemUserId)
                                    void _this._loadManagerData();
                            }
                            else if (role === 'approver_tier1' || role === 'approver_tier2' || role === 'approver_tier3') {
                                if (_this._currentSystemUserId)
                                    void _this._loadApproverData(role);
                                else
                                    void _this._applyRoleFilter(role); // pending IDs at minimum until user resolves
                            }
                            else if (role && role !== 'admin') {
                                void _this._applyRoleFilter(role);
                            }
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        err_1 = _a.sent();
                        this.setState({ error: err_1 instanceof Error ? err_1.message : 'Load failed', loading: false });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    Dashboard.prototype._processDeepLink = function () {
        var _this = this;
        var target = (this.props.deepLinkRequestId || '').trim().toLowerCase();
        if (!target)
            return;
        var pool = this.state.masterRequestData.length > 0 ? this.state.masterRequestData : this.state.records;
        var found = pool.find(function (r) {
            var ref = String(r['sct_sct_reference'] || '').trim().toLowerCase();
            var id = String(_this._getRecordId(r) || '').trim().toLowerCase();
            return ref === target || id === target;
        });
        if (found) {
            this.setState({ viewRecord: found, deepLinkProcessed: true });
        }
    };
    Dashboard.prototype._getRecordId = function (record) {
        if (!this._primaryKey)
            return null;
        var v = record[this._primaryKey];
        return typeof v === 'string' ? v : null;
    };
    // ── Filtering / pagination ────────────────────────────────────────────────
    /**
     * Pure role × tab filter — all 4 inputs explicit so every call site is self-documenting.
     *
     * Lifecycle contract (must never break):
     *   Requester submits → assigned to Manager via _cr2a4_manager_value
     *   → Manager sees it in Review Queue and acts
     *   → acted requests move to Manager's My Requests (processed)
     *   → Requester tracks current status in their My Requests tab
     *
     * managerApprovedIds / approverApprovedIds are read from this.state — they are
     * loaded async and are not caller-supplied data slices.
     *
     * Class-component equivalent of:
     *   function applyFilters(baseData, selectedRole, activeTab, currentUserId) { ... }
     * useEffect([selectedRole, activeTab]) equivalent: componentDidUpdate watches both.
     */
    /**
     * Strict Role Isolation Filter
     *
     * 1. REQUESTER ROLE: Show ONLY requests created by this user (reqId === safeUserId).
     * 2. MANAGER ROLE: Show ONLY requests managed by this user (managerId === safeUserId).
     *    STRICT ISOLATION: Hide the user's personal requests (if reqId === safeUserId return false).
     * 3. TIER APPROVER ROLES: Show requests pending at the tier or approved in logs.
     *    STRICT ISOLATION: Hide the user's personal requests (if reqId === safeUserId return false).
     */
    Dashboard.prototype.applyFilters = function (masterRequestData, selectedRole, selectedTab, currentUserId) {
        var _this = this;
        var _a, _b, _c, _d;
        if (!masterRequestData || masterRequestData.length === 0) {
            return [];
        }
        // Normalize IDs for safe comparison
        var normalizeId = function (id) { return String(id || '').toLowerCase().replace(/[{}]/g, '').trim(); };
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        var ctxAny = this.props.context;
        var safeAadId = normalizeId((_d = (_c = (_b = (_a = ctxAny === null || ctxAny === void 0 ? void 0 : ctxAny.pageContext) === null || _a === void 0 ? void 0 : _a.aadInfo) === null || _b === void 0 ? void 0 : _b.userId) === null || _c === void 0 ? void 0 : _c.toString) === null || _d === void 0 ? void 0 : _d.call(_c));
        var safeUserId = normalizeId(currentUserId) || safeAadId;
        // Normalize role string for robust comparisons (supports 'Requester', 'requester', 'Manager', 'manager', 'Tier 1 Approver', etc.)
        var normRole = (selectedRole || '').trim();
        var normRoleLower = normRole.toLowerCase();
        // Role classification
        var isRequesterRole = normRoleLower === 'requester';
        var isManagerRole = normRoleLower === 'manager';
        var isTierRole = normRoleLower.indexOf('tier') !== -1 || normRoleLower.indexOf('approver') !== -1 || normRole.indexOf('Tier') !== -1;
        // Normalize tab string for robust comparisons (supports 'Pending', 'pending', 'Review Queue', 'Completed', 'completed', 'My Requests', 'myRequests')
        var normTab = (selectedTab || '').trim();
        var normTabLower = normTab.toLowerCase();
        var isPendingTab = normTab === 'Review Queue' || normTab === 'Pending' || normTabLower === 'review queue' || normTabLower === 'pending' || normTabLower === 'reviewqueue' || !normTab;
        var isCompletedOrMyRequests = normTab === 'My Requests' || normTab === 'Completed' || normTabLower === 'my requests' || normTabLower === 'completed' || normTabLower === 'myrequests';
        var hiddenBySodCount = 0;
        var filtered = masterRequestData.filter(function (req) {
            var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t, _u, _v;
            // Normalize IDs for safe comparison
            var rawReq = (_e = (_d = (_c = (_b = (_a = req['_sct_requester_value']) !== null && _a !== void 0 ? _a : req['_createdby_value']) !== null && _b !== void 0 ? _b : req['sct_requester']) !== null && _c !== void 0 ? _c : req['_sct_requestorid_value']) !== null && _d !== void 0 ? _d : req['createdby']) !== null && _e !== void 0 ? _e : '';
            var reqId = rawReq ? String(rawReq).toLowerCase().replace(/[{}]/g, '') : '';
            var createdById = req['_createdby_value'] ? String(req['_createdby_value']).toLowerCase().replace(/[{}]/g, '') : '';
            var rawMgr = (_l = (_k = (_j = (_h = (_g = (_f = req['_sct_manager_value']) !== null && _f !== void 0 ? _f : req['_cr2a4_manager_value']) !== null && _g !== void 0 ? _g : req['_manager_value']) !== null && _h !== void 0 ? _h : req['sct_manager']) !== null && _j !== void 0 ? _j : req['cr2a4_manager']) !== null && _k !== void 0 ? _k : req['manager']) !== null && _l !== void 0 ? _l : '';
            var managerId = rawMgr ? String(rawMgr).toLowerCase().replace(/[{}]/g, '') : '';
            var isDirectReport = Boolean(_this.state.directReportIds && (_this.state.directReportIds.has(reqId) || (createdById && _this.state.directReportIds.has(createdById))));
            if (!managerId && isDirectReport) {
                managerId = safeUserId;
            }
            var mgrFormatted = ((0, DataverseService_1.strVal)(req, '_cr2a4_manager_value') || (0, DataverseService_1.strVal)(req, '_sct_manager_value') || (0, DataverseService_1.strVal)(req, '_manager_value') || (0, DataverseService_1.strVal)(req, 'cr2a4_manager') || (0, DataverseService_1.strVal)(req, 'sct_manager') || (0, DataverseService_1.strVal)(req, 'manager') || '').toLowerCase();
            var currentUserName = (_this.props.context.pageContext.user.displayName || '').toLowerCase().trim();
            var currentUserEmail = (_this.props.context.pageContext.user.email || '').toLowerCase().trim();
            var isMgrByName = Boolean(currentUserName) && (mgrFormatted.indexOf(currentUserName) !== -1 ||
                (Boolean(currentUserEmail) && mgrFormatted.indexOf(currentUserEmail) !== -1));
            var isManagerOfReq = (managerId === safeUserId) || isMgrByName || isDirectReport;
            // Resolve status
            var rawStatus = typeof req['sct_request_status'] === 'number'
                ? req['sct_request_status']
                : parseInt(String((_o = (_m = req['sct_request_status']) !== null && _m !== void 0 ? _m : req['sct_requeststatus']) !== null && _o !== void 0 ? _o : ''), 10);
            var statusLabel = (0, DataverseService_1.strVal)(req, 'sct_request_status') || (0, DataverseService_1.strVal)(req, 'sct_requeststatus') || '';
            var statusLower = statusLabel.toLowerCase();
            var isInReview = statusLower === 'in review' ||
                statusLower.indexOf('in review') !== -1 ||
                rawStatus === 854000003 || rawStatus === 854000002 ||
                req['sct_requeststatus'] === 'In Review' ||
                String((_p = req['sct_requeststatus']) !== null && _p !== void 0 ? _p : '').toLowerCase().indexOf('in review') !== -1;
            var isSubmitted = statusLower === 'submitted' ||
                statusLower.indexOf('submitted') !== -1 ||
                rawStatus === 854000002 || rawStatus === 854000001 ||
                req['sct_requeststatus'] === 'Submitted' ||
                String((_q = req['sct_requeststatus']) !== null && _q !== void 0 ? _q : '').toLowerCase().indexOf('submitted') !== -1;
            var statusString = isInReview ? 'In Review' : (isSubmitted ? 'Submitted' : (statusLabel || String(rawStatus)));
            req['sct_requeststatus'] = statusString;
            // Resolve pending tier
            var rawPt = String(req['cr2a4_pendingtier'] ||
                req['cr2a4_PendingTier'] ||
                req['sct_pendingtier'] ||
                req['pendingtier'] ||
                '').trim();
            if (!rawPt && statusLower.indexOf('tier') !== -1) {
                var m = statusLower.match(/tier\s*([123])/i);
                if (m)
                    rawPt = "Tier ".concat(m[1]);
            }
            req['sct_pendingtier'] = rawPt;
            var ptLower = rawPt.toLowerCase();
            // Record ID for approval log check
            var pk = _this._getRecordId(req);
            var isUserInApprovalLog = req['sct_approvallog_contains_user'] === true ||
                (pk !== null && _this.state.approverApprovedIds.has(pk));
            // Resolve whether the current user is the requester of this record
            var reqFormatted = ((0, DataverseService_1.strVal)(req, '_sct_requester_value') || (0, DataverseService_1.strVal)(req, 'sct_requester') || '').toLowerCase().trim();
            var reqNameFromOdata = ((0, DataverseService_1.strVal)(req, '_sct_requester_value@OData.Community.Display.V1.FormattedValue') || (0, DataverseService_1.strVal)(req, 'requester') || '').toLowerCase().trim();
            var isOwnRequest = Boolean((safeUserId && (reqId === safeUserId || createdById === safeUserId || reqId === safeAadId || createdById === safeAadId)) ||
                (currentUserName && (reqFormatted.indexOf(currentUserName) !== -1 || reqNameFromOdata.indexOf(currentUserName) !== -1)) ||
                (currentUserEmail && (reqFormatted.indexOf(currentUserEmail) !== -1 || reqId.indexOf(currentUserEmail) !== -1)));
            // ── 1. Global Segregation of Duties (SoD) Evaluation ───────────────────
            // When Solo Test Mode is OFF (Strict Mode), an approver cannot self-approve their own request.
            // We flag the record so the UI clearly shows the request at its active tier while locking self-approval.
            var isMyRequestsTab = normTabLower === 'my requests' || normTabLower === 'myrequests';
            var allowSelfApproval = _this.props.soloTestMode !== false;
            var isSodLocked = !allowSelfApproval && !isRequesterRole && !isMyRequestsTab && isPendingTab && isOwnRequest;
            req['_is_sod_locked'] = isSodLocked;
            if (isSodLocked) {
                hiddenBySodCount++;
            }
            // ── 2. The Requester View / "My Requests" Tab ───────────────────────────
            // If activeRole === 'Requester' OR user is on the "My Requests" tab:
            // Return true for all requests authored by this user
            if (isRequesterRole || isMyRequestsTab) {
                return isOwnRequest;
            }
            // ── 3. The Manager View ────────────────────────────────────────────────
            // If activeRole === 'Manager':
            if (isManagerRole) {
                var rawMgr_1 = (_v = (_u = (_t = (_s = (_r = req['_cr2a4_manager_value']) !== null && _r !== void 0 ? _r : req['_sct_manager_value']) !== null && _s !== void 0 ? _s : req['cr2a4_manager']) !== null && _t !== void 0 ? _t : req['sct_manager']) !== null && _u !== void 0 ? _u : req['_manager_value']) !== null && _v !== void 0 ? _v : '';
                var managerId_1 = normalizeId(rawMgr_1);
                var isPendingManagerTier = ptLower === 'manager' || (!rawPt && (rawStatus === 854000002 || rawStatus === 854000001 || statusLower === 'submitted'));
                var isManagerAssigned = (managerId_1 && (managerId_1 === safeUserId || isManagerOfReq))
                    || isMgrByName
                    || isDirectReport;
                if (isPendingTab) {
                    // In Strict SoD Mode, own requests never show in Pending Approval (they live in My Requests)
                    if (!allowSelfApproval && isOwnRequest)
                        return false;
                    return isManagerAssigned && isPendingManagerTier && !(pk !== null && _this.state.managerApprovedIds && _this.state.managerApprovedIds.has(pk));
                }
                if (normTabLower === 'completed') {
                    if (isOwnRequest)
                        return false;
                    return (pk !== null && _this.state.managerApprovedIds && _this.state.managerApprovedIds.has(pk)) || (isManagerAssigned && !isPendingManagerTier);
                }
                return isManagerAssigned && isPendingManagerTier && (!allowSelfApproval ? !isOwnRequest : true);
            }
            // ── 4. The Tier 1, Tier 2, and Tier 3 Views ────────────────────────────
            if (selectedRole.includes('Tier') || selectedRole.includes('approver') || selectedRole.includes('Approver') || normRoleLower.includes('tier') || normRoleLower.includes('approver')) {
                var expectedTierValue = '';
                if (selectedRole.includes('1') || normRoleLower.includes('1'))
                    expectedTierValue = 'Tier 1';
                else if (selectedRole.includes('2') || normRoleLower.includes('2'))
                    expectedTierValue = 'Tier 2';
                else if (selectedRole.includes('3') || normRoleLower.includes('3'))
                    expectedTierValue = 'Tier 3';
                var tierNum = expectedTierValue.replace(/\D/g, '');
                var ptClean = ptLower.replace(/[^a-z0-9]/g, '');
                var expectedTierClean = expectedTierValue.toLowerCase().replace(/[^a-z0-9]/g, '');
                var matchesTier = rawPt === expectedTierValue || ptClean === expectedTierClean || ptClean === tierNum || ptClean === ('tier' + tierNum);
                var matchesStatus = req['sct_requeststatus'] === 'In Review' || isInReview || rawStatus === 854000003 || rawStatus === 854000002;
                if (isPendingTab) {
                    // In Strict SoD Mode, approver's own requests do not clutter their review queue (they view them in My Requests)
                    if (!allowSelfApproval && isOwnRequest)
                        return false;
                    return matchesTier && matchesStatus;
                }
                if (normTabLower === 'completed') {
                    return req['sct_approvallog_contains_user'] === true || isUserInApprovalLog;
                }
                return matchesTier && matchesStatus;
            }
            // Admin role fallback
            if (normRoleLower === 'admin') {
                return true;
            }
            return false; // Fallback
        });
        return filtered;
    };
    Dashboard.prototype._applyFilters = function (baseData, selectedRole, activeTab, currentUserId) {
        return this.applyFilters(baseData, selectedRole, activeTab, currentUserId);
    };
    Dashboard.prototype._filteredRecords = function () {
        var _a = this.state, records = _a.records, masterRequestData = _a.masterRequestData, detectedFields = _a.detectedFields, searchText = _a.searchText, statusFilter = _a.statusFilter, requesterTab = _a.requesterTab;
        var role = this.props.activeRole || 'admin';
        var userId = this._currentSystemUserId;
        // Use masterRequestData if available, otherwise records
        var baseData = (role === 'manager' && masterRequestData.length > 0)
            ? masterRequestData
            : (masterRequestData.length > 0 ? masterRequestData : records);
        // Strict role isolation: applyFilters always returns a fresh slice
        var result = this.applyFilters(baseData, role, requesterTab, userId);
        // ── Status dropdown filter ────────────────────────────────────────────────
        if (statusFilter !== 'all') {
            var normFilter_1 = statusFilter.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
            result = result.filter(function (r) {
                var itemStatus = (0, dashboardUtils_1.getRecordStatusLabel)(r);
                var normItem = itemStatus.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
                return normItem === normFilter_1;
            });
        }
        // ── Full-text search ──────────────────────────────────────────────────────
        if (searchText.trim()) {
            var q_1 = searchText.trim().toLowerCase();
            result = result.filter(function (r) {
                return detectedFields.some(function (f) { return (0, DataverseService_1.strVal)(r, f).toLowerCase().indexOf(q_1) !== -1; });
            });
        }
        // ── Sorting ───────────────────────────────────────────────────────────────
        var _b = this.state, sortField = _b.sortField, sortDesc = _b.sortDesc;
        if (sortField) {
            result.sort(function (a, b) {
                var _a, _b;
                var valA = a[sortField];
                var valB = b[sortField];
                if (sortField.startsWith('_') && sortField.endsWith('_value')) {
                    valA = (_a = a["".concat(sortField, "@OData.Community.Display.V1.FormattedValue")]) !== null && _a !== void 0 ? _a : valA;
                    valB = (_b = b["".concat(sortField, "@OData.Community.Display.V1.FormattedValue")]) !== null && _b !== void 0 ? _b : valB;
                }
                if (sortField === 'sct_estimated_total_value' || sortField.indexOf('amount') !== -1 || sortField.indexOf('value') !== -1) {
                    var numA = parseFloat(String(valA !== null && valA !== void 0 ? valA : 0)) || 0;
                    var numB = parseFloat(String(valB !== null && valB !== void 0 ? valB : 0)) || 0;
                    return sortDesc ? numB - numA : numA - numB;
                }
                if (sortField === 'sct_submitted_date' || sortField === 'createdon' || sortField === 'modifiedon' || sortField === 'sct_date_needed_by') {
                    var dateA = valA ? new Date(String(valA)).getTime() : 0;
                    var dateB = valB ? new Date(String(valB)).getTime() : 0;
                    return sortDesc ? dateB - dateA : dateA - dateB;
                }
                var strA = String(valA !== null && valA !== void 0 ? valA : '').toLowerCase();
                var strB = String(valB !== null && valB !== void 0 ? valB : '').toLowerCase();
                var cmp = strA.localeCompare(strB);
                return sortDesc ? -cmp : cmp;
            });
        }
        return result;
    };
    Dashboard.prototype._exportToCsv = function (records) {
        if (!records || records.length === 0)
            return;
        var headers = [
            'Reference',
            'Status',
            'Pending Tier',
            'Requester',
            'Business Unit',
            'Site',
            'Purchase Type',
            'Estimated Value',
            'Urgency',
            'Date Needed By',
            'Submitted Date'
        ];
        var escapeCsv = function (val) {
            var str = val === null || val === undefined ? '' : String(val);
            if (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1) {
                return "\"".concat(str.replace(/"/g, '""'), "\"");
            }
            return str;
        };
        var rows = records.map(function (r) {
            var ref = (0, DataverseService_1.strVal)(r, 'sct_sct_reference') || (0, DataverseService_1.strVal)(r, 'sct_reference') || '';
            var status = (0, dashboardUtils_1.getRecordStatusLabel)(r);
            var pendingTier = (0, DataverseService_1.strVal)(r, 'cr2a4_pendingtier') || (0, DataverseService_1.strVal)(r, 'sct_pendingtier') || (0, DataverseService_1.strVal)(r, 'cr2a4_PendingTier') || '';
            var requester = (0, DataverseService_1.strVal)(r, '_sct_requester_value@OData.Community.Display.V1.FormattedValue')
                || (0, DataverseService_1.strVal)(r, 'sct_requester@OData.Community.Display.V1.FormattedValue')
                || (0, DataverseService_1.strVal)(r, 'requesterName')
                || (0, DataverseService_1.strVal)(r, 'sct_requester')
                || '';
            var bu = (0, dashboardUtils_1.getChoiceLabel)(r, 'sct_business_unit', dashboardUtils_1.BUSINESS_UNIT_LABELS);
            var site = (0, DataverseService_1.strVal)(r, '_sct_site_value@OData.Community.Display.V1.FormattedValue')
                || (0, DataverseService_1.strVal)(r, 'sct_site@OData.Community.Display.V1.FormattedValue')
                || (0, DataverseService_1.strVal)(r, 'site')
                || '';
            var ptype = (0, dashboardUtils_1.getChoiceLabel)(r, 'sct_purchase_type', dashboardUtils_1.PURCHASE_TYPE_LABELS);
            var rawAmt = r['sct_estimated_total_value'];
            var numAmt = parseFloat(String(rawAmt !== null && rawAmt !== void 0 ? rawAmt : 0));
            var val = !isNaN(numAmt) && numAmt > 0 ? (0, dashboardUtils_1.formatCurrency)(numAmt) : (rawAmt !== undefined && rawAmt !== null ? String(rawAmt) : '$0.00');
            var urgency = (0, dashboardUtils_1.getChoiceLabel)(r, 'sct_urgency', dashboardUtils_1.URGENCY_LABELS);
            var neededBy = r['sct_date_needed_by'] ? (0, dashboardUtils_1.formatDate)(String(r['sct_date_needed_by'])) : '';
            var submitted = r['sct_submitted_date'] ? (0, dashboardUtils_1.formatDate)(String(r['sct_submitted_date'])) : (r['createdon'] ? (0, dashboardUtils_1.formatDate)(String(r['createdon'])) : '');
            return [
                escapeCsv(ref),
                escapeCsv(status),
                escapeCsv(pendingTier),
                escapeCsv(requester),
                escapeCsv(bu),
                escapeCsv(site),
                escapeCsv(ptype),
                escapeCsv(val),
                escapeCsv(urgency),
                escapeCsv(neededBy),
                escapeCsv(submitted)
            ].join(',');
        });
        var csvContent = '\uFEFF' + tslib_1.__spreadArray([headers.join(',')], rows, true).join('\r\n');
        var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        var url = URL.createObjectURL(blob);
        var link = document.createElement('a');
        link.setAttribute('href', url);
        link.setAttribute('download', "Spend_Control_Tracker_Export_".concat(new Date().toISOString().slice(0, 10), ".csv"));
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    };
    Dashboard.prototype._pagedRecords = function (filtered) {
        var _a = this.state, currentPage = _a.currentPage, pageSize = _a.pageSize;
        var start = (currentPage - 1) * pageSize;
        return filtered.slice(start, start + pageSize);
    };
    // ── Stat counts ───────────────────────────────────────────────────────────
    Dashboard.prototype._counts = function (filtered) {
        var _this = this;
        var role = this.props.activeRole || 'admin';
        var userId = this._currentSystemUserId;
        var baseData = role === 'manager' ? this.state.masterRequestData : this.state.records;
        // For approver/manager roles, compute stats across Pending and Completed views
        if (role === 'manager' || role === 'approver_tier1' || role === 'approver_tier2' || role === 'approver_tier3') {
            var pendingSet = this._applyFilters(baseData, role, 'pending', userId);
            var completedSet = this._applyFilters(baseData, role, 'completed', userId);
            var uniqueIds_1 = new Set();
            pendingSet.forEach(function (r) { var id = _this._getRecordId(r); if (id)
                uniqueIds_1.add(id); });
            completedSet.forEach(function (r) { var id = _this._getRecordId(r); if (id)
                uniqueIds_1.add(id); });
            var spend_1 = 0;
            var targetSpendSet = this.state.requesterTab === 'completed' ? completedSet : (pendingSet.length > 0 ? pendingSet : completedSet);
            targetSpendSet.forEach(function (r) {
                var _a;
                var amt = parseFloat(String((_a = r['sct_estimated_total_value']) !== null && _a !== void 0 ? _a : 0));
                if (!isNaN(amt))
                    spend_1 += amt;
            });
            return {
                total: uniqueIds_1.size,
                pending: pendingSet.length,
                approved: completedSet.length,
                spend: spend_1,
            };
        }
        var pending = 0, approved = 0, spend = 0;
        filtered.forEach(function (r) {
            var _a;
            var statusLabel = (0, dashboardUtils_1.getRecordStatusLabel)(r);
            if (statusLabel === 'Submitted' || statusLabel === 'In Review')
                pending++;
            if (statusLabel === 'PO Issued')
                approved++;
            var amt = parseFloat(String((_a = r['sct_estimated_total_value']) !== null && _a !== void 0 ? _a : 0));
            if (!isNaN(amt))
                spend += amt;
        });
        return { total: filtered.length, pending: pending, approved: approved, spend: spend };
    };
    // ── Columns ───────────────────────────────────────────────────────────────
    Dashboard.prototype._buildColumns = function (_statusValues) {
        var _this = this;
        var allFields = this.state.allFields;
        // Only include columns that actually exist in the loaded data
        var visible = DASHBOARD_VISIBLE_COLUMNS.filter(function (col) { return allFields.indexOf(col.field) !== -1 ||
            allFields.some(function (f) { return f === col.field; }); });
        var dataCols = visible.map(function (_a) {
            var field = _a.field, minW = _a.minW, maxW = _a.maxW;
            return ({
                key: field,
                name: fieldLabel(field),
                fieldName: field,
                minWidth: minW,
                maxWidth: maxW,
                isResizable: true,
                isSorted: _this.state.sortField === field,
                isSortedDescending: _this.state.sortField === field && _this.state.sortDesc,
                onColumnClick: function (_ev, col) { return col && _this._onSortColumn(col.key); },
                onRender: function (item) {
                    // sct_submitted_date: fall back to createdon for Draft records (not yet submitted)
                    var rawVal = (0, DataverseService_1.strVal)(item, field);
                    var val = (field === 'sct_submitted_date' && !rawVal)
                        ? (0, DataverseService_1.strVal)(item, 'createdon')
                        : rawVal;
                    if (field === 'sct_business_unit') {
                        val = (0, dashboardUtils_1.getChoiceLabel)(item, 'sct_business_unit', dashboardUtils_1.BUSINESS_UNIT_LABELS);
                    }
                    else if (field === 'sct_purchase_type') {
                        val = (0, dashboardUtils_1.getChoiceLabel)(item, 'sct_purchase_type', dashboardUtils_1.PURCHASE_TYPE_LABELS);
                    }
                    else if (field === 'sct_urgency') {
                        val = (0, dashboardUtils_1.getChoiceLabel)(item, 'sct_urgency', dashboardUtils_1.URGENCY_LABELS);
                    }
                    if (!val)
                        return React.createElement("span", { style: { color: '#C8C6C4' } }, "\u2014");
                    if (field === 'sct_estimated_total_value' || field.indexOf('amount') !== -1) {
                        var num = parseFloat(val.replace(/[^0-9.-]/g, ''));
                        return !isNaN(num)
                            ? React.createElement("span", { style: { fontWeight: 700, color: '#00897B', fontFamily: 'monospace', fontSize: 13 } }, (0, dashboardUtils_1.formatCurrency)(num))
                            : React.createElement("span", null, val);
                    }
                    if (field === 'sct_request_status' || field === 'sct_requeststatus') {
                        var statusText = (0, dashboardUtils_1.getRecordStatusLabel)(item);
                        var pt = (0, DataverseService_1.strVal)(item, 'cr2a4_pendingtier') || (0, DataverseService_1.strVal)(item, 'sct_pendingtier') || (0, DataverseService_1.strVal)(item, 'cr2a4_PendingTier');
                        var display = (statusText === 'In Review' && pt) ? "".concat(statusText, " (").concat(pt, ")") : (statusText || val);
                        // SLA Ageing calculation for pending items
                        var slaBadge = null;
                        if (statusText === 'Submitted' || statusText === 'In Review') {
                            var dateStr = (0, DataverseService_1.strVal)(item, 'sct_submitted_date') || (0, DataverseService_1.strVal)(item, 'createdon');
                            if (dateStr) {
                                var diffMs = Date.now() - new Date(dateStr).getTime();
                                var diffHours = diffMs / (1000 * 60 * 60);
                                var diffDays = Math.floor(diffHours / 24);
                                if (diffHours < 24) {
                                    slaBadge = (React.createElement("span", { style: { fontSize: 10, padding: '1px 5px', borderRadius: 4, background: '#EFF6FF', color: '#1D4ED8', border: '1px solid #BFDBFE', fontWeight: 600, display: 'inline-block', marginTop: 3 } }, "\u26A1 <24h"));
                                }
                                else if (diffDays <= 3) {
                                    slaBadge = (React.createElement("span", { style: { fontSize: 10, padding: '1px 5px', borderRadius: 4, background: '#FEF3C7', color: '#B45309', border: '1px solid #FDE68A', fontWeight: 600, display: 'inline-block', marginTop: 3 } },
                                        "\u23F3 ",
                                        diffDays,
                                        "d"));
                                }
                                else {
                                    slaBadge = (React.createElement("span", { style: { fontSize: 10, padding: '1px 5px', borderRadius: 4, background: '#FEE2E2', color: '#B91C1C', border: '1px solid #FECACA', fontWeight: 600, display: 'inline-block', marginTop: 3 } },
                                        "\u26A0\uFE0F ",
                                        diffDays,
                                        "d overdue"));
                                }
                            }
                        }
                        return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 2 } },
                            React.createElement(StatusPill, { value: display }),
                            slaBadge));
                    }
                    if (field === 'sct_urgency' || field === 'sct_purchase_type' || field === 'sct_business_unit') {
                        return React.createElement(StatusPill, { value: val });
                    }
                    if (field === 'sct_submitted_date' || field === 'createdon' || field.indexOf('date') !== -1) {
                        return React.createElement("span", { style: { color: '#605E5C', fontSize: 12 } }, (0, dashboardUtils_1.formatDate)(val));
                    }
                    return (React.createElement("span", { style: { color: '#201F1E', fontSize: 13 }, title: val.length > 36 ? val : undefined }, val.length > 36 ? val.substring(0, 34) + '…' : val));
                },
            });
        });
        var currentRoleKey = (this.props.activeRole || '').toLowerCase().trim();
        var isApproverRole = (currentRoleKey === 'manager' ||
            currentRoleKey.indexOf('tier') !== -1 ||
            currentRoleKey.indexOf('approver') !== -1);
        var actionsCol = {
            key: '_actions',
            name: 'Actions',
            minWidth: isApproverRole ? 180 : 110,
            maxWidth: isApproverRole ? 200 : 120,
            onRender: function (item) {
                var statusRaw = typeof item['sct_request_status'] === 'number'
                    ? item['sct_request_status']
                    : NaN;
                var statusLabel = (0, DataverseService_1.strVal)(item, 'sct_request_status');
                var isDraft = statusRaw === DSTATUS.DRAFT || statusLabel === 'Draft' || (!statusRaw && isNaN(statusRaw));
                var isReturned = statusRaw === DSTATUS.RETURNED || statusLabel === 'Returned';
                var canEditItem = isDraft || isReturned;
                return (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 4, flexWrap: 'nowrap' } },
                    React.createElement("button", { className: "".concat(AppShell_module_scss_1.default.actionBtn, " ").concat(AppShell_module_scss_1.default.actionBtnView), title: "View details", "aria-label": "View", onClick: function (e) { e.stopPropagation(); _this.setState({ viewRecord: item }); } },
                        React.createElement(Icon_1.Icon, { iconName: "View", style: { fontSize: 14 } })),
                    !isApproverRole && (React.createElement(React.Fragment, null,
                        React.createElement("button", { className: "".concat(AppShell_module_scss_1.default.actionBtn, " ").concat(AppShell_module_scss_1.default.actionBtnEdit), title: canEditItem ? (isReturned ? 'Edit & resubmit returned request' : 'Edit record') : 'Edit only available for Draft or Returned records', "aria-label": "Edit", disabled: !canEditItem, onClick: function (e) { e.stopPropagation(); if (canEditItem)
                                _this._openEdit(item); }, style: !canEditItem ? { opacity: 0.35, cursor: 'not-allowed' } : undefined },
                            React.createElement(Icon_1.Icon, { iconName: "Edit", style: { fontSize: 14 } })),
                        React.createElement("button", { className: "".concat(AppShell_module_scss_1.default.actionBtn, " ").concat(AppShell_module_scss_1.default.actionBtnDelete), title: isDraft ? 'Delete record' : 'Delete only available for Draft records', "aria-label": "Delete", disabled: !isDraft, onClick: function (e) { e.stopPropagation(); if (isDraft)
                                _this.setState({ deleteRecord: item }); }, style: !isDraft ? { opacity: 0.35, cursor: 'not-allowed' } : undefined },
                            React.createElement(Icon_1.Icon, { iconName: "Delete", style: { fontSize: 14 } })))),
                    React.createElement("button", { className: AppShell_module_scss_1.default.actionBtn, title: "View request history", "aria-label": "History", style: { color: '#8764B8' }, onClick: function (e) { e.stopPropagation(); _this.setState({ historyRecord: item }); } },
                        React.createElement(Icon_1.Icon, { iconName: "History", style: { fontSize: 14 } })),
                    (function () {
                        var _a;
                        if (!isApproverRole)
                            return null;
                        if (_this.state.requesterTab === 'completed')
                            return null; // Never on completed tab!
                        var statusRawNum = typeof item['sct_request_status'] === 'number'
                            ? item['sct_request_status']
                            : parseInt(String((_a = item['sct_request_status']) !== null && _a !== void 0 ? _a : ''), 10);
                        var statusStr = (0, dashboardUtils_1.getRecordStatusLabel)(item);
                        var isTerminal = (statusRawNum === DSTATUS.PO_ISSUED ||
                            statusRawNum === DSTATUS.REJECTED ||
                            statusRawNum === DSTATUS.CLOSED ||
                            statusRawNum === DSTATUS.DRAFT ||
                            statusRawNum === DSTATUS.RETURNED ||
                            statusStr === 'PO Issued' ||
                            statusStr === 'Rejected' ||
                            statusStr === 'Closed' ||
                            statusStr === 'Draft' ||
                            statusStr === 'Returned');
                        if (isTerminal)
                            return null;
                        // Check if pending tier matches approver role
                        var itemPt = ((0, DataverseService_1.strVal)(item, 'cr2a4_pendingtier') || (0, DataverseService_1.strVal)(item, 'sct_pendingtier') || '').toLowerCase();
                        var isEligible = false;
                        if (currentRoleKey === 'manager') {
                            isEligible = itemPt === 'manager' || (!itemPt && (statusRawNum === DSTATUS.SUBMITTED || statusStr === 'Submitted'));
                        }
                        else if (currentRoleKey === 'approver_tier1') {
                            isEligible = itemPt === 'tier 1' || itemPt === 'tier1' || itemPt === '1';
                        }
                        else if (currentRoleKey === 'approver_tier2') {
                            isEligible = itemPt === 'tier 2' || itemPt === 'tier2' || itemPt === '2';
                        }
                        else if (currentRoleKey === 'approver_tier3') {
                            isEligible = itemPt === 'tier 3' || itemPt === 'tier3' || itemPt === '3';
                        }
                        else {
                            isEligible = true;
                        }
                        if (!isEligible)
                            return null;
                        if (item['_is_sod_locked'] === true) {
                            return (React.createElement("button", { title: "Segregation of Duties Enforced: You submitted this request. Self-approval is locked in Strict Mode. Enable Solo Test Mode in header to test alone.", "aria-label": "SoD Locked", onClick: function (e) { e.stopPropagation(); _this.setState({ takeActionRecord: item, takeActionComments: '', takeActionError: null }); }, style: {
                                    display: 'inline-flex', alignItems: 'center', gap: 5,
                                    background: '#FFFBEB', color: '#92400E', border: '1px solid #FCD34D',
                                    borderRadius: 20, padding: '5px 12px', fontSize: 11, fontWeight: 600,
                                    cursor: 'pointer', whiteSpace: 'nowrap',
                                } },
                                React.createElement(Icon_1.Icon, { iconName: "Lock", style: { fontSize: 11, color: '#D97706' } }),
                                "SoD Locked"));
                        }
                        return (React.createElement("button", { title: "Review and take action on this request", "aria-label": "Take Action", onClick: function (e) { e.stopPropagation(); _this.setState({ takeActionRecord: item, takeActionComments: '', takeActionError: null }); }, style: {
                                display: 'inline-flex', alignItems: 'center', gap: 5,
                                background: 'linear-gradient(135deg, #0F766E 0%, #115E59 100%)',
                                color: '#fff', border: 'none', borderRadius: 20,
                                padding: '5px 13px', fontSize: 12, fontWeight: 600,
                                cursor: 'pointer', letterSpacing: '0.2px',
                                boxShadow: '0 2px 6px rgba(15,118,110,0.35)',
                                whiteSpace: 'nowrap',
                            } },
                            React.createElement(Icon_1.Icon, { iconName: "WorkFlow", style: { fontSize: 11 } }),
                            "Take Action"));
                    })()));
            },
        };
        return tslib_1.__spreadArray(tslib_1.__spreadArray([], dataCols, true), [actionsCol], false);
    };
    // ── Edit helpers ──────────────────────────────────────────────────────────
    Dashboard.prototype._openEdit = function (record) {
        var _this = this;
        var editValues = {};
        // Text fields: FormattedValue annotation is fine — it IS the display text we want to edit
        EDIT_TEXT_FIELDS.forEach(function (field) {
            if (record[field] !== undefined)
                editValues[field] = (0, DataverseService_1.strVal)(record, field);
        });
        // Number fields: use the RAW numeric value, NOT the FormattedValue annotation.
        // strVal() for sct_estimated_total_value can return "7,500.00" or "$7,500.00" from the
        // OData annotation, which breaks parseFloat (returns 7 or NaN) in _saveEdit.
        EDIT_NUMBER_FIELDS.forEach(function (field) {
            var raw = record[field];
            if (raw !== undefined && raw !== null)
                editValues[field] = String(raw);
        });
        // Date fields: use the RAW OData Edm.Date string ("YYYY-MM-DD"), not the localized
        // FormattedValue annotation ("11/3/2026"), so formatToEdmDate always gets a clean input.
        EDIT_DATE_FIELDS.forEach(function (field) {
            var raw = record[field];
            if (raw !== undefined && raw !== null)
                editValues[field] = String(raw);
        });
        // Boolean fields: use actual JS boolean → "true"/"false" strings that toBoolean() and
        // tb() both handle. Avoids dependence on the "Yes"/"No" annotation locale.
        BOOLEAN_FIELDS.forEach(function (field) {
            var raw = record[field];
            if (raw !== undefined && raw !== null) {
                editValues[field] = typeof raw === 'boolean' ? (raw ? 'true' : 'false') : (0, DataverseService_1.strVal)(record, field);
            }
        });
        // Sourcing panel: pre-populate PO date if one exists
        // Use raw Edm.Date string for reliable parsing
        var poDateRaw = record['sct_po_issue_date'];
        var poDateStr = poDateRaw ? String(poDateRaw) : '';
        var poDate;
        if (poDateStr) {
            var d = new Date(poDateStr);
            if (!isNaN(d.getTime()))
                poDate = d;
        }
        // Snapshot the values at open-time so _saveEdit can diff against them for audit logging
        var editInitialValues = {};
        Object.keys(editValues).forEach(function (k) { editInitialValues[k] = editValues[k]; });
        this.setState({ editRecord: record, editValues: editValues, editInitialValues: editInitialValues, poDate: poDate, resolvedApprovers: null, loadingApprovers: false, saveError: null }, function () { void _this._loadApprovers(record); });
    };
    /** Async — loads resolved approvers from sct_role for the record's site (Phase 8) */
    Dashboard.prototype._loadApprovers = function (record) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var siteRaw, siteId, requesterGuid, resolvedApprovers, _a;
            var _b, _c, _d, _e;
            return tslib_1.__generator(this, function (_f) {
                switch (_f.label) {
                    case 0:
                        if (!this._router)
                            return [2 /*return*/];
                        siteRaw = (_c = (_b = record['_sct_site_value']) !== null && _b !== void 0 ? _b : record['_sct_siteid_value']) !== null && _c !== void 0 ? _c : '';
                        siteId = typeof siteRaw === 'string' ? siteRaw : '';
                        if (!siteId)
                            return [2 /*return*/];
                        requesterGuid = String((_e = (_d = record['_sct_requester_value']) !== null && _d !== void 0 ? _d : record['_sct_requestorid_value']) !== null && _e !== void 0 ? _e : '');
                        this.setState({ loadingApprovers: true });
                        _f.label = 1;
                    case 1:
                        _f.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._router.resolveApprovers(siteId, requesterGuid)];
                    case 2:
                        resolvedApprovers = _f.sent();
                        this.setState({ resolvedApprovers: resolvedApprovers, loadingApprovers: false });
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _f.sent();
                        this.setState({ loadingApprovers: false });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /** Stage action — PATCH only the status (and optional extra fields) with audit and queue sync */
    Dashboard.prototype._stageAction = function (targetStatus, extra) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var editRecord, id, payload, existingPt, err_2;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        editRecord = this.state.editRecord;
                        if (!editRecord || !this._service)
                            return [2 /*return*/];
                        id = this._getRecordId(editRecord);
                        if (!id)
                            return [2 /*return*/];
                        this.setState({ saving: true, saveError: null });
                        payload = tslib_1.__assign({ sct_request_status: targetStatus }, (extra !== null && extra !== void 0 ? extra : {}));
                        // Synchronize cr2a4_pendingtier: clear on terminal states, preserve or set to Manager on In Review
                        if (targetStatus === DSTATUS.PO_ISSUED ||
                            targetStatus === DSTATUS.REJECTED ||
                            targetStatus === DSTATUS.RETURNED ||
                            targetStatus === DSTATUS.CLOSED) {
                            payload['cr2a4_pendingtier'] = null;
                        }
                        else if (targetStatus === DSTATUS.IN_REVIEW) {
                            existingPt = String(editRecord['cr2a4_pendingtier'] || editRecord['sct_pendingtier'] || '').trim();
                            payload['cr2a4_pendingtier'] = existingPt || 'Manager';
                        }
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, this._service.updateRequest(id, payload)];
                    case 2:
                        _b.sent();
                        // Audit log the stage transition
                        return [4 /*yield*/, this._service.createAuditLog(tslib_1.__assign({ sct_action: targetStatus === DSTATUS.PO_ISSUED ? 1 : targetStatus === DSTATUS.REJECTED ? 5 : targetStatus === DSTATUS.RETURNED ? 4 : 2, sct_description: "Stage workbench updated status to ".concat(targetStatus), sct_oldvalue: "Status: ".concat((0, DataverseService_1.strVal)(editRecord, 'sct_request_status')), sct_newvalue: "Status: ".concat(targetStatus, ", Pending Tier: ").concat((_a = payload['cr2a4_pendingtier']) !== null && _a !== void 0 ? _a : 'null'), 'sct_requestid@odata.bind': "/sct_requests(".concat(id, ")") }, (this._currentSystemUserId ? { 'sct_actionby@odata.bind': "/systemusers(".concat(this._currentSystemUserId, ")") } : {}))).catch(function () { })];
                    case 3:
                        // Audit log the stage transition
                        _b.sent();
                        this.setState({ saving: false, editRecord: null, saveSuccess: 'Record updated.' });
                        void this._load();
                        return [3 /*break*/, 5];
                    case 4:
                        err_2 = _b.sent();
                        this.setState({ saving: false, saveError: err_2 instanceof Error ? err_2.message : 'Update failed' });
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    Dashboard.prototype._resubmitReturned = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a, editRecord, editValues, id, payload, graphClient, manager, title, amt, reqName, _b, err_3;
            var _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        _a = this.state, editRecord = _a.editRecord, editValues = _a.editValues;
                        if (!editRecord || !this._service)
                            return [2 /*return*/];
                        id = this._getRecordId(editRecord);
                        if (!id)
                            return [2 /*return*/];
                        this.setState({ saving: true, saveError: null });
                        payload = {
                            sct_request_status: DSTATUS.SUBMITTED,
                            cr2a4_pendingtier: 'Manager',
                            sct_submitted_date: _dateToEdm(new Date()),
                        };
                        EDIT_TEXT_FIELDS.forEach(function (f) {
                            var v = editValues[f];
                            if (v !== undefined && v !== '')
                                payload[f] = v;
                        });
                        EDIT_NUMBER_FIELDS.forEach(function (f) {
                            var v = editValues[f];
                            if (v !== undefined && v !== '') {
                                var n = parseFloat(v.replace(/[^0-9.\-]/g, ''));
                                if (!isNaN(n))
                                    payload[f] = n;
                            }
                        });
                        EDIT_DATE_FIELDS.forEach(function (f) {
                            var v = editValues[f];
                            if (v) {
                                var edm = formatToEdmDate(v);
                                if (edm)
                                    payload[f] = edm;
                            }
                        });
                        BOOLEAN_FIELDS.forEach(function (f) {
                            if (editValues[f] !== undefined) {
                                payload[f] = toBoolean(editValues[f]);
                            }
                        });
                        _d.label = 1;
                    case 1:
                        _d.trys.push([1, 11, , 12]);
                        return [4 /*yield*/, this._service.updateRequest(id, payload)];
                    case 2:
                        _d.sent();
                        // Audit log the resubmission
                        return [4 /*yield*/, this._service.createAuditLog(tslib_1.__assign({ sct_action: 1, sct_description: 'Request resubmitted for approval by requester', sct_oldvalue: "Status: ".concat((0, DataverseService_1.strVal)(editRecord, 'sct_request_status')), sct_newvalue: "Status: Submitted, Pending Tier: Manager", 'sct_requestid@odata.bind': "/sct_requests(".concat(id, ")") }, (this._currentSystemUserId ? { 'sct_actionby@odata.bind': "/systemusers(".concat(this._currentSystemUserId, ")") } : {}))).catch(function () { })];
                    case 3:
                        // Audit log the resubmission
                        _d.sent();
                        _d.label = 4;
                    case 4:
                        _d.trys.push([4, 9, , 10]);
                        return [4 /*yield*/, this.props.context.msGraphClientFactory.getClient('3')];
                    case 5:
                        graphClient = _d.sent();
                        return [4 /*yield*/, (0, GraphMailService_1.getManager)(graphClient)];
                    case 6:
                        manager = _d.sent();
                        if (!manager) return [3 /*break*/, 8];
                        title = (0, DataverseService_1.strVal)(editRecord, 'sct_request_title') || (0, DataverseService_1.strVal)(editRecord, 'sct_sct_reference');
                        amt = String((_c = editRecord['sct_estimated_total_value']) !== null && _c !== void 0 ? _c : '');
                        reqName = this.props.context.pageContext.user.displayName;
                        return [4 /*yield*/, (0, GraphMailService_1.sendManagerApprovalEmail)(graphClient, manager, reqName, title, amt, this.props.context.pageContext.web.absoluteUrl)];
                    case 7:
                        _d.sent();
                        _d.label = 8;
                    case 8: return [3 /*break*/, 10];
                    case 9:
                        _b = _d.sent();
                        return [3 /*break*/, 10];
                    case 10:
                        this.setState({ saving: false, editRecord: null, saveSuccess: 'Request resubmitted for review successfully.' });
                        void this._load();
                        return [3 /*break*/, 12];
                    case 11:
                        err_3 = _d.sent();
                        this.setState({ saving: false, saveError: err_3 instanceof Error ? err_3.message : 'Resubmission failed' });
                        return [3 /*break*/, 12];
                    case 12: return [2 /*return*/];
                }
            });
        });
    };
    Dashboard.prototype._saveEdit = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a, editRecord, editValues, editInitialValues, id, payload, err_4;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = this.state, editRecord = _a.editRecord, editValues = _a.editValues, editInitialValues = _a.editInitialValues;
                        if (!editRecord || !this._service)
                            return [2 /*return*/];
                        id = this._getRecordId(editRecord);
                        if (!id)
                            return [2 /*return*/];
                        this.setState({ saving: true, saveError: null });
                        payload = {};
                        Object.keys(editValues).forEach(function (field) {
                            var newStr = editValues[field];
                            var originalRaw = editRecord[field];
                            if (typeof originalRaw === 'boolean') {
                                // Edm.Boolean — convert "Yes"/"No"/"true"/"false" to actual boolean
                                payload[field] = toBoolean(newStr);
                                return;
                            }
                            if (typeof originalRaw === 'number') {
                                // Edm.Int32 / Edm.Decimal — strip any currency formatting characters before parsing.
                                // _openEdit now loads raw numbers, but guard against formatted strings just in case.
                                var stripped = newStr.replace(/[^0-9.\-]/g, '');
                                var num = parseFloat(stripped);
                                if (!isNaN(num))
                                    payload[field] = num;
                                // If still not parseable (e.g. "Standard" for a choice field display text), skip it —
                                // choice field display text cannot be round-tripped; the edit panel makes them read-only.
                                return;
                            }
                            // Date fields — must be strict YYYY-MM-DD for Edm.Date; localized strings like "2/1/2027" cause 400
                            if (EDIT_DATE_FIELDS.has(field)) {
                                var edm = formatToEdmDate(newStr);
                                if (edm)
                                    payload[field] = edm;
                                // If unparseable, skip the field rather than sending a bad value
                                return;
                            }
                            // Plain string — send as-is if non-empty
                            if (newStr !== '' && newStr !== undefined)
                                payload[field] = newStr;
                        });
                        if (Object.keys(payload).length === 0) {
                            // Nothing to patch — close without a round-trip
                            this.setState({ saving: false, editRecord: null, saveSuccess: 'No changes to save.' });
                            return [2 /*return*/];
                        }
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._service.updateRequest(id, payload)];
                    case 2:
                        _b.sent();
                        // Close panel, show success, then reload the full list from Dataverse so the
                        // table reflects the server's authoritative values (computed fields, triggers, etc.)
                        this.setState({ saving: false, editRecord: null, saveSuccess: 'Record updated successfully.' });
                        void this._load();
                        // Log each changed field to sct_auditlog asynchronously (non-blocking)
                        void (0, auditLogger_1.logAuditChanges)(this._service, id, this._currentSystemUserId, editInitialValues, editValues);
                        return [3 /*break*/, 4];
                    case 3:
                        err_4 = _b.sent();
                        this.setState({
                            saving: false,
                            saveError: err_4 instanceof Error ? err_4.message : 'Update failed',
                        });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    Dashboard.prototype._confirmDelete = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var deleteRecord, id, err_5;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        deleteRecord = this.state.deleteRecord;
                        if (!deleteRecord || !this._service)
                            return [2 /*return*/];
                        id = this._getRecordId(deleteRecord);
                        if (!id)
                            return [2 /*return*/];
                        this.setState({ saving: true });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._service.deleteRequest(id)];
                    case 2:
                        _a.sent();
                        this.setState(function (prev) { return ({
                            saving: false,
                            deleteRecord: null,
                            saveSuccess: 'Record deleted.',
                            records: prev.records.filter(function (r) { return _this._getRecordId(r) !== id; }),
                        }); });
                        return [2 /*return*/];
                    case 3:
                        err_5 = _a.sent();
                        this.setState({
                            saving: false,
                            saveError: err_5 instanceof Error ? err_5.message : 'Delete failed',
                            deleteRecord: null,
                        });
                        return [2 /*return*/];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    // ── Render helpers ────────────────────────────────────────────────────────
    Dashboard.prototype._renderViewPanel = function () {
        var _this = this;
        var _a = this.state, viewRecord = _a.viewRecord, allFields = _a.allFields;
        if (!viewRecord)
            return null;
        // Build view field list: PRIORITY_FIELDS order first, then any remaining sct_ fields.
        // Exclude bare GUIDs, annotation-only keys, and system noise.
        // Booleans are shown here (unlike table columns) — this is a detail view.
        var usable = allFields.filter(function (f) {
            return (f.startsWith('sct_') || f === 'createdon' || f === 'modifiedon') &&
                !isGuidField([viewRecord], f);
        });
        var prioritised = [];
        PRIORITY_FIELDS.forEach(function (p) { if (usable.indexOf(p) !== -1)
            prioritised.push(p); });
        usable.forEach(function (f) { if (prioritised.indexOf(f) === -1)
            prioritised.push(f); });
        var displayFields = prioritised;
        // Reference number as panel title; confirmed PrimaryNameAttribute = sct_sct_reference
        var panelTitle = (0, DataverseService_1.strVal)(viewRecord, 'sct_sct_reference') || 'Record Details';
        // Business status from sct_request_status (confirmed); fall back to statuscode if needed
        var statusLabel = (0, DataverseService_1.strVal)(viewRecord, 'sct_request_status') || (0, DataverseService_1.strVal)(viewRecord, 'statuscode');
        var isReturned = statusLabel === 'Returned';
        var canEdit = statusLabel === 'Draft' || isReturned;
        return (React.createElement(Panel_1.Panel, { isOpen: true, onDismiss: function () { return _this.setState({ viewRecord: null }); }, type: Panel_1.PanelType.custom, customWidth: "min(640px, 95vw)", headerText: panelTitle, closeButtonAriaLabel: "Close", styles: { content: { paddingBottom: 80 } }, isFooterAtBottom: true, onRenderFooterContent: function () { return (React.createElement("div", { style: { display: 'flex', gap: 10, padding: '0 24px 24px' } },
                React.createElement(Button_1.PrimaryButton, { text: isReturned ? 'Edit & Resubmit' : 'Edit Record', iconProps: { iconName: 'Edit' }, disabled: !canEdit, title: canEdit
                        ? (isReturned ? 'Edit and resubmit this returned request' : 'Edit this draft request')
                        : "Only Draft or Returned records can be edited \u2014 this record is \"".concat(statusLabel || 'unknown status', "\""), onClick: function () {
                        var rec = viewRecord;
                        _this.setState({ viewRecord: null }, function () { return _this._openEdit(rec); });
                    }, styles: { root: { borderRadius: 6 } } }),
                React.createElement(Button_1.DefaultButton, { text: "Close", onClick: function () { return _this.setState({ viewRecord: null }); }, styles: { root: { borderRadius: 6 } } }))); } },
            statusLabel && (React.createElement("div", { style: { marginBottom: 16 } },
                React.createElement(StatusPill, { value: statusLabel }))),
            displayFields.map(function (field) {
                var val = (0, DataverseService_1.strVal)(viewRecord, field);
                return (React.createElement("div", { key: field, style: { borderBottom: '1px solid #F3F2F1', padding: '11px 0', display: 'flex', flexDirection: 'column', gap: 3 } },
                    React.createElement(Text_1.Text, { variant: "small", style: { color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.5px', fontWeight: 600 } }, fieldLabel(field)),
                    field.indexOf('status') !== -1 || field.indexOf('type') !== -1 || field.indexOf('basis') !== -1
                        ? React.createElement(StatusPill, { value: val || '—' })
                        : React.createElement(Text_1.Text, { variant: "medium", style: { color: val ? '#201F1E' : '#C8C6C4', wordBreak: 'break-all', lineHeight: '1.5' } }, field.indexOf('amount') !== -1 && val
                            ? (0, dashboardUtils_1.formatCurrency)(parseFloat(val) || 0)
                            : (field === 'createdon' || field === 'modifiedon' || field.indexOf('date') !== -1) && val
                                ? (0, dashboardUtils_1.formatDate)(val)
                                : val || '—')));
            })));
    };
    Dashboard.prototype._renderEditPanel = function () {
        var _this = this;
        var _a, _b;
        var _c = this.state, editRecord = _c.editRecord, editValues = _c.editValues, poDate = _c.poDate, saving = _c.saving, saveError = _c.saveError, resolvedApprovers = _c.resolvedApprovers, loadingApprovers = _c.loadingApprovers;
        if (!editRecord)
            return null;
        // ── Detect which stage panel to show ───────────────────────────────────
        var ref = (0, DataverseService_1.strVal)(editRecord, 'sct_sct_reference');
        var statusLabel = (0, DataverseService_1.strVal)(editRecord, 'sct_request_status'); // FormattedValue
        var statusRaw = editRecord['sct_request_status']; // raw integer
        var site = (0, DataverseService_1.strVal)(editRecord, '_sct_site_value');
        var amt = parseFloat(String((_a = editRecord['sct_estimated_total_value']) !== null && _a !== void 0 ? _a : 0));
        var isNew = editRecord['sct_is_new_supplier'] === true || editRecord['sct_is_new_supplier'] === 'Yes';
        var needsGate = (!isNaN(amt) && amt > D_CFO_THRESHOLD) || isNew;
        var isSubmitted = statusRaw === DSTATUS.SUBMITTED || statusLabel === 'Submitted';
        var isInReview = statusRaw === DSTATUS.IN_REVIEW || statusLabel === 'In Review';
        var isPOIssued = statusRaw === DSTATUS.PO_ISSUED || statusLabel === 'PO Issued';
        // Sourcing: In Review + standard (within threshold, known supplier)
        var isSourceStage = isInReview && !needsGate;
        // FP&A / CFO: In Review + gate required
        var isGateStage = isInReview && needsGate;
        // ── Helpers ───────────────────────────────────────────────────────────
        var setVal = function (field) {
            return function (_, v) {
                return _this.setState(function (s) {
                    var _a;
                    return ({ editValues: tslib_1.__assign(tslib_1.__assign({}, s.editValues), (_a = {}, _a[field] = v !== null && v !== void 0 ? v : '', _a)) });
                });
            };
        };
        var setToggle = function (field) {
            return function (_, checked) {
                return _this.setState(function (s) {
                    var _a;
                    return ({ editValues: tslib_1.__assign(tslib_1.__assign({}, s.editValues), (_a = {}, _a[field] = checked ? 'true' : 'false', _a)) });
                });
            };
        };
        var tv = function (field) { var _a; return (_a = editValues[field]) !== null && _a !== void 0 ? _a : ''; };
        var tb = function (field) { return tv(field) === 'true' || tv(field) === 'Yes'; };
        var ss = { marginBottom: 20, paddingBottom: 16, borderBottom: '1px solid #EDEBE9' };
        var tfS = { root: { marginBottom: 12 }, fieldGroup: { borderRadius: 6 } };
        var lbl = { fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 2 };
        var roVal = { fontSize: 13, color: '#201F1E', fontWeight: 500 };
        var close = function () { return _this.setState({ editRecord: null, saveError: null }); };
        var SectionTitle = function (_a) {
            var label = _a.label, _b = _a.color, color = _b === void 0 ? '#0078D4' : _b;
            return (React.createElement("div", { style: { fontSize: 13, fontWeight: 700, color: color, marginBottom: 12 } }, label));
        };
        // ── Expanded read-only context strip (shown in all panels) ─────────────
        var ContextStrip = function () { return (React.createElement("div", { style: { background: '#F8F8F7', borderRadius: 8, padding: '12px 14px', marginBottom: 20 } },
            React.createElement("div", { style: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px 16px' } }, [
                ['Reference', ref],
                ['Status', statusLabel],
                ['Site', site],
                ['Requester', (0, DataverseService_1.strVal)(editRecord, '_sct_requester_value')],
                ['Est. Value', isNaN(amt) ? '—' : (0, dashboardUtils_1.formatCurrency)(amt)],
                ['Purchase Type', (0, DataverseService_1.strVal)(editRecord, 'sct_purchase_type')],
                ['Urgency', (0, DataverseService_1.strVal)(editRecord, 'sct_urgency')],
                ['New Supplier', isNew ? 'Yes ⚑' : 'No'],
                ['Date Needed By', (0, dashboardUtils_1.formatDate)((0, DataverseService_1.strVal)(editRecord, 'sct_date_needed_by'))],
                ['Submitted Date', (0, dashboardUtils_1.formatDate)((0, DataverseService_1.strVal)(editRecord, 'sct_submitted_date'))],
            ].map(function (_a) {
                var label = _a[0], value = _a[1];
                return (React.createElement("div", { key: label, style: { marginBottom: 6 } },
                    React.createElement("div", { style: lbl }, label),
                    React.createElement("div", { style: roVal }, value || '—')));
            })),
            (0, DataverseService_1.strVal)(editRecord, 'sct_description') && (React.createElement("div", { style: { marginTop: 10, paddingTop: 10, borderTop: '1px solid #EDEBE9' } },
                React.createElement("div", { style: lbl }, "Description"),
                React.createElement("div", { style: { fontSize: 12, color: '#605E5C', lineHeight: '1.5', maxHeight: 72, overflowY: 'auto' } }, (0, DataverseService_1.strVal)(editRecord, 'sct_description')))))); };
        // ── Resolved approvers block (Phase 8) ─────────────────────────────────
        var ApproversBlock = function () { return (React.createElement("div", { style: { marginTop: 16 } },
            React.createElement(SectionTitle, { label: "Schedule A \u2014 Resolved Approvers", color: "#605E5C" }),
            loadingApprovers && (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8, padding: '8px 0' } },
                React.createElement(Spinner_1.Spinner, { size: Spinner_1.SpinnerSize.small }),
                React.createElement(Text_1.Text, { variant: "small", style: { color: '#605E5C' } }, "Resolving from sct_role\u2026"))),
            !loadingApprovers && resolvedApprovers && (React.createElement(React.Fragment, null,
                resolvedApprovers.approvers.length === 0 && (React.createElement(Text_1.Text, { variant: "small", style: { color: '#A19F9D', display: 'block', marginBottom: 8 } }, "No approvers found for this site \u2014 verify sct_role records in Dataverse.")),
                resolvedApprovers.approvers.map(function (a) { return (React.createElement("div", { key: a.level + a.systemUserId, style: { display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', background: '#F8F8F8', borderRadius: 8, marginBottom: 8 } },
                    React.createElement("div", { style: { width: 34, height: 34, borderRadius: 8, background: '#0078D4', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontWeight: 700, fontSize: 11, flexShrink: 0 } }, a.level),
                    React.createElement("div", { style: { flex: 1 } },
                        React.createElement("div", { style: { fontSize: 13, fontWeight: 600, color: '#201F1E' } }, a.name || a.systemUserId),
                        React.createElement("div", { style: { display: 'flex', gap: 5, marginTop: 3 } },
                            a.isDelegate && React.createElement("span", { style: { fontSize: 10, background: '#FFF4CE', color: '#905400', padding: '1px 6px', borderRadius: 8, fontWeight: 700 } }, "DELEGATE"),
                            a.wasEscalated && React.createElement("span", { style: { fontSize: 10, background: '#FDE7E9', color: '#A4262C', padding: '1px 6px', borderRadius: 8, fontWeight: 700 } }, "R10 ESCALATED"))))); }),
                resolvedApprovers.escalations.map(function (e, i) { return (React.createElement(MessageBar_1.MessageBar, { key: i, messageBarType: MessageBar_1.MessageBarType.warning, styles: { root: { borderRadius: 6, marginBottom: 6 } } }, e)); }),
                resolvedApprovers.warnings.map(function (w, i) { return (React.createElement(MessageBar_1.MessageBar, { key: i, messageBarType: MessageBar_1.MessageBarType.info, styles: { root: { borderRadius: 6, marginBottom: 6 } } }, w)); }))))); };
        // ── Pre-approval requirements (Phase 8) ────────────────────────────────
        var preApproval = (_b = this._router) === null || _b === void 0 ? void 0 : _b.calculatePreApprovals(isNaN(amt) ? 0 : amt, isNew);
        // ── Determine panel header text + footer buttons ────────────────────────
        var headerText = ref ? "Edit \u00B7 ".concat(ref) : 'Edit Record';
        var footerContent;
        if (isSubmitted) {
            headerText = ref ? "Stage 3 \u2014 Triage \u00B7 ".concat(ref) : 'Stage 3 — Triage';
            footerContent = (React.createElement("div", { style: { display: 'flex', gap: 10, flexWrap: 'wrap', padding: '0 24px 24px' } },
                React.createElement(Button_1.PrimaryButton, { text: saving ? 'Saving…' : 'Submit to Sourcing', disabled: saving, onClick: function () { return void _this._stageAction(DSTATUS.IN_REVIEW, {
                        sct_cost_center: tv('sct_cost_center') || undefined,
                    }); }, styles: { root: { borderRadius: 6 } } }),
                React.createElement(Button_1.DefaultButton, { text: "Return for Information", disabled: saving, onClick: function () { return void _this._stageAction(DSTATUS.RETURNED); }, styles: { root: { borderRadius: 6, color: '#905400', borderColor: '#C7A333' } } }),
                React.createElement(Button_1.DefaultButton, { text: "Cancel", onClick: close, disabled: saving, styles: { root: { borderRadius: 6 } } })));
        }
        else if (isSourceStage) {
            headerText = ref ? "Stage 4 \u2014 Sourcing \u00B7 ".concat(ref) : 'Stage 4 — Sourcing';
            footerContent = (React.createElement("div", { style: { display: 'flex', gap: 10, flexWrap: 'wrap', padding: '0 24px 24px' } },
                React.createElement(Button_1.PrimaryButton, { text: saving ? 'Saving…' : 'Issue PO', disabled: saving, onClick: function () {
                        var extra = {};
                        var po = tv('sct_po_number').trim();
                        if (po)
                            extra['sct_po_number'] = po;
                        if (poDate)
                            extra['sct_po_issue_date'] = _dateToEdm(poDate);
                        void _this._stageAction(DSTATUS.PO_ISSUED, extra);
                    }, styles: { root: { borderRadius: 6 } } }),
                React.createElement(Button_1.DefaultButton, { text: "Reject", disabled: saving, onClick: function () { return void _this._stageAction(DSTATUS.REJECTED); }, styles: { root: { borderRadius: 6, color: '#A4262C', borderColor: '#A4262C' } } }),
                React.createElement(Button_1.DefaultButton, { text: "Cancel", onClick: close, disabled: saving, styles: { root: { borderRadius: 6 } } })));
        }
        else if (isGateStage) {
            headerText = ref ? "Stage 5 \u2014 FP&A / CFO \u00B7 ".concat(ref) : 'Stage 5 — FP&A / CFO';
            footerContent = (React.createElement("div", { style: { display: 'flex', gap: 10, flexWrap: 'wrap', padding: '0 24px 24px' } },
                React.createElement(Button_1.PrimaryButton, { text: saving ? 'Saving…' : 'Approve — Issue PO', disabled: saving, onClick: function () { return void _this._stageAction(DSTATUS.PO_ISSUED); }, styles: { root: { borderRadius: 6 } } }),
                React.createElement(Button_1.DefaultButton, { text: "Reject", disabled: saving, onClick: function () { return void _this._stageAction(DSTATUS.REJECTED); }, styles: { root: { borderRadius: 6, color: '#A4262C', borderColor: '#A4262C' } } }),
                React.createElement(Button_1.DefaultButton, { text: "Return to Requester", disabled: saving, onClick: function () { return void _this._stageAction(DSTATUS.RETURNED); }, styles: { root: { borderRadius: 6, color: '#905400', borderColor: '#C7A333' } } }),
                React.createElement(Button_1.DefaultButton, { text: "Cancel", onClick: close, disabled: saving, styles: { root: { borderRadius: 6 } } })));
        }
        else if (isPOIssued) {
            headerText = ref ? "Stage 6 \u2014 Receiving \u00B7 ".concat(ref) : 'Stage 6 — Receiving';
            footerContent = (React.createElement("div", { style: { display: 'flex', gap: 10, padding: '0 24px 24px' } },
                React.createElement(Button_1.PrimaryButton, { text: saving ? 'Saving…' : 'Confirm Received & Close', disabled: saving, onClick: function () { return void _this._stageAction(DSTATUS.CLOSED); }, styles: { root: { borderRadius: 6 } } }),
                React.createElement(Button_1.DefaultButton, { text: "Cancel", onClick: close, disabled: saving, styles: { root: { borderRadius: 6 } } })));
        }
        else {
            // Default — standard requester edit (Draft, Returned, or custom status)
            var isReturnedRecord_1 = statusRaw === DSTATUS.RETURNED || statusLabel === 'Returned';
            footerContent = (React.createElement("div", { style: { display: 'flex', gap: 10, flexWrap: 'wrap', padding: '0 24px 24px' } },
                React.createElement(Button_1.PrimaryButton, { text: saving ? 'Submitting…' : (isReturnedRecord_1 ? 'Resubmit for Review' : 'Save Changes'), onClick: function () { return void (isReturnedRecord_1 ? _this._resubmitReturned() : _this._saveEdit()); }, disabled: saving, styles: { root: { borderRadius: 6 } } }),
                !isReturnedRecord_1 && (React.createElement(Button_1.DefaultButton, { text: saving ? 'Submitting…' : 'Submit for Review', disabled: saving, onClick: function () { return void _this._resubmitReturned(); }, styles: { root: { borderRadius: 6, color: '#107C10', borderColor: '#107C10' } } })),
                React.createElement(Button_1.DefaultButton, { text: "Cancel", onClick: close, disabled: saving, styles: { root: { borderRadius: 6 } } })));
        }
        return (React.createElement(Panel_1.Panel, { isOpen: true, onDismiss: close, type: Panel_1.PanelType.custom, customWidth: "min(820px, 95vw)", headerText: headerText, closeButtonAriaLabel: "Close", styles: { content: { paddingBottom: 88 } }, isFooterAtBottom: true, onRenderFooterContent: function () { return footerContent; } },
            saveError && (React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.error, style: { marginBottom: 16 } }, saveError)),
            React.createElement(ContextStrip, null),
            isSubmitted && (React.createElement(React.Fragment, null,
                preApproval && (React.createElement(MessageBar_1.MessageBar, { messageBarType: preApproval.levels.length > 1 ? MessageBar_1.MessageBarType.warning : MessageBar_1.MessageBarType.info, styles: { root: { borderRadius: 6, marginBottom: 14 } } },
                    React.createElement("strong", null,
                        "Pre-Approval Required: ",
                        preApproval.levels.join(' + ')),
                    " \u2014 ",
                    preApproval.description)),
                React.createElement(SectionTitle, { label: "Stage 3 \u2014 Triage Assignment", color: "#0078D4" }),
                React.createElement(TextField_1.TextField, { label: "Cost Centre", value: tv('sct_cost_center'), onChange: setVal('sct_cost_center'), placeholder: "e.g. CC-1042", styles: tfS }),
                React.createElement(ApproversBlock, null))),
            isSourceStage && (React.createElement(React.Fragment, null,
                React.createElement(SectionTitle, { label: "Stage 4 \u2014 Procurement Sourcing", color: "#107C10" }),
                React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.info, styles: { root: { borderRadius: 6, marginBottom: 12 } } },
                    "Within direct Procurement authority (\u2264 $",
                    D_CFO_THRESHOLD.toLocaleString(),
                    ", known supplier). Issue PO directly."),
                React.createElement(TextField_1.TextField, { label: "D365 PO Number", value: tv('sct_po_number'), onChange: setVal('sct_po_number'), placeholder: "e.g. PO-2026-00123", styles: tfS }),
                React.createElement(DatePicker_1.DatePicker, { label: "PO Issue Date", value: poDate, onSelectDate: function (d) { return _this.setState({ poDate: d !== null && d !== void 0 ? d : undefined }); }, firstDayOfWeek: Calendar_1.DayOfWeek.Monday, formatDate: function (d) { return d ? _dateToEdm(d) : ''; }, placeholder: "Select PO issue date\u2026", styles: { root: { marginBottom: 14 } }, allowTextInput: true }))),
            isGateStage && (React.createElement(React.Fragment, null,
                React.createElement(SectionTitle, { label: "Stage 5 \u2014 FP&A / CFO Review", color: "#8764B8" }),
                React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.warning, styles: { root: { borderRadius: 6, marginBottom: 12 } } },
                    React.createElement("strong", null, "CFO approval gate triggered"),
                    " \u2014",
                    !isNaN(amt) && amt > D_CFO_THRESHOLD ? " value ".concat((0, dashboardUtils_1.formatCurrency)(amt), " exceeds $").concat(D_CFO_THRESHOLD.toLocaleString(), ".") : '',
                    isNew ? ' New supplier.' : '',
                    ' ',
                    "Approve to issue PO or Reject."),
                ((0, DataverseService_1.strVal)(editRecord, 'sct_po_number') || (0, DataverseService_1.strVal)(editRecord, '_sct_suggested_supplier_value')) && (React.createElement("div", { style: { background: '#F8F8F7', borderRadius: 8, padding: '10px 14px', marginBottom: 14 } },
                    (0, DataverseService_1.strVal)(editRecord, '_sct_suggested_supplier_value') && (React.createElement("div", { style: { marginBottom: 6 } },
                        React.createElement("div", { style: lbl }, "Supplier"),
                        React.createElement("div", { style: roVal }, (0, DataverseService_1.strVal)(editRecord, '_sct_suggested_supplier_value')))),
                    (0, DataverseService_1.strVal)(editRecord, 'sct_po_number') && (React.createElement("div", null,
                        React.createElement("div", { style: lbl }, "PO Number"),
                        React.createElement("div", { style: roVal }, (0, DataverseService_1.strVal)(editRecord, 'sct_po_number')))))),
                React.createElement(ApproversBlock, null))),
            isPOIssued && (React.createElement(React.Fragment, null,
                React.createElement(SectionTitle, { label: "Stage 6 \u2014 Receiving", color: "#00B294" }),
                React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.info, styles: { root: { borderRadius: 6, marginBottom: 12 } } }, "Confirm that goods or services have been received. This closes the request in the Spend Control Tower."),
                (0, DataverseService_1.strVal)(editRecord, 'sct_po_number') && (React.createElement("div", { style: { background: '#F8F8F7', borderRadius: 8, padding: '10px 14px', marginBottom: 14 } },
                    React.createElement("div", { style: lbl }, "PO Number"),
                    React.createElement("div", { style: roVal }, (0, DataverseService_1.strVal)(editRecord, 'sct_po_number')))))),
            !isSubmitted && !isInReview && !isPOIssued && (React.createElement(React.Fragment, null,
                React.createElement("div", { style: ss },
                    React.createElement(SectionTitle, { label: "Description & Scope" }),
                    React.createElement(TextField_1.TextField, { label: "Description", value: tv('sct_description'), multiline: true, rows: 4, onChange: setVal('sct_description'), styles: tfS }),
                    React.createElement(TextField_1.TextField, { label: "Function / Department", value: tv('sct_function'), onChange: setVal('sct_function'), styles: tfS })),
                React.createElement("div", { style: ss },
                    React.createElement(SectionTitle, { label: "Urgency Details" }),
                    React.createElement(TextField_1.TextField, { label: "Urgency Reason", value: tv('sct_urgency_reason'), multiline: true, rows: 2, onChange: setVal('sct_urgency_reason'), styles: tfS }),
                    React.createElement(TextField_1.TextField, { label: "Emergency Declared By (email)", value: tv('sct_emergency_declared_by_email'), onChange: setVal('sct_emergency_declared_by_email'), styles: tfS })),
                React.createElement("div", { style: ss },
                    React.createElement(SectionTitle, { label: "Financials & Dates" }),
                    React.createElement(TextField_1.TextField, { label: "Estimated Total Value (AUD)", value: tv('sct_estimated_total_value'), onChange: setVal('sct_estimated_total_value'), styles: tfS }),
                    React.createElement(TextField_1.TextField, { label: "Date Needed By (YYYY-MM-DD)", value: tv('sct_date_needed_by'), onChange: setVal('sct_date_needed_by'), styles: tfS }),
                    React.createElement(TextField_1.TextField, { label: "Term Start Date (YYYY-MM-DD)", value: tv('sct_term_start_date'), onChange: setVal('sct_term_start_date'), styles: tfS }),
                    React.createElement(TextField_1.TextField, { label: "Term End Date (YYYY-MM-DD)", value: tv('sct_term_end_date'), onChange: setVal('sct_term_end_date'), styles: tfS })),
                React.createElement("div", { style: ss },
                    React.createElement(SectionTitle, { label: "Logistics" }),
                    React.createElement(TextField_1.TextField, { label: "Delivery Address", value: tv('sct_delivery_address'), multiline: true, rows: 2, onChange: setVal('sct_delivery_address'), styles: tfS }),
                    React.createElement(TextField_1.TextField, { label: "Receiving Contact Email", value: tv('sct_receiving_contact_email'), onChange: setVal('sct_receiving_contact_email'), styles: tfS }),
                    React.createElement(TextField_1.TextField, { label: "Equipment Asset Tag", value: tv('sct_equipment_asset_tag'), onChange: setVal('sct_equipment_asset_tag'), styles: tfS }),
                    React.createElement(TextField_1.TextField, { label: "Customer Name (if recharged)", value: tv('sct_customer_name'), onChange: setVal('sct_customer_name'), styles: tfS })),
                React.createElement("div", { style: { marginBottom: 8 } },
                    React.createElement(SectionTitle, { label: "Flags & Declarations" }),
                    React.createElement(Toggle_1.Toggle, { label: "Is New Supplier", checked: tb('sct_is_new_supplier'), onChange: setToggle('sct_is_new_supplier'), styles: { root: { marginBottom: 8 } } }),
                    React.createElement(Toggle_1.Toggle, { label: "Cost Recharged to Customer", checked: tb('sct_cost_recharged_to_customer'), onChange: setToggle('sct_cost_recharged_to_customer'), styles: { root: { marginBottom: 8 } } }),
                    React.createElement(Toggle_1.Toggle, { label: "Customer Agreement Attached", checked: tb('sct_customer_agreement_attached'), onChange: setToggle('sct_customer_agreement_attached'), styles: { root: { marginBottom: 8 } } }),
                    React.createElement(Toggle_1.Toggle, { label: "Tech or Data Involved", checked: tb('sct_tech_or_data_involved'), onChange: setToggle('sct_tech_or_data_involved'), styles: { root: { marginBottom: 8 } } }),
                    React.createElement(Toggle_1.Toggle, { label: "Supplier on Premises", checked: tb('sct_supplier_on_premises'), onChange: setToggle('sct_supplier_on_premises'), styles: { root: { marginBottom: 8 } } }),
                    React.createElement(Toggle_1.Toggle, { label: "Possible Capital Purchase", checked: tb('sct_possible_capital_purchase'), onChange: setToggle('sct_possible_capital_purchase'), styles: { root: { marginBottom: 8 } } }),
                    React.createElement(Toggle_1.Toggle, { label: "Confidential", checked: tb('sct_is_confidential'), onChange: setToggle('sct_is_confidential'), styles: { root: { marginBottom: 8 } } }),
                    React.createElement(Toggle_1.Toggle, { label: "Retrospective", checked: tb('sct_is_retrospective'), onChange: setToggle('sct_is_retrospective'), styles: { root: { marginBottom: 8 } } }),
                    React.createElement(Toggle_1.Toggle, { label: "Existing Agreement or PO", checked: tb('sct_existing_agreement_or_po'), onChange: setToggle('sct_existing_agreement_or_po'), styles: { root: { marginBottom: 8 } } }),
                    React.createElement(Toggle_1.Toggle, { label: "Declaration \u2013 No Prior Commitment", checked: tb('sct_declaration_no_prior_commitment'), onChange: setToggle('sct_declaration_no_prior_commitment'), styles: { root: { marginBottom: 8 } } }),
                    React.createElement(Toggle_1.Toggle, { label: "Declaration \u2013 No Conflict of Interest", checked: tb('sct_declaration_no_conflict_of_interest'), onChange: setToggle('sct_declaration_no_conflict_of_interest'), styles: { root: { marginBottom: 8 } } }))))));
    };
    Dashboard.prototype._renderDeleteDialog = function () {
        var _this = this;
        var _a = this.state, deleteRecord = _a.deleteRecord, saving = _a.saving;
        if (!deleteRecord)
            return null;
        var titleField = this.state.allFields.find(function (f) { return f.indexOf('name') !== -1 || f.indexOf('title') !== -1; });
        var label = titleField ? (0, DataverseService_1.strVal)(deleteRecord, titleField) : 'this record';
        return (React.createElement(Dialog_1.Dialog, { isOpen: true, hidden: false, onDismiss: function () { return _this.setState({ deleteRecord: null }); }, dialogContentProps: { type: Dialog_1.DialogType.normal, title: 'Delete Record', subText: "Permanently delete \"".concat(label, "\"? This cannot be undone.") }, modalProps: { isBlocking: true } },
            React.createElement(Dialog_1.DialogFooter, null,
                React.createElement(Button_1.PrimaryButton, { text: saving ? 'Deleting…' : 'Delete', onClick: function () { return void _this._confirmDelete(); }, disabled: saving, styles: { root: { background: '#A4262C', border: 'none', borderRadius: 6 } } }),
                React.createElement(Button_1.DefaultButton, { text: "Cancel", onClick: function () { return _this.setState({ deleteRecord: null }); }, disabled: saving, styles: { root: { borderRadius: 6 } } }))));
    };
    Dashboard.prototype._renderFieldInspector = function () {
        var _this = this;
        var _a = this.state, showFieldInspector = _a.showFieldInspector, records = _a.records;
        if (!showFieldInspector || !records.length)
            return null;
        var allKeys = Object.keys(records[0]).filter(function (k) { return !k.startsWith('@'); }).sort();
        return (React.createElement(Panel_1.Panel, { isOpen: true, onDismiss: function () { return _this.setState({ showFieldInspector: false }); }, type: Panel_1.PanelType.custom, customWidth: "min(640px, 95vw)", headerText: "Field Inspector \u2014 ".concat(allKeys.length, " columns from sct_request"), closeButtonAriaLabel: "Close", styles: { content: { paddingBottom: 40 } } },
            React.createElement(Text_1.Text, { variant: "small", style: { color: '#605E5C', marginBottom: 12, display: 'block', lineHeight: '1.6' } },
                "Real logical names returned by your Dataverse ",
                React.createElement("code", { style: { background: '#F3F2F1', padding: '1px 5px', borderRadius: 4 } }, "sct_request"),
                " entity. Use these to update ",
                React.createElement("code", { style: { background: '#F3F2F1', padding: '1px 5px', borderRadius: 4 } }, "src/models/ISCTRequest.ts"),
                "."),
            React.createElement("div", { style: { fontFamily: 'Consolas, monospace', fontSize: 12, background: '#F8F8F8', border: '1px solid #EDEBE9', borderRadius: 8, padding: 16, maxHeight: 520, overflowY: 'auto' } }, allKeys.map(function (k) {
                var sample = records[0][k];
                var type = sample === null ? 'null' : typeof sample === 'number' ? 'number' : typeof sample === 'boolean' ? 'boolean' : 'string';
                var sampleStr = sample !== null && sample !== undefined ? String(sample).substring(0, 50) : '';
                return (React.createElement("div", { key: k, style: { marginBottom: 5, display: 'flex', gap: 8, alignItems: 'baseline' } },
                    React.createElement("span", { style: { color: '#0078D4', minWidth: 260 } }, k),
                    React.createElement("span", { style: { color: '#107C10', minWidth: 60 } }, type),
                    React.createElement("span", { style: { color: '#A19F9D', fontSize: 11 } }, sampleStr)));
            }))));
    };
    Dashboard.prototype._renderPagination = function (totalFiltered) {
        var _this = this;
        var _a = this.state, currentPage = _a.currentPage, pageSize = _a.pageSize;
        var totalPages = Math.max(1, Math.ceil(totalFiltered / pageSize));
        var start = Math.min((currentPage - 1) * pageSize + 1, totalFiltered);
        var end = Math.min(currentPage * pageSize, totalFiltered);
        // Show at most 5 page numbers
        var pages = [];
        if (totalPages <= 5) {
            for (var i = 1; i <= totalPages; i++)
                pages.push(i);
        }
        else if (currentPage <= 3) {
            pages.push(1, 2, 3, 4, 5);
        }
        else if (currentPage >= totalPages - 2) {
            for (var i = totalPages - 4; i <= totalPages; i++)
                pages.push(i);
        }
        else {
            for (var i = currentPage - 2; i <= currentPage + 2; i++)
                pages.push(i);
        }
        return (React.createElement("div", { className: AppShell_module_scss_1.default.pagination },
            React.createElement(Text_1.Text, { variant: "small", className: AppShell_module_scss_1.default.paginationInfo },
                "Showing ",
                totalFiltered === 0 ? 0 : start,
                "\u2013",
                end,
                " of ",
                totalFiltered,
                " records"),
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 4 } },
                React.createElement("div", { className: AppShell_module_scss_1.default.paginationControls },
                    React.createElement("button", { className: AppShell_module_scss_1.default.pageBtn, disabled: currentPage === 1, onClick: function () { return _this.setState({ currentPage: 1 }); }, "aria-label": "First page" }, "\u00AB"),
                    React.createElement("button", { className: AppShell_module_scss_1.default.pageBtn, disabled: currentPage === 1, onClick: function () { return _this.setState(function (s) { return ({ currentPage: s.currentPage - 1 }); }); }, "aria-label": "Previous page" }, "\u2039"),
                    pages.map(function (p) { return (React.createElement("button", { key: p, className: "".concat(AppShell_module_scss_1.default.pageBtn, " ").concat(p === currentPage ? AppShell_module_scss_1.default.pageBtnActive : ''), onClick: function () { return _this.setState({ currentPage: p }); }, "aria-label": "Page ".concat(p), "aria-current": p === currentPage }, p)); }),
                    React.createElement("button", { className: AppShell_module_scss_1.default.pageBtn, disabled: currentPage === totalPages, onClick: function () { return _this.setState(function (s) { return ({ currentPage: s.currentPage + 1 }); }); }, "aria-label": "Next page" }, "\u203A"),
                    React.createElement("button", { className: AppShell_module_scss_1.default.pageBtn, disabled: currentPage === totalPages, onClick: function () { return _this.setState({ currentPage: totalPages }); }, "aria-label": "Last page" }, "\u00BB")),
                React.createElement("select", { className: AppShell_module_scss_1.default.pageSizeSelect, value: pageSize, onChange: function (e) { return _this.setState({ pageSize: Number(e.target.value), currentPage: 1 }); }, "aria-label": "Rows per page" }, PAGE_SIZE_OPTIONS.map(function (n) { return React.createElement("option", { key: n, value: n },
                    n,
                    " / page"); })))));
    };
    // ── History panel ─────────────────────────────────────────────────────────
    Dashboard.prototype._renderHistoryPanel = function () {
        var _this = this;
        var historyRecord = this.state.historyRecord;
        if (!historyRecord || !this._service)
            return null;
        var id = this._getRecordId(historyRecord);
        if (!id)
            return null;
        var ref = (0, DataverseService_1.strVal)(historyRecord, 'sct_sct_reference');
        return (React.createElement(RequestLogsPanel_1.default, { requestId: id, reference: ref, service: this._service, onDismiss: function () { return _this.setState({ historyRecord: null }); } }));
    };
    // ── Take Action (Phase 4) ─────────────────────────────────────────────────
    /** Display label for the current approver tier */
    Dashboard.prototype._tierName = function () {
        switch (this.props.activeRole) {
            case 'manager': return 'Manager';
            case 'approver_tier1': return 'Tier 1';
            case 'approver_tier2': return 'Tier 2';
            case 'approver_tier3': return 'Tier 3';
            default: return 'Approver';
        }
    };
    Dashboard.prototype._renderTakeActionDialog = function () {
        var _this = this;
        var takeActionRecord = this.state.takeActionRecord;
        if (!takeActionRecord || !this._service)
            return null;
        return (React.createElement(common_1.TakeActionModal, { isOpen: takeActionRecord !== null, record: takeActionRecord, service: this._service, activeRole: this.props.activeRole || 'manager', currentSystemUserId: this._currentSystemUserId, context: this.props.context, soloTestMode: this.props.soloTestMode, onDismiss: function () { return _this.setState({ takeActionRecord: null }); }, onSuccess: function (msg) {
                _this.setState({
                    takeActionRecord: null,
                    saveSuccess: msg,
                });
                setTimeout(function () {
                    if (_this.props.activeRole === 'manager') {
                        void _this._loadManagerData();
                    }
                    else {
                        void _this._load();
                    }
                }, 500);
            } }));
    };
    // ── Main render ───────────────────────────────────────────────────────────
    Dashboard.prototype.render = function () {
        var _this = this;
        var _a = this.state, loading = _a.loading, error = _a.error, saveSuccess = _a.saveSuccess, saveError = _a.saveError, records = _a.records, searchText = _a.searchText, statusFilter = _a.statusFilter;
        // Merge static known statuses (always shown) with any live values from Dataverse
        var liveStatusValues = detectStatusValues(records);
        var allStatusValues = dashboardUtils_1.ALL_REQUEST_STATUS_VALUES.slice();
        liveStatusValues.forEach(function (v) {
            if (allStatusValues.indexOf(v) === -1)
                allStatusValues.push(v);
        });
        var filtered = this._filteredRecords();
        var counts = this._counts(filtered);
        var paged = this._pagedRecords(filtered);
        var isStrictSoD = this.props.soloTestMode === false;
        var sodLockedCount = 0;
        var activeRole = this.props.activeRole || 'admin';
        if (isStrictSoD && this.state.requesterTab === 'pending' && activeRole !== 'requester') {
            var baseForSod = this.state.masterRequestData.length > 0 ? this.state.masterRequestData : this.state.records;
            var myRequests = this.applyFilters(baseForSod, activeRole, 'myRequests', this._currentSystemUserId);
            var role_1 = activeRole.toLowerCase();
            sodLockedCount = myRequests.filter(function (r) {
                var _a;
                var rawPt = String(r['cr2a4_pendingtier'] || r['cr2a4_PendingTier'] || r['sct_pendingtier'] || '').toLowerCase();
                var rawStatus = typeof r['sct_request_status'] === 'number' ? r['sct_request_status'] : parseInt(String((_a = r['sct_request_status']) !== null && _a !== void 0 ? _a : ''), 10);
                if (role_1 === 'manager') {
                    return rawPt === 'manager' || (!rawPt && (rawStatus === 854000002 || rawStatus === 854000001));
                }
                if (role_1.indexOf('1') !== -1)
                    return rawPt.indexOf('1') !== -1;
                if (role_1.indexOf('2') !== -1)
                    return rawPt.indexOf('2') !== -1;
                if (role_1.indexOf('3') !== -1)
                    return rawPt.indexOf('3') !== -1;
                return false;
            }).length;
        }
        return (React.createElement("div", null,
            React.createElement("div", { className: AppShell_module_scss_1.default.statRow },
                React.createElement("div", { className: "".concat(AppShell_module_scss_1.default.statCard, " ").concat(AppShell_module_scss_1.default.statCardBlue) },
                    React.createElement(Icon_1.Icon, { iconName: "DocumentSet", style: { fontSize: 20, color: '#0078D4' }, "aria-hidden": "true" }),
                    React.createElement("span", { className: AppShell_module_scss_1.default.statValue }, loading ? '…' : counts.total),
                    React.createElement("span", { className: AppShell_module_scss_1.default.statLabel }, "Total Requests")),
                React.createElement("div", { className: "".concat(AppShell_module_scss_1.default.statCard, " ").concat(AppShell_module_scss_1.default.statCardOrange) },
                    React.createElement(Icon_1.Icon, { iconName: "Clock", style: { fontSize: 20, color: '#F7630C' }, "aria-hidden": "true" }),
                    React.createElement("span", { className: AppShell_module_scss_1.default.statValue }, loading ? '…' : counts.pending),
                    React.createElement("span", { className: AppShell_module_scss_1.default.statLabel }, "Pending Approval")),
                React.createElement("div", { className: "".concat(AppShell_module_scss_1.default.statCard, " ").concat(AppShell_module_scss_1.default.statCardGreen) },
                    React.createElement(Icon_1.Icon, { iconName: "CheckMark", style: { fontSize: 20, color: '#107C10' }, "aria-hidden": "true" }),
                    React.createElement("span", { className: AppShell_module_scss_1.default.statValue }, loading ? '…' : counts.approved),
                    React.createElement("span", { className: AppShell_module_scss_1.default.statLabel }, "Approved")),
                React.createElement("div", { className: "".concat(AppShell_module_scss_1.default.statCard, " ").concat(AppShell_module_scss_1.default.statCardTeal) },
                    React.createElement(Icon_1.Icon, { iconName: "Money", style: { fontSize: 20, color: '#00B294' }, "aria-hidden": "true" }),
                    React.createElement("span", { className: AppShell_module_scss_1.default.statValue, style: { fontSize: 20 } }, loading ? '…' : (0, dashboardUtils_1.formatCurrency)(counts.spend)),
                    React.createElement("span", { className: AppShell_module_scss_1.default.statLabel }, "Est. Total Spend"))),
            saveSuccess && (React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.success, isMultiline: false, onDismiss: function () { return _this.setState({ saveSuccess: null }); }, styles: { root: { borderRadius: 6, marginBottom: 12 } } }, saveSuccess)),
            saveError && (React.createElement(common_1.ErrorBanner, { message: saveError, onDismiss: function () { return _this.setState({ saveError: null }); } })),
            error && (React.createElement(common_1.ErrorBanner, { message: error, onRetry: function () { return void _this._load(); }, onDismiss: function () { return _this.setState({ error: null }); } })),
            this.props.activeRole === 'requester' && (React.createElement(Pivot_1.Pivot, { selectedKey: "myRequests", styles: { root: { marginBottom: 12 } } },
                React.createElement(Pivot_1.PivotItem, { headerText: "My Requests", itemKey: "myRequests", itemIcon: "Contact" }))),
            this.props.activeRole === 'manager' && (React.createElement(Pivot_1.Pivot, { selectedKey: this.state.requesterTab, onLinkClick: function (item) { return item && _this.setState({ requesterTab: item.props.itemKey, currentPage: 1 }); }, styles: { root: { marginBottom: 12 } } },
                React.createElement(Pivot_1.PivotItem, { headerText: "Pending My Approval", itemKey: "pending", itemIcon: "Clock" }),
                React.createElement(Pivot_1.PivotItem, { headerText: "Completed", itemKey: "completed", itemIcon: "CheckMark" }),
                React.createElement(Pivot_1.PivotItem, { headerText: "My Requests", itemKey: "myRequests", itemIcon: "Contact" }))),
            (this.props.activeRole === 'approver_tier1' ||
                this.props.activeRole === 'approver_tier2' ||
                this.props.activeRole === 'approver_tier3') && (React.createElement(Pivot_1.Pivot, { selectedKey: this.state.requesterTab, onLinkClick: function (item) { return item && _this.setState({ requesterTab: item.props.itemKey, currentPage: 1 }); }, styles: { root: { marginBottom: 12 } } },
                React.createElement(Pivot_1.PivotItem, { headerText: "Pending Approval", itemKey: "pending", itemIcon: "Clock" }),
                React.createElement(Pivot_1.PivotItem, { headerText: "Completed", itemKey: "completed", itemIcon: "CheckMark" }),
                React.createElement(Pivot_1.PivotItem, { headerText: "My Requests", itemKey: "myRequests", itemIcon: "Contact" }))),
            sodLockedCount > 0 && (React.createElement("div", { style: {
                    background: '#FFFBEB',
                    border: '1px solid #FCD34D',
                    borderRadius: 8,
                    padding: '12px 18px',
                    marginBottom: 14,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    gap: 12,
                    boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                } },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10 } },
                    React.createElement(Icon_1.Icon, { iconName: "ShieldAlert", style: { fontSize: 20, color: '#D97706' } }),
                    React.createElement("div", null,
                        React.createElement("strong", { style: { color: '#92400E', fontSize: 13 } }, "Segregation of Duties (SoD) Active:"),
                        ' ',
                        React.createElement("span", { style: { color: '#78350F', fontSize: 13 } },
                            sodLockedCount,
                            " pending request",
                            sodLockedCount > 1 ? 's' : '',
                            " submitted by you ",
                            sodLockedCount > 1 ? 'are' : 'is',
                            " tracked under ",
                            React.createElement("strong", null, "My Requests"),
                            " with self-approval locked to maintain audit compliance."))),
                this.props.onToggleSoloTestMode && (React.createElement("button", { type: "button", onClick: this.props.onToggleSoloTestMode, style: {
                        background: '#D97706', color: '#ffffff', border: 'none', borderRadius: 6,
                        padding: '6px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                        display: 'inline-flex', alignItems: 'center', gap: 6,
                        boxShadow: '0 1px 2px rgba(0,0,0,0.1)',
                    } },
                    React.createElement(Icon_1.Icon, { iconName: "Unlock", style: { fontSize: 12 } }),
                    "Enable Solo Test Mode to Review")))),
            React.createElement("div", { className: AppShell_module_scss_1.default.card },
                React.createElement("div", { className: AppShell_module_scss_1.default.toolbar },
                    React.createElement("div", { className: AppShell_module_scss_1.default.searchBox, role: "search" },
                        React.createElement(Icon_1.Icon, { iconName: "Search", style: { fontSize: 14, color: '#A19F9D', flexShrink: 0 }, "aria-hidden": "true" }),
                        React.createElement("input", { type: "search", placeholder: "Search all fields\u2026", value: searchText, "aria-label": "Search requests", onChange: function (e) { return _this.setState({ searchText: e.target.value, currentPage: 1 }); } }),
                        searchText && (React.createElement("button", { style: { border: 'none', background: 'none', cursor: 'pointer', padding: 0, color: '#A19F9D', lineHeight: 1 }, onClick: function () { return _this.setState({ searchText: '', currentPage: 1 }); }, "aria-label": "Clear search" },
                            React.createElement(Icon_1.Icon, { iconName: "Cancel", style: { fontSize: 12 } })))),
                    React.createElement("select", { className: AppShell_module_scss_1.default.filterSelect, value: statusFilter, "aria-label": "Filter by status", onChange: function (e) { return _this.setState({ statusFilter: e.target.value, currentPage: 1 }); } },
                        React.createElement("option", { value: "all" }, "All Statuses"),
                        allStatusValues.map(function (v) { return React.createElement("option", { key: v, value: v }, v); })),
                    React.createElement("div", { className: AppShell_module_scss_1.default.toolbarSpacer }),
                    React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement(Icon_1.Icon, { iconName: "DocumentSet", style: { color: '#0078D4' }, "aria-hidden": "true" }),
                        React.createElement(Text_1.Text, { variant: "mediumPlus", style: { fontWeight: 600, color: '#201F1E' } }, this.props.activeRole === 'requester' ? 'My Requests'
                            : this.props.activeRole === 'manager' ? (this.state.requesterTab === 'completed' ? 'Completed — Manager' : this.state.requesterTab === 'myRequests' ? 'My Submitted Requests' : 'Pending My Approval')
                                : this.props.activeRole === 'approver_tier1' ? (this.state.requesterTab === 'completed' ? 'Completed — Tier 1' : this.state.requesterTab === 'myRequests' ? 'My Submitted Requests' : 'Pending — Tier 1')
                                    : this.props.activeRole === 'approver_tier2' ? (this.state.requesterTab === 'completed' ? 'Completed — Tier 2' : this.state.requesterTab === 'myRequests' ? 'My Submitted Requests' : 'Pending — Tier 2')
                                        : this.props.activeRole === 'approver_tier3' ? (this.state.requesterTab === 'completed' ? 'Completed — Tier 3' : this.state.requesterTab === 'myRequests' ? 'My Submitted Requests' : 'Pending — Tier 3')
                                            : 'All Requests'),
                        this.state.roleFilterLoading && (React.createElement("span", { style: { fontSize: 11, color: '#A19F9D', fontStyle: 'italic' } }, "filtering\u2026")),
                        !loading && !this.state.roleFilterLoading && (React.createElement("span", { style: { background: '#EBF3FB', color: '#0078D4', fontSize: 12, fontWeight: 600, padding: '2px 10px', borderRadius: 12 } }, filtered.length))),
                    React.createElement("div", { className: AppShell_module_scss_1.default.toolbarSpacer }),
                    React.createElement(Button_1.DefaultButton, { text: "Export CSV", iconProps: { iconName: 'ExcelDocument' }, title: "Download currently filtered requests as CSV", disabled: loading || filtered.length === 0, onClick: function () { return _this._exportToCsv(filtered); }, styles: { root: { borderRadius: 6, height: 32, borderColor: '#EDEBE9' } } }),
                    React.createElement(Button_1.DefaultButton, { text: "Refresh", iconProps: { iconName: 'Refresh' }, onClick: function () { return void _this._load(); }, disabled: loading, styles: { root: { borderRadius: 6, height: 32 } } })),
                React.createElement("div", { style: { padding: 0 } }, loading ? (React.createElement(common_1.LoadingSpinner, { label: "Loading from Dataverse\u2026" })) : filtered.length === 0 && !error ? (React.createElement(common_1.EmptyState, { message: searchText || statusFilter !== 'all' ? 'No matching records' : 'No requests found', subMessage: searchText || statusFilter !== 'all' ? 'Try adjusting your search or filter.' : 'Submit a new spend request to get started.', iconName: "DocumentSearch", actionLabel: searchText || statusFilter !== 'all' ? 'Clear filters' : undefined, onAction: searchText || statusFilter !== 'all' ? function () { return _this.setState({ searchText: '', statusFilter: 'all', currentPage: 1 }); } : undefined })) : (React.createElement("div", { style: { overflowX: 'auto', width: '100%' } },
                    React.createElement(DetailsList_1.DetailsList, { items: paged, columns: this._buildColumns(allStatusValues), layoutMode: DetailsList_1.DetailsListLayoutMode.justified, selectionMode: DetailsList_1.SelectionMode.none, styles: { root: { minWidth: 600 } }, ariaLabelForGrid: "Spend control requests" })))),
                !loading && filtered.length > 0 && this._renderPagination(filtered.length)),
            this._renderViewPanel(),
            this._renderEditPanel(),
            this._renderDeleteDialog(),
            this._renderFieldInspector(),
            this._renderHistoryPanel(),
            this._renderTakeActionDialog()));
    };
    return Dashboard;
}(React.Component));
exports.default = Dashboard;
//# sourceMappingURL=Dashboard.js.map