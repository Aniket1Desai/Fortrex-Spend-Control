import { __assign, __awaiter, __extends, __generator } from "tslib";
import * as React from 'react';
import { strVal } from '../../services/DataverseService';
import { handleApprovalAction } from '../../services/approvalHandler';
import { ApprovalRoutingService } from '../../services/ApprovalRoutingService';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react/lib/Calendar';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { DetailsList, DetailsListLayoutMode, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { LoadingSpinner, EmptyState, ErrorBanner, TakeActionModal } from '../common';
import styles from '../shell/AppShell.module.scss';
import { formatCurrency, formatDate } from '../dashboard/dashboardUtils';
// ── Status option-set integers — confirmed from Power Apps sct_Request_Status (854000001 to 854000007) ──
// Source: Power Apps > Fortrex Spend Control T. > Tables > SCT Request > Columns > Status
var STATUS = {
    DRAFT: 854000001, // Draft
    SUBMITTED: 854000002, // Submitted
    IN_REVIEW: 854000003, // In Review   (covers Triage → Sourcing → FP&A → CFO)
    PO_ISSUED: 854000004, // PO Issued   (Procurement complete, awaiting receipt)
    REJECTED: 854000005, // Rejected
    RETURNED: 854000006, // Returned    (returned to requester for information)
    CLOSED: 854000007, // Closed      (received and complete)
};
var CFO_THRESHOLD = 10000;
// Note: Dataverse has 7 statuses only (confirmed from Power Apps).
// FP&A and CFO are conceptual sub-stages that BOTH operate while status = "In Review".
// They are distinguished in the UI by filtering on estimated value / new-supplier flag.
// Both resolve by issuing a PO → status transitions to "PO Issued" (854000003).
var STAGES = [
    {
        id: 'triage',
        stage: '3',
        title: 'SCT Desk — Triage',
        desc: 'Assign cost centre and advance submitted requests into the review pipeline',
        iconName: 'Sort',
        color: '#0078D4',
        bg: '#EBF3FB',
        watchLabels: ['Submitted'],
        primaryAction: 'Submit to Sourcing',
        primaryStatus: STATUS.IN_REVIEW,
        altAction: 'Return for Information',
        altStatus: STATUS.RETURNED,
    },
    {
        id: 'sourcing',
        stage: '4',
        title: 'Procurement — Sourcing',
        desc: 'Collect quotes and raise PO for standard requests (value ≤ $10k, known supplier)',
        iconName: 'ShoppingCart',
        color: '#107C10',
        bg: '#DFF6DD',
        watchLabels: ['In Review'],
        primaryAction: 'Raise PO',
        primaryStatus: STATUS.PO_ISSUED,
        altAction: 'Reject',
        altStatus: STATUS.REJECTED,
    },
    {
        id: 'fpa',
        stage: '5a',
        title: 'FP&A Review',
        desc: 'Post-sourcing review for requests >$10k or involving a new supplier — still "In Review" in Dataverse',
        iconName: 'BarChart4',
        color: '#8764B8',
        bg: '#F4F0FF',
        watchLabels: ['In Review'],
        primaryAction: 'FP&A Approved — Issue PO',
        primaryStatus: STATUS.PO_ISSUED,
        altAction: 'Reject',
        altStatus: STATUS.REJECTED,
    },
    {
        id: 'cfo',
        stage: '5b',
        title: 'CFO Approval',
        desc: 'L4 CFO gate — requests >$10k or new supplier requiring CFO sign-off before PO can be issued',
        iconName: 'Permissions',
        color: '#A4262C',
        bg: '#FDE7E9',
        watchLabels: ['In Review'],
        primaryAction: 'CFO Approved — Issue PO',
        primaryStatus: STATUS.PO_ISSUED,
        altAction: 'Reject',
        altStatus: STATUS.REJECTED,
    },
    {
        id: 'receiving',
        stage: '6',
        title: 'Receiving',
        desc: 'Confirm goods or services received and close the spend request',
        iconName: 'Completed',
        color: '#00B294',
        bg: '#E6F7F4',
        watchLabels: ['PO Issued'],
        primaryAction: 'Confirm Received & Close',
        primaryStatus: STATUS.CLOSED,
    },
];
// ── Date helpers (no padStart — not in this project's TS lib target) ──────────
function dateToEdm(d) {
    var y = d.getFullYear();
    var mo = d.getMonth() + 1;
    var da = d.getDate();
    var mm = mo < 10 ? ('0' + mo) : String(mo);
    var dd = da < 10 ? ('0' + da) : String(da);
    return y + '-' + mm + '-' + dd;
}
function edmToDate(str) {
    if (!str)
        return undefined;
    var d = new Date(str);
    return isNaN(d.getTime()) ? undefined : d;
}
// ── Status pill ───────────────────────────────────────────────────────────────
var StatusPill = function (_a) {
    var value = _a.value, color = _a.color, bg = _a.bg;
    return (React.createElement("span", { style: { background: bg, color: color, fontSize: 11, fontWeight: 700, padding: '3px 10px', borderRadius: 12, whiteSpace: 'nowrap' } }, value || '—'));
};
// ── Approver card ─────────────────────────────────────────────────────────────
var ApproverCard = function (_a) {
    var level = _a.level, name = _a.name, isDelegate = _a.isDelegate, wasEscalated = _a.wasEscalated;
    return (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', background: '#F8F8F8', borderRadius: 8, marginBottom: 8 } },
        React.createElement("div", { style: { width: 36, height: 36, borderRadius: 8, background: '#0078D4', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontWeight: 700, fontSize: 12, flexShrink: 0 } }, level),
        React.createElement("div", { style: { flex: 1 } },
            React.createElement("div", { style: { fontSize: 13, fontWeight: 600, color: '#201F1E' } }, name),
            React.createElement("div", { style: { display: 'flex', gap: 6, marginTop: 3 } },
                isDelegate && React.createElement("span", { style: { fontSize: 10, background: '#FFF4CE', color: '#905400', padding: '1px 7px', borderRadius: 8, fontWeight: 700 } }, "DELEGATE"),
                wasEscalated && React.createElement("span", { style: { fontSize: 10, background: '#FDE7E9', color: '#A4262C', padding: '1px 7px', borderRadius: 8, fontWeight: 700 } }, "R10 ESCALATED")))));
};
// ── Main component ────────────────────────────────────────────────────────────
var ReviewQueue = /** @class */ (function (_super) {
    __extends(ReviewQueue, _super);
    function ReviewQueue(props) {
        var _this = _super.call(this, props) || this;
        _this._primaryKey = null;
        _this._service = props.dataverseService;
        _this._router = new ApprovalRoutingService(props.dataverseService);
        _this.state = {
            activeStage: 'triage',
            records: [],
            loading: false,
            error: null,
            actionRecord: null,
            actionValues: {},
            poDate: undefined,
            saving: false,
            saveError: null,
            saveSuccess: null,
            preApproval: null,
            postSourcing: null,
            resolvedApprovers: null,
            loadingApprovers: false,
            pendingStageRows: [],
            loadingPending: false,
            systemUserId: null,
            managerApprovedIds: new Set(),
            directReportIds: new Set(),
            managerActionRecord: null,
            managerComments: '',
            managerSaving: false,
            managerError: null,
        };
        return _this;
    }
    ReviewQueue.prototype.componentDidMount = function () {
        void this._load();
        void this._resolveUser();
    };
    ReviewQueue.prototype.componentDidUpdate = function (prevProps) {
        var _this = this;
        if (prevProps.activeRole !== this.props.activeRole) {
            this.setState({ managerActionRecord: null, managerComments: '', managerError: null, saveSuccess: null });
            void this._load();
            if (!this.state.systemUserId) {
                void this._resolveUser();
            }
            else if (this.props.activeRole === 'manager') {
                void this._service.getDirectReportUserIds(this.state.systemUserId).then(function (reports) {
                    _this.setState({
                        directReportIds: new Set(reports.map(function (id) { return id.toLowerCase().replace(/[{}]/g, ''); })),
                    });
                }).catch(function () { });
                void this._loadManagerApprovedIds(this.state.systemUserId);
            }
        }
    };
    ReviewQueue.prototype._resolveUser = function () {
        return __awaiter(this, void 0, void 0, function () {
            var ctx, aadId, sysUserId, user, user, reports, _a;
            var _b, _c, _d, _e, _f, _g, _h, _j, _k;
            return __generator(this, function (_l) {
                switch (_l.label) {
                    case 0:
                        _l.trys.push([0, 8, , 9]);
                        ctx = this.props.context;
                        aadId = (_f = (_e = (_d = (_c = (_b = ctx.pageContext) === null || _b === void 0 ? void 0 : _b.aadInfo) === null || _c === void 0 ? void 0 : _c.userId) === null || _d === void 0 ? void 0 : _d.toString) === null || _e === void 0 ? void 0 : _e.call(_d)) !== null && _f !== void 0 ? _f : '';
                        sysUserId = null;
                        if (!aadId) return [3 /*break*/, 2];
                        return [4 /*yield*/, this._service.getCurrentSystemUser(aadId)];
                    case 1:
                        user = _l.sent();
                        sysUserId = (_g = user === null || user === void 0 ? void 0 : user.systemuserid) !== null && _g !== void 0 ? _g : null;
                        _l.label = 2;
                    case 2:
                        if (!(!sysUserId && ((_j = (_h = ctx.pageContext) === null || _h === void 0 ? void 0 : _h.user) === null || _j === void 0 ? void 0 : _j.email))) return [3 /*break*/, 4];
                        return [4 /*yield*/, this._service.getSystemUserByEmail(ctx.pageContext.user.email)];
                    case 3:
                        user = _l.sent();
                        sysUserId = (_k = user === null || user === void 0 ? void 0 : user.systemuserid) !== null && _k !== void 0 ? _k : null;
                        _l.label = 4;
                    case 4:
                        if (!sysUserId) return [3 /*break*/, 7];
                        reports = [];
                        if (!(this.props.activeRole === 'manager')) return [3 /*break*/, 6];
                        return [4 /*yield*/, this._service.getDirectReportUserIds(sysUserId).catch(function () { return []; })];
                    case 5:
                        reports = _l.sent();
                        _l.label = 6;
                    case 6:
                        this.setState({
                            systemUserId: sysUserId,
                            directReportIds: new Set(reports.map(function (id) { return id.toLowerCase().replace(/[{}]/g, ''); })),
                        });
                        if (this.props.activeRole === 'manager') {
                            void this._loadManagerApprovedIds(sysUserId);
                        }
                        _l.label = 7;
                    case 7: return [3 /*break*/, 9];
                    case 8:
                        _a = _l.sent();
                        return [3 /*break*/, 9];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    ReviewQueue.prototype._loadManagerApprovedIds = function (systemUserId) {
        return __awaiter(this, void 0, void 0, function () {
            var userId, ids, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        userId = systemUserId !== null && systemUserId !== void 0 ? systemUserId : this.state.systemUserId;
                        if (!userId)
                            return [2 /*return*/];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._service.getMyApprovedRequestIds(userId)];
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
    // ── Data ──────────────────────────────────────────────────────────────────
    ReviewQueue.prototype._load = function () {
        return __awaiter(this, void 0, void 0, function () {
            var records_1, guidFields, err_1;
            var _a, _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        this.setState({ loading: true, error: null, saveSuccess: null });
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._service.getRawRequests()];
                    case 2:
                        records_1 = _c.sent();
                        if (records_1.length && !this._primaryKey) {
                            guidFields = Object.keys(records_1[0]).filter(function (k) {
                                return !k.startsWith('@') && !k.startsWith('_') &&
                                    typeof records_1[0][k] === 'string' &&
                                    /^[0-9a-f]{8}-[0-9a-f]{4}-/.test(String(records_1[0][k]));
                            });
                            this._primaryKey = (_b = (_a = guidFields.find(function (k) { return k.endsWith('requestid'); })) !== null && _a !== void 0 ? _a : guidFields[0]) !== null && _b !== void 0 ? _b : null;
                        }
                        this.setState({ records: records_1, loading: false });
                        return [3 /*break*/, 4];
                    case 3:
                        err_1 = _c.sent();
                        this.setState({ error: err_1 instanceof Error ? err_1.message : 'Load failed', loading: false });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    ReviewQueue.prototype._getId = function (r) {
        if (!this._primaryKey)
            return null;
        var v = r[this._primaryKey];
        return typeof v === 'string' ? v : null;
    };
    // ── Queue filtering ───────────────────────────────────────────────────────
    ReviewQueue.prototype._queueFor = function (stageDef) {
        return this.state.records.filter(function (r) {
            var _a, _b;
            var label = String((_a = r['sct_request_status@OData.Community.Display.V1.FormattedValue']) !== null && _a !== void 0 ? _a : '');
            var matchesStatus = stageDef.watchLabels.some(function (w) { return label.toLowerCase() === w.toLowerCase(); });
            if (!matchesStatus)
                return false;
            // FP&A and CFO both watch "In Review" but only show high-value / new-supplier requests.
            // Sourcing shows the remainder (standard, within-threshold requests).
            var amt = parseFloat(String((_b = r['sct_estimated_total_value']) !== null && _b !== void 0 ? _b : 0));
            var isNew = r['sct_is_new_supplier'] === true || r['sct_is_new_supplier'] === 'Yes';
            var needsGate = (!isNaN(amt) && amt > CFO_THRESHOLD) || isNew;
            if (stageDef.id === 'sourcing')
                return !needsGate; // standard only
            if (stageDef.id === 'fpa')
                return needsGate; // above threshold / new supplier
            if (stageDef.id === 'cfo')
                return needsGate; // same population as FP&A
            return true; // triage + receiving: no extra filter
        });
    };
    // ── Pending stage calculation ─────────────────────────────────────────────
    ReviewQueue.prototype._calculatePendingStage = function (logs) {
        // Filter to approved logs only (sct_approvalstatus === 1)
        var approved = logs.filter(function (l) { return l['sct_approvalstatus'] === 1; });
        if (approved.length === 0) {
            // No approvals recorded yet — request sits with the requester's manager
            return [
                { stage: 'Pending with Manager', role: "Requester's Line Manager" },
            ];
        }
        // Determine stage based on tiers approved in approval logs
        // Dataverse sct_tiername: 4 = Manager, 1 = Tier 1, 2 = Tier 2, 3 = Tier 3
        var tiersApproved = new Set(approved.map(function (l) { return l['sct_tiername']; }));
        if (tiersApproved.has(3)) {
            return [{ stage: 'All Stages Approved', role: 'Awaiting PO Issuance' }];
        }
        else if (tiersApproved.has(2)) {
            return [{ stage: 'Pending with Tier 3 / CFO', role: 'Tier 3 / CFO Approver Group' }];
        }
        else if (tiersApproved.has(1)) {
            return [{ stage: 'Pending with Tier 2', role: 'Tier 2 AAD Approver Group' }];
        }
        else if (tiersApproved.has(4)) {
            return [{ stage: 'Pending with Tier 1', role: 'Tier 1 AAD Approver Group' }];
        }
        else {
            return [{ stage: 'Pending with Manager', role: "Requester's Line Manager" }];
        }
    };
    // ── Open action panel ─────────────────────────────────────────────────────
    ReviewQueue.prototype._openAction = function (record, stageDef) {
        var _this = this;
        var _a;
        var poDateStr = strVal(record, 'sct_po_issue_date');
        var statusLabel = String((_a = record['sct_request_status@OData.Community.Display.V1.FormattedValue']) !== null && _a !== void 0 ? _a : '').toLowerCase();
        var isInReview = statusLabel === 'in review';
        this.setState({
            actionRecord: record,
            actionValues: { sct_cost_center: strVal(record, 'sct_cost_center'), sct_po_number: strVal(record, 'sct_po_number') },
            poDate: poDateStr ? edmToDate(poDateStr) : undefined,
            saveError: null,
            preApproval: null,
            postSourcing: null,
            resolvedApprovers: null,
            loadingApprovers: false,
            pendingStageRows: [],
            loadingPending: isInReview,
        }, function () {
            void _this._loadRoutingData(record, stageDef);
            if (isInReview)
                void _this._loadPendingStage(record);
        });
    };
    // ── Fetch approval logs and derive pending stage ───────────────────────────
    ReviewQueue.prototype._loadPendingStage = function (record) {
        return __awaiter(this, void 0, void 0, function () {
            var id, logs, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        id = this._getId(record);
                        if (!id) {
                            this.setState({ loadingPending: false });
                            return [2 /*return*/];
                        }
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._service.getApprovalLogs(id)];
                    case 2:
                        logs = _b.sent();
                        this.setState({
                            pendingStageRows: this._calculatePendingStage(Array.isArray(logs) ? logs : []),
                            loadingPending: false,
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _b.sent();
                        this.setState({ loadingPending: false });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    // ── Load routing data for the opened action panel ─────────────────────────
    ReviewQueue.prototype._loadRoutingData = function (record, stageDef) {
        return __awaiter(this, void 0, void 0, function () {
            var estimatedValue, isNewSupplier, siteId, preApproval, postSourcing, requesterGuid, resolvedApprovers, _a;
            var _b, _c, _d, _e;
            return __generator(this, function (_f) {
                switch (_f.label) {
                    case 0:
                        estimatedValue = parseFloat(String((_b = record['sct_estimated_total_value']) !== null && _b !== void 0 ? _b : 0));
                        isNewSupplier = record['sct_is_new_supplier'] === true || record['sct_is_new_supplier'] === 'Yes';
                        siteId = String(record['_sct_site_value'] ? record['_sct_site_value'] : (_c = record['_sct_siteid_value']) !== null && _c !== void 0 ? _c : '');
                        preApproval = this._router.calculatePreApprovals(estimatedValue, isNewSupplier);
                        postSourcing = this._router.calculatePostSourcingRequirements(estimatedValue, isNewSupplier);
                        this.setState({ preApproval: preApproval, postSourcing: postSourcing });
                        if (!(stageDef.id === 'triage' || stageDef.id === 'fpa' || stageDef.id === 'cfo')) return [3 /*break*/, 4];
                        requesterGuid = String((_e = (_d = record['_sct_requester_value']) !== null && _d !== void 0 ? _d : record['_sct_requestorid_value']) !== null && _e !== void 0 ? _e : '');
                        if (!siteId) return [3 /*break*/, 4];
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
    // ── Save action ───────────────────────────────────────────────────────────
    ReviewQueue.prototype._takeAction = function (stageDef, targetStatus) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, actionRecord, actionValues, poDate, systemUserId, id, payload, existingPt, cc, po, currentStatusLabel, stageName, err_2;
            var _b, _c, _d;
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        _a = this.state, actionRecord = _a.actionRecord, actionValues = _a.actionValues, poDate = _a.poDate, systemUserId = _a.systemUserId;
                        if (!actionRecord)
                            return [2 /*return*/];
                        id = this._getId(actionRecord);
                        if (!id)
                            return [2 /*return*/];
                        this.setState({ saving: true, saveError: null });
                        payload = { sct_request_status: targetStatus };
                        // Synchronize cr2a4_pendingtier: clear on termination states, preserve or set on In Review
                        if (targetStatus === STATUS.PO_ISSUED ||
                            targetStatus === STATUS.REJECTED ||
                            targetStatus === STATUS.RETURNED ||
                            targetStatus === STATUS.CLOSED) {
                            payload['cr2a4_pendingtier'] = null;
                        }
                        else if (targetStatus === STATUS.IN_REVIEW) {
                            existingPt = String(actionRecord['cr2a4_pendingtier'] || actionRecord['sct_pendingtier'] || '').trim();
                            payload['cr2a4_pendingtier'] = existingPt || 'Manager';
                        }
                        cc = (_b = actionValues['sct_cost_center']) === null || _b === void 0 ? void 0 : _b.trim();
                        if (cc)
                            payload['sct_cost_center'] = cc;
                        po = (_c = actionValues['sct_po_number']) === null || _c === void 0 ? void 0 : _c.trim();
                        if (po)
                            payload['sct_po_number'] = po;
                        if (poDate)
                            payload['sct_po_issue_date'] = dateToEdm(poDate);
                        _e.label = 1;
                    case 1:
                        _e.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, this._service.updateRequest(id, payload)];
                    case 2:
                        _e.sent();
                        currentStatusLabel = strVal(actionRecord, 'sct_request_status');
                        stageName = stageDef.title;
                        return [4 /*yield*/, this._service.createAuditLog(__assign({ sct_action: targetStatus === STATUS.PO_ISSUED ? 1 : targetStatus === STATUS.REJECTED ? 5 : targetStatus === STATUS.RETURNED ? 4 : 2, sct_description: "Stage ".concat(stageDef.stage, " (").concat(stageName, ") executed: ").concat(stageDef.primaryAction), sct_oldvalue: "Status: ".concat(currentStatusLabel), sct_newvalue: "Status: ".concat(targetStatus, ", Pending Tier: ").concat((_d = payload['cr2a4_pendingtier']) !== null && _d !== void 0 ? _d : 'null'), 'sct_requestid@odata.bind': "/sct_requests(".concat(id, ")") }, (systemUserId ? { 'sct_actionby@odata.bind': "/systemusers(".concat(systemUserId, ")") } : {}))).catch(function () { })];
                    case 3:
                        _e.sent();
                        this.setState({ saving: false, actionRecord: null, saveSuccess: 'Record updated successfully.' });
                        void this._load();
                        return [3 /*break*/, 5];
                    case 4:
                        err_2 = _e.sent();
                        this.setState({ saving: false, saveError: err_2 instanceof Error ? err_2.message : 'Save failed' });
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    // ── Helpers ───────────────────────────────────────────────────────────────
    ReviewQueue.prototype._estimatedValue = function (r) {
        var _a;
        return parseFloat(String((_a = r['sct_estimated_total_value']) !== null && _a !== void 0 ? _a : 0));
    };
    ReviewQueue.prototype._isNewSupplier = function (r) {
        return r['sct_is_new_supplier'] === true || r['sct_is_new_supplier'] === 'Yes';
    };
    // ── Approver queue methods ─────────────────────────────────────────────────
    /** Records waiting for approval by the active approver role */
    ReviewQueue.prototype._approverQueue = function () {
        var _this = this;
        var _a, _b, _c, _d;
        var _e = this.state, records = _e.records, systemUserId = _e.systemUserId, directReportIds = _e.directReportIds, managerApprovedIds = _e.managerApprovedIds;
        var role = this.props.activeRole || 'manager';
        var myId = systemUserId ? systemUserId.toLowerCase().replace(/[{}]/g, '') : '';
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        var ctxAny = this.props.context;
        var aadId = (((_d = (_c = (_b = (_a = ctxAny === null || ctxAny === void 0 ? void 0 : ctxAny.pageContext) === null || _a === void 0 ? void 0 : _a.aadInfo) === null || _b === void 0 ? void 0 : _b.userId) === null || _c === void 0 ? void 0 : _c.toString) === null || _d === void 0 ? void 0 : _d.call(_c)) || '').toLowerCase().replace(/[{}]/g, '');
        var isDevOrWorkbench = typeof window !== 'undefined' && (window.location.href.indexOf('workbench') !== -1 ||
            window.location.href.indexOf('devrishabhsoft') !== -1 ||
            window.location.hostname.indexOf('localhost') !== -1 ||
            window.location.hostname.indexOf('127.0.0.1') !== -1);
        return records.filter(function (r) {
            var _a, _b, _c, _d, _e, _f, _g, _h, _j;
            var rawReq = String((_c = (_b = (_a = r['_sct_requester_value']) !== null && _a !== void 0 ? _a : r['_createdby_value']) !== null && _b !== void 0 ? _b : r['sct_requester']) !== null && _c !== void 0 ? _c : '');
            var reqId = rawReq.toLowerCase().replace(/[{}]/g, '');
            var createdById = r['_createdby_value'] ? String(r['_createdby_value']).toLowerCase().replace(/[{}]/g, '') : '';
            var userDisplayName = (_this.props.context.pageContext.user.displayName || '').toLowerCase().trim();
            var userEmail = (_this.props.context.pageContext.user.email || '').toLowerCase().trim();
            var reqFormatted = (strVal(r, '_sct_requester_value') || strVal(r, 'sct_requester') || '').toLowerCase().trim();
            var reqNameFromOdata = (strVal(r, '_sct_requester_value@OData.Community.Display.V1.FormattedValue') || strVal(r, 'requester') || '').toLowerCase().trim();
            var isOwnRequest = Boolean((myId && (reqId === myId || createdById === myId || reqId === aadId || createdById === aadId)) ||
                (userDisplayName && (reqFormatted.indexOf(userDisplayName) !== -1 || reqNameFromOdata.indexOf(userDisplayName) !== -1)) ||
                (userEmail && (reqFormatted.indexOf(userEmail) !== -1 || reqId.indexOf(userEmail) !== -1)));
            // 1. Global Segregation of Duties (SoD) Guard:
            var allowSelfApproval = Boolean(_this.props.soloTestMode);
            if (!allowSelfApproval && isOwnRequest) {
                return false;
            }
            var rawStatus = typeof r['sct_request_status'] === 'number'
                ? r['sct_request_status']
                : parseInt(String((_d = r['sct_request_status']) !== null && _d !== void 0 ? _d : ''), 10);
            var statusLabel = strVal(r, 'sct_request_status').toLowerCase();
            var isReview = rawStatus === 854000003 || rawStatus === 854000002 || statusLabel === 'in review' || statusLabel.indexOf('in review') !== -1;
            var ptRaw = String(r['cr2a4_pendingtier'] ||
                r['cr2a4_PendingTier'] ||
                r['cr2a4_pending_tier'] ||
                r['sct_pendingtier'] ||
                r['pendingtier'] ||
                '').trim();
            if (!ptRaw && statusLabel.indexOf('tier') !== -1) {
                var m = statusLabel.match(/tier\s*([123])/i);
                if (m)
                    ptRaw = "Tier ".concat(m[1]);
            }
            var pt = ptRaw.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
            var normRole = (role || 'manager').toLowerCase().trim();
            var isManager = normRole === 'manager' || normRole.indexOf('manager') !== -1;
            var isTier1 = normRole === 'approver_tier1' || normRole.indexOf('tier 1') !== -1 || normRole.indexOf('tier1') !== -1;
            var isTier2 = normRole === 'approver_tier2' || normRole.indexOf('tier 2') !== -1 || normRole.indexOf('tier2') !== -1;
            var isTier3 = normRole === 'approver_tier3' || normRole.indexOf('tier 3') !== -1 || normRole.indexOf('tier3') !== -1;
            if (isManager) {
                var isPendingStatus = rawStatus === 854000002 || rawStatus === 854000003 || rawStatus === 854000001 || statusLabel === 'in review' || statusLabel === 'submitted';
                if (!isPendingStatus)
                    return false;
                var isPendingWithManager = pt === 'manager' || ((rawStatus === 854000002 || rawStatus === 854000001) && !pt) || (statusLabel === 'submitted' && !pt);
                if (!isPendingWithManager)
                    return false;
                var pk = _this._getId(r);
                if (pk && managerApprovedIds.has(pk))
                    return false;
                var rawMgr = (_j = (_h = (_g = (_f = (_e = r['_sct_manager_value']) !== null && _e !== void 0 ? _e : r['_cr2a4_manager_value']) !== null && _f !== void 0 ? _f : r['_manager_value']) !== null && _g !== void 0 ? _g : r['sct_manager']) !== null && _h !== void 0 ? _h : r['cr2a4_manager']) !== null && _j !== void 0 ? _j : '';
                var mgrId = rawMgr ? String(rawMgr).toLowerCase().replace(/[{}]/g, '') : '';
                var mgrFormatted = (strVal(r, '_cr2a4_manager_value') || strVal(r, '_sct_manager_value') || strVal(r, '_manager_value') || strVal(r, 'cr2a4_manager') || strVal(r, 'sct_manager') || strVal(r, 'manager') || '').toLowerCase();
                var isMgrByName = Boolean(userDisplayName) && (mgrFormatted.indexOf(userDisplayName) !== -1 ||
                    (Boolean(userEmail) && mgrFormatted.indexOf(userEmail) !== -1));
                var isDirectReport = Boolean(directReportIds && (directReportIds.has(reqId) || (createdById && directReportIds.has(createdById))));
                var isManagerOfReq = (mgrId && (mgrId === myId || (aadId && mgrId === aadId.toLowerCase())))
                    || isMgrByName
                    || isDirectReport;
                if (!isManagerOfReq)
                    return false;
                return true;
            }
            if (isTier1) {
                if (!isReview)
                    return false;
                return pt === 'tier1' || pt.indexOf('tier1') !== -1 || statusLabel.indexOf('tier 1') !== -1 || statusLabel.indexOf('tier1') !== -1 || pt === '1';
            }
            if (isTier2) {
                if (!isReview)
                    return false;
                return pt === 'tier2' || pt.indexOf('tier2') !== -1 || statusLabel.indexOf('tier 2') !== -1 || statusLabel.indexOf('tier2') !== -1 || pt === '2';
            }
            if (isTier3) {
                if (!isReview)
                    return false;
                return pt === 'tier3' || pt.indexOf('tier3') !== -1 || statusLabel.indexOf('tier 3') !== -1 || statusLabel.indexOf('tier3') !== -1 || pt === '3';
            }
            return false;
        });
    };
    /** Take action (approve / reject / return) on a record from the approver queue */
    ReviewQueue.prototype._handleManagerAction = function (decision) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, managerActionRecord, managerComments, id, graphClient, _b, role, rawReq, result, successMsg, err_3;
            var _c, _d, _e;
            return __generator(this, function (_f) {
                switch (_f.label) {
                    case 0:
                        _a = this.state, managerActionRecord = _a.managerActionRecord, managerComments = _a.managerComments;
                        if (!managerActionRecord)
                            return [2 /*return*/];
                        id = this._getId(managerActionRecord);
                        if (!id)
                            return [2 /*return*/];
                        this.setState({ managerSaving: true, managerError: null });
                        _f.label = 1;
                    case 1:
                        _f.trys.push([1, 8, , 9]);
                        graphClient = void 0;
                        _f.label = 2;
                    case 2:
                        _f.trys.push([2, 4, , 5]);
                        return [4 /*yield*/, this.props.context.msGraphClientFactory.getClient('3')];
                    case 3:
                        graphClient = _f.sent();
                        return [3 /*break*/, 5];
                    case 4:
                        _b = _f.sent();
                        return [3 /*break*/, 5];
                    case 5:
                        role = this.props.activeRole || 'manager';
                        rawReq = managerActionRecord
                            ? String((_e = (_d = (_c = managerActionRecord['_sct_requester_value']) !== null && _c !== void 0 ? _c : managerActionRecord['sct_requester']) !== null && _d !== void 0 ? _d : managerActionRecord['_createdby_value']) !== null && _e !== void 0 ? _e : '')
                            : '';
                        return [4 /*yield*/, handleApprovalAction({
                                requestId: id,
                                role: role,
                                action: decision,
                                comments: managerComments,
                                service: this._service,
                                currentSystemUserId: this.state.systemUserId,
                                requesterId: rawReq,
                                graphClient: graphClient,
                                webPartUrl: this.props.context.pageContext.web.absoluteUrl,
                                approverDisplayName: this.props.context.pageContext.user.displayName,
                                currentUserEmail: this.props.context.pageContext.user.email || undefined,
                                soloTestMode: Boolean(this.props.soloTestMode),
                            })];
                    case 6:
                        result = _f.sent();
                        successMsg = result.emailDispatched
                            ? 'Action saved and notification email sent.'
                            : 'Action saved successfully.';
                        this.setState({ managerActionRecord: null, managerComments: '', managerSaving: false, saveSuccess: successMsg });
                        // Wait 500ms for Dataverse to commit the PATCH before re-fetching
                        return [4 /*yield*/, new Promise(function (resolve) { return setTimeout(resolve, 500); })];
                    case 7:
                        // Wait 500ms for Dataverse to commit the PATCH before re-fetching
                        _f.sent();
                        void this._load();
                        // Refresh approved IDs if manager acted so the record immediately disappears from Pending queue
                        if (role === 'manager') {
                            void this._loadManagerApprovedIds();
                        }
                        return [3 /*break*/, 9];
                    case 8:
                        err_3 = _f.sent();
                        this.setState({
                            managerSaving: false,
                            managerError: err_3 instanceof Error ? err_3.message : 'Action failed. Please try again.',
                        });
                        return [3 /*break*/, 9];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    ReviewQueue.prototype._renderManagerQueue = function () {
        var _this = this;
        var _a = this.state, loading = _a.loading, systemUserId = _a.systemUserId, managerActionRecord = _a.managerActionRecord, managerComments = _a.managerComments, managerSaving = _a.managerSaving, managerError = _a.managerError;
        var role = this.props.activeRole || 'manager';
        var queue = this._approverQueue();
        var queueTitle = role === 'manager' ? 'Manager Review Queue'
            : role === 'approver_tier1' ? 'Tier 1 Review Queue'
                : role === 'approver_tier2' ? 'Tier 2 Review Queue'
                    : role === 'approver_tier3' ? 'Tier 3 Review Queue'
                        : 'Approver Review Queue';
        var queueDesc = role === 'manager' ? 'Requests assigned to you as line manager — pending your approval'
            : role === 'approver_tier1' ? 'Requests awaiting Tier 1 approval — status: In Review'
                : role === 'approver_tier2' ? 'Requests awaiting Tier 2 approval — status: In Review'
                    : role === 'approver_tier3' ? 'Requests awaiting Tier 3 / Final approval — status: In Review'
                        : 'Requests awaiting review';
        var headerIcon = role === 'manager' ? 'People' : 'WorkFlow';
        var themeColor = role === 'manager' ? '#0078D4' : role === 'approver_tier1' ? '#00897B' : role === 'approver_tier2' ? '#0277BD' : '#8764B8';
        var columns = [
            {
                key: 'ref', name: 'Reference', minWidth: 110, maxWidth: 140, isResizable: true,
                onRender: function (r) { return (React.createElement("span", { style: { fontWeight: 700, color: '#0078D4', fontFamily: 'monospace', fontSize: 12 } }, strVal(r, 'sct_sct_reference') || '—')); },
            },
            {
                key: 'requester', name: 'Requester', minWidth: 120, maxWidth: 180, isResizable: true,
                onRender: function (r) { return React.createElement("span", { style: { fontSize: 13 } }, strVal(r, '_sct_requester_value') || '—'); },
            },
            {
                key: 'site', name: 'Site', minWidth: 100, maxWidth: 150, isResizable: true,
                onRender: function (r) { return React.createElement("span", { style: { fontSize: 13 } }, strVal(r, '_sct_site_value') || '—'); },
            },
            {
                key: 'type', name: 'Type', minWidth: 90, maxWidth: 130, isResizable: true,
                onRender: function (r) { return React.createElement("span", { style: { fontSize: 13 } }, strVal(r, 'sct_purchase_type') || '—'); },
            },
            {
                key: 'value', name: 'Est. Value', minWidth: 100, maxWidth: 130, isResizable: true,
                onRender: function (r) {
                    var amt = _this._estimatedValue(r);
                    if (isNaN(amt))
                        return React.createElement("span", { style: { color: '#C8C6C4' } }, "\u2014");
                    return (React.createElement("span", { style: { fontWeight: 700, color: '#107C10', fontFamily: 'monospace', fontSize: 13 } }, formatCurrency(amt)));
                },
            },
            {
                key: 'submitted', name: 'Submitted', minWidth: 100, maxWidth: 130, isResizable: true,
                onRender: function (r) {
                    var d = strVal(r, 'sct_submitted_date') || strVal(r, 'createdon');
                    return React.createElement("span", { style: { fontSize: 12, color: '#605E5C' } }, d ? formatDate(d) : '—');
                },
            },
            {
                key: '_act', name: '', minWidth: 95, maxWidth: 100,
                onRender: function (r) { return (React.createElement("button", { title: "Review and take action", onClick: function (e) { e.stopPropagation(); _this.setState({ managerActionRecord: r, managerComments: '', managerError: null }); }, style: {
                        display: 'inline-flex', alignItems: 'center', gap: 5,
                        background: 'linear-gradient(135deg, #0078D4 0%, #005A9E 100%)',
                        color: '#fff', border: 'none', borderRadius: 20,
                        padding: '5px 13px', fontSize: 12, fontWeight: 600,
                        cursor: 'pointer', letterSpacing: '0.2px',
                        boxShadow: '0 1px 4px rgba(0,120,212,0.35)',
                        whiteSpace: 'nowrap',
                    } },
                    React.createElement(Icon, { iconName: "WorkFlow", style: { fontSize: 11 } }),
                    "Review")); },
            },
        ];
        return (React.createElement("div", null,
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16, padding: '12px 16px', background: '#EBF3FB', borderRadius: 10 } },
                React.createElement("div", { style: { width: 44, height: 44, borderRadius: 10, background: themeColor, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 } },
                    React.createElement(Icon, { iconName: headerIcon, style: { fontSize: 22, color: '#fff' } })),
                React.createElement("div", { style: { flex: 1 } },
                    React.createElement(Text, { variant: "mediumPlus", style: { color: '#201F1E', fontWeight: 700, display: 'block' } }, queueTitle),
                    React.createElement(Text, { variant: "small", style: { color: '#605E5C' } }, queueDesc)),
                React.createElement("span", { style: { background: themeColor, color: '#fff', fontSize: 18, fontWeight: 700, padding: '4px 14px', borderRadius: 12 } }, queue.length)),
            !systemUserId && !loading && (React.createElement(MessageBar, { messageBarType: MessageBarType.warning, styles: { root: { borderRadius: 6, marginBottom: 12 } } }, "Resolving your Dataverse user account\u2026")),
            queue.length === 0 && !loading
                ? React.createElement(EmptyState, { message: "No pending reviews", subMessage: "All requests assigned to you are up to date.", iconName: headerIcon })
                : React.createElement(DetailsList, { items: queue, columns: columns, layoutMode: DetailsListLayoutMode.justified, selectionMode: SelectionMode.none, styles: { root: { minWidth: 0 } }, ariaLabelForGrid: "".concat(queueTitle, " grid") }),
            React.createElement(TakeActionModal, { isOpen: managerActionRecord !== null, record: managerActionRecord, service: this._service, activeRole: this.props.activeRole || 'manager', currentSystemUserId: this.state.systemUserId, context: this.props.context, onDismiss: function () { return _this.setState({ managerActionRecord: null }); }, onSuccess: function (msg) {
                    _this.setState({
                        managerActionRecord: null,
                        saveSuccess: msg,
                    });
                    setTimeout(function () {
                        void _this._load();
                        if ((_this.props.activeRole || 'manager') === 'manager') {
                            void _this._loadManagerApprovedIds();
                        }
                    }, 500);
                } })));
    };
    // ── Render: Stage tabs ─────────────────────────────────────────────────────
    ReviewQueue.prototype._renderStageTabs = function () {
        var _this = this;
        var activeStage = this.state.activeStage;
        return (React.createElement("div", { style: { display: 'flex', gap: 4, marginBottom: 20, flexWrap: 'wrap' } }, STAGES.map(function (s) {
            var count = _this._queueFor(s).length;
            var isActive = activeStage === s.id;
            return (React.createElement("button", { key: s.id, onClick: function () { return _this.setState({ activeStage: s.id, saveSuccess: null }); }, style: {
                    display: 'flex', alignItems: 'center', gap: 8,
                    padding: '8px 16px', borderRadius: 8,
                    border: isActive ? "2px solid ".concat(s.color) : '2px solid #EDEBE9',
                    background: isActive ? s.bg : '#FFFFFF',
                    color: isActive ? s.color : '#605E5C',
                    fontWeight: isActive ? 700 : 500, fontSize: 13, cursor: 'pointer',
                }, "aria-selected": isActive },
                React.createElement(Icon, { iconName: s.iconName, style: { fontSize: 15, color: isActive ? s.color : '#A19F9D' } }),
                React.createElement("span", null,
                    "Stage ",
                    s.stage,
                    ": ",
                    s.title.split('—')[0].trim()),
                !_this.state.loading && count > 0 && (React.createElement("span", { style: { background: s.color, color: '#fff', fontSize: 11, fontWeight: 700, padding: '1px 7px', borderRadius: 10 } }, count))));
        })));
    };
    // ── Render: Workbench table ────────────────────────────────────────────────
    ReviewQueue.prototype._renderWorkbench = function (stageDef) {
        var _this = this;
        var queue = this._queueFor(stageDef);
        var columns = [
            {
                key: 'ref', name: 'Reference', minWidth: 110, maxWidth: 140, isResizable: true,
                onRender: function (r) { return (React.createElement("span", { style: { fontWeight: 700, color: '#0078D4', fontFamily: 'monospace', fontSize: 12 } }, strVal(r, 'sct_sct_reference') || '—')); },
            },
            {
                key: 'requester', name: 'Requester', minWidth: 120, maxWidth: 180, isResizable: true,
                onRender: function (r) { return React.createElement("span", { style: { fontSize: 13 } }, strVal(r, '_sct_requester_value') || '—'); },
            },
            {
                key: 'site', name: 'Site', minWidth: 100, maxWidth: 160, isResizable: true,
                onRender: function (r) { return React.createElement("span", { style: { fontSize: 13 } }, strVal(r, '_sct_site_value') || '—'); },
            },
            {
                key: 'type', name: 'Type', minWidth: 90, maxWidth: 130, isResizable: true,
                onRender: function (r) { return React.createElement("span", { style: { fontSize: 13 } }, strVal(r, 'sct_purchase_type') || '—'); },
            },
            {
                key: 'value', name: 'Est. Value', minWidth: 100, maxWidth: 130, isResizable: true,
                onRender: function (r) {
                    var amt = _this._estimatedValue(r);
                    if (isNaN(amt))
                        return React.createElement("span", { style: { color: '#C8C6C4' } }, "\u2014");
                    var flag = amt > CFO_THRESHOLD;
                    return (React.createElement("span", { style: { fontWeight: 700, color: flag ? '#A4262C' : '#107C10', fontFamily: 'monospace', fontSize: 13 } },
                        formatCurrency(amt),
                        flag && React.createElement("span", { title: "Exceeds $10k \u2014 CFO route" }, " \u2691")));
                },
            },
            {
                key: 'supplier', name: 'Supplier', minWidth: 110, maxWidth: 160, isResizable: true,
                onRender: function (r) {
                    return _this._isNewSupplier(r)
                        ? React.createElement(StatusPill, { value: "New Supplier", color: "#905400", bg: "#FFF4CE" })
                        : React.createElement("span", { style: { fontSize: 13 } }, strVal(r, '_sct_suggested_supplier_value') || '—');
                },
            },
            {
                key: 'submitted', name: 'Submitted', minWidth: 100, maxWidth: 130, isResizable: true,
                onRender: function (r) {
                    var d = strVal(r, 'sct_submitted_date') || strVal(r, 'createdon');
                    return React.createElement("span", { style: { fontSize: 12, color: '#605E5C' } }, d ? formatDate(d) : '—');
                },
            },
            {
                key: '_act', name: '', minWidth: 80, maxWidth: 80,
                onRender: function (r) { return (React.createElement(PrimaryButton, { text: "Action", onClick: function (e) { e.stopPropagation(); _this._openAction(r, stageDef); }, styles: { root: { borderRadius: 6, height: 28, fontSize: 12, padding: '0 12px' } } })); },
            },
        ];
        return (React.createElement("div", null,
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16, padding: '12px 16px', background: stageDef.bg, borderRadius: 10 } },
                React.createElement("div", { style: { width: 44, height: 44, borderRadius: 10, background: stageDef.color, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 } },
                    React.createElement(Icon, { iconName: stageDef.iconName, style: { fontSize: 22, color: '#fff' } })),
                React.createElement("div", { style: { flex: 1 } },
                    React.createElement(Text, { variant: "mediumPlus", style: { color: '#201F1E', fontWeight: 700, display: 'block' } },
                        "Stage ",
                        stageDef.stage,
                        ": ",
                        stageDef.title),
                    React.createElement(Text, { variant: "small", style: { color: '#605E5C' } }, stageDef.desc)),
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                    React.createElement("span", { style: { background: stageDef.color, color: '#fff', fontSize: 18, fontWeight: 700, padding: '4px 14px', borderRadius: 12 } }, queue.length),
                    React.createElement(Text, { variant: "small", style: { color: stageDef.color, fontWeight: 600 } }, queue.length === 1 ? 'item' : 'items'))),
            React.createElement("div", { style: { fontSize: 11, color: '#A19F9D', marginBottom: 10 } },
                "Monitors status: ",
                stageDef.watchLabels.map(function (l) { return "\"".concat(l, "\""); }).join(', '),
                (stageDef.id === 'fpa' || stageDef.id === 'cfo') &&
                    React.createElement("span", { style: { color: '#605E5C' } },
                        ' · ',
                        "Filtered to requests > $",
                        CFO_THRESHOLD.toLocaleString(),
                        " or new supplier")),
            queue.length === 0
                ? React.createElement(EmptyState, { message: "Queue is empty", subMessage: "No requests at the ".concat(stageDef.watchLabels.join(' / '), " stage."), iconName: stageDef.iconName })
                : React.createElement(DetailsList, { items: queue, columns: columns, layoutMode: DetailsListLayoutMode.justified, selectionMode: SelectionMode.none, styles: { root: { minWidth: 0 } }, ariaLabelForGrid: "".concat(stageDef.title, " queue") })));
    };
    // ── Render: Action panel ───────────────────────────────────────────────────
    // ── Render: Pending Stage Tracker ────────────────────────────────────────
    ReviewQueue.prototype._renderPendingTracker = function () {
        var _a;
        var _b = this.state, loadingPending = _b.loadingPending, pendingStageRows = _b.pendingStageRows, actionRecord = _b.actionRecord;
        if (!actionRecord)
            return null;
        var statusLabel = String((_a = actionRecord['sct_request_status@OData.Community.Display.V1.FormattedValue']) !== null && _a !== void 0 ? _a : '').toLowerCase();
        if (statusLabel !== 'in review')
            return null;
        return (React.createElement("div", { style: { marginTop: 16, marginBottom: 4 } },
            React.createElement("div", { style: { fontSize: 13, fontWeight: 700, color: '#0078D4', marginBottom: 10 } }, "Pending Stage Tracker"),
            loadingPending ? (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8, padding: '10px 0' } },
                React.createElement(Spinner, { size: SpinnerSize.small }),
                React.createElement("span", { style: { fontSize: 12, color: '#605E5C' } }, "Calculating current stage\u2026"))) : (React.createElement("table", { style: {
                    width: '100%', borderCollapse: 'collapse',
                    border: '1px solid #EDEBE9', borderRadius: 8, overflow: 'hidden',
                    fontSize: 13,
                } },
                React.createElement("thead", null,
                    React.createElement("tr", { style: { background: '#EBF3FB' } },
                        React.createElement("th", { style: {
                                padding: '8px 14px', textAlign: 'left',
                                fontSize: 11, fontWeight: 700, color: '#0078D4',
                                textTransform: 'uppercase', letterSpacing: '0.5px',
                                borderBottom: '1px solid #EDEBE9',
                            } }, "Pending Stage"),
                        React.createElement("th", { style: {
                                padding: '8px 14px', textAlign: 'left',
                                fontSize: 11, fontWeight: 700, color: '#0078D4',
                                textTransform: 'uppercase', letterSpacing: '0.5px',
                                borderBottom: '1px solid #EDEBE9',
                            } }, "Expected Role / Group"))),
                React.createElement("tbody", null, pendingStageRows.map(function (row, i) { return (React.createElement("tr", { key: i, style: { background: i % 2 === 0 ? '#FFFFFF' : '#F8FFFE' } },
                    React.createElement("td", { style: { padding: '10px 14px', color: '#201F1E', fontWeight: 600, borderBottom: '1px solid #EDEBE9' } },
                        React.createElement("span", { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                            React.createElement(Icon, { iconName: "Clock", style: { fontSize: 13, color: '#0078D4', flexShrink: 0 } }),
                            row.stage)),
                    React.createElement("td", { style: { padding: '10px 14px', color: '#605E5C', borderBottom: '1px solid #EDEBE9' } }, row.role))); }))))));
    };
    // ── Render: Action panel ───────────────────────────────────────────────────
    ReviewQueue.prototype._renderActionPanel = function (stageDef) {
        var _this = this;
        var _a, _b;
        var _c = this.state, actionRecord = _c.actionRecord, actionValues = _c.actionValues, poDate = _c.poDate, saving = _c.saving, saveError = _c.saveError, preApproval = _c.preApproval, postSourcing = _c.postSourcing, resolvedApprovers = _c.resolvedApprovers, loadingApprovers = _c.loadingApprovers;
        if (!actionRecord)
            return null;
        var ref = strVal(actionRecord, 'sct_sct_reference');
        var status = strVal(actionRecord, 'sct_request_status');
        var site = strVal(actionRecord, '_sct_site_value');
        var amt = this._estimatedValue(actionRecord);
        var isNew = this._isNewSupplier(actionRecord);
        var needsCFO = amt > CFO_THRESHOLD || isNew;
        var setVal = function (field) {
            return function (_, v) {
                return _this.setState(function (s) {
                    var _a;
                    return ({ actionValues: __assign(__assign({}, s.actionValues), (_a = {}, _a[field] = v !== null && v !== void 0 ? v : '', _a)) });
                });
            };
        };
        var tfStyles = { root: { marginBottom: 14 }, fieldGroup: { borderRadius: 6 } };
        var close = function () { return _this.setState({ actionRecord: null, saveError: null }); };
        var sectionHeader = function (label, color) { return (React.createElement("div", { style: { fontSize: 13, fontWeight: 700, color: color, marginTop: 16, marginBottom: 10 } }, label)); };
        return (React.createElement(Panel, { isOpen: true, onDismiss: close, type: PanelType.medium, headerText: ref ? "".concat(stageDef.title, " \u00B7 ").concat(ref) : stageDef.title, closeButtonAriaLabel: "Close", styles: { content: { paddingBottom: 88 } }, isFooterAtBottom: true, onRenderFooterContent: function () { return (React.createElement("div", { style: { display: 'flex', gap: 10, flexWrap: 'wrap', padding: '0 24px 24px' } },
                React.createElement(PrimaryButton, { text: saving ? 'Saving…' : stageDef.primaryAction, disabled: saving, onClick: function () { return void _this._takeAction(stageDef, stageDef.primaryStatus); }, styles: { root: { borderRadius: 6 } } }),
                stageDef.altAction && stageDef.altStatus !== undefined && (React.createElement(DefaultButton, { text: stageDef.altAction, disabled: saving, onClick: function () { return void _this._takeAction(stageDef, stageDef.altStatus); }, styles: { root: {
                            borderRadius: 6,
                            color: stageDef.altAction === 'Reject' ? '#A4262C' : '#905400',
                            borderColor: stageDef.altAction === 'Reject' ? '#A4262C' : '#C7A333',
                        } } })),
                React.createElement(DefaultButton, { text: "Cancel", onClick: close, disabled: saving, styles: { root: { borderRadius: 6 } } }))); } },
            saveError && (React.createElement(MessageBar, { messageBarType: MessageBarType.error, style: { marginBottom: 14 } }, saveError)),
            React.createElement("div", { style: { background: '#F8F8F7', borderRadius: 8, padding: '12px 14px', marginBottom: 16 } },
                React.createElement("div", { style: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px 16px' } }, [
                    ['Reference', ref],
                    ['Status', status],
                    ['Site', site],
                    ['Estimated Value', isNaN(amt) ? '—' : formatCurrency(amt)],
                    ['Requester', strVal(actionRecord, '_sct_requester_value')],
                    ['Purchase Type', strVal(actionRecord, 'sct_purchase_type')],
                    ['Urgency', strVal(actionRecord, 'sct_urgency')],
                    ['New Supplier', isNew ? 'Yes ⚑' : 'No'],
                ].map(function (_a) {
                    var label = _a[0], value = _a[1];
                    return (React.createElement("div", { key: label, style: { marginBottom: 6 } },
                        React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 2 } }, label),
                        React.createElement("div", { style: { fontSize: 13, color: '#201F1E', fontWeight: 500 } }, value || '—')));
                })),
                strVal(actionRecord, 'sct_description') && (React.createElement("div", { style: { marginTop: 10, paddingTop: 10, borderTop: '1px solid #EDEBE9' } },
                    React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', marginBottom: 4 } }, "Description"),
                    React.createElement("div", { style: { fontSize: 12, color: '#605E5C', lineHeight: '1.5', maxHeight: 72, overflowY: 'auto' } }, strVal(actionRecord, 'sct_description'))))),
            (stageDef.id === 'triage') && preApproval && (React.createElement(React.Fragment, null,
                sectionHeader('Pre-Approval Requirements', stageDef.color),
                React.createElement(MessageBar, { messageBarType: preApproval.levels.length > 1 ? MessageBarType.warning : MessageBarType.info, styles: { root: { borderRadius: 6, marginBottom: 10 } } },
                    React.createElement("strong", null,
                        "Required: ",
                        preApproval.levels.join(' + ')),
                    " \u2014 ",
                    preApproval.description))),
            stageDef.id === 'triage' && (React.createElement(React.Fragment, null,
                sectionHeader('Triage Assignment', stageDef.color),
                React.createElement(TextField, { label: "Cost Centre", value: (_a = actionValues['sct_cost_center']) !== null && _a !== void 0 ? _a : '', onChange: setVal('sct_cost_center'), placeholder: "e.g. CC-1042", styles: tfStyles }))),
            stageDef.id === 'sourcing' && (React.createElement(React.Fragment, null,
                sectionHeader('PO Details', stageDef.color),
                React.createElement(TextField, { label: "D365 PO Number", value: (_b = actionValues['sct_po_number']) !== null && _b !== void 0 ? _b : '', onChange: setVal('sct_po_number'), placeholder: "e.g. PO-2026-00123", styles: tfStyles }),
                React.createElement(DatePicker, { label: "PO Issue Date", value: poDate, onSelectDate: function (d) { return _this.setState({ poDate: d !== null && d !== void 0 ? d : undefined }); }, firstDayOfWeek: DayOfWeek.Monday, formatDate: function (d) { return d ? dateToEdm(d) : ''; }, placeholder: "Select PO issue date\u2026", styles: { root: { marginBottom: 14 } }, allowTextInput: true }),
                postSourcing && (React.createElement(MessageBar, { messageBarType: postSourcing.needsFPAReview ? MessageBarType.warning : MessageBarType.info, styles: { root: { borderRadius: 6, marginBottom: 10 } } },
                    React.createElement("strong", null, postSourcing.needsFPAReview ? 'FP&A + CFO approval required' : 'Direct PO authority'),
                    ' — ',
                    postSourcing.description)),
                (postSourcing === null || postSourcing === void 0 ? void 0 : postSourcing.needsLegalReview) && (React.createElement(MessageBar, { messageBarType: MessageBarType.severeWarning, styles: { root: { borderRadius: 6 } } }, "\u2696 Legal Review required \u2014 value exceeds $50,000.")))),
            stageDef.id === 'fpa' && (React.createElement(React.Fragment, null,
                sectionHeader('FP&A Decision', stageDef.color),
                needsCFO
                    ? React.createElement(MessageBar, { messageBarType: MessageBarType.warning, styles: { root: { borderRadius: 6, marginBottom: 10 } } },
                        "CFO approval is required \u2014 ",
                        isNaN(amt) ? '' : formatCurrency(amt),
                        isNew ? ' / new supplier' : '',
                        ". Use ",
                        React.createElement("strong", null, "Escalate to CFO"),
                        " unless you hold delegated authority.")
                    : React.createElement(MessageBar, { messageBarType: MessageBarType.info, styles: { root: { borderRadius: 6, marginBottom: 10 } } }, "Within FP&A direct approval authority. You may approve or escalate."))),
            stageDef.id === 'cfo' && (React.createElement(React.Fragment, null,
                sectionHeader('CFO Decision', stageDef.color),
                React.createElement(MessageBar, { messageBarType: MessageBarType.warning, styles: { root: { borderRadius: 6, marginBottom: 10 } } }, "Approving sets status to \"Approved\". Rejecting sets status to \"Rejected\"."))),
            stageDef.id === 'receiving' && (React.createElement(React.Fragment, null,
                sectionHeader('Goods/Services Receipt', stageDef.color),
                React.createElement(MessageBar, { messageBarType: MessageBarType.info, styles: { root: { borderRadius: 6 } } }, "Confirming receipt closes this request in the Spend Control Tower."))),
            this._renderPendingTracker(),
            (stageDef.id === 'triage' || stageDef.id === 'fpa' || stageDef.id === 'cfo') && (React.createElement(React.Fragment, null,
                sectionHeader('Resolved Approvers', '#605E5C'),
                loadingApprovers && (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8, padding: '8px 0' } },
                    React.createElement(Spinner, { size: SpinnerSize.small }),
                    React.createElement(Text, { variant: "small", style: { color: '#605E5C' } }, "Resolving approvers from sct_role\u2026"))),
                !loadingApprovers && resolvedApprovers && (React.createElement(React.Fragment, null,
                    resolvedApprovers.approvers.length === 0 && (React.createElement(Text, { variant: "small", style: { color: '#A19F9D', display: 'block', marginBottom: 8 } }, "No approvers resolved \u2014 verify sct_role records for this site in Dataverse.")),
                    resolvedApprovers.approvers.map(function (a) { return (React.createElement(ApproverCard, { key: a.level + a.systemUserId, level: a.level, name: a.name, isDelegate: a.isDelegate, wasEscalated: a.wasEscalated })); }),
                    resolvedApprovers.escalations.map(function (e, i) { return (React.createElement(MessageBar, { key: i, messageBarType: MessageBarType.warning, styles: { root: { borderRadius: 6, marginBottom: 6 } } }, e)); }),
                    resolvedApprovers.warnings.map(function (w, i) { return (React.createElement(MessageBar, { key: i, messageBarType: MessageBarType.info, styles: { root: { borderRadius: 6, marginBottom: 6 } } }, w)); })))))));
    };
    // ── Main render ───────────────────────────────────────────────────────────
    ReviewQueue.prototype.render = function () {
        var _this = this;
        var _a = this.state, activeStage = _a.activeStage, loading = _a.loading, error = _a.error, saveSuccess = _a.saveSuccess;
        var role = (this.props.activeRole || 'manager').toLowerCase().trim();
        var isApproverRole = (role === 'manager' || role.indexOf('manager') !== -1 ||
            role.indexOf('tier') !== -1 || role.indexOf('approver') !== -1);
        var stageDef = STAGES.find(function (s) { return s.id === activeStage; });
        var queueTitle = (role === 'manager' || role.indexOf('manager') !== -1) ? 'Manager Review Queue'
            : (role === 'approver_tier1' || role.indexOf('tier 1') !== -1 || role.indexOf('tier1') !== -1) ? 'Tier 1 Review Queue'
                : (role === 'approver_tier2' || role.indexOf('tier 2') !== -1 || role.indexOf('tier2') !== -1) ? 'Tier 2 Review Queue'
                    : (role === 'approver_tier3' || role.indexOf('tier 3') !== -1 || role.indexOf('tier3') !== -1) ? 'Tier 3 Review Queue'
                        : 'Spend Control Tower';
        var queueSubtitle = (role === 'manager' || role.indexOf('manager') !== -1)
            ? 'Requests assigned to you as line manager, pending your review'
            : (role === 'approver_tier1' || role.indexOf('tier 1') !== -1 || role.indexOf('tier1') !== -1)
                ? 'Requests awaiting Tier 1 approval — status: In Review'
                : (role === 'approver_tier2' || role.indexOf('tier 2') !== -1 || role.indexOf('tier2') !== -1)
                    ? 'Requests awaiting Tier 2 approval — status: In Review'
                    : (role === 'approver_tier3' || role.indexOf('tier 3') !== -1 || role.indexOf('tier3') !== -1)
                        ? 'Requests awaiting Tier 3 / Final approval — status: In Review'
                        : 'Stage workbenches · Schedule A approval routing engine';
        var headerIcon = role === 'manager' ? 'People' : isApproverRole ? 'WorkFlow' : 'Checklist';
        return (React.createElement("div", null,
            React.createElement("div", { className: styles.card, style: { marginBottom: 20 } },
                React.createElement("div", { className: styles.cardBody, style: { display: 'flex', alignItems: 'center', gap: 16 } },
                    React.createElement("div", { style: {
                            width: 56, height: 56, borderRadius: 12,
                            background: isApproverRole ? 'linear-gradient(135deg, #0078D4, #1BA1E2)' : 'linear-gradient(135deg, #8764B8, #0078D4)',
                            display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0
                        } },
                        React.createElement(Icon, { iconName: headerIcon, style: { fontSize: 26, color: '#ffffff' } })),
                    React.createElement("div", { style: { flex: 1 } },
                        React.createElement(Text, { variant: "xLarge", style: { color: '#201F1E', fontWeight: 700, display: 'block' } }, queueTitle),
                        React.createElement(Text, { variant: "medium", style: { color: '#605E5C' } }, queueSubtitle)),
                    React.createElement(DefaultButton, { text: "Refresh", iconProps: { iconName: 'Refresh' }, onClick: function () { return void _this._load(); }, disabled: loading, styles: { root: { borderRadius: 6, height: 32 } } }))),
            saveSuccess && (React.createElement(MessageBar, { messageBarType: MessageBarType.success, isMultiline: false, onDismiss: function () { return _this.setState({ saveSuccess: null }); }, styles: { root: { borderRadius: 6, marginBottom: 12 } } }, saveSuccess)),
            error && (React.createElement(ErrorBanner, { message: error, onRetry: function () { return void _this._load(); }, onDismiss: function () { return _this.setState({ error: null }); } })),
            React.createElement("div", { className: styles.card },
                React.createElement("div", { className: styles.cardBody }, loading
                    ? React.createElement(LoadingSpinner, { label: "Loading queue from Dataverse\u2026" })
                    : isApproverRole
                        ? this._renderManagerQueue()
                        : (React.createElement(React.Fragment, null,
                            this._renderStageTabs(),
                            this._renderWorkbench(stageDef))))),
            !isApproverRole && this._renderActionPanel(stageDef)));
    };
    return ReviewQueue;
}(React.Component));
export default ReviewQueue;
//# sourceMappingURL=ReviewQueue.js.map