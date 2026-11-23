"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var Icon_1 = require("@fluentui/react/lib/Icon");
var Text_1 = require("@fluentui/react/lib/Text");
var SearchBox_1 = require("@fluentui/react/lib/SearchBox");
var Panel_1 = require("@fluentui/react/lib/Panel");
var Dialog_1 = require("@fluentui/react/lib/Dialog");
var TextField_1 = require("@fluentui/react/lib/TextField");
var Spinner_1 = require("@fluentui/react/lib/Spinner");
var MessageBar_1 = require("@fluentui/react/lib/MessageBar");
var Button_1 = require("@fluentui/react/lib/Button");
var GraphMailService_1 = require("../../services/GraphMailService");
var DataverseService_1 = require("../../services/DataverseService");
var common_1 = require("../common");
var AppShell_module_scss_1 = tslib_1.__importDefault(require("../shell/AppShell.module.scss"));
var dashboardUtils_1 = require("../dashboard/dashboardUtils");
var LogsContent_1 = tslib_1.__importDefault(require("../history/LogsContent"));
// ── Confirmed status constants (Power Apps sct_Request_Status) ─────────────────
var MY_STATUS_FILTERS = [
    { label: 'All', color: '#605E5C', bg: '#F3F2F1' },
    { label: 'Draft', color: '#605E5C', bg: '#EDEBE9' },
    { label: 'Submitted', color: '#0078D4', bg: '#DEECF9' },
    { label: 'In Review', color: '#FF8C00', bg: '#FFF4CE' },
    { label: 'PO Issued', color: '#00B294', bg: '#E1FAF4' },
    { label: 'Rejected', color: '#A4262C', bg: '#FDE7E9' },
    { label: 'Returned', color: '#8764B8', bg: '#F4F0FA' },
    { label: 'Closed', color: '#107C10', bg: '#DFF6DD' },
];
var STATUS_COLOR = {
    'Draft': '#605E5C',
    'Submitted': '#0078D4',
    'In Review': '#FF8C00',
    'PO Issued': '#00B294',
    'Rejected': '#A4262C',
    'Returned': '#8764B8',
    'Closed': '#107C10',
};
var STATUS_BG = {
    'Draft': '#EDEBE9',
    'Submitted': '#DEECF9',
    'In Review': '#FFF4CE',
    'PO Issued': '#E1FAF4',
    'Rejected': '#FDE7E9',
    'Returned': '#F4F0FA',
    'Closed': '#DFF6DD',
};
// Stage progress for visual indicator — maps status label to step index (0-based)
var STAGE_STEP = {
    'Draft': 0, 'Submitted': 1, 'In Review': 2, 'PO Issued': 3, 'Closed': 4,
    'Returned': -1, 'Rejected': -1,
};
var STAGE_LABELS = ['Draft', 'Submitted', 'In Review', 'PO Issued', 'Closed'];
var MyRequests = /** @class */ (function (_super) {
    tslib_1.__extends(MyRequests, _super);
    function MyRequests(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            records: [], loading: true, error: null,
            systemUserId: null, activeFilter: 'All', searchText: '', viewRecord: null,
            resubmitRecord: null, resubmitDescription: '', resubmitSaving: false,
            resubmitSuccess: null, resubmitError: null,
        };
        return _this;
    }
    MyRequests.prototype.componentDidMount = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._loadData()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    MyRequests.prototype.componentDidUpdate = function (prevProps) {
        var _this = this;
        if (prevProps.activeRole !== this.props.activeRole) {
            this.setState({ records: [], loading: true }, function () {
                void _this._loadData();
            });
        }
    };
    MyRequests.prototype._loadData = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var ctx, aadId, sysUserId, user, user, records, e_1;
            var _a, _b, _c, _d, _e, _f, _g, _h, _j;
            return tslib_1.__generator(this, function (_k) {
                switch (_k.label) {
                    case 0:
                        this.setState({ loading: true, error: null });
                        _k.label = 1;
                    case 1:
                        _k.trys.push([1, 7, , 8]);
                        ctx = this.props.context;
                        aadId = (_e = (_d = (_c = (_b = (_a = ctx.pageContext) === null || _a === void 0 ? void 0 : _a.aadInfo) === null || _b === void 0 ? void 0 : _b.userId) === null || _c === void 0 ? void 0 : _c.toString) === null || _d === void 0 ? void 0 : _d.call(_c)) !== null && _e !== void 0 ? _e : '';
                        sysUserId = null;
                        if (!aadId) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.props.dataverseService.getCurrentSystemUser(aadId)];
                    case 2:
                        user = _k.sent();
                        sysUserId = (_f = user === null || user === void 0 ? void 0 : user.systemuserid) !== null && _f !== void 0 ? _f : null;
                        _k.label = 3;
                    case 3:
                        if (!(!sysUserId && ((_h = (_g = ctx.pageContext) === null || _g === void 0 ? void 0 : _g.user) === null || _h === void 0 ? void 0 : _h.email))) return [3 /*break*/, 5];
                        return [4 /*yield*/, this.props.dataverseService.getSystemUserByEmail(ctx.pageContext.user.email)];
                    case 4:
                        user = _k.sent();
                        sysUserId = (_j = user === null || user === void 0 ? void 0 : user.systemuserid) !== null && _j !== void 0 ? _j : null;
                        _k.label = 5;
                    case 5:
                        if (!sysUserId) {
                            this.setState({ loading: false, error: 'Could not resolve your Dataverse user account. Contact your administrator.' });
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this.props.dataverseService.getMyRequests(sysUserId)];
                    case 6:
                        records = _k.sent();
                        this.setState({ records: records, systemUserId: sysUserId, loading: false });
                        return [3 /*break*/, 8];
                    case 7:
                        e_1 = _k.sent();
                        this.setState({ loading: false, error: e_1 instanceof Error ? e_1.message : String(e_1) });
                        return [3 /*break*/, 8];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    MyRequests.prototype._getStatusLabel = function (record) {
        return (0, dashboardUtils_1.getRecordStatusLabel)(record);
    };
    // sct_submitted_date confirmed from Power Apps (schema: sct_Submitted_Date, logical: sct_submitted_date)
    MyRequests.prototype._getSubmittedDate = function (record) {
        var submitted = record['sct_submitted_date'];
        if (submitted)
            return (0, dashboardUtils_1.formatDate)(String(submitted));
        var created = record['createdon'];
        if (created)
            return (0, dashboardUtils_1.formatDate)(String(created));
        return '—';
    };
    MyRequests.prototype._getFiltered = function () {
        var _this = this;
        var _a = this.state, records = _a.records, activeFilter = _a.activeFilter, searchText = _a.searchText;
        var q = searchText.trim().toLowerCase();
        return records.filter(function (r) {
            var statusLabel = _this._getStatusLabel(r);
            if (activeFilter !== 'All' && statusLabel !== activeFilter)
                return false;
            if (q) {
                var ref = (0, DataverseService_1.strVal)(r, 'sct_sct_reference').toLowerCase();
                var desc = (0, DataverseService_1.strVal)(r, 'sct_description').toLowerCase();
                var site = (0, DataverseService_1.strVal)(r, '_sct_site_value').toLowerCase();
                var type = (0, DataverseService_1.strVal)(r, 'sct_purchase_type').toLowerCase();
                if (!ref.includes(q) && !desc.includes(q) && !site.includes(q) && !type.includes(q))
                    return false;
            }
            return true;
        });
    };
    MyRequests.prototype._countFor = function (label) {
        var _this = this;
        if (label === 'All')
            return this.state.records.length;
        return this.state.records.filter(function (r) { return _this._getStatusLabel(r) === label; }).length;
    };
    MyRequests.prototype._handleResubmit = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a, resubmitRecord, resubmitDescription, reqId, payload, graphClient, manager, title, amt, reqName, _b, err_1, errMsg;
            var _c, _d, _e;
            return tslib_1.__generator(this, function (_f) {
                switch (_f.label) {
                    case 0:
                        _a = this.state, resubmitRecord = _a.resubmitRecord, resubmitDescription = _a.resubmitDescription;
                        if (!resubmitRecord)
                            return [2 /*return*/];
                        reqId = String((_d = (_c = resubmitRecord['sct_requestid']) !== null && _c !== void 0 ? _c : resubmitRecord['id']) !== null && _d !== void 0 ? _d : '');
                        if (!reqId)
                            return [2 /*return*/];
                        this.setState({ resubmitSaving: true, resubmitError: null });
                        _f.label = 1;
                    case 1:
                        _f.trys.push([1, 12, , 13]);
                        payload = {
                            sct_request_status: 854000002, // Submitted (Confirmed Dataverse choice value)
                            cr2a4_pendingtier: 'Manager',
                            sct_description: resubmitDescription.trim() || (0, DataverseService_1.strVal)(resubmitRecord, 'sct_description'),
                        };
                        return [4 /*yield*/, this.props.dataverseService.updateRequest(reqId, payload)];
                    case 2:
                        _f.sent();
                        // Audit log the resubmission event
                        return [4 /*yield*/, this.props.dataverseService.createAuditLog(tslib_1.__assign({ sct_action: 1, sct_description: 'Request resubmitted for Line Manager approval by requester', sct_oldvalue: 'Status: Returned, Pending Tier: None', sct_newvalue: 'Status: Submitted, Pending Tier: Manager', 'sct_requestid@odata.bind': "/sct_requests(".concat(reqId, ")") }, (this.state.systemUserId ? { 'sct_actionby@odata.bind': "/systemusers(".concat(this.state.systemUserId, ")") } : {}))).catch(function () { })];
                    case 3:
                        // Audit log the resubmission event
                        _f.sent();
                        _f.label = 4;
                    case 4:
                        _f.trys.push([4, 9, , 10]);
                        return [4 /*yield*/, this.props.context.msGraphClientFactory.getClient('3')];
                    case 5:
                        graphClient = _f.sent();
                        return [4 /*yield*/, (0, GraphMailService_1.getManager)(graphClient)];
                    case 6:
                        manager = _f.sent();
                        if (!(manager && manager.mail)) return [3 /*break*/, 8];
                        title = (0, DataverseService_1.strVal)(resubmitRecord, 'sct_request_title') || (0, DataverseService_1.strVal)(resubmitRecord, 'sct_sct_reference');
                        amt = String((_e = resubmitRecord['sct_estimated_total_value']) !== null && _e !== void 0 ? _e : '');
                        reqName = this.props.context.pageContext.user.displayName;
                        return [4 /*yield*/, (0, GraphMailService_1.sendManagerApprovalEmail)(graphClient, manager, reqName, title, amt, this.props.context.pageContext.web.absoluteUrl)];
                    case 7:
                        _f.sent();
                        _f.label = 8;
                    case 8: return [3 /*break*/, 10];
                    case 9:
                        _b = _f.sent();
                        return [3 /*break*/, 10];
                    case 10:
                        this.setState({
                            resubmitSaving: false,
                            resubmitRecord: null,
                            viewRecord: null,
                            resubmitSuccess: 'Request resubmitted successfully! It is now pending Line Manager review.',
                        });
                        return [4 /*yield*/, this._loadData()];
                    case 11:
                        _f.sent();
                        return [3 /*break*/, 13];
                    case 12:
                        err_1 = _f.sent();
                        errMsg = err_1 instanceof Error ? err_1.message : 'Resubmission failed. Please try again.';
                        this.setState({ resubmitSaving: false, resubmitError: errMsg });
                        return [3 /*break*/, 13];
                    case 13: return [2 /*return*/];
                }
            });
        });
    };
    MyRequests.prototype.render = function () {
        var _this = this;
        var _a = this.state, loading = _a.loading, error = _a.error, activeFilter = _a.activeFilter, searchText = _a.searchText, viewRecord = _a.viewRecord;
        var filtered = this._getFiltered();
        if (loading)
            return React.createElement(common_1.LoadingSpinner, { label: "Loading your requests\u2026" });
        return (React.createElement("div", null,
            this.state.resubmitSuccess && (React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.success, onDismiss: function () { return _this.setState({ resubmitSuccess: null }); }, styles: { root: { marginBottom: 16, borderRadius: 6 } } }, this.state.resubmitSuccess)),
            error && (React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.error, style: { marginBottom: 16, borderRadius: 6 } },
                error,
                React.createElement("button", { style: { marginLeft: 12, background: 'none', border: 'none', color: '#A4262C', cursor: 'pointer', textDecoration: 'underline', fontFamily: 'inherit', fontSize: 13 }, onClick: function () { return void _this._loadData(); } }, "Retry"))),
            React.createElement("div", { className: AppShell_module_scss_1.default.card, style: { marginBottom: 16 } },
                React.createElement("div", { className: AppShell_module_scss_1.default.cardBody, style: { paddingTop: 14, paddingBottom: 14 } },
                    React.createElement("div", { style: { display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' } },
                        React.createElement(Text_1.Text, { variant: "small", style: { color: '#605E5C', fontWeight: 600, marginRight: 4 } }, "Filter:"),
                        MY_STATUS_FILTERS.map(function (f) {
                            var count = _this._countFor(f.label);
                            var active = activeFilter === f.label;
                            return (React.createElement("button", { key: f.label, onClick: function () { return _this.setState({ activeFilter: f.label }); }, style: {
                                    background: active ? f.color : f.bg,
                                    color: active ? '#fff' : f.color,
                                    border: active ? "2px solid ".concat(f.color) : '2px solid transparent',
                                    borderRadius: 16, padding: '4px 12px',
                                    fontSize: 12, fontWeight: 600, cursor: 'pointer',
                                    fontFamily: 'inherit', display: 'flex', alignItems: 'center', gap: 5,
                                    transition: 'all 0.15s',
                                } },
                                f.label,
                                React.createElement("span", { style: {
                                        background: active ? 'rgba(255,255,255,0.25)' : 'rgba(0,0,0,0.10)',
                                        borderRadius: 10, padding: '0 5px', fontSize: 11,
                                    } }, count)));
                        }),
                        React.createElement("div", { style: { flex: 1, minWidth: 200 } },
                            React.createElement(SearchBox_1.SearchBox, { placeholder: "Search reference, description, site\u2026", value: searchText, onChange: function (_, v) { return _this.setState({ searchText: v !== null && v !== void 0 ? v : '' }); }, styles: { root: { borderRadius: 6 } } }))))),
            React.createElement("div", { className: AppShell_module_scss_1.default.card },
                React.createElement("div", { className: AppShell_module_scss_1.default.cardHeader },
                    React.createElement("div", { className: AppShell_module_scss_1.default.cardTitle },
                        React.createElement("div", { className: AppShell_module_scss_1.default.cardTitleIcon, style: { background: '#DFF6DD' }, "aria-hidden": "true" },
                            React.createElement(Icon_1.Icon, { iconName: "DocumentSearch", style: { color: '#107C10', fontSize: 16 } })),
                        "My Requests",
                        React.createElement("span", { style: { marginLeft: 8, fontSize: 12, color: '#A19F9D', fontWeight: 400 } },
                            filtered.length,
                            " of ",
                            this.state.records.length)),
                    React.createElement("button", { onClick: function () { return void _this._loadData(); }, style: { background: 'none', border: 'none', cursor: 'pointer', color: '#0078D4', display: 'flex', alignItems: 'center', gap: 4, fontSize: 13, fontFamily: 'inherit' }, title: "Refresh" },
                        React.createElement(Icon_1.Icon, { iconName: "Refresh", style: { fontSize: 14 } }),
                        " Refresh")),
                React.createElement("div", { className: AppShell_module_scss_1.default.cardBody }, filtered.length === 0 ? (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '48px 0', gap: 12 } },
                    React.createElement(Icon_1.Icon, { iconName: "DocumentReply", style: { fontSize: 48, color: '#C8C6C4' }, "aria-hidden": "true" }),
                    React.createElement(Text_1.Text, { variant: "mediumPlus", style: { color: '#A19F9D', fontWeight: 500, textAlign: 'center' } }, this.state.records.length === 0 ? 'You have no requests yet' : 'No requests match your filter'),
                    React.createElement(Text_1.Text, { variant: "small", style: { color: '#C8C6C4', textAlign: 'center' } }, this.state.records.length === 0 ? 'Use the New Request tab to submit your first request.' : 'Try clearing the search or selecting All.'))) : (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 12 } }, filtered.map(function (r) { return _this._renderCard(r); }))))),
            viewRecord && this._renderDetailPanel(viewRecord),
            React.createElement(Dialog_1.Dialog, { hidden: !this.state.resubmitRecord, onDismiss: function () { if (!_this.state.resubmitSaving)
                    _this.setState({ resubmitRecord: null }); }, dialogContentProps: {
                    type: Dialog_1.DialogType.normal,
                    title: "Revise & Resubmit \u2014 ".concat(this.state.resubmitRecord ? (0, DataverseService_1.strVal)(this.state.resubmitRecord, 'sct_sct_reference') : ''),
                    subText: 'Update the request description/justification based on approver feedback. Once resubmitted, the request will restart from the Line Manager approval stage.',
                }, modalProps: { isBlocking: true } },
                this.state.resubmitError && (React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.error, styles: { root: { marginBottom: 12, borderRadius: 6 } } }, this.state.resubmitError)),
                React.createElement(TextField_1.TextField, { label: "Updated Description / Justification", multiline: true, rows: 4, value: this.state.resubmitDescription, onChange: function (_, val) { return _this.setState({ resubmitDescription: val || '' }); }, placeholder: "Add clarifications or changes requested by approvers...", disabled: this.state.resubmitSaving, required: true }),
                React.createElement(Dialog_1.DialogFooter, null,
                    React.createElement(Button_1.PrimaryButton, { text: this.state.resubmitSaving ? 'Resubmitting…' : 'Resubmit Request', iconProps: this.state.resubmitSaving ? undefined : { iconName: 'Send' }, onClick: function () { return void _this._handleResubmit(); }, disabled: this.state.resubmitSaving || !this.state.resubmitDescription.trim(), styles: { root: { borderRadius: 6, background: '#107C10', borderColor: '#107C10' } } }, this.state.resubmitSaving && React.createElement(Spinner_1.Spinner, { size: Spinner_1.SpinnerSize.xSmall, style: { marginRight: 6 } })),
                    React.createElement(Button_1.DefaultButton, { text: "Cancel", onClick: function () { return _this.setState({ resubmitRecord: null }); }, disabled: this.state.resubmitSaving, styles: { root: { borderRadius: 6 } } })))));
    };
    MyRequests.prototype._renderCard = function (r) {
        var _this = this;
        var _a, _b, _c, _d, _e, _f;
        var ref = (0, DataverseService_1.strVal)(r, 'sct_sct_reference') || '(no reference)';
        var statusLabel = this._getStatusLabel(r);
        var site = (0, DataverseService_1.strVal)(r, '_sct_site_value') || (0, DataverseService_1.strVal)(r, 'sct_site');
        var desc = (0, DataverseService_1.strVal)(r, 'sct_description');
        var amtRaw = parseFloat(String((_a = r['sct_estimated_total_value']) !== null && _a !== void 0 ? _a : ''));
        var amt = !isNaN(amtRaw) ? (0, dashboardUtils_1.formatCurrency)(amtRaw) : '—';
        var type = (0, DataverseService_1.strVal)(r, 'sct_purchase_type') || (0, DataverseService_1.strVal)(r, 'sct_requesttype');
        var submittedOn = this._getSubmittedDate(r);
        var poNum = (0, DataverseService_1.strVal)(r, 'sct_po_number');
        var isRejected = statusLabel === 'Rejected';
        var isReturned = statusLabel === 'Returned';
        var stepIndex = (_b = STAGE_STEP[statusLabel]) !== null && _b !== void 0 ? _b : -1;
        var pt = (0, DataverseService_1.strVal)(r, 'cr2a4_pendingtier');
        var displayStatus = (statusLabel === 'In Review' && pt) ? "".concat(statusLabel, " (").concat(pt, ")") : statusLabel;
        var pill = (React.createElement("span", { style: {
                background: (_c = STATUS_BG[statusLabel]) !== null && _c !== void 0 ? _c : '#EDEBE9',
                color: (_d = STATUS_COLOR[statusLabel]) !== null && _d !== void 0 ? _d : '#605E5C',
                borderRadius: 12, padding: '2px 10px',
                fontSize: 11, fontWeight: 700, whiteSpace: 'nowrap',
            } }, displayStatus));
        return (React.createElement("div", { key: String((_e = r['sct_requestid']) !== null && _e !== void 0 ? _e : ref), onClick: function () { return _this.setState({ viewRecord: r }); }, style: {
                border: "1px solid ".concat(isRejected ? '#F3CFCF' : isReturned ? '#D9D0E8' : '#EDEBE9'),
                borderLeft: "4px solid ".concat((_f = STATUS_COLOR[statusLabel]) !== null && _f !== void 0 ? _f : '#EDEBE9'),
                borderRadius: 8, padding: '14px 16px', cursor: 'pointer',
                background: '#FAFAFA', transition: 'box-shadow 0.15s',
            }, onMouseEnter: function (e) { return (e.currentTarget.style.boxShadow = '0 2px 8px rgba(0,0,0,0.10)'); }, onMouseLeave: function (e) { return (e.currentTarget.style.boxShadow = 'none'); } },
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6, flexWrap: 'wrap' } },
                React.createElement(Text_1.Text, { variant: "medium", style: { fontWeight: 700, color: '#201F1E' } }, ref),
                pill,
                site && (React.createElement("span", { style: { fontSize: 12, color: '#605E5C', display: 'flex', alignItems: 'center', gap: 4 } },
                    React.createElement(Icon_1.Icon, { iconName: "MapPin", style: { fontSize: 11 } }),
                    site)),
                React.createElement("span", { style: { marginLeft: 'auto', fontSize: 11, color: '#A19F9D' } },
                    React.createElement(Icon_1.Icon, { iconName: "Clock", style: { fontSize: 11, marginRight: 4 } }),
                    "Submitted: ",
                    submittedOn)),
            desc && (React.createElement("div", { style: { fontSize: 13, color: '#605E5C', marginBottom: 8, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '90%' } }, desc)),
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap', fontSize: 12, color: '#605E5C' } },
                amt !== '—' && (React.createElement("span", { style: { display: 'flex', alignItems: 'center', gap: 4 } },
                    React.createElement(Icon_1.Icon, { iconName: "Money", style: { fontSize: 12, color: '#0078D4' } }),
                    React.createElement("strong", { style: { color: '#201F1E' } }, amt))),
                type && (React.createElement("span", { style: { display: 'flex', alignItems: 'center', gap: 4 } },
                    React.createElement(Icon_1.Icon, { iconName: "Tag", style: { fontSize: 11 } }),
                    type)),
                poNum && (React.createElement("span", { style: { display: 'flex', alignItems: 'center', gap: 4, color: '#00B294', fontWeight: 600 } },
                    React.createElement(Icon_1.Icon, { iconName: "CheckMark", style: { fontSize: 11 } }),
                    "PO: ",
                    poNum)),
                React.createElement("span", { style: { marginLeft: 'auto', color: '#0078D4', fontSize: 12, display: 'flex', alignItems: 'center', gap: 4 } },
                    "View details ",
                    React.createElement(Icon_1.Icon, { iconName: "ChevronRight", style: { fontSize: 10 } }))),
            stepIndex >= 0 && (React.createElement("div", { style: { marginTop: 10, display: 'flex', alignItems: 'center', gap: 4 } }, STAGE_LABELS.map(function (s, i) {
                var _a;
                return (React.createElement(React.Fragment, { key: s },
                    React.createElement("div", { style: {
                            flex: 1, height: 4, borderRadius: 2,
                            background: i <= stepIndex ? ((_a = STATUS_COLOR[statusLabel]) !== null && _a !== void 0 ? _a : '#0078D4') : '#EDEBE9',
                            transition: 'background 0.3s',
                        }, title: s })));
            }))),
            (isRejected || isReturned) && (React.createElement("div", { style: { marginTop: 8, fontSize: 12, color: STATUS_COLOR[statusLabel], fontWeight: 600, display: 'flex', alignItems: 'center', gap: 4 } },
                React.createElement(Icon_1.Icon, { iconName: isRejected ? 'StatusErrorFull' : 'ReturnToSession', style: { fontSize: 12 } }),
                isRejected ? 'This request was rejected.' : 'This request was returned — please revise and resubmit.'))));
    };
    MyRequests.prototype._cancelRequest = function (rec) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var id, err_2;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        id = String((_a = rec['sct_requestid']) !== null && _a !== void 0 ? _a : '');
                        if (!id)
                            return [2 /*return*/];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, this.props.dataverseService.updateRequest(id, {
                                sct_request_status: 854000007, // Closed (Confirmed Dataverse choice value)
                                cr2a4_pendingtier: null,
                            })];
                    case 2:
                        _b.sent();
                        return [4 /*yield*/, this.props.dataverseService.createAuditLog({
                                'sct_requestid@odata.bind': "/sct_requests(".concat(id, ")"),
                                sct_action: 4,
                                sct_description: 'Request cancelled and withdrawn by Requester',
                                sct_oldvalue: 'Status: Submitted',
                                sct_newvalue: 'Status: Closed, Pending Tier: None',
                            }).catch(function () { })];
                    case 3:
                        _b.sent();
                        this.setState({ viewRecord: null });
                        void this._loadData();
                        return [3 /*break*/, 5];
                    case 4:
                        err_2 = _b.sent();
                        this.setState({ error: err_2 instanceof Error ? err_2.message : 'Failed to withdraw request.' });
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    MyRequests.prototype._renderDetailPanel = function (r) {
        var _this = this;
        var _a, _b, _c, _d;
        var ref = (0, DataverseService_1.strVal)(r, 'sct_sct_reference') || '(no reference)';
        var statusLabel = this._getStatusLabel(r);
        var site = (0, DataverseService_1.strVal)(r, '_sct_site_value');
        var desc = (0, DataverseService_1.strVal)(r, 'sct_description');
        var amtRaw = parseFloat(String((_a = r['sct_estimated_total_value']) !== null && _a !== void 0 ? _a : ''));
        var amt = !isNaN(amtRaw) ? (0, dashboardUtils_1.formatCurrency)(amtRaw) : '—';
        var close = function () { return _this.setState({ viewRecord: null }); };
        var lbl = { fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 2 };
        var val = { fontSize: 13, color: '#201F1E', fontWeight: 500, marginBottom: 10 };
        var Row = function (_a) {
            var label = _a.label, value = _a.value;
            return (React.createElement("div", null,
                React.createElement("div", { style: lbl }, label),
                React.createElement("div", { style: val }, value || '—')));
        };
        var requestId = String((_b = r['sct_requestid']) !== null && _b !== void 0 ? _b : '');
        var canEdit = statusLabel === 'Draft';
        return (React.createElement(Panel_1.Panel, { isOpen: true, onDismiss: close, type: Panel_1.PanelType.custom, customWidth: "min(720px, 96vw)", headerText: "Request: ".concat(ref), closeButtonAriaLabel: "Close", isFooterAtBottom: true, onRenderFooterContent: function () { return (React.createElement("div", { style: { display: 'flex', gap: 10, padding: '0 24px 24px' } },
                statusLabel === 'Returned' ? (React.createElement(Button_1.PrimaryButton, { text: "Revise in Form", iconProps: { iconName: 'Edit' }, onClick: function () {
                        var _a, _b;
                        close();
                        (_b = (_a = _this.props).onEditRequest) === null || _b === void 0 ? void 0 : _b.call(_a, r);
                    }, styles: { root: { borderRadius: 6, background: '#0078D4', borderColor: '#005A9E' } } })) : (React.createElement(Button_1.PrimaryButton, { text: "Edit Draft", iconProps: { iconName: 'Edit' }, disabled: !canEdit, title: canEdit ? 'Edit this draft request' : "Records with status \"".concat(statusLabel, "\" cannot be edited"), onClick: function () {
                        var _a, _b;
                        close();
                        (_b = (_a = _this.props).onEditRequest) === null || _b === void 0 ? void 0 : _b.call(_a, r);
                    }, styles: { root: { borderRadius: 6 } } })),
                (statusLabel === 'Submitted' || statusLabel === 'Draft') && (React.createElement(Button_1.DefaultButton, { text: "Withdraw Request", iconProps: { iconName: 'Cancel' }, onClick: function () {
                        if (window.confirm("Are you sure you want to withdraw request ".concat(ref, "? It will be closed and removed from review."))) {
                            void _this._cancelRequest(r);
                        }
                    }, styles: { root: { borderRadius: 6, color: '#A4262C', borderColor: '#FDE7E9' } } })),
                React.createElement(Button_1.DefaultButton, { text: "Close", onClick: close, styles: { root: { borderRadius: 6 } } }))); }, styles: { content: { paddingBottom: 60 } } },
            React.createElement("div", { style: { marginBottom: 16 } },
                React.createElement("span", { style: { background: (_c = STATUS_BG[statusLabel]) !== null && _c !== void 0 ? _c : '#EDEBE9', color: (_d = STATUS_COLOR[statusLabel]) !== null && _d !== void 0 ? _d : '#605E5C', borderRadius: 12, padding: '4px 14px', fontSize: 13, fontWeight: 700 } }, statusLabel)),
            React.createElement("div", { style: { background: '#F8F8F7', borderRadius: 8, padding: '12px 14px', marginBottom: 20 } },
                React.createElement("div", { style: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '4px 16px' } },
                    React.createElement(Row, { label: "Reference", value: ref }),
                    React.createElement(Row, { label: "Status", value: statusLabel }),
                    React.createElement(Row, { label: "Site", value: site }),
                    React.createElement(Row, { label: "Purchase Type", value: (0, DataverseService_1.strVal)(r, 'sct_purchase_type') || (0, DataverseService_1.strVal)(r, 'sct_requesttype') }),
                    React.createElement(Row, { label: "Estimated Value", value: amt }),
                    React.createElement(Row, { label: "Urgency", value: (0, DataverseService_1.strVal)(r, 'sct_urgency') }),
                    React.createElement(Row, { label: "Submitted Date", value: this._getSubmittedDate(r) }),
                    React.createElement(Row, { label: "Date Needed By", value: (0, dashboardUtils_1.formatDate)((0, DataverseService_1.strVal)(r, 'sct_date_needed_by')) }),
                    React.createElement(Row, { label: "PO Number", value: (0, DataverseService_1.strVal)(r, 'sct_po_number') }),
                    React.createElement(Row, { label: "PO Issue Date", value: (0, dashboardUtils_1.formatDate)((0, DataverseService_1.strVal)(r, 'sct_po_issue_date')) }))),
            desc && (React.createElement("div", { style: { marginBottom: 16 } },
                React.createElement("div", { style: lbl }, "Description"),
                React.createElement("div", { style: { fontSize: 13, color: '#201F1E', lineHeight: '1.6' } }, desc))),
            (0, DataverseService_1.strVal)(r, 'sct_urgency_reason') && (React.createElement("div", { style: { marginBottom: 16 } },
                React.createElement("div", { style: lbl }, "Urgency Reason"),
                React.createElement("div", { style: { fontSize: 13, color: '#201F1E', lineHeight: '1.6' } }, (0, DataverseService_1.strVal)(r, 'sct_urgency_reason')))),
            React.createElement("div", { style: { borderTop: '1px solid #EDEBE9', paddingTop: 14 } },
                React.createElement("div", { style: lbl }, "Flags"),
                React.createElement("div", { style: { display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 6 } }, [
                    ['sct_is_new_supplier', 'New Supplier', '#905400', '#FFF4CE'],
                    ['sct_is_retrospective', 'Retrospective', '#605E5C', '#EDEBE9'],
                    ['sct_is_confidential', 'Confidential', '#A4262C', '#FDE7E9'],
                    ['sct_tech_or_data_involved', 'Tech/Data', '#0078D4', '#DEECF9'],
                    ['sct_possible_capital_purchase', 'Capital Purchase', '#8764B8', '#F4F0FA'],
                    ['sct_cost_recharged_to_customer', 'Recharged to Customer', '#00B294', '#E1FAF4'],
                ].map(function (_a) {
                    var field = _a[0], label = _a[1], color = _a[2], bg = _a[3];
                    var v = r[field];
                    var on = v === true || v === 'Yes' || v === '1';
                    if (!on)
                        return null;
                    return (React.createElement("span", { key: field, style: { background: bg, color: color, borderRadius: 12, padding: '2px 10px', fontSize: 11, fontWeight: 700 } }, label));
                }))),
            requestId && (React.createElement("div", { style: { marginTop: 24 } },
                React.createElement("div", { style: {
                        display: 'flex', alignItems: 'center', gap: 10,
                        marginBottom: 14, paddingBottom: 10, borderBottom: '2px solid #EDEBE9',
                    } },
                    React.createElement("div", { style: {
                            width: 28, height: 28, borderRadius: 6, flexShrink: 0,
                            background: 'linear-gradient(135deg, #107C10, #8764B8)',
                            display: 'flex', alignItems: 'center', justifyContent: 'center',
                        } },
                        React.createElement(Icon_1.Icon, { iconName: "History", style: { fontSize: 14, color: '#fff' } })),
                    React.createElement("span", { style: { fontSize: 15, fontWeight: 700, color: '#201F1E' } }, "Request History")),
                React.createElement(LogsContent_1.default, { requestId: requestId, service: this.props.dataverseService })))));
    };
    return MyRequests;
}(React.Component));
exports.default = MyRequests;
//# sourceMappingURL=MyRequests.js.map