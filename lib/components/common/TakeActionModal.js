import { __assign, __awaiter, __extends, __generator } from "tslib";
import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { IconButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { strVal } from '../../services/DataverseService';
import { handleApprovalAction, checkStageDecisionLock } from '../../services/approvalHandler';
import { formatCurrency, formatDate, formatDateTime } from '../dashboard/dashboardUtils';
var TakeActionModal = /** @class */ (function (_super) {
    __extends(TakeActionModal, _super);
    function TakeActionModal(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            comments: '',
            commentsError: null,
            isSubmitting: false,
            submitError: null,
            loadingHistory: false,
            historyLogs: [],
            historyError: null,
            showHistory: true,
        };
        return _this;
    }
    TakeActionModal.prototype.componentDidMount = function () {
        if (this.props.isOpen && this.props.record) {
            void this._loadHistory();
        }
    };
    TakeActionModal.prototype.componentDidUpdate = function (prevProps) {
        if (this.props.isOpen && (!prevProps.isOpen || prevProps.record !== this.props.record)) {
            this.setState({
                comments: '',
                commentsError: null,
                isSubmitting: false,
                submitError: null,
            });
            void this._loadHistory();
        }
    };
    TakeActionModal.prototype._getRecordId = function (r) {
        var _a, _b, _c;
        if (!r)
            return '';
        return String((_c = (_b = (_a = r['sct_requestid']) !== null && _a !== void 0 ? _a : r['sct_requests_id']) !== null && _b !== void 0 ? _b : r['id']) !== null && _c !== void 0 ? _c : '');
    };
    TakeActionModal.prototype._loadHistory = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _a, record, service, reqId, logs, err_1, errMsg;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = this.props, record = _a.record, service = _a.service;
                        if (!record)
                            return [2 /*return*/];
                        reqId = this._getRecordId(record);
                        if (!reqId)
                            return [2 /*return*/];
                        this.setState({ loadingHistory: true, historyError: null });
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, service.getApprovalLogs(reqId)];
                    case 2:
                        logs = _b.sent();
                        this.setState({ historyLogs: logs || [], loadingHistory: false });
                        return [3 /*break*/, 4];
                    case 3:
                        err_1 = _b.sent();
                        errMsg = err_1 instanceof Error ? err_1.message : String(err_1);
                        this.setState({ historyError: errMsg, loadingHistory: false });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /** Role title for display */
    TakeActionModal.prototype._roleDisplay = function () {
        var r = (this.props.activeRole || '').toLowerCase();
        if (r === 'manager')
            return 'Line Manager';
        if (r.indexOf('1') !== -1)
            return 'Tier 1 Approver';
        if (r.indexOf('2') !== -1)
            return 'Tier 2 Approver';
        if (r.indexOf('3') !== -1)
            return 'Tier 3 Approver';
        return 'Approver';
    };
    /** Target destination text for the approve button */
    TakeActionModal.prototype._approveActionText = function () {
        var r = (this.props.activeRole || '').toLowerCase();
        if (r === 'manager')
            return 'Approve & Route to Tier 1';
        if (r.indexOf('1') !== -1)
            return 'Approve & Route to Tier 2';
        if (r.indexOf('2') !== -1)
            return 'Approve & Route to Tier 3';
        if (r.indexOf('3') !== -1)
            return 'Final Approval (Issue PO)';
        return 'Approve Request';
    };
    /** Current pending stage text */
    TakeActionModal.prototype._currentStageInfo = function () {
        var record = this.props.record;
        if (!record)
            return { label: 'Line Manager Review', currentStep: 0 };
        var pt = String(record['cr2a4_pendingtier'] || record['cr2a4_PendingTier'] || record['sct_pendingtier'] || record['pendingtier'] || '').toLowerCase();
        if (pt.indexOf('3') !== -1)
            return { label: 'Tier 3 Approver Review', currentStep: 3 };
        if (pt.indexOf('2') !== -1)
            return { label: 'Tier 2 Approver Review', currentStep: 2 };
        if (pt.indexOf('1') !== -1)
            return { label: 'Tier 1 Approver Review', currentStep: 1 };
        return { label: 'Line Manager Review', currentStep: 0 };
    };
    TakeActionModal.prototype._handleDecision = function (decision) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, record, service, currentSystemUserId, activeRole, context, comments, requestId, trimmedComments, graphClient, _b, rawReq, result, roleLabel, successMsg, next, err_2, errMsg;
            var _c, _d, _e;
            return __generator(this, function (_f) {
                switch (_f.label) {
                    case 0:
                        _a = this.props, record = _a.record, service = _a.service, currentSystemUserId = _a.currentSystemUserId, activeRole = _a.activeRole, context = _a.context;
                        comments = this.state.comments;
                        if (!record)
                            return [2 /*return*/];
                        requestId = this._getRecordId(record);
                        if (!requestId) {
                            this.setState({ submitError: 'Invalid request: Missing record ID.' });
                            return [2 /*return*/];
                        }
                        trimmedComments = (comments || '').trim();
                        // Enforce mandatory rationale for Reject and Return
                        if ((decision === 'reject' || decision === 'return') && trimmedComments.length === 0) {
                            this.setState({
                                commentsError: "Comments are mandatory when taking a \"".concat(decision === 'reject' ? 'Reject' : 'Return for Correction', "\" action. Please provide a clear reason for the requester."),
                            });
                            return [2 /*return*/];
                        }
                        this.setState({ isSubmitting: true, submitError: null, commentsError: null });
                        _f.label = 1;
                    case 1:
                        _f.trys.push([1, 7, , 8]);
                        graphClient = void 0;
                        _f.label = 2;
                    case 2:
                        _f.trys.push([2, 4, , 5]);
                        return [4 /*yield*/, context.msGraphClientFactory.getClient('3')];
                    case 3:
                        graphClient = _f.sent();
                        return [3 /*break*/, 5];
                    case 4:
                        _b = _f.sent();
                        return [3 /*break*/, 5];
                    case 5:
                        rawReq = String((_e = (_d = (_c = record['_sct_requester_value']) !== null && _c !== void 0 ? _c : record['sct_requester']) !== null && _d !== void 0 ? _d : record['_createdby_value']) !== null && _e !== void 0 ? _e : '');
                        return [4 /*yield*/, handleApprovalAction({
                                requestId: requestId,
                                role: activeRole || 'manager',
                                action: decision,
                                comments: trimmedComments,
                                service: service,
                                currentSystemUserId: currentSystemUserId,
                                requesterId: rawReq,
                                graphClient: graphClient,
                                webPartUrl: context.pageContext.web.absoluteUrl,
                                approverDisplayName: context.pageContext.user.displayName,
                                currentUserEmail: context.pageContext.user.email || undefined,
                                soloTestMode: this.props.soloTestMode,
                            })];
                    case 6:
                        result = _f.sent();
                        roleLabel = this._roleDisplay();
                        successMsg = '';
                        if (decision === 'approve') {
                            next = (activeRole || '').toLowerCase().indexOf('3') !== -1 ? 'Approved and PO Issued' : 'advanced to the next approval tier';
                            successMsg = "Request successfully approved by ".concat(roleLabel, " and ").concat(next, ".");
                        }
                        else if (decision === 'reject') {
                            successMsg = "Request successfully rejected by ".concat(roleLabel, ". Workflow terminated.");
                        }
                        else {
                            successMsg = "Request returned for correction by ".concat(roleLabel, ". The requester may now revise and resubmit.");
                        }
                        if (result.emailDispatched) {
                            successMsg += ' Automated notification email dispatched.';
                        }
                        this.setState({ isSubmitting: false });
                        this.props.onSuccess(successMsg);
                        return [3 /*break*/, 8];
                    case 7:
                        err_2 = _f.sent();
                        errMsg = err_2 instanceof Error ? err_2.message : 'Approval action failed. Please try again.';
                        this.setState({ isSubmitting: false, submitError: errMsg });
                        return [3 /*break*/, 8];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    TakeActionModal.prototype.render = function () {
        var _this = this;
        var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l;
        var _m = this.props, isOpen = _m.isOpen, record = _m.record, onDismiss = _m.onDismiss;
        var _o = this.state, comments = _o.comments, commentsError = _o.commentsError, isSubmitting = _o.isSubmitting, submitError = _o.submitError, loadingHistory = _o.loadingHistory, historyLogs = _o.historyLogs, historyError = _o.historyError, showHistory = _o.showHistory;
        if (!isOpen || !record)
            return null;
        var ref = strVal(record, 'sct_sct_reference') || 'Request';
        var title = strVal(record, 'sct_request_title') || '—';
        var desc = strVal(record, 'sct_description');
        var reqName = strVal(record, '_sct_requester_value@OData.Community.Display.V1.FormattedValue')
            || strVal(record, 'sct_requester@OData.Community.Display.V1.FormattedValue')
            || strVal(record, 'requesterName')
            || strVal(record, 'sct_requester')
            || '—';
        var bu = strVal(record, 'sct_business_unit') || strVal(record, 'sct_businessunit') || '—';
        var costCenter = strVal(record, 'sct_cost_center') || strVal(record, 'sct_costcenter') || '—';
        var site = strVal(record, '_sct_site_value') || '—';
        var ptype = strVal(record, 'sct_purchase_type') || '—';
        var urgency = strVal(record, 'sct_urgency') || 'Standard';
        var dateNeeded = strVal(record, 'sct_date_needed_by');
        var amtRaw = parseFloat(String((_a = record['sct_estimated_total_value']) !== null && _a !== void 0 ? _a : 0));
        var amountFormatted = !isNaN(amtRaw) && amtRaw > 0 ? formatCurrency(amtRaw) : '—';
        var stageInfo = this._currentStageInfo();
        var roleTitle = this._roleDisplay();
        var approveText = this._approveActionText();
        var decisionLock = checkStageDecisionLock(record, this.props.activeRole, historyLogs);
        // Segregation of Duties (SoD) evaluation
        var rawReq = String((_d = (_c = (_b = record['_sct_requester_value']) !== null && _b !== void 0 ? _b : record['sct_requester']) !== null && _c !== void 0 ? _c : record['_createdby_value']) !== null && _d !== void 0 ? _d : '');
        var reqId = rawReq.toLowerCase().replace(/[{}]/g, '');
        var myId = (this.props.currentSystemUserId || '').toLowerCase().replace(/[{}]/g, '');
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        var ctxAny = this.props.context;
        var aadId = (((_h = (_g = (_f = (_e = ctxAny === null || ctxAny === void 0 ? void 0 : ctxAny.pageContext) === null || _e === void 0 ? void 0 : _e.aadInfo) === null || _f === void 0 ? void 0 : _f.userId) === null || _g === void 0 ? void 0 : _g.toString) === null || _h === void 0 ? void 0 : _h.call(_g)) || '').toLowerCase().replace(/[{}]/g, '');
        var myName = (((_l = (_k = (_j = this.props.context) === null || _j === void 0 ? void 0 : _j.pageContext) === null || _k === void 0 ? void 0 : _k.user) === null || _l === void 0 ? void 0 : _l.displayName) || '').toLowerCase().trim();
        var reqDisplayName = strVal(record, '_sct_requester_value@OData.Community.Display.V1.FormattedValue').toLowerCase().trim();
        var isOwnRequest = (Boolean(myId) && reqId === myId) || (Boolean(aadId) && reqId === aadId) || (Boolean(myName) && reqDisplayName === myName);
        var isSodViolation = !this.props.soloTestMode && isOwnRequest;
        var canAct = !isSubmitting && !decisionLock.isLocked && !isSodViolation;
        var STAGES = ['Manager', 'Tier 1', 'Tier 2', 'Tier 3', 'Approved'];
        var btnBase = {
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 7,
            padding: '10px 18px', borderRadius: 6, border: 'none',
            fontSize: 13, fontWeight: 600, cursor: isSubmitting ? 'not-allowed' : 'pointer',
            opacity: isSubmitting ? 0.6 : 1, transition: 'all 0.15s ease-in-out',
            minWidth: 130, height: 38,
        };
        return (React.createElement(Panel, { isOpen: this.props.isOpen && record !== null, onDismiss: function () { return !isSubmitting && onDismiss(); }, type: PanelType.custom, customWidth: "min(800px, 96vw)", hasCloseButton: false, isFooterAtBottom: true, onRenderNavigationContent: function () { return (React.createElement("div", { style: {
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '16px 24px',
                    borderBottom: '1px solid #EDEBE9',
                    background: '#FFFFFF',
                    width: '100%',
                    boxSizing: 'border-box',
                } },
                React.createElement("div", null,
                    React.createElement("div", { style: { fontSize: 18, fontWeight: 700, color: '#201F1E', display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement(Icon, { iconName: "CheckList", style: { color: '#0078D4', fontSize: 18 } }),
                        React.createElement("span", null,
                            "Take Action \u2014 ",
                            ref)),
                    React.createElement("div", { style: { fontSize: 12, color: '#605E5C', marginTop: 2 } },
                        "Review and decide on this spend request as ",
                        React.createElement("strong", null, roleTitle))),
                React.createElement(IconButton, { iconProps: { iconName: 'Cancel' }, ariaLabel: "Close", title: "Close panel", disabled: isSubmitting, onClick: onDismiss, styles: {
                        root: { color: '#605E5C', borderRadius: 4 },
                        rootHovered: { color: '#201F1E', background: '#F3F2F1' },
                    } }))); }, styles: {
                main: { boxShadow: '-4px 0 24px rgba(0,0,0,0.12)' },
                scrollableContent: { display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' },
                content: { flex: 1, overflowY: 'auto', paddingTop: 16, paddingBottom: 8, paddingLeft: 24, paddingRight: 24 },
                footer: { borderTop: 'none', padding: 0 },
            }, onRenderFooterContent: function () { return (React.createElement("div", { style: {
                    display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'flex-end',
                    padding: '14px 24px', background: '#FAFAFA', borderTop: '1px solid #EDEBE9',
                } },
                React.createElement("button", { style: __assign(__assign({}, btnBase), { background: canAct ? '#107C10' : '#EDEBE9', color: canAct ? '#FFFFFF' : '#A19F9D', boxShadow: canAct && !isSubmitting ? '0 2px 6px rgba(16,124,16,0.25)' : 'none', cursor: canAct ? 'pointer' : 'not-allowed' }), disabled: !canAct, title: canAct ? approveText : (decisionLock.lockMessage || 'Actions are locked for this stage'), onClick: function () { return void _this._handleDecision('approve'); } },
                    isSubmitting ? React.createElement(Spinner, { size: SpinnerSize.xSmall, styles: { root: { display: 'inline-flex' } } }) : React.createElement(Icon, { iconName: "CheckMark", style: { fontSize: 12 } }),
                    isSubmitting ? 'Saving…' : approveText),
                React.createElement("button", { style: __assign(__assign({}, btnBase), { background: canAct ? '#FFF4CE' : '#EDEBE9', color: canAct ? '#7A4F01' : '#A19F9D', border: canAct ? '1.5px solid #C8961F' : '1px solid #EDEBE9', cursor: canAct ? 'pointer' : 'not-allowed' }), disabled: !canAct, title: canAct ? 'Return for Correction' : (decisionLock.lockMessage || 'Actions are locked for this stage'), onClick: function () { return void _this._handleDecision('return'); } },
                    isSubmitting ? React.createElement(Spinner, { size: SpinnerSize.xSmall, styles: { root: { display: 'inline-flex' } } }) : React.createElement(Icon, { iconName: "ReturnToSession", style: { fontSize: 12 } }),
                    isSubmitting ? 'Saving…' : 'Return for Correction'),
                React.createElement("button", { style: __assign(__assign({}, btnBase), { background: canAct ? '#FDE7E9' : '#EDEBE9', color: canAct ? '#A4262C' : '#A19F9D', border: canAct ? '1.5px solid #A4262C' : '1px solid #EDEBE9', cursor: canAct ? 'pointer' : 'not-allowed' }), disabled: !canAct, title: canAct ? 'Reject' : (decisionLock.lockMessage || 'Actions are locked for this stage'), onClick: function () { return void _this._handleDecision('reject'); } },
                    isSubmitting ? React.createElement(Spinner, { size: SpinnerSize.xSmall, styles: { root: { display: 'inline-flex' } } }) : React.createElement(Icon, { iconName: "Cancel", style: { fontSize: 12 } }),
                    isSubmitting ? 'Saving…' : 'Reject'),
                React.createElement("button", { style: __assign(__assign({}, btnBase), { background: '#F3F2F1', color: '#323130', border: '1px solid #C8C6C4', minWidth: 90 }), disabled: isSubmitting, onClick: onDismiss }, "Cancel"))); } },
            submitError && (React.createElement(MessageBar, { messageBarType: MessageBarType.error, isMultiline: true, onDismiss: function () { return _this.setState({ submitError: null }); }, styles: { root: { borderRadius: 6, marginBottom: 14 } } }, submitError)),
            decisionLock.isLocked && (React.createElement("div", { style: {
                    background: decisionLock.previousDecision ? '#F0FDF4' : '#FFFBEB',
                    border: "1.5px solid ".concat(decisionLock.previousDecision ? '#86EFAC' : '#FDE68A'),
                    borderRadius: 8,
                    padding: '12px 16px',
                    marginBottom: 16,
                } },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 } },
                    React.createElement(Icon, { iconName: decisionLock.previousDecision ? 'CompletedSolid' : 'LockSolid', style: { fontSize: 16, color: decisionLock.previousDecision ? '#16A34A' : '#D97706' } }),
                    React.createElement("span", { style: { fontSize: 13, fontWeight: 700, color: decisionLock.previousDecision ? '#166534' : '#92400E' } }, decisionLock.previousDecision
                        ? "You have already submitted your decision for this stage (".concat(decisionLock.stageName, ").")
                        : decisionLock.lockMessage)),
                decisionLock.previousDecision && (React.createElement("div", { style: { fontSize: 12, color: '#1E293B', marginLeft: 24, lineHeight: 1.6 } },
                    React.createElement("div", null,
                        "Decision: ",
                        React.createElement("strong", { style: { color: decisionLock.previousDecision.action === 'Approved' ? '#16A34A' : decisionLock.previousDecision.action === 'Rejected' ? '#DC2626' : '#D97706' } }, decisionLock.previousDecision.action),
                        " by ",
                        React.createElement("strong", null, decisionLock.previousDecision.approverName),
                        " on ",
                        React.createElement("strong", null, decisionLock.previousDecision.timestamp),
                        "."),
                    decisionLock.previousDecision.comments && (React.createElement("div", { style: { fontStyle: 'italic', marginTop: 2 } },
                        "\u201C",
                        decisionLock.previousDecision.comments,
                        "\u201D")),
                    React.createElement("div", { style: { marginTop: 4, fontSize: 11, color: '#64748B' } }, "Further actions for this stage are permanently locked in this approval cycle."))))),
            isSodViolation && (React.createElement("div", { style: {
                    background: '#FFFBEB',
                    border: '1.5px solid #FCD34D',
                    borderRadius: 8,
                    padding: '12px 16px',
                    marginBottom: 16,
                } },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                    React.createElement(Icon, { iconName: "ShieldAlert", style: { fontSize: 18, color: '#D97706' } }),
                    React.createElement("span", { style: { fontSize: 13, fontWeight: 700, color: '#92400E' } }, "Segregation of Duties (SoD) Enforced \u2014 Strict Mode")),
                React.createElement("div", { style: { fontSize: 12, color: '#78350F', marginTop: 4, lineHeight: 1.5, marginLeft: 26 } },
                    "You submitted this procurement request. Corporate governance strictly prohibits requesters from approving their own requests. To simulate and test this workflow alone, toggle ",
                    React.createElement("strong", null, "Solo Test Mode: ON"),
                    " in the top header."))),
            this.props.soloTestMode && isOwnRequest && (React.createElement("div", { style: {
                    background: '#F0FDF4',
                    border: '1.5px solid #86EFAC',
                    borderRadius: 8,
                    padding: '10px 14px',
                    marginBottom: 16,
                    display: 'flex', alignItems: 'center', gap: 8,
                } },
                React.createElement(Icon, { iconName: "Unlock", style: { fontSize: 16, color: '#16A34A' } }),
                React.createElement("span", { style: { fontSize: 12, color: '#166534', fontWeight: 600 } }, "Solo Test Mode Active \u2014 Self-approval is enabled for solo development and end-to-end verification."))),
            React.createElement("div", { style: {
                    display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                    background: '#F0F6FF', border: '1px solid #C7E0F4', borderRadius: 8,
                    padding: '10px 14px', marginBottom: 16,
                } },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                    React.createElement(Icon, { iconName: "Shield", style: { color: '#0078D4', fontSize: 16 } }),
                    React.createElement(Text, { variant: "small", style: { color: '#004578', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.4px' } },
                        "Acting Role: ",
                        roleTitle)),
                React.createElement("span", { style: {
                        background: '#0078D4', color: '#fff', fontSize: 11, fontWeight: 700,
                        padding: '3px 10px', borderRadius: 12, letterSpacing: '0.3px',
                    } }, stageInfo.label)),
            React.createElement("div", { style: {
                    display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                    background: '#FAF9F8', border: '1px solid #EDEBE9', borderRadius: 8,
                    padding: '10px 16px', marginBottom: 18,
                } }, STAGES.map(function (s, idx) {
                var isCurrent = idx === stageInfo.currentStep;
                var isDone = idx < stageInfo.currentStep;
                var color = isCurrent ? '#0078D4' : isDone ? '#107C10' : '#8A8886';
                var bg = isCurrent ? '#DEECF9' : isDone ? '#DFF6DD' : '#F3F2F1';
                return (React.createElement(React.Fragment, { key: s },
                    React.createElement("div", { style: { display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 3 } },
                        React.createElement("div", { style: {
                                width: 22, height: 22, borderRadius: '50%', background: bg,
                                color: color,
                                display: 'flex', alignItems: 'center', justifyContent: 'center',
                                fontSize: 11, fontWeight: 700, border: "1.5px solid ".concat(color),
                            } }, isDone ? React.createElement(Icon, { iconName: "CheckMark", style: { fontSize: 11 } }) : idx + 1),
                        React.createElement("span", { style: { fontSize: 11, fontWeight: isCurrent ? 700 : 500, color: color } }, s)),
                    idx < STAGES.length - 1 && (React.createElement("div", { style: {
                            flex: 1, height: 2, margin: '0 6px',
                            background: idx < stageInfo.currentStep ? '#107C10' : '#EDEBE9',
                        } }))));
            })),
            React.createElement("div", { style: {
                    background: '#FFFFFF', border: '1px solid #E1DFDD', borderRadius: 8,
                    padding: '14px 18px', marginBottom: 18, boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                } },
                React.createElement("div", { style: { fontSize: 15, fontWeight: 700, color: '#201F1E', marginBottom: 12 } }, title || ref),
                React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px 16px' } },
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' } }, "Requester"),
                        React.createElement("div", { style: { fontSize: 13, fontWeight: 600, color: '#323130' } }, reqName)),
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' } }, "Business Unit"),
                        React.createElement("div", { style: { fontSize: 13, fontWeight: 600, color: '#323130' } }, bu)),
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' } }, "Estimated Value"),
                        React.createElement("div", { style: { fontSize: 14, fontWeight: 700, color: '#107C10' } }, amountFormatted)),
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' } }, "Site / Location"),
                        React.createElement("div", { style: { fontSize: 13, color: '#323130' } }, site)),
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' } }, "Purchase Type"),
                        React.createElement("div", { style: { fontSize: 13, color: '#323130' } }, ptype)),
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' } }, "Urgency"),
                        React.createElement("div", { style: {
                                fontSize: 12, fontWeight: 700,
                                color: urgency === 'Emergency' ? '#A4262C' : urgency === 'Urgent' ? '#C8961F' : '#107C10',
                            } }, urgency)),
                    costCenter && costCenter !== '—' && (React.createElement("div", null,
                        React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' } }, "Cost Center"),
                        React.createElement("div", { style: { fontSize: 13, color: '#323130' } }, costCenter))),
                    dateNeeded && (React.createElement("div", null,
                        React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' } }, "Date Needed"),
                        React.createElement("div", { style: { fontSize: 13, color: '#323130' } }, formatDate(dateNeeded))))),
                desc && (React.createElement("div", { style: { marginTop: 12, paddingTop: 10, borderTop: '1px solid #F3F2F1' } },
                    React.createElement("div", { style: { fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px', marginBottom: 4 } }, "Requirement Justification"),
                    React.createElement("div", { style: { fontSize: 12, color: '#494543', lineHeight: 1.5, background: '#F8F9FA', padding: '8px 10px', borderRadius: 4 } }, desc)))),
            React.createElement("div", { style: {
                    border: '1px solid #EDEBE9', borderRadius: 8, marginBottom: 18, overflow: 'hidden',
                } },
                React.createElement("button", { onClick: function () { return _this.setState({ showHistory: !showHistory }); }, style: {
                        width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                        padding: '9px 14px', background: '#FAF9F8', border: 'none', cursor: 'pointer',
                        textAlign: 'left',
                    } },
                    React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 6 } },
                        React.createElement(Icon, { iconName: "History", style: { color: '#605E5C', fontSize: 13 } }),
                        React.createElement("span", { style: { fontSize: 12, fontWeight: 700, color: '#323130' } }, "Previous Approval History"),
                        React.createElement("span", { style: {
                                fontSize: 11, background: '#EDEBE9', color: '#605E5C', fontWeight: 600,
                                padding: '1px 7px', borderRadius: 10, marginLeft: 4,
                            } }, historyLogs.length)),
                    React.createElement(Icon, { iconName: showHistory ? 'ChevronUp' : 'ChevronDown', style: { fontSize: 11, color: '#605E5C' } })),
                showHistory && (React.createElement("div", { style: { padding: '10px 14px', background: '#FFFFFF', maxHeight: 260, overflowY: 'auto' } },
                    loadingHistory && (React.createElement("div", { style: { padding: '8px 0', textAlign: 'center' } },
                        React.createElement(Spinner, { size: SpinnerSize.small, label: "Loading approval history\u2026" }))),
                    historyError && (React.createElement(MessageBar, { messageBarType: MessageBarType.warning, styles: { root: { borderRadius: 4 } } },
                        "Could not load approval logs: ",
                        historyError)),
                    !loadingHistory && !historyError && historyLogs.length === 0 && (React.createElement("div", { style: { fontSize: 12, color: '#605E5C', fontStyle: 'italic', padding: '4px 0' } }, "No previous approvals recorded. This request is at the initial approval stage.")),
                    !loadingHistory && historyLogs.map(function (log, idx) {
                        var _a, _b;
                        var logTierRaw = log['sct_tiername'];
                        var logTierName = logTierRaw === 4 ? 'Manager'
                            : logTierRaw === 1 ? 'Tier 1'
                                : logTierRaw === 2 ? 'Tier 2'
                                    : logTierRaw === 3 ? 'Tier 3'
                                        : strVal(log, 'sct_tiername') || 'Approver';
                        var statusRaw = log['sct_approvalstatus'];
                        var isApproved = statusRaw === 1;
                        var isRejected = statusRaw === 2;
                        var statusBadgeBg = isApproved ? '#DFF6DD' : isRejected ? '#FDE7E9' : '#FFF4CE';
                        var statusBadgeColor = isApproved ? '#107C10' : isRejected ? '#A4262C' : '#7A4F01';
                        var statusText = isApproved ? 'Approved' : isRejected ? 'Rejected' : 'Returned';
                        var approver = strVal(log, '_sct_approverid_value') || strVal(log, '_createdby_value') || 'Approver';
                        var commentsText = strVal(log, 'sct_comments') || '—';
                        var rawDate = ((_b = (_a = log['createdon']) !== null && _a !== void 0 ? _a : log['sct_decisionon']) !== null && _b !== void 0 ? _b : '');
                        var dateText = formatDateTime(rawDate);
                        return (React.createElement("div", { key: idx, style: {
                                display: 'flex', flexDirection: 'column', gap: 4,
                                padding: '8px 10px', borderRadius: 6, background: '#F8F9FA',
                                border: '1px solid #EDEBE9', marginBottom: idx < historyLogs.length - 1 ? 8 : 0,
                            } },
                            React.createElement("div", { style: { display: 'flex', alignItems: 'center', justifyContent: 'space-between' } },
                                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 6 } },
                                    React.createElement("span", { style: { fontSize: 12, fontWeight: 700, color: '#201F1E' } }, approver),
                                    React.createElement("span", { style: { fontSize: 10, background: '#E1DFDD', color: '#323130', padding: '1px 6px', borderRadius: 4, fontWeight: 600 } }, logTierName)),
                                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                                    React.createElement("span", { style: {
                                            fontSize: 10, fontWeight: 700, padding: '2px 8px', borderRadius: 10,
                                            background: statusBadgeBg, color: statusBadgeColor,
                                        } }, statusText),
                                    React.createElement("span", { style: { fontSize: 11, color: '#8A8886' } }, dateText))),
                            commentsText && commentsText !== '—' && (React.createElement("div", { style: { fontSize: 12, color: '#494543', fontStyle: 'italic', paddingLeft: 4 } },
                                "\u201C",
                                commentsText,
                                "\u201D"))));
                    })))),
            React.createElement(TextField, { label: "Decision Comments", required: canAct, multiline: true, rows: 3, placeholder: canAct ? "Enter your rationale, conditions, or instructions for this decision…" : "Decision comments are locked because action has already been taken or this stage is inactive.", value: comments, onChange: function (_, v) { return _this.setState({ comments: v !== null && v !== void 0 ? v : '', commentsError: null }); }, errorMessage: commentsError || undefined, disabled: !canAct, description: canAct ? "Mandatory for Reject and Return for Correction decisions." : undefined, styles: { root: { marginBottom: 12 }, fieldGroup: { borderRadius: 6 } } })));
    };
    return TakeActionModal;
}(React.Component));
export default TakeActionModal;
//# sourceMappingURL=TakeActionModal.js.map