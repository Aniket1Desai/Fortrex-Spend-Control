import { __awaiter, __extends, __generator } from "tslib";
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import styles from './AppShell.module.scss';
import Dashboard from '../dashboard/Dashboard';
import NewRequestForm from '../newRequest/NewRequestForm';
import MyRequests from '../myRequests/MyRequests';
import ReviewQueue from '../reviewQueue/ReviewQueue';
var TABS = [
    { id: 'dashboard', label: 'Dashboard', iconName: 'ViewDashboard', ariaLabel: 'Dashboard tab' },
    { id: 'newRequest', label: 'New Request', iconName: 'Add', ariaLabel: 'New Request tab' },
    { id: 'myRequests', label: 'My Requests', iconName: 'DocumentSearch', ariaLabel: 'My Requests tab' },
    { id: 'reviewQueue', label: 'Review Queue', iconName: 'Checklist', ariaLabel: 'Review Queue tab' },
];
var TIER1_KEY = 'approver_tier1';
// ── Azure AD Security Group IDs (confirmed from Entra admin center) ───────────
var TIER1_GROUP_ID = 'fdbf07a0-548b-4ccf-875b-900cb1b24a1c'; // SCT-Fortrex_Approvers-Tier1
var TIER2_GROUP_ID = '6a8a4077-c78f-4af2-ae5e-12428d1dcbc3'; // SCT-Fortrex_Approvers-Tier2
var TIER3_GROUP_ID = 'e63dbc4f-6ad2-4969-a518-b229836756af'; // SCT-Fortrex_Approvers-Tier3
var AppShell = /** @class */ (function (_super) {
    __extends(AppShell, _super);
    function AppShell(props) {
        var _this = _super.call(this, props) || this;
        // ── Role Switcher badge renderer ─────────────────────────────────────────────
        _this._onRenderRoleOption = function (option, defaultRender) {
            if (!option)
                return null;
            var count = _this.state.rolePendingCounts[option.key] || (option.key === TIER1_KEY ? _this.state.tier1PendingCount : 0);
            if (count > 0) {
                return (React.createElement("div", { style: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', width: '100%', gap: 8 } },
                    React.createElement("span", { style: { flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, option.text),
                    React.createElement("span", { style: {
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            minWidth: 20, height: 20, padding: '0 6px',
                            borderRadius: 10,
                            background: '#D13438',
                            color: '#ffffff',
                            fontSize: 11, fontWeight: 700, lineHeight: 1,
                            flexShrink: 0,
                        } }, count > 99 ? '99+' : count)));
            }
            return defaultRender ? defaultRender(option) : React.createElement("span", null, option.text);
        };
        // ── Selected-item renderer — mirrors badge in the collapsed trigger ──────────
        _this._onRenderRoleTitle = function (options, defaultRender) {
            if (!options || !options.length)
                return null;
            var option = options[0];
            var count = _this.state.rolePendingCounts[option.key] || (option.key === TIER1_KEY ? _this.state.tier1PendingCount : 0);
            if (count > 0) {
                return (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 6 } },
                    React.createElement("span", null, option.text),
                    React.createElement("span", { style: {
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            minWidth: 18, height: 18, padding: '0 5px',
                            borderRadius: 9,
                            background: '#D13438', color: '#ffffff',
                            fontSize: 10, fontWeight: 700, lineHeight: 1,
                            flexShrink: 0,
                        } }, count > 99 ? '99+' : count)));
            }
            return defaultRender ? defaultRender(options) : React.createElement("span", null, option.text);
        };
        _this._handleTabClick = function (targetTab) {
            if (_this.state.activeTab === targetTab)
                return;
            // Check if user has entered data in the new request form ("on half way")
            if (_this.state.activeTab === 'newRequest' && _this.state.isFormDirty) {
                _this.setState({
                    showLeaveConfirmDialog: true,
                    pendingTabSwitch: targetTab,
                    pendingRoleSwitch: null,
                });
                return;
            }
            _this.setState({
                activeTab: targetTab,
                editingRequest: targetTab === 'newRequest' ? null : _this.state.editingRequest,
            });
        };
        _this._handleRoleChange = function (option) {
            if (!option)
                return;
            var newRole = option.key;
            var currentTab = _this.state.activeTab;
            var newTab = currentTab;
            // If switching to Requester while on Review Queue, redirect to dashboard.
            // (New Request is now available for all roles).
            if (newRole === 'requester' && currentTab === 'reviewQueue')
                newTab = 'dashboard';
            // If switching role while new request form is dirty
            if (currentTab === 'newRequest' && _this.state.isFormDirty && newTab !== 'newRequest') {
                _this.setState({
                    showLeaveConfirmDialog: true,
                    pendingTabSwitch: newTab,
                    pendingRoleSwitch: newRole,
                });
                return;
            }
            _this.setState({ activeRole: newRole, activeTab: newTab });
        };
        _this._confirmLeaveTab = function () {
            var _a = _this.state, pendingTabSwitch = _a.pendingTabSwitch, pendingRoleSwitch = _a.pendingRoleSwitch;
            _this.setState({
                showLeaveConfirmDialog: false,
                isFormDirty: false,
                activeTab: pendingTabSwitch || 'dashboard',
                activeRole: pendingRoleSwitch || _this.state.activeRole,
                pendingTabSwitch: null,
                pendingRoleSwitch: null,
                editingRequest: null,
            });
        };
        _this._cancelLeaveConfirm = function () {
            _this.setState({
                showLeaveConfirmDialog: false,
                pendingTabSwitch: null,
                pendingRoleSwitch: null,
            });
        };
        var defaultTestMode = true;
        try {
            if (typeof window !== 'undefined' && window.localStorage) {
                var stored = window.localStorage.getItem('sct_solo_test_mode');
                if (stored !== null)
                    defaultTestMode = stored === 'true';
            }
        }
        catch ( /* ignore */_a) { /* ignore */ }
        _this.state = {
            activeTab: 'dashboard',
            activeRole: 'requester',
            tier1PendingCount: 0,
            rolePendingCounts: {},
            roleOptions: [{ key: 'requester', text: 'Requester' }],
            rolesLoading: true,
            userRoles: ['requester'],
            soloTestMode: defaultTestMode,
            editingRequest: null,
            deepLinkRequestId: null,
            isFormDirty: false,
            showLeaveConfirmDialog: false,
            pendingTabSwitch: null,
            pendingRoleSwitch: null,
        };
        return _this;
    }
    AppShell.prototype.componentDidMount = function () {
        void this._loadUserRoles();
        if (!this.props.serviceError) {
            void this._loadAllPendingCounts();
        }
        // Deep-link auto-routing (e.g. from Outlook approval email: ?requestId=SCT-2026-xxxx)
        try {
            if (typeof window !== 'undefined' && window.location && window.location.search) {
                var params = new URLSearchParams(window.location.search);
                var reqParam = params.get('requestId') || params.get('sctRef') || params.get('id');
                if (reqParam) {
                    this.setState({ deepLinkRequestId: reqParam.trim() });
                }
            }
        }
        catch ( /* ignore */_a) { /* ignore */ }
    };
    AppShell.prototype._loadUserRoles = function () {
        return __awaiter(this, void 0, void 0, function () {
            var ctx, aadOid, normAadOid, userEmail, userDisplayName, systemUser, _a, _b, systemUserId, adIsManager, graphClient, _c, reportsResponse, _d, dataverseIsManager, mgrPending, isManager, userGroupIds_1, transResp, err_1, checkResp, err_2, dvTier1_1, dvTier2_1, dvTier3_1, roleConfigs, sId_1, matchingRoles, _e, isTier1, isTier2, isTier3, opts, seenKeys, uniqueOpts, _i, opts_1, opt, k, validKeys, newActiveRole, _f;
            var _g, _h, _j, _k, _l, _m, _o, _p, _q, _r;
            return __generator(this, function (_s) {
                switch (_s.label) {
                    case 0:
                        ctx = this.props.context;
                        aadOid = (_k = (_j = (_h = (_g = ctx.pageContext.aadInfo) === null || _g === void 0 ? void 0 : _g.userId) === null || _h === void 0 ? void 0 : _h.toString()) === null || _j === void 0 ? void 0 : _j.toLowerCase()) !== null && _k !== void 0 ? _k : '';
                        normAadOid = aadOid.replace(/[{}-]/g, '');
                        userEmail = ((_m = (_l = ctx.pageContext.user) === null || _l === void 0 ? void 0 : _l.email) !== null && _m !== void 0 ? _m : '').toLowerCase().trim();
                        userDisplayName = ((_p = (_o = ctx.pageContext.user) === null || _o === void 0 ? void 0 : _o.displayName) !== null && _p !== void 0 ? _p : '').toLowerCase().trim();
                        _s.label = 1;
                    case 1:
                        _s.trys.push([1, 30, , 31]);
                        return [4 /*yield*/, this.props.dataverseService.getCurrentSystemUser(aadOid)];
                    case 2:
                        if (!((_q = _s.sent()) !== null && _q !== void 0)) return [3 /*break*/, 3];
                        _a = _q;
                        return [3 /*break*/, 7];
                    case 3:
                        if (!userEmail) return [3 /*break*/, 5];
                        return [4 /*yield*/, this.props.dataverseService.getSystemUserByEmail(userEmail)];
                    case 4:
                        _b = _s.sent();
                        return [3 /*break*/, 6];
                    case 5:
                        _b = null;
                        _s.label = 6;
                    case 6:
                        _a = (_b);
                        _s.label = 7;
                    case 7:
                        systemUser = _a;
                        systemUserId = (_r = systemUser === null || systemUser === void 0 ? void 0 : systemUser.systemuserid) !== null && _r !== void 0 ? _r : null;
                        adIsManager = false;
                        graphClient = void 0;
                        _s.label = 8;
                    case 8:
                        _s.trys.push([8, 10, , 11]);
                        return [4 /*yield*/, ctx.msGraphClientFactory.getClient('3')];
                    case 9:
                        graphClient = _s.sent();
                        return [3 /*break*/, 11];
                    case 10:
                        _c = _s.sent();
                        return [3 /*break*/, 11];
                    case 11:
                        if (!graphClient) return [3 /*break*/, 15];
                        _s.label = 12;
                    case 12:
                        _s.trys.push([12, 14, , 15]);
                        return [4 /*yield*/, graphClient
                                .api('/me/directReports')
                                .select('id')
                                .top(1)
                                .get()];
                    case 13:
                        reportsResponse = _s.sent();
                        if (((reportsResponse === null || reportsResponse === void 0 ? void 0 : reportsResponse.value) || []).length > 0) {
                            adIsManager = true;
                        }
                        return [3 /*break*/, 15];
                    case 14:
                        _d = _s.sent();
                        return [3 /*break*/, 15];
                    case 15:
                        dataverseIsManager = false;
                        if (!systemUserId) return [3 /*break*/, 18];
                        return [4 /*yield*/, this.props.dataverseService.hasDirectReportsInDataverse(systemUserId)];
                    case 16:
                        dataverseIsManager = _s.sent();
                        if (!!dataverseIsManager) return [3 /*break*/, 18];
                        return [4 /*yield*/, this.props.dataverseService.getManagerPendingCount(systemUserId)];
                    case 17:
                        mgrPending = _s.sent();
                        if (mgrPending > 0)
                            dataverseIsManager = true;
                        _s.label = 18;
                    case 18:
                        isManager = adIsManager || dataverseIsManager;
                        userGroupIds_1 = new Set();
                        if (!graphClient) return [3 /*break*/, 25];
                        _s.label = 19;
                    case 19:
                        _s.trys.push([19, 21, , 22]);
                        return [4 /*yield*/, graphClient
                                .api('/me/transitiveMemberOf')
                                .select('id,displayName')
                                .top(999)
                                .get()];
                    case 20:
                        transResp = _s.sent();
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        ((transResp === null || transResp === void 0 ? void 0 : transResp.value) || []).forEach(function (g) {
                            if (g === null || g === void 0 ? void 0 : g.id)
                                userGroupIds_1.add(String(g.id).toLowerCase());
                        });
                        return [3 /*break*/, 22];
                    case 21:
                        err_1 = _s.sent();
                        console.warn('[AppShell] Graph transitiveMemberOf check warning:', err_1);
                        return [3 /*break*/, 22];
                    case 22:
                        _s.trys.push([22, 24, , 25]);
                        return [4 /*yield*/, graphClient.api('/me/checkMemberGroups').post({
                                groupIds: [TIER1_GROUP_ID, TIER2_GROUP_ID, TIER3_GROUP_ID],
                            })];
                    case 23:
                        checkResp = _s.sent();
                        ((checkResp === null || checkResp === void 0 ? void 0 : checkResp.value) || []).forEach(function (id) {
                            if (id)
                                userGroupIds_1.add(String(id).toLowerCase());
                        });
                        return [3 /*break*/, 25];
                    case 24:
                        err_2 = _s.sent();
                        console.warn('[AppShell] Graph checkMemberGroups check warning:', err_2);
                        return [3 /*break*/, 25];
                    case 25:
                        dvTier1_1 = false;
                        dvTier2_1 = false;
                        dvTier3_1 = false;
                        if (!systemUserId) return [3 /*break*/, 29];
                        _s.label = 26;
                    case 26:
                        _s.trys.push([26, 28, , 29]);
                        return [4 /*yield*/, this.props.dataverseService.getRoleConfigs()];
                    case 27:
                        roleConfigs = _s.sent();
                        sId_1 = systemUserId.toLowerCase().replace(/[{}-]/g, '');
                        matchingRoles = roleConfigs.filter(function (r) {
                            var _a, _b;
                            var pId = String((_a = r['_sct_personid_value']) !== null && _a !== void 0 ? _a : '').toLowerCase().replace(/[{}-]/g, '');
                            var dId = String((_b = r['_sct_delegateid_value']) !== null && _b !== void 0 ? _b : '').toLowerCase().replace(/[{}-]/g, '');
                            return pId === sId_1 || dId === sId_1 || pId === normAadOid || dId === normAadOid;
                        });
                        matchingRoles.forEach(function (r) {
                            var _a;
                            var lvl = String((_a = r['sct_approvallevel']) !== null && _a !== void 0 ? _a : '').toUpperCase();
                            if (lvl === 'L1')
                                dvTier1_1 = true;
                            if (lvl === 'L2')
                                dvTier2_1 = true;
                            if (lvl === 'L3' || lvl === 'L4')
                                dvTier3_1 = true;
                        });
                        return [3 /*break*/, 29];
                    case 28:
                        _e = _s.sent();
                        return [3 /*break*/, 29];
                    case 29:
                        isTier1 = userGroupIds_1.has(TIER1_GROUP_ID.toLowerCase()) || dvTier1_1;
                        isTier2 = userGroupIds_1.has(TIER2_GROUP_ID.toLowerCase()) || dvTier2_1;
                        isTier3 = userGroupIds_1.has(TIER3_GROUP_ID.toLowerCase()) || dvTier3_1;
                        opts = [{ key: 'requester', text: 'Requester' }];
                        if (isManager)
                            opts.push({ key: 'manager', text: 'Manager' });
                        if (isTier1)
                            opts.push({ key: TIER1_KEY, text: 'Approver — Tier 1' });
                        if (isTier2)
                            opts.push({ key: 'approver_tier2', text: 'Approver — Tier 2' });
                        if (isTier3)
                            opts.push({ key: 'approver_tier3', text: 'Approver — Tier 3' });
                        seenKeys = new Set();
                        uniqueOpts = [];
                        for (_i = 0, opts_1 = opts; _i < opts_1.length; _i++) {
                            opt = opts_1[_i];
                            k = String(opt.key);
                            if (!seenKeys.has(k)) {
                                seenKeys.add(k);
                                uniqueOpts.push(opt);
                            }
                        }
                        validKeys = new Set(uniqueOpts.map(function (o) { return o.key; }));
                        newActiveRole = validKeys.has(this.state.activeRole) ? this.state.activeRole : 'requester';
                        this.setState({
                            roleOptions: uniqueOpts,
                            userRoles: uniqueOpts.map(function (o) { return o.key; }),
                            activeRole: newActiveRole,
                            rolesLoading: false,
                        });
                        // Refresh pending counts for all assigned roles
                        void this._loadAllPendingCounts(systemUserId !== null && systemUserId !== void 0 ? systemUserId : undefined);
                        return [3 /*break*/, 31];
                    case 30:
                        _f = _s.sent();
                        this.setState({ rolesLoading: false });
                        return [3 /*break*/, 31];
                    case 31: return [2 /*return*/];
                }
            });
        });
    };
    /** Loads pending counts across all tiers and manager to display notification badges */
    AppShell.prototype._loadAllPendingCounts = function (systemUserId) {
        return __awaiter(this, void 0, void 0, function () {
            var counts, _a, t1Count, t2Count, t3Count, mgrCount, _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        counts = {};
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, Promise.all([
                                this.props.dataverseService.getTierPendingRequestIds(1).then(function (ids) { return ids.length; }).catch(function () { return 0; }),
                                this.props.dataverseService.getTierPendingRequestIds(2).then(function (ids) { return ids.length; }).catch(function () { return 0; }),
                                this.props.dataverseService.getTierPendingRequestIds(3).then(function (ids) { return ids.length; }).catch(function () { return 0; }),
                                systemUserId ? this.props.dataverseService.getManagerPendingCount(systemUserId) : 0,
                            ])];
                    case 2:
                        _a = _c.sent(), t1Count = _a[0], t2Count = _a[1], t3Count = _a[2], mgrCount = _a[3];
                        counts['manager'] = mgrCount;
                        counts[TIER1_KEY] = t1Count;
                        counts['approver_tier2'] = t2Count;
                        counts['approver_tier3'] = t3Count;
                        this.setState({
                            tier1PendingCount: t1Count,
                            rolePendingCounts: counts,
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        _b = _c.sent();
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    AppShell.prototype._getInitials = function (displayName) {
        return displayName
            .split(' ')
            .slice(0, 2)
            .map(function (part) { return part.charAt(0).toUpperCase(); })
            .join('');
    };
    AppShell.prototype._renderTabContent = function () {
        var _this = this;
        // When the Dataverse tables are unreachable, render nothing — no data fetches
        if (this.props.serviceError) {
            return (React.createElement("div", { style: {
                    display: 'flex', flexDirection: 'column', alignItems: 'center',
                    justifyContent: 'center', height: '100%', gap: 14, padding: 40,
                    color: '#605E5C', textAlign: 'center',
                } },
                React.createElement(Icon, { iconName: "PlugDisconnected", style: { fontSize: 48, color: '#C8C6C4' } }),
                React.createElement("div", { style: { fontSize: 16, fontWeight: 600, color: '#323130' } }, "No data available"),
                React.createElement("div", { style: { fontSize: 13, maxWidth: 420, lineHeight: 1.6 } }, "This web part cannot connect to the Dataverse environment. Open the property pane, verify the URL, and ensure the Fortrex SCT solution is imported and published.")));
        }
        var _a = this.props, dataverseService = _a.dataverseService, context = _a.context;
        switch (this.state.activeTab) {
            case 'dashboard':
                return (React.createElement(Dashboard, { dataverseService: dataverseService, context: context, activeRole: this.state.activeRole, soloTestMode: this.state.soloTestMode, onToggleSoloTestMode: function () {
                        var nextVal = !_this.state.soloTestMode;
                        try {
                            if (typeof window !== 'undefined' && window.localStorage) {
                                window.localStorage.setItem('sct_solo_test_mode', String(nextVal));
                            }
                        }
                        catch ( /* ignore */_a) { /* ignore */ }
                        _this.setState({ soloTestMode: nextVal });
                    }, deepLinkRequestId: this.state.deepLinkRequestId }));
            case 'newRequest':
                return (React.createElement(NewRequestForm, { dataverseService: dataverseService, context: context, activeRole: this.state.activeRole, initialRecord: this.state.editingRequest, onDirtyChange: function (isDirty) { return _this.setState({ isFormDirty: isDirty }); }, onFinished: function () { return _this.setState({ activeTab: 'myRequests', editingRequest: null, isFormDirty: false }); } }));
            case 'myRequests':
                return (React.createElement(MyRequests, { dataverseService: dataverseService, context: context, activeRole: this.state.activeRole, onEditRequest: function (rec) { return _this.setState({ activeTab: 'newRequest', editingRequest: rec }); } }));
            case 'reviewQueue':
                return React.createElement(ReviewQueue, { dataverseService: dataverseService, context: context, activeRole: this.state.activeRole, soloTestMode: this.state.soloTestMode });
            default:
                return (React.createElement(Dashboard, { dataverseService: dataverseService, context: context, activeRole: this.state.activeRole, soloTestMode: this.state.soloTestMode, onToggleSoloTestMode: function () {
                        var nextVal = !_this.state.soloTestMode;
                        try {
                            if (typeof window !== 'undefined' && window.localStorage) {
                                window.localStorage.setItem('sct_solo_test_mode', String(nextVal));
                            }
                        }
                        catch ( /* ignore */_a) { /* ignore */ }
                        _this.setState({ soloTestMode: nextVal });
                    } }));
        }
    };
    AppShell.prototype.render = function () {
        var _this = this;
        var _a = this.props, userDisplayName = _a.userDisplayName, serviceError = _a.serviceError;
        var activeTab = this.state.activeTab;
        var initials = this._getInitials(userDisplayName || 'User');
        var isDisabled = !!serviceError;
        // 1. Use the deduplicated dynamic role options from state (derived from Azure AD & Dataverse)
        var dynamicRoleOptions = (this.state.roleOptions && this.state.roleOptions.length > 0)
            ? this.state.roleOptions
            : [{ key: 'requester', text: 'Requester' }];
        return (React.createElement("div", { className: styles.root, role: "main" },
            isDisabled && (React.createElement("div", { style: {
                    display: 'flex', alignItems: 'flex-start', gap: 12,
                    background: '#FFF4CE', borderBottom: '2px solid #F7C948',
                    padding: '12px 20px', flexShrink: 0,
                } },
                React.createElement(Icon, { iconName: "Warning", style: { fontSize: 18, color: '#6B3A00', marginTop: 2, flexShrink: 0 } }),
                React.createElement("div", null,
                    React.createElement("div", { style: { fontWeight: 700, color: '#6B3A00', fontSize: 13, marginBottom: 2 } }, "Dataverse tables not available"),
                    React.createElement("div", { style: { color: '#6B3A00', fontSize: 12, lineHeight: 1.5 } }, serviceError)))),
            React.createElement("header", { className: styles.header },
                React.createElement("div", { className: styles.headerTop },
                    React.createElement("div", { className: styles.headerBrand },
                        React.createElement("div", { className: styles.headerIcon, "aria-hidden": "true" },
                            React.createElement(Icon, { iconName: "PaymentCard" })),
                        React.createElement("div", null,
                            React.createElement("div", { className: styles.headerTitle }, "Fortrex Spend Control"),
                            React.createElement("div", { style: { fontSize: 11, color: '#94A3B8', fontWeight: 500, marginTop: 1, letterSpacing: '0.2px' } }, "Procurement & Budget Governance Platform"))),
                    this.state.rolesLoading ? (React.createElement("span", { style: { color: 'rgba(255,255,255,0.6)', fontSize: 12, fontStyle: 'italic' } }, "Loading roles\u2026")) : dynamicRoleOptions.length <= 1 ? (React.createElement("span", { style: {
                            color: '#fff', fontSize: 13, padding: '5px 12px',
                            background: 'rgba(255,255,255,0.12)',
                            border: '1px solid rgba(255,255,255,0.3)',
                            borderRadius: 6,
                        } }, "Role: Requester")) : (React.createElement(Dropdown, { options: dynamicRoleOptions, selectedKey: this.state.activeRole, disabled: isDisabled, onRenderOption: this._onRenderRoleOption, onRenderTitle: this._onRenderRoleTitle, onChange: function (_, option) { return _this._handleRoleChange(option); }, ariaLabel: "Switch active role", styles: {
                            root: { minWidth: 180 },
                            title: { borderRadius: 6, fontSize: 13, background: 'rgba(255,255,255,0.12)', color: '#fff', border: '1px solid rgba(255,255,255,0.3)', height: 32, lineHeight: '30px' },
                            caretDownWrapper: { lineHeight: '32px', color: '#fff' },
                            dropdownItem: { fontSize: 13 },
                            dropdownItemSelected: { fontSize: 13 },
                        } })),
                    React.createElement("button", { type: "button", onClick: function () {
                            var nextVal = !_this.state.soloTestMode;
                            try {
                                if (typeof window !== 'undefined' && window.localStorage) {
                                    window.localStorage.setItem('sct_solo_test_mode', String(nextVal));
                                }
                            }
                            catch ( /* ignore */_a) { /* ignore */ }
                            _this.setState({ soloTestMode: nextVal });
                        }, title: this.state.soloTestMode
                            ? 'Solo Test Mode is ON: Self-approval is enabled so you can test all stages alone. Click to enforce strict SoD.'
                            : 'Strict Segregation of Duties: Self-approval is blocked. Click to enable Solo Test Mode.', style: {
                            display: 'inline-flex',
                            alignItems: 'center',
                            gap: 6,
                            padding: '0 12px',
                            height: 32,
                            borderRadius: 16,
                            fontSize: 12,
                            fontWeight: 600,
                            cursor: 'pointer',
                            border: this.state.soloTestMode ? '1px solid #107C10' : '1px solid rgba(255,255,255,0.4)',
                            background: this.state.soloTestMode ? '#DFF6DD' : 'rgba(255,255,255,0.12)',
                            color: this.state.soloTestMode ? '#107C10' : '#ffffff',
                            transition: 'all 0.2s ease',
                        } },
                        React.createElement(Icon, { iconName: this.state.soloTestMode ? 'Unlock' : 'Lock', style: { fontSize: 13 } }),
                        React.createElement("span", null, this.state.soloTestMode ? 'Solo Test: ON' : 'SoD: Strict')),
                    React.createElement("div", { className: styles.headerUser, "aria-label": "Signed in as ".concat(userDisplayName) },
                        React.createElement("div", { className: styles.headerUserAvatar, "aria-hidden": "true" }, initials),
                        React.createElement("span", null, userDisplayName))),
                React.createElement("nav", { className: styles.tabBar, role: "tablist", "aria-label": "Main navigation" }, TABS.filter(function (tab) {
                    var role = _this.state.activeRole;
                    // "Review Queue" not shown for Requester (they have no review responsibilities)
                    if (tab.id === 'reviewQueue' && role === 'requester')
                        return false;
                    return true;
                }).map(function (tab) { return (React.createElement("button", { key: tab.id, role: "tab", "aria-selected": !isDisabled && activeTab === tab.id, "aria-label": tab.ariaLabel, "aria-controls": "tabpanel-".concat(tab.id), id: "tab-".concat(tab.id), disabled: isDisabled, className: [
                        styles.tabItem,
                        !isDisabled && activeTab === tab.id ? styles.tabItemActive : '',
                    ].join(' '), style: isDisabled ? { opacity: 0.4, cursor: 'not-allowed', pointerEvents: 'none' } : undefined, onClick: function () { return !isDisabled && _this._handleTabClick(tab.id); } },
                    React.createElement(Icon, { iconName: tab.iconName, style: { fontSize: 15 }, "aria-hidden": "true" }),
                    tab.label,
                    tab.badge !== undefined && tab.badge > 0 && (React.createElement("span", { className: styles.tabBadge }, tab.badge)))); }))),
            React.createElement("main", { className: styles.content, id: "tabpanel-".concat(activeTab), role: "tabpanel", "aria-labelledby": "tab-".concat(activeTab), "aria-hidden": isDisabled }, this._renderTabContent()),
            React.createElement(Dialog, { hidden: !this.state.showLeaveConfirmDialog, onDismiss: this._cancelLeaveConfirm, dialogContentProps: {
                    type: DialogType.normal,
                    title: 'Unsaved Request in Progress',
                    subText: 'You are currently filling out a new request. If you leave this tab now, your entered information will be discarded. Are you sure you want to leave?',
                }, modalProps: {
                    isBlocking: true,
                    styles: { main: { maxWidth: 440, borderRadius: 8 } },
                } },
                React.createElement(DialogFooter, null,
                    React.createElement(PrimaryButton, { text: "Leave & Discard", onClick: this._confirmLeaveTab, styles: { root: { background: '#D13438', borderColor: '#D13438', borderRadius: 6 } } }),
                    React.createElement(DefaultButton, { text: "Stay & Keep Editing", onClick: this._cancelLeaveConfirm, styles: { root: { borderRadius: 6 } } })))));
    };
    return AppShell;
}(React.Component));
export default AppShell;
//# sourceMappingURL=AppShell.js.map