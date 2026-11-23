import { __assign, __awaiter, __extends, __generator, __spreadArray } from "tslib";
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react/lib/Calendar';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { strVal } from '../../services/DataverseService';
import { LoadingSpinner } from '../common';
import styles from '../shell/AppShell.module.scss';
import { getManager, sendGraphEmail, generateApprovalEmailHTML, getGroupMemberEmails, } from '../../services/GraphMailService';
var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
// ── Azure AD Security Group IDs for Direct Multi-Tier Routing ────────────────
var TIER1_GROUP_ID = 'fdbf07a0-548b-4ccf-875b-900cb1b24a1c'; // SCT-Fortrex_Approvers-Tier1
var TIER2_GROUP_ID = '6a8a4077-c78f-4af2-ae5e-12428d1dcbc3'; // SCT-Fortrex_Approvers-Tier2
var TIER3_GROUP_ID = 'e63dbc4f-6ad2-4969-a518-b229836756af'; // SCT-Fortrex_Approvers-Tier3
// ── Choice option sets ──────────────────────────────────────────────────────────
var BUSINESS_UNIT_OPTIONS = [
    { key: 'PSSI', text: 'PSSI' },
    { key: 'PCI', text: 'PCI' },
    { key: 'SAFE', text: 'SAFE' },
    { key: 'CCCS', text: 'CCCS' },
];
var COST_IS_FOR_OPTIONS = [
    { key: 'This Site', text: 'This Site' },
    { key: 'A Customer Job', text: 'A Customer Job' },
    { key: 'Several Sites', text: 'Several Sites' },
    { key: 'Company-Wide', text: 'Company-Wide' },
];
var PURCHASE_TYPE_OPTIONS = [
    { key: 'Goods', text: 'Goods' },
    { key: 'Repair', text: 'Repair' },
    { key: 'Rental', text: 'Rental' },
    { key: 'Service', text: 'Service' },
    { key: 'Software', text: 'Software' },
    { key: 'Temporary Labor', text: 'Temporary Labor' },
    { key: 'Other', text: 'Other' },
];
var PRICING_BASIS_OPTIONS = [
    { key: 'Fixed Price', text: 'Fixed Price' },
    { key: 'Rate Card', text: 'Rate Card' },
    { key: 'Time and Materials with Cap', text: 'Time and Materials with Cap' },
    { key: 'Usage-Based Estimate', text: 'Usage-Based Estimate' },
];
var URGENCY_OPTIONS = [
    { key: 'Standard', text: 'Standard' },
    { key: 'Urgent', text: 'Urgent' },
    { key: 'Emergency', text: 'Emergency' },
];
// ── Helpers ────────────────────────────────────────────────────────────────────
// ── Dataverse choice-field integer mappings ────────────────────────────────────
// Choice fields in Dataverse are Edm.Int32. Publisher prefix sct_ uses base 854000000.
// Confirmed from actual Dataverse record: sct_urgency=854000000→"Standard",
// sct_purchase_type=854000000→"Goods", sct_business_unit=854000000→"PSSI", etc.
var URGENCY_INT = {
    'Standard': 854000000,
    'Urgent': 854000001,
    'Emergency': 854000002,
};
var PURCHASE_TYPE_INT = {
    'Goods': 854000000,
    'Repair': 854000001,
    'Rental': 854000002,
    'Service': 854000003,
    'Software': 854000004,
    'Temporary Labor': 854000005,
    'Other': 854000006,
};
var PRICING_BASIS_INT = {
    'Fixed Price': 854000000,
    'Rate Card': 854000001,
    'Time and Materials with Cap': 854000002,
    'Usage-Based Estimate': 854000003,
};
var COST_IS_FOR_INT = {
    'This Site': 854000000,
    'A Customer Job': 854000001,
    'Several Sites': 854000002,
    'Company-Wide': 854000003,
};
var BUSINESS_UNIT_INT = {
    'PSSI': 854000000,
    'PCI': 854000001,
    'SAFE': 854000002,
    'CCCS': 854000003,
};
/** Returns the integer option-set value if known; otherwise undefined (field not sent). */
function choiceInt(map, label) {
    return map[label];
}
function readFileAsBase64(file) {
    return new Promise(function (resolve, reject) {
        var reader = new FileReader();
        reader.onload = function () {
            var _a;
            var result = reader.result;
            resolve((_a = result.split(',')[1]) !== null && _a !== void 0 ? _a : result);
        };
        reader.onerror = function () { return reject(new Error("Failed to read ".concat(file.name))); };
        reader.readAsDataURL(file);
    });
}
/** Edm.Date — Dataverse DateOnly columns require strict YYYY-MM-DD, no time component.
 *  Uses local date components to avoid UTC midnight shift (e.g. Oct 31 → Oct 30).
 *  Avoids padStart — not available in the project's TS lib target. */
function dateOnly(d) {
    if (!d)
        return undefined;
    var y = d.getFullYear();
    var mo = d.getMonth() + 1;
    var da = d.getDate();
    var mm = mo < 10 ? ('0' + mo) : String(mo);
    var dd = da < 10 ? ('0' + da) : String(da);
    return y + '-' + mm + '-' + dd;
}
var Section = function (_a) {
    var iconName = _a.iconName, iconBg = _a.iconBg, iconColor = _a.iconColor, title = _a.title, badge = _a.badge, children = _a.children;
    return (React.createElement("div", { className: styles.card, style: { marginBottom: 16 } },
        React.createElement("div", { className: styles.cardHeader },
            React.createElement("div", { className: styles.cardTitle },
                React.createElement("div", { className: styles.cardTitleIcon, style: { background: iconBg }, "aria-hidden": "true" },
                    React.createElement(Icon, { iconName: iconName, style: { color: iconColor, fontSize: 16 } })),
                title,
                badge && (React.createElement("span", { style: { marginLeft: 10, background: iconBg, color: iconColor, fontSize: 11, fontWeight: 700, padding: '2px 10px', borderRadius: 12, letterSpacing: '0.3px' } }, badge)))),
        React.createElement("div", { className: styles.cardBody }, children)));
};
var fieldStyles = { root: { marginBottom: 16 }, fieldGroup: { borderRadius: 6 } };
var dropdownStyles = { root: { marginBottom: 16 }, dropdown: { borderRadius: 6 } };
var checkStyles = { root: { marginBottom: 14 } };
// ── Row layout helper ─────────────────────────────────────────────────────────
var Row = function (_a) {
    var children = _a.children, _b = _a.cols, cols = _b === void 0 ? 2 : _b;
    return (React.createElement("div", { style: { display: 'grid', gridTemplateColumns: "repeat(".concat(cols, ", 1fr)"), gap: '0 24px' } }, children));
};
function calcApprovalGuidance(isNewSupplier, amount) {
    if (isNewSupplier) {
        return {
            levels: ['L1 — Budget Owner', 'L2 — Senior Manager', 'L3 — ELT Leader'],
            color: '#A4262C', bg: '#FDE7E9', iconName: 'Warning',
        };
    }
    if (amount <= 2500) {
        return {
            levels: ['L1 — Budget Owner'],
            color: '#107C10', bg: '#DFF6DD', iconName: 'CheckMark',
        };
    }
    if (amount <= 10000) {
        return {
            levels: ['L2 — Senior Manager'],
            color: '#905400', bg: '#FFF4CE', iconName: 'Info',
        };
    }
    return {
        levels: ['L3 — ELT Leader'],
        color: '#A4262C', bg: '#FDE7E9', iconName: 'Warning',
    };
}
// ── Main component ─────────────────────────────────────────────────────────────
var NewRequestForm = /** @class */ (function (_super) {
    __extends(NewRequestForm, _super);
    function NewRequestForm(props) {
        var _this = _super.call(this, props) || this;
        _this._service = null;
        _this._fileInputRef = React.createRef();
        _this._handleBeforeUnload = function (e) {
            if (_this._isDirty(_this.state)) {
                e.preventDefault();
                e.returnValue = '';
            }
        };
        _this.state = {
            sites: [], suppliers: [],
            loadingMeta: false, metaError: null, sitesError: null,
            requesterGuid: null, requesterName: '', requesterEmail: '', userLookupError: null,
            managerGuid: null, managerName: '', managerEmail: '', managerNavProp: 'cr2a4_Manager',
            businessUnit: '', selectedSiteId: '', selectedSiteName: '',
            autoCostCentre: '', autoDivision: '', autoBudgetOwner: '',
            sctFunction: '', costIsFor: '',
            requestTitle: '', description: '', purchaseType: '',
            isNewSupplier: false, selectedSupplierId: '',
            newSupplierName: '', newSupplierLegalName: '',
            newSupplierContactName: '', newSupplierContactEmail: '',
            newSupplierContactPhone: '', newSupplierW9Notes: '',
            estimatedAmount: '', termStartDate: null, termEndDate: null,
            pricingBasis: '', dateNeededBy: null,
            urgency: '', urgencyReason: '', emergencyDeclaredBy: '',
            deliverTo: '', receivingContact: '', equipmentAssetTag: '',
            costRechargedToCustomer: false, customerName: '',
            customerAgreementAttached: false, techOrDataInvolved: false,
            supplierOnPremises: false, possibleCapitalPurchase: false,
            isConfidential: false, isRetrospective: false,
            declarationNoPriorCommitment: false,
            declarationNoConflictOfInterest: false,
            attachments: [], uploadingFiles: false,
            submitting: false, isSubmitting: false, savingDraft: false,
            submitError: null, submitSuccess: null, savedId: null,
            errors: {},
            currentStep: 0,
        };
        return _this;
    }
    NewRequestForm.prototype.componentDidMount = function () {
        this._service = this.props.dataverseService;
        void this._loadMeta();
        if (typeof window !== 'undefined') {
            window.addEventListener('beforeunload', this._handleBeforeUnload);
        }
    };
    NewRequestForm.prototype.componentWillUnmount = function () {
        if (typeof window !== 'undefined') {
            window.removeEventListener('beforeunload', this._handleBeforeUnload);
        }
        if (this.props.onDirtyChange) {
            this.props.onDirtyChange(false);
        }
    };
    NewRequestForm.prototype.componentDidUpdate = function (prevProps, prevState) {
        if (this.props.onDirtyChange) {
            var wasDirty = this._isDirty(prevState);
            var isNowDirty = this._isDirty(this.state);
            if (wasDirty !== isNowDirty) {
                this.props.onDirtyChange(isNowDirty);
            }
        }
    };
    NewRequestForm.prototype._isDirty = function (s) {
        if (s.submitSuccess || s.savedId)
            return false;
        if (s.currentStep > 0)
            return true;
        if ((s.requestTitle || '').trim().length > 0)
            return true;
        if ((s.description || '').trim().length > 0)
            return true;
        if ((s.purchaseType || '').length > 0)
            return true;
        if ((s.businessUnit || '').length > 0)
            return true;
        if ((s.selectedSiteId || '').length > 0)
            return true;
        if ((s.estimatedAmount || '').trim().length > 0)
            return true;
        if ((s.urgency || '').length > 0)
            return true;
        if ((s.newSupplierName || '').trim().length > 0)
            return true;
        if ((s.selectedSupplierId || '').length > 0)
            return true;
        if ((s.attachments || []).length > 0)
            return true;
        if ((s.deliverTo || '').trim().length > 0)
            return true;
        if ((s.receivingContact || '').trim().length > 0)
            return true;
        if ((s.customerName || '').trim().length > 0)
            return true;
        if ((s.equipmentAssetTag || '').trim().length > 0)
            return true;
        if ((s.sctFunction || '').trim().length > 0)
            return true;
        if ((s.costIsFor || '').length > 0)
            return true;
        return false;
    };
    /** Resolves workflow routing based on submitter's active authority / role */
    NewRequestForm.prototype._getRoutingConfig = function () {
        var role = (this.props.activeRole || 'requester').toLowerCase().trim();
        if (role === 'manager' || role.indexOf('manager') !== -1) {
            return {
                submittingAs: 'Manager',
                targetPendingTier: 'Tier 1',
                targetStatus: 854000003, // In Review
                skippedStagesDesc: 'Line Manager stage auto-bypassed (Submitting as Department Manager)',
                routeDisplay: 'Tier 1 → Tier 2 → Tier 3 → PO Issued',
                targetGroupId: TIER1_GROUP_ID,
            };
        }
        if (role === 'approver_tier1' || role.indexOf('tier 1') !== -1 || role.indexOf('tier1') !== -1) {
            return {
                submittingAs: 'Tier 1 Approver',
                targetPendingTier: 'Tier 2',
                targetStatus: 854000003, // In Review
                skippedStagesDesc: 'Line Manager & Tier 1 stages auto-bypassed (Submitting as Tier 1 Approver)',
                routeDisplay: 'Tier 2 → Tier 3 → PO Issued',
                targetGroupId: TIER2_GROUP_ID,
            };
        }
        if (role === 'approver_tier2' || role.indexOf('tier 2') !== -1 || role.indexOf('tier2') !== -1) {
            return {
                submittingAs: 'Tier 2 Approver',
                targetPendingTier: 'Tier 3',
                targetStatus: 854000003, // In Review
                skippedStagesDesc: 'Line Manager, Tier 1 & Tier 2 stages auto-bypassed (Submitting as Tier 2 Approver)',
                routeDisplay: 'Tier 3 → PO Issued',
                targetGroupId: TIER3_GROUP_ID,
            };
        }
        if (role === 'approver_tier3' || role.indexOf('tier 3') !== -1 || role.indexOf('tier3') !== -1) {
            return {
                submittingAs: 'Tier 3 Approver',
                targetPendingTier: 'Tier 3',
                targetStatus: 854000003, // In Review
                skippedStagesDesc: 'Executive Request (Submitting as Tier 3 Approver)',
                routeDisplay: 'Peer Tier 3 Review → PO Issued',
                targetGroupId: TIER3_GROUP_ID,
            };
        }
        // Standard Requester
        return {
            submittingAs: 'Requester',
            targetPendingTier: 'Manager',
            targetStatus: 854000002, // Submitted
            skippedStagesDesc: 'Standard routing to Line Manager',
            routeDisplay: 'Line Manager → Tier 1 → Tier 2 → Tier 3 → PO Issued',
        };
    };
    // ── Data loading ─────────────────────────────────────────────────────────────
    NewRequestForm.prototype._loadMeta = function () {
        return __awaiter(this, void 0, void 0, function () {
            var ctx_1, aadId_1, spEmail_1, resolveUser, _a, sites, suppliers, userResult, managerNavProp, managerGuid, managerName, managerEmail, graphClient, managerData, mgr, mgr, graphErr_1, u, _b, _c, sysMgr, initialUpdate_1, rec_1, id, siteId_1, foundSite, err_1;
            var _this = this;
            var _d, _e, _f, _g, _h, _j, _k, _l, _m, _o;
            return __generator(this, function (_p) {
                switch (_p.label) {
                    case 0:
                        if (!this._service)
                            return [2 /*return*/];
                        this.setState({ loadingMeta: true, metaError: null, sitesError: null, userLookupError: null });
                        _p.label = 1;
                    case 1:
                        _p.trys.push([1, 24, , 25]);
                        ctx_1 = this.props.context;
                        aadId_1 = (_f = (_e = (_d = ctx_1 === null || ctx_1 === void 0 ? void 0 : ctx_1.pageContext) === null || _d === void 0 ? void 0 : _d.aadInfo) === null || _e === void 0 ? void 0 : _e.userId) === null || _f === void 0 ? void 0 : _f.toString();
                        spEmail_1 = (_h = (_g = ctx_1 === null || ctx_1 === void 0 ? void 0 : ctx_1.pageContext) === null || _g === void 0 ? void 0 : _g.user) === null || _h === void 0 ? void 0 : _h.email;
                        resolveUser = function () { return __awaiter(_this, void 0, void 0, function () {
                            var u, u;
                            var _a, _b, _c;
                            return __generator(this, function (_d) {
                                switch (_d.label) {
                                    case 0:
                                        if (!aadId_1) return [3 /*break*/, 2];
                                        return [4 /*yield*/, this._service.getCurrentSystemUser(aadId_1)];
                                    case 1:
                                        u = _d.sent();
                                        if (u)
                                            return [2 /*return*/, { guid: u.systemuserid, name: u.fullname, email: u.internalemailaddress, err: null }];
                                        _d.label = 2;
                                    case 2:
                                        if (!spEmail_1) return [3 /*break*/, 4];
                                        return [4 /*yield*/, this._service.getSystemUserByEmail(spEmail_1)];
                                    case 3:
                                        u = _d.sent();
                                        if (u)
                                            return [2 /*return*/, { guid: u.systemuserid, name: u.fullname, email: u.internalemailaddress, err: null }];
                                        _d.label = 4;
                                    case 4: return [2 /*return*/, {
                                            guid: null,
                                            name: (_c = (_b = (_a = ctx_1 === null || ctx_1 === void 0 ? void 0 : ctx_1.pageContext) === null || _a === void 0 ? void 0 : _a.user) === null || _b === void 0 ? void 0 : _b.displayName) !== null && _c !== void 0 ? _c : '',
                                            email: spEmail_1 !== null && spEmail_1 !== void 0 ? spEmail_1 : '',
                                            err: 'Could not resolve your Dataverse user record — requester will not be bound.',
                                        }];
                                }
                            });
                        }); };
                        return [4 /*yield*/, Promise.all([
                                this._service.getSitesRaw(true).catch(function (e) {
                                    var msg = e instanceof Error ? e.message : 'Failed to load sites';
                                    _this.setState({ sitesError: msg });
                                    return [];
                                }),
                                this._service.getSuppliers(true).catch(function () { return []; }),
                                resolveUser(),
                                // cr2a4_manager = logical name of the Manager lookup (publisher prefix cr2a4_).
                                // getNavPropName returns: (1) relationship nav prop, (2) attribute SchemaName 'cr2a4_Manager',
                                // (3) the logical name as last resort. The .catch() hardcodes the known schema name.
                                this._service.getNavPropName('sct_request', 'cr2a4_manager').catch(function () { return 'cr2a4_Manager'; }),
                            ])];
                    case 2:
                        _a = _p.sent(), sites = _a[0], suppliers = _a[1], userResult = _a[2], managerNavProp = _a[3];
                        managerGuid = null;
                        managerName = '';
                        managerEmail = '';
                        _p.label = 3;
                    case 3:
                        _p.trys.push([3, 12, , 13]);
                        return [4 /*yield*/, this.props.context.msGraphClientFactory.getClient('3')];
                    case 4:
                        graphClient = _p.sent();
                        return [4 /*yield*/, graphClient.api('/me/manager').version('v1.0').get().catch(function () { return null; })];
                    case 5:
                        managerData = _p.sent();
                        if (!managerData) return [3 /*break*/, 11];
                        // Extract manager email (check mail or userPrincipalName)
                        managerEmail = (managerData.mail || managerData.userPrincipalName || '').trim();
                        managerName = managerData.displayName || managerEmail;
                        if (!(managerEmail && this._service)) return [3 /*break*/, 7];
                        return [4 /*yield*/, this._service.getSystemUserIdByEmail(managerEmail).catch(function () { return null; })];
                    case 6:
                        managerGuid = _p.sent();
                        _p.label = 7;
                    case 7:
                        if (!(!managerGuid && managerData.id && this._service)) return [3 /*break*/, 9];
                        return [4 /*yield*/, this._service.getCurrentSystemUser(managerData.id).catch(function () { return null; })];
                    case 8:
                        mgr = _p.sent();
                        managerGuid = (_j = mgr === null || mgr === void 0 ? void 0 : mgr.systemuserid) !== null && _j !== void 0 ? _j : null;
                        _p.label = 9;
                    case 9:
                        if (!(!managerGuid && managerEmail && this._service)) return [3 /*break*/, 11];
                        return [4 /*yield*/, this._service.getSystemUserByEmail(managerEmail).catch(function () { return null; })];
                    case 10:
                        mgr = _p.sent();
                        managerGuid = (_k = mgr === null || mgr === void 0 ? void 0 : mgr.systemuserid) !== null && _k !== void 0 ? _k : null;
                        _p.label = 11;
                    case 11: return [3 /*break*/, 13];
                    case 12:
                        graphErr_1 = _p.sent();
                        console.warn('[NewRequestForm] Graph /me/manager lookup warning:', graphErr_1);
                        return [3 /*break*/, 13];
                    case 13:
                        if (!(!managerGuid && userResult.guid && this._service)) return [3 /*break*/, 23];
                        return [4 /*yield*/, this._service.getCurrentSystemUser(aadId_1 !== null && aadId_1 !== void 0 ? aadId_1 : '').catch(function () { return null; })];
                    case 14:
                        if (!((_l = _p.sent()) !== null && _l !== void 0)) return [3 /*break*/, 15];
                        _b = _l;
                        return [3 /*break*/, 19];
                    case 15:
                        if (!spEmail_1) return [3 /*break*/, 17];
                        return [4 /*yield*/, this._service.getSystemUserByEmail(spEmail_1).catch(function () { return null; })];
                    case 16:
                        _c = _p.sent();
                        return [3 /*break*/, 18];
                    case 17:
                        _c = null;
                        _p.label = 18;
                    case 18:
                        _b = (_c);
                        _p.label = 19;
                    case 19:
                        u = _b;
                        managerGuid = (_m = u === null || u === void 0 ? void 0 : u._parentsystemuserid_value) !== null && _m !== void 0 ? _m : null;
                        if (!!managerGuid) return [3 /*break*/, 21];
                        return [4 /*yield*/, this._service.getManagerSystemUserId(userResult.guid).catch(function () { return null; })];
                    case 20:
                        managerGuid = _p.sent();
                        _p.label = 21;
                    case 21:
                        if (!managerGuid) return [3 /*break*/, 23];
                        return [4 /*yield*/, this._service.getSystemUserById(managerGuid).catch(function () { return null; })];
                    case 22:
                        sysMgr = _p.sent();
                        if (sysMgr) {
                            if (!managerName)
                                managerName = sysMgr.fullname;
                            if (!managerEmail)
                                managerEmail = sysMgr.internalemailaddress;
                        }
                        _p.label = 23;
                    case 23:
                        initialUpdate_1 = {};
                        rec_1 = this.props.initialRecord;
                        if (rec_1) {
                            id = String((_o = rec_1['sct_requestid']) !== null && _o !== void 0 ? _o : '');
                            siteId_1 = strVal(rec_1, '_sct_site_value') || strVal(rec_1, 'sct_site');
                            foundSite = sites.find(function (s) { return _this._siteId(s) === siteId_1; });
                            initialUpdate_1 = {
                                savedId: id || null,
                                requestTitle: strVal(rec_1, 'sct_request_title') || strVal(rec_1, 'sct_title'),
                                description: strVal(rec_1, 'sct_description'),
                                businessUnit: strVal(rec_1, 'sct_business_unit'),
                                selectedSiteId: siteId_1,
                                selectedSiteName: foundSite ? this._siteLabel(foundSite) : '',
                                autoCostCentre: strVal(rec_1, 'sct_cost_center') || (foundSite ? strVal(foundSite, 'sct_costcenter') : ''),
                                autoDivision: strVal(rec_1, 'sct_division') || (foundSite ? strVal(foundSite, 'sct_region') : ''),
                                autoBudgetOwner: strVal(rec_1, 'sct_budgetowner') || (foundSite ? strVal(foundSite, '_sct_budgetownerid_value') : ''),
                                purchaseType: strVal(rec_1, 'sct_purchase_type'),
                                isNewSupplier: Boolean(rec_1['sct_is_new_supplier']),
                                selectedSupplierId: strVal(rec_1, '_sct_suggested_supplier_value'),
                                newSupplierName: strVal(rec_1, 'sct_proposed_supplier_name'),
                                estimatedAmount: strVal(rec_1, 'sct_estimated_total_value'),
                                urgency: strVal(rec_1, 'sct_urgency'),
                                urgencyReason: strVal(rec_1, 'sct_urgency_reason'),
                                equipmentAssetTag: strVal(rec_1, 'sct_equipment_asset_tag'),
                                costRechargedToCustomer: Boolean(rec_1['sct_cost_recharged_to_customer']),
                                customerName: strVal(rec_1, 'sct_customer_name'),
                                declarationNoPriorCommitment: Boolean(rec_1['sct_declaration_no_prior_commitment']),
                                declarationNoConflictOfInterest: Boolean(rec_1['sct_declaration_no_conflict_of_interest']),
                            };
                        }
                        this.setState({
                            sites: sites,
                            suppliers: suppliers,
                            loadingMeta: false,
                            requesterGuid: userResult.guid,
                            requesterName: userResult.name,
                            requesterEmail: userResult.email,
                            userLookupError: userResult.err,
                            managerGuid: managerGuid,
                            managerName: managerName,
                            managerEmail: managerEmail,
                            managerNavProp: (managerNavProp && managerNavProp.toLowerCase() !== 'sct_manager') ? managerNavProp : 'cr2a4_Manager',
                        }, function () {
                            if (rec_1) {
                                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                _this.setState(initialUpdate_1);
                            }
                        });
                        return [3 /*break*/, 25];
                    case 24:
                        err_1 = _p.sent();
                        this.setState({
                            metaError: err_1 instanceof Error ? err_1.message : 'Failed to load form data',
                            loadingMeta: false,
                        });
                        return [3 /*break*/, 25];
                    case 25: return [2 /*return*/];
                }
            });
        });
    };
    // ── Site selection — auto-fill cost centre / division / budget owner ──────────
    /** Resolve site display name — PrimaryNameAttribute confirmed as sct_site_name */
    NewRequestForm.prototype._siteLabel = function (site) {
        var _a;
        return strVal(site, 'sct_site_name') ||
            strVal(site, 'sct_name') ||
            strVal(site, 'sct_sitename') ||
            strVal(site, 'sct_sitecode') ||
            String((_a = site['sct_sites_masterid']) !== null && _a !== void 0 ? _a : '');
    };
    /** Resolve site primary key — PrimaryIdAttribute confirmed as sct_sites_masterid */
    NewRequestForm.prototype._siteId = function (site) {
        var _a, _b;
        return String((_b = (_a = site['sct_sites_masterid']) !== null && _a !== void 0 ? _a : site['sct_sitesmasterid']) !== null && _b !== void 0 ? _b : '');
    };
    NewRequestForm.prototype._onSiteSelected = function (siteId) {
        var _this = this;
        var site = this.state.sites.find(function (s) { return _this._siteId(s) === siteId; });
        if (!site)
            return;
        var rawBudgetOwner = strVal(site, '_sct_budgetownerid_value') ||
            strVal(site, 'sct_budgetowner') ||
            strVal(site, '_sct_budgetowner_value');
        this.setState({
            selectedSiteId: siteId,
            selectedSiteName: this._siteLabel(site),
            autoCostCentre: strVal(site, 'sct_costcenter') || strVal(site, 'sct_cost_center'),
            autoDivision: strVal(site, 'sct_region') || strVal(site, 'sct_division'),
            autoBudgetOwner: rawBudgetOwner,
            businessUnit: this.state.businessUnit || strVal(site, 'sct_businessunit'),
        });
        // If budget owner is a raw GUID, resolve display name from Dataverse systemusers
        var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
        if (rawBudgetOwner && GUID_RE.test(rawBudgetOwner) && this._service) {
            void this._service.getSystemUserById(rawBudgetOwner).then(function (user) {
                if (user && user.fullname) {
                    _this.setState({ autoBudgetOwner: user.fullname });
                }
            }).catch(function () { });
        }
    };
    // ── Validation ────────────────────────────────────────────────────────────────
    NewRequestForm.prototype._validate = function () {
        var s = this.state;
        var errs = {};
        // ── Section 1: Header ─────────────────────────────────────────────────────
        if (!s.businessUnit)
            errs['businessUnit'] = 'Business unit is required.';
        if (!s.selectedSiteId)
            errs['selectedSiteId'] = 'Site / Location is required.';
        // ── Section 2: Requirement ────────────────────────────────────────────────
        if (!s.description.trim())
            errs['description'] = 'Description / Justification is required.';
        if (!s.purchaseType)
            errs['purchaseType'] = 'Purchase type is required.';
        if (s.isNewSupplier && !s.newSupplierName.trim())
            errs['newSupplierName'] = 'Proposed supplier name is required for new supplier requests.';
        // ── Section 3: Financials ─────────────────────────────────────────────────
        var rawAmt = s.estimatedAmount ? String(s.estimatedAmount).trim() : '';
        var cleanAmt = rawAmt.replace(/[$,]/g, '').trim();
        var isValidPositiveNumber = /^[0-9]+(\.[0-9]{1,2})?$/.test(cleanAmt);
        var amt = parseFloat(cleanAmt);
        if (!rawAmt || isNaN(amt) || amt <= 0 || !isValidPositiveNumber)
            errs['estimatedAmount'] = 'Amount must be greater than zero and contain only valid positive numbers.';
        if (!s.dateNeededBy) {
            errs['dateNeededBy'] = 'Date needed by is required.';
        }
        else {
            var today = new Date();
            today.setHours(0, 0, 0, 0);
            var neededDate = new Date(s.dateNeededBy);
            neededDate.setHours(0, 0, 0, 0);
            if (neededDate < today && !s.isRetrospective) {
                errs['dateNeededBy'] = 'Date needed by cannot be in the past (unless Retrospective Spend is declared).';
            }
        }
        if (!s.urgency)
            errs['urgency'] = 'Urgency is required.';
        if ((s.urgency === 'Urgent' || s.urgency === 'Emergency') && !s.urgencyReason.trim())
            errs['urgencyReason'] = 'Urgency reason is required for Urgent / Emergency requests.';
        if (s.urgency === 'Emergency' && !s.emergencyDeclaredBy.trim())
            errs['emergencyDeclaredBy'] = 'The person declaring the emergency must be named.';
        // ── Section 4: Logistics ──────────────────────────────────────────────────
        if ((s.purchaseType === 'Repair' || s.purchaseType === 'Rental') && !s.equipmentAssetTag.trim())
            errs['equipmentAssetTag'] = 'Asset tag / serial number is required for Repair or Rental.';
        if (s.costRechargedToCustomer && !s.customerName.trim())
            errs['customerName'] = 'Customer name is required when cost is recharged.';
        if (s.costRechargedToCustomer && !s.customerAgreementAttached)
            errs['customerAgreementAttached'] = 'Customer agreement or quote must be attached when cost is recharged to a customer.';
        // ── Section 5: Declarations (R1 + R10) ───────────────────────────────────
        if (!s.declarationNoPriorCommitment)
            errs['declarationNoPriorCommitment'] = 'You must confirm no prior commitment has been made (Rule R1).';
        if (!s.declarationNoConflictOfInterest)
            errs['declarationNoConflictOfInterest'] = 'You must declare no conflict of interest (Rule R10).';
        // ── Section 6: Approval evidence attachment ───────────────────────────────
        var validAmt = !isNaN(amt) && amt > 0 && isValidPositiveNumber;
        if ((validAmt || s.isNewSupplier) && s.attachments.length === 0) {
            var guidance = calcApprovalGuidance(s.isNewSupplier, validAmt ? amt : 0);
            errs['attachments'] = s.isNewSupplier
                ? 'Pre-approval evidence from L1, L2 and L3 approvers must be attached for new supplier requests.'
                : 'Written pre-approval evidence from ' + guidance.levels.join(' + ') + ' must be attached before submission.';
        }
        if (Object.keys(errs).length > 0) {
            var firstErrKey = errs['estimatedAmount'] ? 'estimatedAmount' : Object.keys(errs)[0];
            var firstErrMsg = errs[firstErrKey];
            this.setState({ errors: errs, submitError: "Validation error: ".concat(firstErrMsg) });
            return false;
        }
        this.setState({ errors: errs, submitError: null });
        return true;
    };
    // ── Payload builder ───────────────────────────────────────────────────────────
    NewRequestForm.prototype._buildPayload = function (isDraft) {
        var s = this.state;
        var payload = {};
        // ── Text fields — names confirmed from actual Dataverse record ────────────
        // NOTE: sct_request has NO sct_name field. PrimaryNameAttribute is sct_sct_reference
        // (autonumber, auto-generated by Dataverse — never include it in the payload).
        // Build description — append new-supplier details inline (no dedicated schema fields for them)
        var descParts = [];
        if (s.description.trim())
            descParts.push(s.description.trim());
        if (s.isNewSupplier && s.newSupplierName.trim()) {
            descParts.push('[New Supplier] ' + s.newSupplierName.trim() +
                (s.newSupplierLegalName.trim() ? ' | Legal: ' + s.newSupplierLegalName.trim() : '') +
                (s.newSupplierContactName.trim() ? ' | Contact: ' + s.newSupplierContactName.trim() : '') +
                (s.newSupplierContactEmail.trim() ? ' <' + s.newSupplierContactEmail.trim() + '>' : '') +
                (s.newSupplierContactPhone.trim() ? ' ' + s.newSupplierContactPhone.trim() : '') +
                (s.newSupplierW9Notes.trim() ? ' | W-9: ' + s.newSupplierW9Notes.trim() : ''));
        }
        if (descParts.length)
            payload['sct_description'] = descParts.join('\n\n');
        if (s.sctFunction.trim())
            payload['sct_function'] = s.sctFunction.trim();
        if (s.urgencyReason.trim())
            payload['sct_urgency_reason'] = s.urgencyReason.trim();
        if (s.emergencyDeclaredBy.trim())
            payload['sct_emergency_declared_by_email'] = s.emergencyDeclaredBy.trim();
        if (s.deliverTo.trim())
            payload['sct_delivery_address'] = s.deliverTo.trim();
        if (s.equipmentAssetTag.trim())
            payload['sct_equipment_asset_tag'] = s.equipmentAssetTag.trim();
        if (s.customerName.trim())
            payload['sct_customer_name'] = s.customerName.trim();
        if (s.autoCostCentre)
            payload['sct_cost_center'] = s.autoCostCentre;
        // ── Choice fields → Edm.Int32, base 854000000 (confirmed from Dataverse record) ──
        var buInt = choiceInt(BUSINESS_UNIT_INT, s.businessUnit);
        if (buInt !== undefined)
            payload['sct_business_unit'] = buInt;
        var ciInt = choiceInt(COST_IS_FOR_INT, s.costIsFor);
        if (ciInt !== undefined)
            payload['sct_cost_is_for'] = ciInt;
        var ptInt = choiceInt(PURCHASE_TYPE_INT, s.purchaseType);
        if (ptInt !== undefined)
            payload['sct_purchase_type'] = ptInt;
        var pbInt = choiceInt(PRICING_BASIS_INT, s.pricingBasis);
        if (pbInt !== undefined)
            payload['sct_pricing_basis'] = pbInt;
        var ugInt = choiceInt(URGENCY_INT, s.urgency);
        if (ugInt !== undefined)
            payload['sct_urgency'] = ugInt;
        // ── Currency / numeric ────────────────────────────────────────────────────
        var amt = parseFloat(s.estimatedAmount);
        if (!isNaN(amt))
            payload['sct_estimated_total_value'] = amt;
        // ── Dates — Edm.Date fields require YYYY-MM-DD only, no time/timezone component ──
        var termStart = dateOnly(s.termStartDate);
        var termEnd = dateOnly(s.termEndDate);
        var needed = dateOnly(s.dateNeededBy);
        if (termStart)
            payload['sct_term_start_date'] = termStart;
        if (termEnd)
            payload['sct_term_end_date'] = termEnd;
        if (needed)
            payload['sct_date_needed_by'] = needed;
        // ── Booleans — names confirmed from actual Dataverse record ───────────────
        payload['sct_is_new_supplier'] = s.isNewSupplier;
        payload['sct_cost_recharged_to_customer'] = s.costRechargedToCustomer;
        payload['sct_customer_agreement_attached'] = s.customerAgreementAttached;
        payload['sct_tech_or_data_involved'] = s.techOrDataInvolved;
        payload['sct_supplier_on_premises'] = s.supplierOnPremises;
        payload['sct_possible_capital_purchase'] = s.possibleCapitalPurchase;
        payload['sct_is_confidential'] = s.isConfidential;
        payload['sct_is_retrospective'] = s.isRetrospective;
        payload['sct_declaration_no_prior_commitment'] = s.declarationNoPriorCommitment;
        payload['sct_declaration_no_conflict_of_interest'] = s.declarationNoConflictOfInterest;
        // ── Lookup nav-prop bindings ──────────────────────────────────────────────────
        //
        // Rules for every @odata.bind property:
        //   1. Only include if the GUID is a valid, non-empty GUID string.
        //   2. Never send a null/undefined value — omit the key entirely.
        //   3. Never send the bare logical name (lowercase) as the property key —
        //      always use the schema/nav-prop name so Dataverse doesn't 400.
        //
        var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
        // Requester — sct_Requester nav prop (schema name 'sct_Requester')
        if (s.requesterGuid && GUID_RE.test(s.requesterGuid))
            payload['sct_Requester@odata.bind'] = "/systemusers(".concat(s.requesterGuid, ")");
        // Site — sct_Site nav prop
        if (s.selectedSiteId && GUID_RE.test(s.selectedSiteId))
            payload['sct_Site@odata.bind'] = "/sct_sites_masters(".concat(s.selectedSiteId, ")");
        // Suggested Supplier — sct_Suggested_Supplier nav prop
        if (!s.isNewSupplier && s.selectedSupplierId && GUID_RE.test(s.selectedSupplierId))
            payload['sct_Suggested_Supplier@odata.bind'] = "/sct_suppliers(".concat(s.selectedSupplierId, ")");
        // Manager — automatically set manager lookup field using cr2a4_manager@odata.bind
        if (s.managerGuid && GUID_RE.test(s.managerGuid)) {
            payload['cr2a4_manager@odata.bind'] = "/systemusers(".concat(s.managerGuid, ")");
            var navProp = (s.managerNavProp && s.managerNavProp.toLowerCase() !== 'sct_manager')
                ? s.managerNavProp
                : 'cr2a4_Manager';
            if (navProp !== 'cr2a4_manager') {
                payload["".concat(navProp, "@odata.bind")] = "/systemusers(".concat(s.managerGuid, ")");
            }
        }
        // Canonical status & pending tier (resolved via dynamic routing based on submitter):
        var routing = this._getRoutingConfig();
        payload['sct_request_status'] = isDraft ? 854000001 : routing.targetStatus;
        payload['cr2a4_pendingtier'] = isDraft ? null : routing.targetPendingTier;
        return payload;
    };
    // ── Submit ────────────────────────────────────────────────────────────────────
    NewRequestForm.prototype._submit = function (isDraft) {
        return __awaiter(this, void 0, void 0, function () {
            var key, payload, newId_1, isExisting, submitErr_1, errMsg, cleanPayload_1, routing_1, auditDesc, auditPayload, auditErr_1, routing_2, submitMessage, err_2;
            var _a, _b, _c;
            var _this = this;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        if (!isDraft && !this._validate())
                            return [2 /*return*/];
                        if (!this._service)
                            return [2 /*return*/];
                        key = isDraft ? 'savingDraft' : 'submitting';
                        this.setState((_a = {}, _a[key] = true, _a.isSubmitting = true, _a.submitError = null, _a));
                        _d.label = 1;
                    case 1:
                        _d.trys.push([1, 17, , 18]);
                        payload = this._buildPayload(isDraft);
                        isExisting = Boolean(this.state.savedId && GUID_RE.test(this.state.savedId));
                        if (!isExisting) return [3 /*break*/, 3];
                        newId_1 = this.state.savedId;
                        return [4 /*yield*/, this._service.updateRequest(newId_1, payload)];
                    case 2:
                        _d.sent();
                        return [3 /*break*/, 9];
                    case 3:
                        _d.trys.push([3, 5, , 9]);
                        return [4 /*yield*/, this._service.createRequest(payload)];
                    case 4:
                        newId_1 = _d.sent();
                        return [3 /*break*/, 9];
                    case 5:
                        submitErr_1 = _d.sent();
                        errMsg = submitErr_1 instanceof Error ? submitErr_1.message : String(submitErr_1);
                        if (!(errMsg.indexOf('manager') !== -1 || errMsg.indexOf('cr2a4_Manager') !== -1 || errMsg.indexOf('sct_Manager') !== -1)) return [3 /*break*/, 7];
                        console.warn('[NewRequestForm] Retrying request creation without manager lookup bind:', errMsg);
                        cleanPayload_1 = __assign({}, payload);
                        Object.keys(cleanPayload_1).forEach(function (k) {
                            if (k.toLowerCase().indexOf('manager@odata.bind') !== -1) {
                                delete cleanPayload_1[k];
                            }
                        });
                        return [4 /*yield*/, this._service.createRequest(cleanPayload_1)];
                    case 6:
                        newId_1 = _d.sent();
                        return [3 /*break*/, 8];
                    case 7: throw submitErr_1;
                    case 8: return [3 /*break*/, 9];
                    case 9:
                        if (!!isDraft) return [3 /*break*/, 14];
                        routing_1 = this._getRoutingConfig();
                        auditDesc = isExisting
                            ? "Request revised and resubmitted by ".concat(routing_1.submittingAs, " (Pending: ").concat(routing_1.targetPendingTier, ")")
                            : routing_1.submittingAs === 'Requester'
                                ? 'Request submitted for Line Manager approval'
                                : "Request submitted by ".concat(routing_1.submittingAs, ". ").concat(routing_1.skippedStagesDesc, "; routed directly to ").concat(routing_1.targetPendingTier, ".");
                        auditPayload = {
                            sct_action: isExisting ? 2 : 1, // 2 = Edited/Resubmitted, 1 = Submitted
                            sct_description: auditDesc,
                            sct_oldvalue: 'Status: Draft, Pending Tier: None',
                            sct_newvalue: "Status: ".concat(routing_1.targetStatus === 854000002 ? 'Submitted' : 'In Review', ", Pending Tier: ").concat(routing_1.targetPendingTier),
                            'sct_requestid@odata.bind': "/sct_requests(".concat(newId_1, ")"),
                        };
                        if (this.state.requesterGuid && GUID_RE.test(this.state.requesterGuid)) {
                            auditPayload['sct_actionby@odata.bind'] = "/systemusers(".concat(this.state.requesterGuid, ")");
                        }
                        _d.label = 10;
                    case 10:
                        _d.trys.push([10, 12, , 14]);
                        return [4 /*yield*/, this._service.createAuditLog(auditPayload)];
                    case 11:
                        _d.sent();
                        return [3 /*break*/, 14];
                    case 12:
                        auditErr_1 = _d.sent();
                        console.warn('[NewRequestForm] Primary audit log failed, retrying without actionby:', auditErr_1);
                        delete auditPayload['sct_actionby@odata.bind'];
                        return [4 /*yield*/, this._service.createAuditLog(auditPayload).catch(function () { })];
                    case 13:
                        _d.sent();
                        return [3 /*break*/, 14];
                    case 14:
                        if (!(this.state.attachments.length > 0)) return [3 /*break*/, 16];
                        this.setState({ uploadingFiles: true });
                        return [4 /*yield*/, Promise.all(this.state.attachments.map(function (a) {
                                return _this._service.uploadAnnotation(newId_1, a.name, a.mimeType, a.base64);
                            }))];
                    case 15:
                        _d.sent();
                        this.setState({ uploadingFiles: false });
                        _d.label = 16;
                    case 16:
                        routing_2 = this._getRoutingConfig();
                        submitMessage = isDraft
                            ? 'Draft saved. You can retrieve it from your Requests.'
                            : routing_2.submittingAs === 'Requester'
                                ? 'Request submitted successfully! It is now pending Manager approval.'
                                : "Request submitted successfully! Routed directly to ".concat(routing_2.targetPendingTier, " for approval.");
                        this.setState((_b = {},
                            _b[key] = false,
                            _b.isSubmitting = false,
                            _b.submitSuccess = submitMessage,
                            _b.savedId = newId_1,
                            _b));
                        // ── Send approval email to reviewer (Line Manager or Target Tier Group) ──
                        if (!isDraft) {
                            void (function () { return __awaiter(_this, void 0, void 0, function () {
                                var graphClient, currentUserEmail_1, siteUrl, createdRec, _a, refNumber, requestSummary, managerEmail, mgr, _b, _c, mgrGraph, toEmails, ccEmails, subject, htmlBody, groupEmails, toEmails, ccEmails, subject, htmlBody, mailErr_1;
                                var _d, _e;
                                return __generator(this, function (_f) {
                                    switch (_f.label) {
                                        case 0:
                                            _f.trys.push([0, 22, , 23]);
                                            return [4 /*yield*/, this.props.context.msGraphClientFactory.getClient('3')];
                                        case 1:
                                            graphClient = _f.sent();
                                            currentUserEmail_1 = this.props.context.pageContext.user.email || this.state.requesterEmail || '';
                                            siteUrl = this.props.context.pageContext.web.absoluteUrl;
                                            createdRec = null;
                                            if (!this._service) return [3 /*break*/, 5];
                                            _f.label = 2;
                                        case 2:
                                            _f.trys.push([2, 4, , 5]);
                                            return [4 /*yield*/, this._service.getRawRequestById(newId_1)];
                                        case 3:
                                            createdRec = _f.sent();
                                            return [3 /*break*/, 5];
                                        case 4:
                                            _a = _f.sent();
                                            return [3 /*break*/, 5];
                                        case 5:
                                            refNumber = createdRec
                                                ? (strVal(createdRec, 'sct_sct_reference') || strVal(createdRec, 'sct_referencenumber') || newId_1)
                                                : '(New Submission)';
                                            requestSummary = createdRec || {
                                                sct_sct_reference: refNumber,
                                                sct_referencenumber: newId_1,
                                                sct_request_title: this.state.requestTitle,
                                                sct_estimated_total_value: this.state.estimatedAmount,
                                                _sct_requester_value: this.state.requesterName || this.props.context.pageContext.user.displayName,
                                                sct_business_unit: this.state.businessUnit,
                                                _sct_site_value: this.state.selectedSiteName || this.state.selectedSiteId,
                                                sct_description: this.state.description,
                                            };
                                            if (!(routing_2.targetPendingTier === 'Manager')) return [3 /*break*/, 18];
                                            managerEmail = '';
                                            if (!(this.state.managerGuid && this._service)) return [3 /*break*/, 13];
                                            return [4 /*yield*/, this._service.getSystemUserById(this.state.managerGuid).catch(function () { return null; })];
                                        case 6:
                                            if (!((_d = _f.sent()) !== null && _d !== void 0)) return [3 /*break*/, 7];
                                            _c = _d;
                                            return [3 /*break*/, 9];
                                        case 7: return [4 /*yield*/, this._service.getCurrentSystemUser(this.state.managerGuid).catch(function () { return null; })];
                                        case 8:
                                            _c = _f.sent();
                                            _f.label = 9;
                                        case 9:
                                            if (!((_e = _c) !== null && _e !== void 0)) return [3 /*break*/, 10];
                                            _b = _e;
                                            return [3 /*break*/, 12];
                                        case 10: return [4 /*yield*/, this._service.getSystemUserByEmail(this.state.managerGuid).catch(function () { return null; })];
                                        case 11:
                                            _b = _f.sent();
                                            _f.label = 12;
                                        case 12:
                                            mgr = _b;
                                            if (mgr && mgr.internalemailaddress) {
                                                managerEmail = mgr.internalemailaddress;
                                            }
                                            _f.label = 13;
                                        case 13:
                                            if (!!managerEmail) return [3 /*break*/, 15];
                                            return [4 /*yield*/, getManager(graphClient).catch(function () { return null; })];
                                        case 14:
                                            mgrGraph = _f.sent();
                                            if (mgrGraph && mgrGraph.mail) {
                                                managerEmail = mgrGraph.mail;
                                            }
                                            _f.label = 15;
                                        case 15:
                                            toEmails = managerEmail ? [managerEmail] : [];
                                            ccEmails = currentUserEmail_1 ? [currentUserEmail_1] : [];
                                            subject = "[SCT Approval Required] ".concat(refNumber, " \u2014 Line Manager Approval");
                                            htmlBody = generateApprovalEmailHTML(requestSummary, [], siteUrl, 'Line Manager Review & Approval Required');
                                            if (!(toEmails.length > 0)) return [3 /*break*/, 17];
                                            return [4 /*yield*/, sendGraphEmail(graphClient, toEmails, ccEmails, subject, htmlBody)];
                                        case 16:
                                            _f.sent();
                                            _f.label = 17;
                                        case 17: return [3 /*break*/, 21];
                                        case 18:
                                            if (!routing_2.targetGroupId) return [3 /*break*/, 21];
                                            return [4 /*yield*/, getGroupMemberEmails(graphClient, routing_2.targetGroupId).catch(function () { return []; })];
                                        case 19:
                                            groupEmails = _f.sent();
                                            toEmails = groupEmails.filter(function (e) { return Boolean(e) && e.toLowerCase() !== currentUserEmail_1.toLowerCase(); });
                                            ccEmails = currentUserEmail_1 ? [currentUserEmail_1] : [];
                                            subject = "[SCT Approval Required] ".concat(refNumber, " \u2014 ").concat(routing_2.targetPendingTier, " Review");
                                            htmlBody = generateApprovalEmailHTML(requestSummary, [], siteUrl, "".concat(routing_2.targetPendingTier, " Review & Approval Required (Originated by ").concat(routing_2.submittingAs, ")"));
                                            if (!(toEmails.length > 0)) return [3 /*break*/, 21];
                                            return [4 /*yield*/, sendGraphEmail(graphClient, toEmails, ccEmails, subject, htmlBody)];
                                        case 20:
                                            _f.sent();
                                            _f.label = 21;
                                        case 21: return [3 /*break*/, 23];
                                        case 22:
                                            mailErr_1 = _f.sent();
                                            console.warn('[NewRequestForm] Notification email dispatch warning:', mailErr_1);
                                            return [3 /*break*/, 23];
                                        case 23: return [2 /*return*/];
                                    }
                                });
                            }); })();
                        }
                        return [3 /*break*/, 18];
                    case 17:
                        err_2 = _d.sent();
                        this.setState((_c = {},
                            _c[key] = false,
                            _c.isSubmitting = false,
                            _c.submitError = err_2 instanceof Error ? "Network error: ".concat(err_2.message) : 'Network error: Please try again.',
                            _c));
                        return [3 /*break*/, 18];
                    case 18: return [2 /*return*/];
                }
            });
        });
    };
    // ── Attachment handling ───────────────────────────────────────────────────────
    NewRequestForm.prototype._onFilesSelected = function (files) {
        return __awaiter(this, void 0, void 0, function () {
            var added, i, file, base64, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!files || files.length === 0)
                            return [2 /*return*/];
                        added = [];
                        i = 0;
                        _b.label = 1;
                    case 1:
                        if (!(i < files.length)) return [3 /*break*/, 6];
                        file = files[i];
                        _b.label = 2;
                    case 2:
                        _b.trys.push([2, 4, , 5]);
                        return [4 /*yield*/, readFileAsBase64(file)];
                    case 3:
                        base64 = _b.sent();
                        added.push({ name: file.name, mimeType: file.type || 'application/octet-stream', base64: base64, sizeKb: Math.round(file.size / 1024) });
                        return [3 /*break*/, 5];
                    case 4:
                        _a = _b.sent();
                        return [3 /*break*/, 5];
                    case 5:
                        i++;
                        return [3 /*break*/, 1];
                    case 6:
                        this.setState(function (prev) { return ({ attachments: __spreadArray(__spreadArray([], prev.attachments, true), added, true) }); });
                        if (this._fileInputRef.current)
                            this._fileInputRef.current.value = '';
                        return [2 /*return*/];
                }
            });
        });
    };
    NewRequestForm.prototype._removeAttachment = function (index) {
        this.setState(function (prev) { return ({ attachments: prev.attachments.filter(function (_, i) { return i !== index; }) }); });
    };
    // ── Render helpers ────────────────────────────────────────────────────────────
    NewRequestForm.prototype._err = function (field) {
        return this.state.errors[field] || undefined;
    };
    NewRequestForm.prototype._clrErr = function (field) {
        if (this.state.errors[field]) {
            var errs = __assign({}, this.state.errors);
            delete errs[field];
            this.setState({ errors: errs });
        }
    };
    // ── Section 1: Header & Requester ─────────────────────────────────────────────
    NewRequestForm.prototype._renderHeaderSection = function () {
        var _this = this;
        var _a = this.state, sites = _a.sites, sitesError = _a.sitesError, selectedSiteId = _a.selectedSiteId, businessUnit = _a.businessUnit, sctFunction = _a.sctFunction, costIsFor = _a.costIsFor, autoCostCentre = _a.autoCostCentre, autoDivision = _a.autoDivision, autoBudgetOwner = _a.autoBudgetOwner;
        // Build dropdown options from raw ODataRecord[] — field names resolved at runtime
        var siteOptions = sites
            .map(function (s) { return ({ key: _this._siteId(s), text: _this._siteLabel(s) }); })
            .filter(function (o) { return o.key; });
        var showMultiSiteNote = costIsFor === 'Several Sites' || costIsFor === 'Company-Wide';
        var _b = this.state, requesterGuid = _b.requesterGuid, requesterName = _b.requesterName, requesterEmail = _b.requesterEmail, userLookupError = _b.userLookupError, loadingMeta = _b.loadingMeta;
        var routing = this._getRoutingConfig();
        return (React.createElement(Section, { iconName: "Contact", iconBg: "#EBF3FB", iconColor: "#0078D4", title: "Requester & Site Information" },
            sitesError && (React.createElement(MessageBar, { messageBarType: MessageBarType.warning, styles: { root: { borderRadius: 6, marginBottom: 12 } } },
                "Could not load sites: ",
                sitesError,
                ". Check Dataverse permissions and entity set name.")),
            userLookupError && (React.createElement(MessageBar, { messageBarType: MessageBarType.warning, styles: { root: { borderRadius: 6, marginBottom: 12 } } }, userLookupError)),
            React.createElement("div", { style: {
                    display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 12,
                    padding: '12px 18px', background: '#F0F6FF', borderRadius: 8, border: '1px solid #C7E0F4', marginBottom: 16
                } },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10 } },
                    React.createElement(Icon, { iconName: "Flow", style: { fontSize: 18, color: '#0078D4' } }),
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontSize: 12, fontWeight: 700, color: '#0078D4', textTransform: 'uppercase', letterSpacing: '0.4px' } },
                            "Routing Workflow \u2014 Submitting as ",
                            routing.submittingAs),
                        React.createElement("div", { style: { fontSize: 13, color: '#201F1E', marginTop: 2, fontWeight: 500 } }, routing.routeDisplay))),
                routing.submittingAs !== 'Requester' && (React.createElement("span", { style: {
                        fontSize: 12, fontWeight: 600, color: '#107C10', background: '#DFF6DD',
                        padding: '4px 12px', borderRadius: 12, border: '1px solid #92C353'
                    } },
                    "Direct to ",
                    routing.targetPendingTier))),
            React.createElement("div", { style: { display: 'flex', gap: 16, marginBottom: 16, background: '#F8F8F8', borderRadius: 8, padding: '12px 16px', border: '1px solid #EDEBE9', alignItems: 'center' } },
                React.createElement("div", { style: { flex: 1 } },
                    React.createElement(Text, { variant: "small", style: { color: '#A19F9D', display: 'block', marginBottom: 2 } }, "REQUESTER"),
                    loadingMeta
                        ? React.createElement(Text, { variant: "small", style: { color: '#605E5C' } }, "Resolving user\u2026")
                        : React.createElement(React.Fragment, null,
                            React.createElement(Text, { variant: "medium", style: { color: '#201F1E', fontWeight: 600, display: 'block' } }, requesterName || this.props.context.pageContext.user.displayName),
                            requesterEmail && (React.createElement(Text, { variant: "small", style: { color: '#605E5C' } }, requesterEmail)))),
                React.createElement("div", { style: { flex: 1 } },
                    React.createElement(Text, { variant: "small", style: { color: '#A19F9D', display: 'block', marginBottom: 2 } }, "DATAVERSE BINDING"),
                    requesterGuid
                        ? React.createElement("span", { style: { background: '#DFF6DD', color: '#107C10', fontSize: 11, fontWeight: 700, padding: '3px 10px', borderRadius: 12 } }, "Resolved \u2713")
                        : React.createElement("span", { style: { background: '#FFF4CE', color: '#905400', fontSize: 11, fontWeight: 700, padding: '3px 10px', borderRadius: 12 } }, loadingMeta ? 'Resolving…' : 'Not bound')),
                React.createElement("div", { style: { flex: 1 } },
                    React.createElement(Text, { variant: "small", style: { color: '#A19F9D', display: 'block', marginBottom: 2 } }, "SCT REFERENCE"),
                    React.createElement("span", { style: { background: '#EDEBE9', color: '#605E5C', fontSize: 12, fontWeight: 600, padding: '3px 12px', borderRadius: 12 } }, "Draft \u2014 assigned on submission"))),
            React.createElement(Row, null,
                React.createElement(TextField, { label: "Line Manager", value: routing.submittingAs !== 'Requester'
                        ? "Auto-bypassed (".concat(routing.submittingAs, ")")
                        : (this.state.managerName ? "".concat(this.state.managerName).concat(this.state.managerEmail ? " (".concat(this.state.managerEmail, ")") : '') : (this.state.loadingMeta ? 'Fetching manager from Azure AD…' : 'No manager assigned')), readOnly: true, disabled: true, description: routing.submittingAs !== 'Requester'
                        ? routing.skippedStagesDesc
                        : 'Automatically fetched from Azure Active Directory (read-only)', styles: fieldStyles })),
            React.createElement(Row, null,
                React.createElement(Dropdown, { label: "Business Unit", required: true, placeholder: "Select business unit", options: BUSINESS_UNIT_OPTIONS, selectedKey: businessUnit || undefined, onChange: function (_, opt) { var _a; _this.setState({ businessUnit: String((_a = opt === null || opt === void 0 ? void 0 : opt.key) !== null && _a !== void 0 ? _a : '') }); _this._clrErr('businessUnit'); }, errorMessage: this._err('businessUnit'), styles: dropdownStyles }),
                React.createElement(Dropdown, { label: "Site / Location", required: true, placeholder: "Select site", options: siteOptions, selectedKey: selectedSiteId || undefined, onChange: function (_, opt) { var _a; _this._onSiteSelected(String((_a = opt === null || opt === void 0 ? void 0 : opt.key) !== null && _a !== void 0 ? _a : '')); _this._clrErr('selectedSiteId'); }, errorMessage: this._err('selectedSiteId'), styles: dropdownStyles })),
            (autoCostCentre || autoDivision || autoBudgetOwner) && (React.createElement("div", { style: { display: 'flex', flexWrap: 'wrap', gap: 24, background: '#F0F6FF', borderRadius: 8, padding: '12px 16px', marginBottom: 16, border: '1px solid #C7E0F4' } },
                autoCostCentre && (React.createElement("div", null,
                    React.createElement(Text, { variant: "small", style: { color: '#0078D4', display: 'block', fontWeight: 600 } }, "Cost Centre"),
                    React.createElement(Text, { variant: "medium", style: { color: '#201F1E' } }, autoCostCentre))),
                autoDivision && (React.createElement("div", null,
                    React.createElement(Text, { variant: "small", style: { color: '#0078D4', display: 'block', fontWeight: 600 } }, "Division / Region"),
                    React.createElement(Text, { variant: "medium", style: { color: '#201F1E' } }, autoDivision))),
                autoBudgetOwner && (React.createElement("div", null,
                    React.createElement(Text, { variant: "small", style: { color: '#0078D4', display: 'block', fontWeight: 600 } }, "L1 Budget Owner"),
                    React.createElement(Text, { variant: "medium", style: { color: '#201F1E' } }, autoBudgetOwner))),
                React.createElement("div", { style: { alignSelf: 'center' } },
                    React.createElement(Text, { variant: "small", style: { color: '#605E5C' } },
                        React.createElement(Icon, { iconName: "Lock", style: { fontSize: 11, marginRight: 4 } }),
                        "Auto-filled from Site Master \u2014 not editable by requester")))),
            React.createElement(Row, null,
                React.createElement(TextField, { label: "Function", placeholder: "e.g. Operations, IT, Finance", value: sctFunction, onChange: function (_, v) { return _this.setState({ sctFunction: v !== null && v !== void 0 ? v : '' }); }, styles: fieldStyles }),
                React.createElement(Dropdown, { label: "Cost Is For", placeholder: "Select scope", options: COST_IS_FOR_OPTIONS, selectedKey: costIsFor || undefined, onChange: function (_, opt) { var _a; return _this.setState({ costIsFor: String((_a = opt === null || opt === void 0 ? void 0 : opt.key) !== null && _a !== void 0 ? _a : '') }); }, styles: dropdownStyles })),
            showMultiSiteNote && (React.createElement(MessageBar, { messageBarType: MessageBarType.info, styles: { root: { borderRadius: 6, marginBottom: 16 } } }, "This request spans multiple sites or business units. Attach a cost allocation breakdown with your supporting documents."))));
    };
    // ── Section 2: Requirement & Supplier ─────────────────────────────────────────
    NewRequestForm.prototype._renderRequirementSection = function () {
        var _this = this;
        var _a = this.state, requestTitle = _a.requestTitle, description = _a.description, purchaseType = _a.purchaseType, isNewSupplier = _a.isNewSupplier, selectedSupplierId = _a.selectedSupplierId, newSupplierName = _a.newSupplierName, newSupplierLegalName = _a.newSupplierLegalName, newSupplierContactName = _a.newSupplierContactName, newSupplierContactEmail = _a.newSupplierContactEmail, newSupplierContactPhone = _a.newSupplierContactPhone, newSupplierW9Notes = _a.newSupplierW9Notes, suppliers = _a.suppliers;
        var supplierOptions = suppliers.map(function (s) { return ({
            key: s.sct_supplierid,
            text: s.sct_supplier_name || s.sct_suppliername || s.sct_supplierid,
        }); });
        var isRepairOrRental = purchaseType === 'Repair' || purchaseType === 'Rental';
        var isTempLabor = purchaseType === 'Temporary Labor';
        return (React.createElement(Section, { iconName: "BuildDefinition", iconBg: "#F4F0FF", iconColor: "#8764B8", title: "Requirement & Supplier" },
            React.createElement(TextField, { label: "Request Title", required: true, placeholder: "Brief, specific title for this spend request", value: requestTitle, onChange: function (_, v) { _this.setState({ requestTitle: v !== null && v !== void 0 ? v : '' }); _this._clrErr('requestTitle'); }, errorMessage: this._err('requestTitle'), styles: fieldStyles }),
            React.createElement(Row, null,
                React.createElement(Dropdown, { label: "Purchase Type", required: true, placeholder: "Select type", options: PURCHASE_TYPE_OPTIONS, selectedKey: purchaseType || undefined, onChange: function (_, opt) { var _a; _this.setState({ purchaseType: String((_a = opt === null || opt === void 0 ? void 0 : opt.key) !== null && _a !== void 0 ? _a : '') }); _this._clrErr('purchaseType'); }, errorMessage: this._err('purchaseType'), styles: dropdownStyles }),
                React.createElement("div", null)),
            isTempLabor && (React.createElement(MessageBar, { messageBarType: MessageBarType.warning, styles: { root: { borderRadius: 6, marginBottom: 16 } } }, "Temporary Labour requests at a unionized site require prior union compliance review. Confirm with HR / Site Manager before submission.")),
            isRepairOrRental && (React.createElement(MessageBar, { messageBarType: MessageBarType.info, styles: { root: { borderRadius: 6, marginBottom: 16 } } }, "Repair and Rental requests require an equipment asset tag or serial number in Section 4.")),
            React.createElement(TextField, { label: "Description / Justification", required: true, placeholder: "Describe what is being purchased, why it is needed, and expected business outcome", multiline: true, rows: 4, value: description, onChange: function (_, v) { _this.setState({ description: v !== null && v !== void 0 ? v : '' }); _this._clrErr('description'); }, errorMessage: this._err('description'), styles: fieldStyles }),
            React.createElement(Toggle, { label: "Is this a new / proposed supplier?", checked: isNewSupplier, onChange: function (_, checked) { return _this.setState({ isNewSupplier: !!checked, selectedSupplierId: '', newSupplierName: '' }); }, onText: "Yes \u2014 proposed / new supplier", offText: "No \u2014 select from approved supplier list", styles: { root: { marginBottom: 16 } } }),
            !isNewSupplier ? (React.createElement(Dropdown, { label: "Suggested Supplier", placeholder: "Select an approved supplier", options: supplierOptions, selectedKey: selectedSupplierId || undefined, onChange: function (_, opt) { var _a; return _this.setState({ selectedSupplierId: String((_a = opt === null || opt === void 0 ? void 0 : opt.key) !== null && _a !== void 0 ? _a : '') }); }, styles: dropdownStyles })) : (React.createElement("div", { style: { background: '#FFF4CE', borderRadius: 8, padding: '12px 16px', marginBottom: 16, border: '1px solid #FFB900' } },
                React.createElement(Text, { variant: "small", style: { color: '#7A4800', display: 'block', marginBottom: 12, fontWeight: 600 } },
                    React.createElement(Icon, { iconName: "Warning", style: { marginRight: 6 } }),
                    "New supplier details below do not automatically make this supplier Active or Preferred. Procurement review and approval is required before any commitment can be made."),
                React.createElement(TextField, { label: "Proposed Supplier Name", required: true, placeholder: "Trading / display name", value: newSupplierName, onChange: function (_, v) { _this.setState({ newSupplierName: v !== null && v !== void 0 ? v : '' }); _this._clrErr('newSupplierName'); }, errorMessage: this._err('newSupplierName'), styles: fieldStyles }),
                (function () {
                    var trimmed = (newSupplierName || '').trim().toLowerCase();
                    if (trimmed.length >= 3) {
                        var dup = _this.state.suppliers.find(function (sup) {
                            var sName = (sup.sct_supplier_name || sup.sct_suppliername || '').toLowerCase();
                            return Boolean(sName && (sName.indexOf(trimmed) !== -1 || trimmed.indexOf(sName) !== -1));
                        });
                        if (dup) {
                            var matchName = dup.sct_supplier_name || dup.sct_suppliername || 'Existing Vendor';
                            return (React.createElement(MessageBar, { messageBarType: MessageBarType.warning, styles: { root: { borderRadius: 6, margin: '8px 0' } } },
                                "Approved vendor ",
                                React.createElement("strong", null, matchName),
                                " already exists in the system. Verify if this request can use the approved vendor."));
                        }
                    }
                    return null;
                })(),
                React.createElement(TextField, { label: "Legal / Registered Name", placeholder: "Full registered legal entity name", value: newSupplierLegalName, onChange: function (_, v) { return _this.setState({ newSupplierLegalName: v !== null && v !== void 0 ? v : '' }); }, styles: fieldStyles }),
                React.createElement(Row, null,
                    React.createElement(TextField, { label: "Contact Name", placeholder: "Primary contact person", value: newSupplierContactName, onChange: function (_, v) { return _this.setState({ newSupplierContactName: v !== null && v !== void 0 ? v : '' }); }, styles: fieldStyles }),
                    React.createElement(TextField, { label: "Contact Email", placeholder: "contact@supplier.com", type: "email", value: newSupplierContactEmail, onChange: function (_, v) { return _this.setState({ newSupplierContactEmail: v !== null && v !== void 0 ? v : '' }); }, styles: fieldStyles })),
                React.createElement(Row, null,
                    React.createElement(TextField, { label: "Contact Phone", placeholder: "+61 ...", value: newSupplierContactPhone, onChange: function (_, v) { return _this.setState({ newSupplierContactPhone: v !== null && v !== void 0 ? v : '' }); }, styles: fieldStyles }),
                    React.createElement(TextField, { label: "W-9 / Bank Details Notes", placeholder: "Note status of tax form and bank details verification", value: newSupplierW9Notes, onChange: function (_, v) { return _this.setState({ newSupplierW9Notes: v !== null && v !== void 0 ? v : '' }); }, styles: fieldStyles }))))));
    };
    // ── Section 3: Financials, Term, Urgency ──────────────────────────────────────
    NewRequestForm.prototype._renderFinancialSection = function () {
        var _this = this;
        var _a = this.state, estimatedAmount = _a.estimatedAmount, termStartDate = _a.termStartDate, termEndDate = _a.termEndDate, pricingBasis = _a.pricingBasis, dateNeededBy = _a.dateNeededBy, urgency = _a.urgency, urgencyReason = _a.urgencyReason, emergencyDeclaredBy = _a.emergencyDeclaredBy, isNewSupplier = _a.isNewSupplier;
        var amt = parseFloat(estimatedAmount) || 0;
        var guidance = calcApprovalGuidance(isNewSupplier, amt);
        var showUrgencyReason = urgency === 'Urgent' || urgency === 'Emergency';
        var isEmergency = urgency === 'Emergency';
        return (React.createElement(Section, { iconName: "Money", iconBg: "#DFF6DD", iconColor: "#107C10", title: "Financials, Term & Urgency" },
            (amt > 0 || isNewSupplier) && (React.createElement("div", { style: { background: guidance.bg, borderRadius: 8, padding: '12px 16px', marginBottom: 16, border: "1px solid ".concat(guidance.color, "40") } },
                React.createElement(Text, { variant: "small", style: { color: guidance.color, fontWeight: 700, display: 'block', marginBottom: 6 } },
                    React.createElement(Icon, { iconName: guidance.iconName, style: { marginRight: 6, fontSize: 13 } }),
                    "Pre-Approval Evidence Required"),
                React.createElement("div", { style: { display: 'flex', gap: 8, flexWrap: 'wrap' } }, guidance.levels.map(function (l) { return (React.createElement("span", { key: l, style: { background: 'rgba(255,255,255,0.7)', color: guidance.color, fontSize: 12, fontWeight: 600, padding: '3px 10px', borderRadius: 10, border: "1px solid ".concat(guidance.color, "50") } }, l)); })),
                React.createElement(Text, { variant: "small", style: { color: guidance.color, marginTop: 6, display: 'block' } },
                    "Written pre-approval from the above level",
                    guidance.levels.length > 1 ? 's' : '',
                    " must be attached before submission."))),
            React.createElement(Row, null,
                React.createElement(TextField, { label: "Estimated Total Value (AUD)", required: true, placeholder: "0.00", prefix: "$", type: "number", value: estimatedAmount, onChange: function (_, v) { _this.setState({ estimatedAmount: v !== null && v !== void 0 ? v : '' }); _this._clrErr('estimatedAmount'); }, errorMessage: this._err('estimatedAmount'), description: "Full value over the entire intended term or contract period", styles: fieldStyles }),
                React.createElement(Dropdown, { label: "Pricing Basis", placeholder: "Select pricing basis", options: PRICING_BASIS_OPTIONS, selectedKey: pricingBasis || undefined, onChange: function (_, opt) { var _a; return _this.setState({ pricingBasis: String((_a = opt === null || opt === void 0 ? void 0 : opt.key) !== null && _a !== void 0 ? _a : '') }); }, styles: dropdownStyles })),
            React.createElement(Row, null,
                React.createElement(DatePicker, { label: "Term Start Date", placeholder: "Select start date", firstDayOfWeek: DayOfWeek.Monday, value: termStartDate !== null && termStartDate !== void 0 ? termStartDate : undefined, onSelectDate: function (d) { return _this.setState({ termStartDate: d !== null && d !== void 0 ? d : null }); }, styles: { root: { marginBottom: 16 } } }),
                React.createElement(DatePicker, { label: "Term End Date", placeholder: "Select end date", firstDayOfWeek: DayOfWeek.Monday, value: termEndDate !== null && termEndDate !== void 0 ? termEndDate : undefined, onSelectDate: function (d) { return _this.setState({ termEndDate: d !== null && d !== void 0 ? d : null }); }, styles: { root: { marginBottom: 16 } } })),
            React.createElement(Row, null,
                React.createElement(DatePicker, { label: "Date Needed By", isRequired: true, placeholder: "Select required date", firstDayOfWeek: DayOfWeek.Monday, value: dateNeededBy !== null && dateNeededBy !== void 0 ? dateNeededBy : undefined, onSelectDate: function (d) { _this.setState({ dateNeededBy: d !== null && d !== void 0 ? d : null }); _this._clrErr('dateNeededBy'); }, styles: { root: { marginBottom: 16 } } }),
                React.createElement(Dropdown, { label: "Urgency", required: true, placeholder: "Select urgency", options: URGENCY_OPTIONS, selectedKey: urgency || undefined, onChange: function (_, opt) { var _a; _this.setState({ urgency: String((_a = opt === null || opt === void 0 ? void 0 : opt.key) !== null && _a !== void 0 ? _a : '') }); _this._clrErr('urgency'); }, errorMessage: this._err('urgency'), styles: dropdownStyles })),
            showUrgencyReason && (React.createElement(TextField, { label: "Urgency Reason", required: true, placeholder: "Explain why this request is urgent or an emergency", multiline: true, rows: 3, value: urgencyReason, onChange: function (_, v) { _this.setState({ urgencyReason: v !== null && v !== void 0 ? v : '' }); _this._clrErr('urgencyReason'); }, errorMessage: this._err('urgencyReason'), styles: fieldStyles })),
            isEmergency && (React.createElement(MessageBar, { messageBarType: MessageBarType.severeWarning, styles: { root: { borderRadius: 6, marginBottom: 16 } } },
                React.createElement("strong", null, "Emergency Purchase Cap:"),
                " Emergency spend is limited to ",
                React.createElement("strong", null, "$25,000 AUD and 7 days duration"),
                " unless an approved written override from CFO is attached. The person declaring the emergency must be recorded below.")),
            isEmergency && (React.createElement(TextField, { label: "Emergency Declared By", required: true, placeholder: "Full name and title of person authorising this emergency", value: emergencyDeclaredBy, onChange: function (_, v) { _this.setState({ emergencyDeclaredBy: v !== null && v !== void 0 ? v : '' }); _this._clrErr('emergencyDeclaredBy'); }, errorMessage: this._err('emergencyDeclaredBy'), styles: fieldStyles }))));
    };
    // ── Section 4: Logistics & Conditional Triggers ───────────────────────────────
    NewRequestForm.prototype._renderLogisticsSection = function () {
        var _this = this;
        var _a = this.state, deliverTo = _a.deliverTo, receivingContact = _a.receivingContact, equipmentAssetTag = _a.equipmentAssetTag, purchaseType = _a.purchaseType, costRechargedToCustomer = _a.costRechargedToCustomer, customerName = _a.customerName, customerAgreementAttached = _a.customerAgreementAttached, techOrDataInvolved = _a.techOrDataInvolved, supplierOnPremises = _a.supplierOnPremises, possibleCapitalPurchase = _a.possibleCapitalPurchase, isConfidential = _a.isConfidential, isRetrospective = _a.isRetrospective;
        var isRepairOrRental = purchaseType === 'Repair' || purchaseType === 'Rental';
        return (React.createElement(Section, { iconName: "DeliveryTruck", iconBg: "#FFF4CE", iconColor: "#905400", title: "Logistics & Conditions" },
            React.createElement(Row, null,
                React.createElement(TextField, { label: "Deliver To / Delivery Address", placeholder: "Site address or internal delivery location", multiline: true, rows: 2, value: deliverTo, onChange: function (_, v) { return _this.setState({ deliverTo: v !== null && v !== void 0 ? v : '' }); }, styles: fieldStyles }),
                React.createElement(TextField, { label: "Receiving Contact", placeholder: "Name of person receiving goods / services", value: receivingContact, onChange: function (_, v) { return _this.setState({ receivingContact: v !== null && v !== void 0 ? v : '' }); }, styles: fieldStyles })),
            React.createElement(TextField, { label: "Equipment Asset Tag / Serial Number".concat(isRepairOrRental ? ' *' : ''), placeholder: "Asset tag or serial number for the item being repaired or rented", value: equipmentAssetTag, onChange: function (_, v) { _this.setState({ equipmentAssetTag: v !== null && v !== void 0 ? v : '' }); _this._clrErr('equipmentAssetTag'); }, errorMessage: this._err('equipmentAssetTag'), description: isRepairOrRental ? 'Required for Repair and Rental purchase types' : 'If applicable', styles: fieldStyles }),
            React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '0 24px', marginTop: 8 } },
                React.createElement(Toggle, { label: "Cost Recharged to Customer?", checked: costRechargedToCustomer, onChange: function (_, c) { return _this.setState({ costRechargedToCustomer: !!c }); }, onText: "Yes", offText: "No", styles: { root: { marginBottom: 16 } } }),
                React.createElement(Toggle, { label: "Technology or Data Involved?", checked: techOrDataInvolved, onChange: function (_, c) { return _this.setState({ techOrDataInvolved: !!c }); }, onText: "Yes \u2014 CIO approval required (Rule R17)", offText: "No", styles: { root: { marginBottom: 16 } } }),
                React.createElement(Toggle, { label: "Supplier Working on Premises?", checked: supplierOnPremises, onChange: function (_, c) { return _this.setState({ supplierOnPremises: !!c }); }, onText: "Yes \u2014 Risk COI check required", offText: "No", styles: { root: { marginBottom: 16 } } }),
                React.createElement(Toggle, { label: "Possible Capital Purchase?", checked: possibleCapitalPurchase, onChange: function (_, c) { return _this.setState({ possibleCapitalPurchase: !!c }); }, onText: "Yes \u2014 Finance capital review triggered", offText: "No", styles: { root: { marginBottom: 16 } } }),
                React.createElement(Toggle, { label: "Confidential / Legal / HR Matter?", checked: isConfidential, onChange: function (_, c) { return _this.setState({ isConfidential: !!c }); }, onText: "Yes \u2014 item-level access restriction applied", offText: "No", styles: { root: { marginBottom: 16 } } }),
                React.createElement(Toggle, { label: "Retrospective Request?", checked: isRetrospective, onChange: function (_, c) { return _this.setState({ isRetrospective: !!c }); }, onText: "Yes \u2014 commitment may precede approval", offText: "No", styles: { root: { marginBottom: 16 } } })),
            costRechargedToCustomer && (React.createElement("div", { style: { background: '#EBF3FB', borderRadius: 8, padding: '12px 16px', marginTop: 8, border: '1px solid #C7E0F4' } },
                React.createElement(Row, null,
                    React.createElement(TextField, { label: "Customer Name", required: true, placeholder: "Customer or project being recharged", value: customerName, onChange: function (_, v) { _this.setState({ customerName: v !== null && v !== void 0 ? v : '' }); _this._clrErr('customerName'); }, errorMessage: this._err('customerName'), styles: fieldStyles }),
                    React.createElement("div", { style: { paddingTop: 36 } },
                        React.createElement(Checkbox, { label: "Customer agreement / quote attached", checked: customerAgreementAttached, onChange: function (_, c) { _this.setState({ customerAgreementAttached: !!c }); _this._clrErr('customerAgreementAttached'); }, styles: checkStyles }),
                        this._err('customerAgreementAttached') && (React.createElement(Text, { variant: "small", style: { color: '#A4262C', display: 'block', marginTop: 4 } }, this._err('customerAgreementAttached'))))))),
            techOrDataInvolved && (React.createElement(MessageBar, { messageBarType: MessageBarType.info, styles: { root: { borderRadius: 6, marginTop: 12 } } }, "CIO approval (Rule R17) has been triggered. The SCT desk will route this request accordingly after submission.")),
            supplierOnPremises && (React.createElement(MessageBar, { messageBarType: MessageBarType.warning, styles: { root: { borderRadius: 6, marginTop: 8 } } }, "Risk Management COI (Certificate of Insurance) check has been triggered. Ensure the supplier's insurance certificate is on file before work begins.")),
            possibleCapitalPurchase && (React.createElement(MessageBar, { messageBarType: MessageBarType.info, styles: { root: { borderRadius: 6, marginTop: 8 } } }, "Finance capital review has been triggered. FP&A will assess whether this spend should be capitalised.")),
            isRetrospective && (React.createElement(MessageBar, { messageBarType: MessageBarType.severeWarning, styles: { root: { borderRadius: 6, marginTop: 8 } } }, "Retrospective requests indicate a commitment may have been made before SCT approval. This requires escalation and ratification review."))));
    };
    // ── Section 5: Declarations ────────────────────────────────────────────────────
    NewRequestForm.prototype._renderDeclarationsSection = function () {
        var _this = this;
        var _a = this.state, declarationNoPriorCommitment = _a.declarationNoPriorCommitment, declarationNoConflictOfInterest = _a.declarationNoConflictOfInterest;
        return (React.createElement(Section, { iconName: "Shield", iconBg: "#FDE7E9", iconColor: "#A4262C", title: "Mandatory Declarations" },
            React.createElement(Text, { variant: "small", style: { color: '#605E5C', display: 'block', marginBottom: 16, lineHeight: '1.6' } }, "Both declarations must be confirmed before this request can be submitted. False declarations may result in disciplinary action."),
            React.createElement("div", { style: { background: '#FDE7E9', borderRadius: 8, padding: '16px 20px', border: '1px solid #F4B8BC' } },
                React.createElement(Checkbox, { label: "Rule R1 \u2014 No Prior Commitment: I confirm that no commitment, agreement, or purchase order has been made or placed with any supplier prior to obtaining the necessary SCT approvals.", checked: declarationNoPriorCommitment, onChange: function (_, c) { _this.setState({ declarationNoPriorCommitment: !!c }); _this._clrErr('declarationNoPriorCommitment'); }, styles: { root: { marginBottom: 14 }, label: { color: '#201F1E', lineHeight: '1.5' } } }),
                this._err('declarationNoPriorCommitment') && (React.createElement(Text, { variant: "small", style: { color: '#A4262C', marginBottom: 10, display: 'block' } }, this._err('declarationNoPriorCommitment'))),
                React.createElement(Checkbox, { label: "Rule R10 \u2014 No Conflict of Interest: I declare that I have no personal, financial, or professional conflict of interest with any supplier relevant to this request.", checked: declarationNoConflictOfInterest, onChange: function (_, c) { _this.setState({ declarationNoConflictOfInterest: !!c }); _this._clrErr('declarationNoConflictOfInterest'); }, styles: { root: { marginBottom: 0 }, label: { color: '#201F1E', lineHeight: '1.5' } } }),
                this._err('declarationNoConflictOfInterest') && (React.createElement(Text, { variant: "small", style: { color: '#A4262C', marginTop: 4, display: 'block' } }, this._err('declarationNoConflictOfInterest'))))));
    };
    // ── Section 6: Attachments ────────────────────────────────────────────────────
    NewRequestForm.prototype._renderAttachmentsSection = function () {
        var _this = this;
        var _a = this.state, attachments = _a.attachments, isNewSupplier = _a.isNewSupplier, estimatedAmount = _a.estimatedAmount;
        var amt = parseFloat(estimatedAmount) || 0;
        var guidance = calcApprovalGuidance(isNewSupplier, amt);
        return (React.createElement(Section, { iconName: "Attach", iconBg: "#E6F7F4", iconColor: "#00B294", title: "Supporting Documents & Evidence" },
            (amt > 0 || isNewSupplier) && (React.createElement("div", { style: { background: guidance.bg, borderRadius: 6, padding: '10px 14px', marginBottom: 14, border: "1px solid ".concat(guidance.color, "40") } },
                React.createElement(Text, { variant: "small", style: { color: guidance.color, fontWeight: 600 } },
                    "Required attachments for this request: ",
                    guidance.levels.join(', '),
                    " pre-approval evidence",
                    isNewSupplier ? ', plus supplier contact details and any quotes / proposals' : '',
                    "."))),
            React.createElement(Text, { variant: "small", style: { color: '#605E5C', display: 'block', marginBottom: 12 } }, "Attach quotes, supplier proposals, written approvals, customer agreements, and any other supporting evidence."),
            React.createElement("input", { type: "file", multiple: true, ref: this._fileInputRef, style: { display: 'none' }, onChange: function (e) { return void _this._onFilesSelected(e.target.files); }, "aria-hidden": "true" }),
            React.createElement(DefaultButton, { text: "Attach Files", iconProps: { iconName: 'Attach' }, onClick: function () { var _a; return (_a = _this._fileInputRef.current) === null || _a === void 0 ? void 0 : _a.click(); }, styles: { root: { borderRadius: 6, marginBottom: 16 } } }),
            this._err('attachments') && (React.createElement(MessageBar, { messageBarType: MessageBarType.error, styles: { root: { borderRadius: 6, marginBottom: 12 } } }, this._err('attachments'))),
            attachments.length > 0 && (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 8 } }, attachments.map(function (a, i) { return (React.createElement("div", { key: i, style: { display: 'flex', alignItems: 'center', gap: 10, background: '#F8F8F8', borderRadius: 6, padding: '8px 12px', border: '1px solid #EDEBE9' } },
                React.createElement(Icon, { iconName: "Page", style: { color: '#0078D4', fontSize: 16, flexShrink: 0 } }),
                React.createElement("div", { style: { flex: 1, minWidth: 0 } },
                    React.createElement(Text, { variant: "small", style: { color: '#201F1E', fontWeight: 600, display: 'block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, a.name),
                    React.createElement(Text, { variant: "small", style: { color: '#A19F9D' } },
                        a.sizeKb,
                        " KB")),
                React.createElement(IconButton, { iconProps: { iconName: 'Cancel' }, title: "Remove", ariaLabel: "Remove attachment", onClick: function () { return _this._removeAttachment(i); }, styles: { root: { color: '#A4262C', width: 28, height: 28 } } }))); })))));
    };
    // ── Wizard helpers ────────────────────────────────────────────────────────────
    NewRequestForm.prototype._canGoNext = function (step) {
        var s = this.state;
        switch (step) {
            case 0: return !!s.businessUnit && !!s.selectedSiteId;
            case 1: {
                var rawAmt = s.estimatedAmount ? String(s.estimatedAmount).trim() : '';
                var amt = parseFloat(rawAmt.replace(/[$,]/g, ''));
                var validAmt = !isNaN(amt) && amt > 0;
                var validSupplier = !s.isNewSupplier ? true : !!s.newSupplierName.trim();
                return validAmt && !!s.description.trim() && !!s.purchaseType && validSupplier && !!s.dateNeededBy && !!s.urgency;
            }
            case 2: {
                var validRecharge = !s.costRechargedToCustomer || (!!s.customerName.trim() && s.customerAgreementAttached);
                var validAsset = (s.purchaseType !== 'Repair' && s.purchaseType !== 'Rental') || !!s.equipmentAssetTag.trim();
                return validRecharge && validAsset && s.declarationNoPriorCommitment && s.declarationNoConflictOfInterest;
            }
            case 3: return true;
            default: return true;
        }
    };
    NewRequestForm.prototype._renderReviewStep = function () {
        var s = this.state;
        var field = function (label, value) { return (React.createElement("div", { style: { display: 'flex', gap: 8, marginBottom: 10 } },
            React.createElement(Text, { variant: "small", style: { color: '#A19F9D', minWidth: 180, flexShrink: 0 } }, label),
            React.createElement(Text, { variant: "small", style: { color: '#201F1E', fontWeight: 600, wordBreak: 'break-word' } }, value || '—'))); };
        var bool = function (v) { return v ? 'Yes' : 'No'; };
        return (React.createElement("div", { className: styles.card, style: { marginBottom: 16 } },
            React.createElement("div", { className: styles.cardBody },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20 } },
                    React.createElement(Icon, { iconName: "DocumentSearch", style: { fontSize: 20, color: '#0078D4' } }),
                    React.createElement(Text, { variant: "large", style: { fontWeight: 700, color: '#201F1E' } }, "Review Your Request")),
                React.createElement("div", { style: { marginBottom: 18 } },
                    React.createElement(Text, { variant: "mediumPlus", style: { fontWeight: 600, color: '#0078D4', display: 'block', marginBottom: 10, borderBottom: '2px solid #EBF3FB', paddingBottom: 6 } }, "General Info"),
                    field('Requester', s.requesterName || '—'),
                    field('Line Manager', s.managerName ? "".concat(s.managerName).concat(s.managerEmail ? " (".concat(s.managerEmail, ")") : '') : '—'),
                    field('Business Unit', s.businessUnit),
                    field('Site', s.selectedSiteName),
                    field('SCT Function', s.sctFunction),
                    field('Cost Is For', s.costIsFor),
                    field('Purchase Type', s.purchaseType)),
                React.createElement("div", { style: { marginBottom: 18 } },
                    React.createElement(Text, { variant: "mediumPlus", style: { fontWeight: 600, color: '#0078D4', display: 'block', marginBottom: 10, borderBottom: '2px solid #EBF3FB', paddingBottom: 6 } }, "Financials"),
                    field('Estimated Amount', s.estimatedAmount ? "$".concat(parseFloat(s.estimatedAmount).toLocaleString()) : '—'),
                    field('Cost Centre', s.autoCostCentre),
                    field('Pricing Basis', s.pricingBasis),
                    field('Date Needed By', s.dateNeededBy ? s.dateNeededBy.toLocaleDateString() : '—'),
                    field('Urgency', s.urgency)),
                React.createElement("div", { style: { marginBottom: 18 } },
                    React.createElement(Text, { variant: "mediumPlus", style: { fontWeight: 600, color: '#0078D4', display: 'block', marginBottom: 10, borderBottom: '2px solid #EBF3FB', paddingBottom: 6 } }, "Supplier & Compliance"),
                    field('Request Title', s.requestTitle),
                    field('New Supplier?', bool(s.isNewSupplier)),
                    !s.isNewSupplier && field('Selected Supplier', s.selectedSupplierId ? 'Linked' : '—'),
                    s.isNewSupplier && field('New Supplier Name', s.newSupplierName),
                    field('Tech/Data Involved?', bool(s.techOrDataInvolved)),
                    field('Supplier On Premises?', bool(s.supplierOnPremises)),
                    field('Capital Purchase?', bool(s.possibleCapitalPurchase)),
                    field('Attachments', s.attachments.length > 0 ? "".concat(s.attachments.length, " file(s)") : 'None')),
                React.createElement("div", null,
                    React.createElement(Text, { variant: "mediumPlus", style: { fontWeight: 600, color: '#0078D4', display: 'block', marginBottom: 10, borderBottom: '2px solid #EBF3FB', paddingBottom: 6 } }, "Declarations"),
                    field('No Prior Commitment (R1)', s.declarationNoPriorCommitment ? '✓ Confirmed' : '✗ Not confirmed'),
                    field('No Conflict of Interest (R10)', s.declarationNoConflictOfInterest ? '✓ Confirmed' : '✗ Not confirmed')))));
    };
    // ── Success screen ─────────────────────────────────────────────────────────────
    NewRequestForm.prototype._renderSuccess = function () {
        var _this = this;
        var _a = this.state, savedId = _a.savedId, submitSuccess = _a.submitSuccess;
        return (React.createElement("div", { className: styles.card, style: { marginBottom: 16 } },
            React.createElement("div", { className: styles.cardBody, style: { textAlign: 'center', padding: '48px 24px' } },
                React.createElement("div", { style: { width: 72, height: 72, borderRadius: '50%', background: '#DFF6DD', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 20px' } },
                    React.createElement(Icon, { iconName: "CheckMark", style: { fontSize: 32, color: '#107C10' } })),
                React.createElement(Text, { variant: "xLarge", style: { color: '#107C10', fontWeight: 700, display: 'block', marginBottom: 10 } }, submitSuccess),
                savedId && (React.createElement(Text, { variant: "medium", style: { color: '#605E5C', display: 'block', marginBottom: 24 } },
                    "Dataverse ID: ",
                    React.createElement("code", { style: { background: '#F3F2F1', padding: '2px 8px', borderRadius: 4, fontFamily: 'monospace', fontSize: 12 } }, savedId))),
                React.createElement("div", { style: { display: 'flex', gap: 12, justifyContent: 'center', marginTop: 16 } },
                    React.createElement(PrimaryButton, { text: "Submit Another Request", iconProps: { iconName: 'Add' }, onClick: function () { return _this.setState({
                            submitSuccess: null, savedId: null, errors: {},
                            requestTitle: '', description: '', businessUnit: '',
                            selectedSiteId: '', selectedSiteName: '', autoCostCentre: '', autoDivision: '', autoBudgetOwner: '',
                            sctFunction: '', costIsFor: '', purchaseType: '',
                            isNewSupplier: false, selectedSupplierId: '',
                            newSupplierName: '', newSupplierLegalName: '', newSupplierContactName: '',
                            newSupplierContactEmail: '', newSupplierContactPhone: '', newSupplierW9Notes: '',
                            estimatedAmount: '', termStartDate: null, termEndDate: null, pricingBasis: '',
                            dateNeededBy: null, urgency: '', urgencyReason: '', emergencyDeclaredBy: '',
                            deliverTo: '', receivingContact: '', equipmentAssetTag: '',
                            costRechargedToCustomer: false, customerName: '', customerAgreementAttached: false,
                            techOrDataInvolved: false, supplierOnPremises: false, possibleCapitalPurchase: false,
                            isConfidential: false, isRetrospective: false,
                            declarationNoPriorCommitment: false, declarationNoConflictOfInterest: false,
                            attachments: [], submitError: null,
                            currentStep: 0,
                        }); }, styles: { root: { borderRadius: 6 } } }),
                    this.props.onFinished && (React.createElement(DefaultButton, { text: "View My Requests", iconProps: { iconName: 'DocumentSearch' }, onClick: this.props.onFinished, styles: { root: { borderRadius: 6 } } }))))));
    };
    // ── Main render ───────────────────────────────────────────────────────────────
    NewRequestForm.prototype.render = function () {
        var _this = this;
        var dataverseService = this.props.dataverseService;
        var _a = this.state, loadingMeta = _a.loadingMeta, metaError = _a.metaError, submitSuccess = _a.submitSuccess, submitting = _a.submitting, savingDraft = _a.savingDraft, submitError = _a.submitError, uploadingFiles = _a.uploadingFiles;
        if (!dataverseService)
            return React.createElement(LoadingSpinner, { label: "Connecting to Dataverse..." });
        if (loadingMeta)
            return React.createElement(LoadingSpinner, { label: "Loading sites and suppliers\u2026" });
        if (metaError) {
            return (React.createElement("div", { className: styles.card, style: { marginBottom: 16 } },
                React.createElement("div", { className: styles.cardBody },
                    React.createElement(MessageBar, { messageBarType: MessageBarType.error, styles: { root: { borderRadius: 6, marginBottom: 12 } } }, metaError),
                    React.createElement(DefaultButton, { text: "Retry", iconProps: { iconName: 'Refresh' }, onClick: function () { return void _this._loadMeta(); }, styles: { root: { borderRadius: 6 } } }))));
        }
        if (submitSuccess)
            return this._renderSuccess();
        var isWorking = submitting || savingDraft || uploadingFiles;
        // Hero header
        var hero = (React.createElement("div", { className: styles.card, style: {
                marginBottom: 16,
                background: 'linear-gradient(135deg, #0F172A 0%, #1E293B 100%)',
                color: '#ffffff',
                border: 'none',
                boxShadow: '0 4px 16px rgba(15, 23, 42, 0.12)',
            } },
            React.createElement("div", { className: styles.cardBody, style: { display: 'flex', alignItems: 'center', gap: 16, padding: '20px 24px' } },
                React.createElement("div", { style: {
                        width: 48,
                        height: 48,
                        borderRadius: 12,
                        background: 'linear-gradient(135deg, #0078D4, #2563EB)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        flexShrink: 0,
                        boxShadow: '0 2px 8px rgba(0, 120, 212, 0.4)',
                    } },
                    React.createElement(Icon, { iconName: "Add", style: { fontSize: 24, color: '#ffffff' } })),
                React.createElement("div", null,
                    React.createElement(Text, { variant: "xLarge", style: { color: '#ffffff', fontWeight: 700, display: 'block', letterSpacing: '-0.2px' } }, "New Spend Control Request"),
                    React.createElement(Text, { variant: "medium", style: { color: '#94A3B8', fontSize: 13 } },
                        "Step ",
                        this.state.currentStep + 1,
                        " of 4: Complete required procurement details. Reference number is auto-generated upon submission.")))));
        var currentStep = this.state.currentStep;
        var STEPS = [
            { key: 0, title: 'Location & Requester', sub: 'Site, Unit & Line Manager' },
            { key: 1, title: 'Requirement & Spend', sub: 'Vendor, Amount & Type' },
            { key: 2, title: 'Governance & Logistics', sub: 'Asset, Recharges & Rules' },
            { key: 3, title: 'Review & Submit', sub: 'Evidence & Pre-Approvals' },
        ];
        var canNext = this._canGoNext(currentStep);
        // ── Step indicator ──
        var stepIndicator = (React.createElement("div", { style: {
                background: '#ffffff',
                borderRadius: 12,
                border: '1px solid #E2E8F0',
                padding: '14px 20px',
                marginBottom: 20,
                boxShadow: '0 1px 4px rgba(15, 23, 42, 0.04)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                flexWrap: 'wrap',
                gap: 8,
            } }, STEPS.map(function (sDef, idx) {
            var isDone = idx < currentStep;
            var isActive = idx === currentStep;
            var isClickable = idx <= currentStep || (idx === currentStep + 1 && canNext);
            return (React.createElement(React.Fragment, { key: idx },
                React.createElement("div", { onClick: function () {
                        if (isClickable)
                            _this.setState({ currentStep: idx });
                    }, style: {
                        display: 'flex',
                        alignItems: 'center',
                        gap: 10,
                        cursor: isClickable ? 'pointer' : 'default',
                        opacity: isActive || isDone ? 1 : 0.6,
                        padding: '6px 12px',
                        borderRadius: 8,
                        background: isActive ? '#EFF6FF' : 'transparent',
                        border: isActive ? '1px solid #BFDBFE' : '1px solid transparent',
                        transition: 'all 0.2s ease',
                    } },
                    React.createElement("div", { style: {
                            width: 30,
                            height: 30,
                            borderRadius: '50%',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            fontWeight: 700,
                            fontSize: 13,
                            background: isDone ? '#107C10' : isActive ? '#0078D4' : '#E2E8F0',
                            color: isDone || isActive ? '#ffffff' : '#64748B',
                            flexShrink: 0,
                            boxShadow: isActive ? '0 2px 6px rgba(0, 120, 212, 0.35)' : 'none',
                        } }, isDone ? React.createElement(Icon, { iconName: "CheckMark", style: { fontSize: 13 } }) : idx + 1),
                    React.createElement("div", null,
                        React.createElement("div", { style: {
                                fontSize: 13,
                                fontWeight: isActive ? 700 : 600,
                                color: isActive ? '#0078D4' : isDone ? '#107C10' : '#334155',
                                lineHeight: 1.2,
                            } }, sDef.title),
                        React.createElement("div", { style: { fontSize: 11, color: '#64748B', marginTop: 2 } }, sDef.sub))),
                idx < STEPS.length - 1 && (React.createElement("div", { style: {
                        flex: 1,
                        height: 2,
                        minWidth: 16,
                        background: idx < currentStep ? '#107C10' : '#E2E8F0',
                        borderRadius: 1,
                    } }))));
        })));
        // ── Step content ──
        var stepContent = (function () {
            switch (currentStep) {
                case 0: return _this._renderHeaderSection();
                case 1: return (React.createElement(React.Fragment, null,
                    _this._renderRequirementSection(),
                    _this._renderFinancialSection()));
                case 2: return (React.createElement(React.Fragment, null,
                    _this._renderLogisticsSection(),
                    _this._renderDeclarationsSection()));
                case 3: return (React.createElement(React.Fragment, null,
                    _this._renderAttachmentsSection(),
                    _this._renderReviewStep()));
                default: return null;
            }
        })();
        // ── Navigation footer ──
        var navFooter = (React.createElement("div", { style: { display: 'flex', gap: 12, padding: '16px 0 32px', alignItems: 'center' } },
            currentStep > 0 && (React.createElement(DefaultButton, { text: "Back", iconProps: { iconName: 'ChevronLeft' }, disabled: isWorking, onClick: function () { return _this.setState({ currentStep: currentStep - 1 }); }, styles: { root: { borderRadius: 6, height: 40 } } })),
            currentStep < STEPS.length - 1 && (React.createElement(PrimaryButton, { text: "Next \u2192", iconProps: { iconName: 'ChevronRight' }, disabled: !canNext || isWorking, onClick: function () { return _this.setState({ currentStep: currentStep + 1 }); }, styles: { root: { borderRadius: 6, height: 40, minWidth: 120 } } })),
            currentStep === STEPS.length - 1 && (React.createElement(React.Fragment, null,
                React.createElement(PrimaryButton, { text: (submitting || this.state.isSubmitting) ? 'Submitting…' : 'Submit Request', iconProps: { iconName: (submitting || this.state.isSubmitting) ? undefined : 'Send' }, disabled: isWorking || this.state.isSubmitting || submitting || !this.state.declarationNoPriorCommitment || !this.state.declarationNoConflictOfInterest, title: (!this.state.declarationNoPriorCommitment || !this.state.declarationNoConflictOfInterest)
                        ? 'Both mandatory declarations (R1 and R10) must be checked before submitting'
                        : undefined, onClick: function () { return void _this._submit(false); }, onRenderText: (submitting || this.state.isSubmitting) ? function () { return (React.createElement("span", { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement(Spinner, { size: SpinnerSize.xSmall }),
                        " Submitting\u2026")); } : undefined, styles: { root: { borderRadius: 6, height: 40, minWidth: 160 } } }),
                React.createElement(DefaultButton, { text: savingDraft ? 'Saving…' : 'Save as Draft', iconProps: { iconName: savingDraft ? undefined : 'Save' }, disabled: isWorking, onClick: function () { return void _this._submit(true); }, onRenderText: savingDraft ? function () { return (React.createElement("span", { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement(Spinner, { size: SpinnerSize.xSmall }),
                        " Saving\u2026")); } : undefined, styles: { root: { borderRadius: 6, height: 40 } } }))),
            currentStep < STEPS.length - 1 && (React.createElement(DefaultButton, { text: savingDraft ? 'Saving…' : 'Save Draft', iconProps: { iconName: 'Save' }, disabled: isWorking, onClick: function () { return void _this._submit(true); }, styles: { root: { borderRadius: 6, height: 40 } } })),
            uploadingFiles && (React.createElement("span", { style: { display: 'flex', alignItems: 'center', gap: 8, color: '#605E5C', fontSize: 13 } },
                React.createElement(Spinner, { size: SpinnerSize.xSmall }),
                " Uploading attachments\u2026"))));
        return (React.createElement("div", null,
            hero,
            stepIndicator,
            submitError && (React.createElement(MessageBar, { messageBarType: MessageBarType.error, isMultiline: false, onDismiss: function () { return _this.setState({ submitError: null }); }, styles: { root: { borderRadius: 6, marginBottom: 12 } } }, submitError)),
            stepContent,
            currentStep === STEPS.length - 1 && Object.keys(this.state.errors).length > 0 && (React.createElement("div", { style: { background: '#FDE7E9', border: '1px solid #F4B8BC', borderRadius: 8, padding: '14px 18px', marginBottom: 16 } },
                React.createElement(Text, { variant: "small", style: { color: '#A4262C', fontWeight: 700, display: 'block', marginBottom: 8 } },
                    React.createElement(Icon, { iconName: "ErrorBadge", style: { marginRight: 6, fontSize: 14 } }),
                    "Please fix the following before submitting:"),
                React.createElement("ul", { style: { margin: 0, paddingLeft: 20 } }, Object.keys(this.state.errors).map(function (key, i) { return (React.createElement("li", { key: i, style: { color: '#A4262C', fontSize: 13, marginBottom: 4 } }, _this.state.errors[key])); })))),
            navFooter));
    };
    return NewRequestForm;
}(React.Component));
export default NewRequestForm;
//# sourceMappingURL=NewRequestForm.js.map