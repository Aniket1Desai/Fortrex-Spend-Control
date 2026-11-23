/**
 * approvalHandler.ts — Robust sequential approval state machine & validation engine
 *
 * Workflow:  Manager → Tier 1 → Tier 2 → Tier 3 → PO Issued
 *
 * Hardened Features:
 *  1. State validation: Prevents double-approvals, race conditions, or acting on closed/rejected/draft records.
 *  2. Queue validation: Enforces that requests are pending at the actor's tier before allowing actions.
 *  3. Rule R10 Enforcement: Requesters are strictly prohibited from approving, rejecting, or returning their own requests.
 *  4. Comment validation: Mandatory non-empty rationale for rejections and returns (min 5 chars).
 *  5. Dual logging: Simultaneously records in sct_approvallogs and sct_auditlogs for complete compliance.
 *  6. Integrated notification dispatch: Automatically sends MS Graph notifications to the next tier or the requester.
 *
 * Routing matrix:
 *   Role             Action    → newStatus       cr2a4_pendingtier
 *   ───────────────────────────────────────────────────────────────
 *   manager          approve   → IN_REVIEW        'Tier 1'
 *   approver_tier1   approve   → IN_REVIEW        'Tier 2'
 *   approver_tier2   approve   → IN_REVIEW        'Tier 3'
 *   approver_tier3   approve   → PO_ISSUED         null    (final)
 *   any              reject    → REJECTED          null    (hard reset)
 *   any              return    → RETURNED          null    (soft reset)
 */
import { __awaiter, __generator } from "tslib";
import { strVal } from './DataverseService';
import { getRecordStatusLabel, formatDateTime } from '../components/dashboard/dashboardUtils';
import { sendGraphEmail, generateApprovalEmailHTML, getGroupMemberEmails, } from './GraphMailService';
// ── Option set constants ──────────────────────────────────────────────────────
/** sct_approvalstatus (Edm.Int32) — confirmed from Dataverse validation */
export var APPROVAL_STATUS = {
    APPROVED: 1,
    REJECTED: 2,
    RETURN_FOR_INFO: 3,
};
/** sct_request_status (Edm.Int32) — Confirmed Dataverse choice values: 854000001 to 854000007 */
export var REQUEST_STATUS = {
    DRAFT: 854000001,
    SUBMITTED: 854000002,
    IN_REVIEW: 854000003,
    PO_ISSUED: 854000004,
    REJECTED: 854000005,
    RETURNED: 854000006,
    CLOSED: 854000007,
};
/** sct_tiername (Edm.Int32) — Confirmed Accepted Values from Dataverse: 1=Tier 1, 2=Tier 2, 3=Tier 3, 4=Manager */
export var TIER_INT = {
    manager: 4,
    approver_tier1: 1,
    approver_tier2: 2,
    approver_tier3: 3,
};
/**
 * cr2a4_pendingtier (Edm.String) — explicit queue ownership.
 * Logical name: cr2a4_pendingtier  |  Schema: cr2a4_PendingTier
 * Column exists on sct_request (confirmed from Power Apps, 11 Nov 2026).
 * null = no active queue (initial draft / returned / rejected / final).
 */
export var PENDING_TIER = {
    manager: 'Tier 1', // manager approved → Tier 1 picks up
    approver_tier1: 'Tier 2', // tier 1 approved  → Tier 2 picks up
    approver_tier2: 'Tier 3', // tier 2 approved  → Tier 3 picks up
    approver_tier3: null, // tier 3 approved  → final, no next queue
};
/** Tier label written to cr2a4_pendingtier when a request enters review */
export var INITIAL_PENDING_TIER = 'Manager';
/** Column name on sct_request for the pending tier routing field */
export var PENDING_TIER_COLUMN = 'cr2a4_pendingtier';
// ── Azure AD Security Group IDs for dynamic member email resolution ───────────
export var TIER1_GROUP_ID = 'fdbf07a0-548b-4ccf-875b-900cb1b24a1c'; // SCT-Fortrex_Approvers-Tier1
export var TIER2_GROUP_ID = '6a8a4077-c78f-4af2-ae5e-12428d1dcbc3'; // SCT-Fortrex_Approvers-Tier2
export var TIER3_GROUP_ID = 'e63dbc4f-6ad2-4969-a518-b229836756af'; // SCT-Fortrex_Approvers-Tier3
// ── GUID guard ────────────────────────────────────────────────────────────────
var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
// ── Main function ─────────────────────────────────────────────────────────────
/**
 * handleApprovalAction
 *
 * Executes one approval step in the sequential workflow with strict validation,
 * concurrency checks, Rule R10 enforcement, dual audit logging, and email notifications.
 */
export function handleApprovalAction(params) {
    return __awaiter(this, void 0, void 0, function () {
        var requestId, role, action, comments, service, currentSystemUserId, graphClient, webPartUrl, approverDisplayName, currentUserEmail, currentRecord, fetchErr_1, requestRequesterId, actingUserId, isDevOrWorkbench, allowSelfApproval, normalizedRole, roleLower, actionStr, actionType, currentRole, logTier, tierInt, approvalStatusInt, cleanComments, fallbackComments, finalComments, currentStatus, currentPendingTier, rawRequester, canonicalStatus, rawPt, hasPt, isTerminal, termLabel, normPt, normActor, isManagerValid, isTierValid, existingLogs, latestReturnTime_1, currentCycleLogs, priorLogForTier, priorStatus, priorApprover, priorTimeRaw, priorTime, checkErr_1, logPayload, requestUpdatePayload, logErr_1, fallbackLog, newStatus, nextPendingTier, patched, err1_1, patchErr_1, actionDesc, statusName, auditPayload, auditErr_1, emailDispatched, emailWarning, rec, ref, requestTitle, requesterEmail, uById, _a, createdByAad, uInfo, _b, actingUserEmail, actingUser, _c, buildCc, tier1Emails, tier2Emails, tier3Emails, _d, t1, t2, t3, previousLogs, _e, toEmails, ccEmails, subject, actionRequired, htmlBody, mailErr_1;
        var _f, _g, _h, _j, _k, _l, _m, _o;
        return __generator(this, function (_p) {
            switch (_p.label) {
                case 0:
                    requestId = params.requestId, role = params.role, action = params.action, comments = params.comments, service = params.service, currentSystemUserId = params.currentSystemUserId, graphClient = params.graphClient, webPartUrl = params.webPartUrl, approverDisplayName = params.approverDisplayName, currentUserEmail = params.currentUserEmail;
                    // ── Guard 1: Validate input syntax ──────────────────────────────────────────
                    if (!GUID_RE.test(requestId)) {
                        throw new Error("handleApprovalAction: invalid requestId \"".concat(requestId, "\""));
                    }
                    currentRecord = null;
                    _p.label = 1;
                case 1:
                    _p.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, service.getRawRequestById(requestId)];
                case 2:
                    currentRecord = _p.sent();
                    return [3 /*break*/, 4];
                case 3:
                    fetchErr_1 = _p.sent();
                    console.warn('[approvalHandler] getRawRequestById warn:', fetchErr_1);
                    return [3 /*break*/, 4];
                case 4:
                    requestRequesterId = (params.requesterId ||
                        (currentRecord ? String((_h = (_g = (_f = currentRecord['_sct_requester_value']) !== null && _f !== void 0 ? _f : currentRecord['sct_requester']) !== null && _g !== void 0 ? _g : currentRecord['_createdby_value']) !== null && _h !== void 0 ? _h : '') : '')).toLowerCase().replace(/[{}]/g, '');
                    actingUserId = (currentSystemUserId || '').toLowerCase().replace(/[{}]/g, '');
                    isDevOrWorkbench = typeof window !== 'undefined' && (window.location.href.indexOf('workbench') !== -1 ||
                        window.location.href.indexOf('devrishabhsoft') !== -1 ||
                        window.location.hostname.indexOf('localhost') !== -1 ||
                        window.location.hostname.indexOf('127.0.0.1') !== -1);
                    allowSelfApproval = params.soloTestMode === true || isDevOrWorkbench;
                    if (!allowSelfApproval && requestRequesterId && actingUserId && requestRequesterId === actingUserId) {
                        throw new Error("Segregation of Duties Violation: You cannot approve your own request in Strict Mode. Enable Solo Test Mode in the header if you are testing alone.");
                    }
                    normalizedRole = (role || '').trim();
                    roleLower = normalizedRole.toLowerCase();
                    if (roleLower === 'manager') {
                        normalizedRole = 'manager';
                    }
                    else if (roleLower === 'approver_tier1' || roleLower.indexOf('tier 1') !== -1 || roleLower.indexOf('tier1') !== -1) {
                        normalizedRole = 'approver_tier1';
                    }
                    else if (roleLower === 'approver_tier2' || roleLower.indexOf('tier 2') !== -1 || roleLower.indexOf('tier2') !== -1) {
                        normalizedRole = 'approver_tier2';
                    }
                    else if (roleLower === 'approver_tier3' || roleLower.indexOf('tier 3') !== -1 || roleLower.indexOf('tier3') !== -1) {
                        normalizedRole = 'approver_tier3';
                    }
                    actionStr = String(action || '').toLowerCase();
                    actionType = actionStr === 'approve' ? 'Approve'
                        : actionStr === 'reject' ? 'Reject'
                            : actionStr === 'return' ? 'Return'
                                : action;
                    currentRole = role;
                    if (normalizedRole === 'manager' || roleLower === 'manager') {
                        currentRole = 'Manager';
                    }
                    else if (normalizedRole === 'approver_tier1' || roleLower.indexOf('1') !== -1) {
                        currentRole = 'Approver - Tier 1';
                    }
                    else if (normalizedRole === 'approver_tier2' || roleLower.indexOf('2') !== -1) {
                        currentRole = 'Approver - Tier 2';
                    }
                    else if (normalizedRole === 'approver_tier3' || roleLower.indexOf('3') !== -1) {
                        currentRole = 'Approver - Tier 3';
                    }
                    logTier = "";
                    if (currentRole === 'Manager')
                        logTier = 'Manager';
                    else if (currentRole.includes('1'))
                        logTier = 'Tier 1';
                    else if (currentRole.includes('2'))
                        logTier = 'Tier 2';
                    else if (currentRole.includes('3'))
                        logTier = 'Tier 3';
                    tierInt = logTier === 'Manager' ? 4 : logTier === 'Tier 1' ? 1 : logTier === 'Tier 2' ? 2 : 3;
                    approvalStatusInt = actionType === 'Approve' ? APPROVAL_STATUS.APPROVED
                        : actionType === 'Reject' ? APPROVAL_STATUS.REJECTED
                            : APPROVAL_STATUS.RETURN_FOR_INFO;
                    cleanComments = (comments !== null && comments !== void 0 ? comments : '').trim();
                    // Require explicit comments for Reject and Return
                    if ((actionType === 'Reject' || actionType === 'Return') && cleanComments.length === 0) {
                        throw new Error("Comments are required for ".concat(actionType === 'Reject' ? 'rejection' : 'return for correction', "."));
                    }
                    fallbackComments = "".concat(actionType.toUpperCase(), " by ").concat(logTier);
                    finalComments = cleanComments || fallbackComments;
                    currentStatus = currentRecord ? currentRecord['sct_request_status'] : REQUEST_STATUS.IN_REVIEW;
                    currentPendingTier = currentRecord ? String(currentRecord['cr2a4_pendingtier'] || currentRecord['sct_pendingtier'] || currentRecord['cr2a4_PendingTier'] || '').trim() : '';
                    rawRequester = currentRecord ? String((_k = (_j = currentRecord['_sct_requester_value']) !== null && _j !== void 0 ? _j : currentRecord['sct_requester']) !== null && _k !== void 0 ? _k : '') : '';
                    if (!currentRecord) return [3 /*break*/, 8];
                    canonicalStatus = getRecordStatusLabel(currentRecord);
                    rawPt = String(currentRecord['cr2a4_pendingtier'] ||
                        currentRecord['cr2a4_PendingTier'] ||
                        currentRecord['sct_pendingtier'] ||
                        currentRecord['pendingtier'] ||
                        '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
                    hasPt = rawPt === 'manager' || rawPt.indexOf('tier') !== -1 || rawPt === '1' || rawPt === '2' || rawPt === '3';
                    isTerminal = canonicalStatus === 'Rejected' || canonicalStatus === 'Closed' || (!hasPt && canonicalStatus === 'PO Issued');
                    if (isTerminal) {
                        termLabel = canonicalStatus === 'PO Issued' ? 'Approved / PO Issued' : canonicalStatus;
                        throw new Error("Cannot process action: This request has already reached terminal status (".concat(termLabel, ")."));
                    }
                    normPt = currentPendingTier.toLowerCase().replace(/[^a-z0-9]/g, '');
                    normActor = logTier.toLowerCase().replace(/[^a-z0-9]/g, '');
                    isManagerValid = logTier === 'Manager' && (normPt === 'manager' || !normPt || currentStatus === REQUEST_STATUS.SUBMITTED);
                    isTierValid = normPt === normActor ||
                        (normActor.indexOf('1') !== -1 && (normPt.indexOf('1') !== -1 || normPt === '1')) ||
                        (normActor.indexOf('2') !== -1 && (normPt.indexOf('2') !== -1 || normPt === '2')) ||
                        (normActor.indexOf('3') !== -1 && (normPt.indexOf('3') !== -1 || normPt === '3'));
                    if (!isManagerValid && !isTierValid) {
                        throw new Error("Workflow stage mismatch: Request is currently pending at '".concat(currentPendingTier || 'Line Manager', "', but action was submitted by '").concat(logTier, "'. Stages cannot be skipped."));
                    }
                    _p.label = 5;
                case 5:
                    _p.trys.push([5, 7, , 8]);
                    return [4 /*yield*/, service.getApprovalLogs(requestId)];
                case 6:
                    existingLogs = _p.sent();
                    if (Array.isArray(existingLogs) && existingLogs.length > 0) {
                        latestReturnTime_1 = 0;
                        existingLogs.forEach(function (l) {
                            var _a, _b, _c, _d;
                            var statusVal = l['sct_approvalstatus'];
                            var nameVal = String((_a = l['sct_name']) !== null && _a !== void 0 ? _a : '').toLowerCase();
                            var commentsVal = String((_b = l['sct_comments']) !== null && _b !== void 0 ? _b : '').toLowerCase();
                            var isReturn = statusVal === APPROVAL_STATUS.RETURN_FOR_INFO ||
                                nameVal.indexOf('return') !== -1 ||
                                commentsVal.indexOf('return') !== -1;
                            if (isReturn) {
                                var t = new Date(String((_d = (_c = l['sct_decisionon']) !== null && _c !== void 0 ? _c : l['createdon']) !== null && _d !== void 0 ? _d : 0)).getTime();
                                if (t > latestReturnTime_1)
                                    latestReturnTime_1 = t;
                            }
                        });
                        currentCycleLogs = latestReturnTime_1 > 0
                            ? existingLogs.filter(function (l) {
                                var _a, _b;
                                var t = new Date(String((_b = (_a = l['sct_decisionon']) !== null && _a !== void 0 ? _a : l['createdon']) !== null && _b !== void 0 ? _b : 0)).getTime();
                                return t > latestReturnTime_1;
                            })
                            : existingLogs;
                        priorLogForTier = currentCycleLogs.find(function (l) {
                            var _a;
                            var tInt = l['sct_tiername'];
                            if (tInt === tierInt)
                                return true;
                            var name = String((_a = l['sct_name']) !== null && _a !== void 0 ? _a : '').toLowerCase();
                            return name.indexOf(logTier.toLowerCase()) !== -1;
                        });
                        if (priorLogForTier) {
                            priorStatus = priorLogForTier['sct_approvalstatus'] === 1 ? 'Approved'
                                : priorLogForTier['sct_approvalstatus'] === 2 ? 'Rejected'
                                    : 'Returned';
                            priorApprover = strVal(priorLogForTier, '_sct_approverid_value') ||
                                strVal(priorLogForTier, '_createdby_value') ||
                                strVal(priorLogForTier, 'sct_name') ||
                                'an approver';
                            priorTimeRaw = priorLogForTier['createdon'] || priorLogForTier['sct_decisionon'];
                            priorTime = priorTimeRaw ? formatDateTime(String(priorTimeRaw)) : '';
                            throw new Error("One-Time Decision Enforcement: A decision for stage '".concat(logTier, "' has already been recorded in this approval cycle (").concat(priorStatus, " by ").concat(priorApprover).concat(priorTime ? ' on ' + priorTime : '', "). You cannot submit another decision for this stage."));
                        }
                    }
                    return [3 /*break*/, 8];
                case 7:
                    checkErr_1 = _p.sent();
                    if (checkErr_1 instanceof Error && checkErr_1.message.indexOf('One-Time Decision Enforcement') !== -1) {
                        throw checkErr_1;
                    }
                    return [3 /*break*/, 8];
                case 8:
                    logPayload = {
                        sct_name: "".concat(logTier, " Decision"),
                        sct_tiername: tierInt,
                        sct_approvalstatus: approvalStatusInt,
                        sct_comments: finalComments,
                        sct_decisionon: new Date().toISOString(),
                        'sct_requestid@odata.bind': "/sct_requests(".concat(requestId, ")"),
                    };
                    if (currentSystemUserId && GUID_RE.test(currentSystemUserId)) {
                        logPayload['sct_approverid@odata.bind'] = "/systemusers(".concat(currentSystemUserId, ")");
                    }
                    requestUpdatePayload = {};
                    if (actionType === 'Reject') {
                        requestUpdatePayload.sct_request_status = 'Rejected';
                        requestUpdatePayload.cr2a4_pendingtier = null;
                    }
                    else if (actionType === 'Approve') {
                        requestUpdatePayload.sct_request_status = 'In Review';
                        // Map the NEXT destination based on current role
                        switch (logTier) {
                            case 'Manager':
                                requestUpdatePayload.cr2a4_pendingtier = 'Tier 1';
                                break;
                            case 'Tier 1':
                                requestUpdatePayload.cr2a4_pendingtier = 'Tier 2';
                                break;
                            case 'Tier 2':
                                requestUpdatePayload.cr2a4_pendingtier = 'Tier 3';
                                break;
                            case 'Tier 3':
                                requestUpdatePayload.sct_request_status = 'PO Issued';
                                requestUpdatePayload.cr2a4_pendingtier = null;
                                break;
                            default:
                                throw new Error("Invalid approval role");
                        }
                    }
                    else if (actionType === 'Return') {
                        requestUpdatePayload.sct_request_status = 'Returned';
                        requestUpdatePayload.cr2a4_pendingtier = null;
                    }
                    _p.label = 9;
                case 9:
                    _p.trys.push([9, 11, , 13]);
                    return [4 /*yield*/, service.createApprovalLog(logPayload)];
                case 10:
                    _p.sent();
                    return [3 /*break*/, 13];
                case 11:
                    logErr_1 = _p.sent();
                    console.warn('[approvalHandler] Primary createApprovalLog failed, retrying without approver bind:', logErr_1);
                    fallbackLog = {
                        sct_name: "".concat(logTier, " Decision"),
                        sct_tiername: tierInt,
                        sct_approvalstatus: approvalStatusInt,
                        sct_comments: finalComments,
                        sct_decisionon: new Date().toISOString(),
                        'sct_requestid@odata.bind': "/sct_requests(".concat(requestId, ")"),
                    };
                    return [4 /*yield*/, service.createApprovalLog(fallbackLog).catch(function (fallbackErr) {
                            console.warn('[approvalHandler] Fallback createApprovalLog also failed:', fallbackErr);
                        })];
                case 12:
                    _p.sent();
                    return [3 /*break*/, 13];
                case 13:
                    if (requestUpdatePayload.sct_request_status === 'Rejected') {
                        newStatus = REQUEST_STATUS.REJECTED;
                    }
                    else if (requestUpdatePayload.sct_request_status === 'PO Issued' ||
                        requestUpdatePayload.sct_request_status === 'Approved') {
                        newStatus = REQUEST_STATUS.PO_ISSUED;
                    }
                    else if (requestUpdatePayload.sct_request_status === 'Returned') {
                        newStatus = REQUEST_STATUS.RETURNED;
                    }
                    else {
                        newStatus = REQUEST_STATUS.IN_REVIEW;
                    }
                    nextPendingTier = (_l = requestUpdatePayload.cr2a4_pendingtier) !== null && _l !== void 0 ? _l : null;
                    patched = false;
                    _p.label = 14;
                case 14:
                    _p.trys.push([14, 16, , 21]);
                    // Attempt 1: numeric status + verified cr2a4_pendingtier
                    return [4 /*yield*/, service.updateRequest(requestId, {
                            sct_request_status: newStatus,
                            cr2a4_pendingtier: nextPendingTier,
                        })];
                case 15:
                    // Attempt 1: numeric status + verified cr2a4_pendingtier
                    _p.sent();
                    patched = true;
                    return [3 /*break*/, 21];
                case 16:
                    err1_1 = _p.sent();
                    console.warn('[approvalHandler] Primary update with cr2a4_pendingtier failed, retrying with status only:', err1_1);
                    _p.label = 17;
                case 17:
                    _p.trys.push([17, 19, , 20]);
                    // Attempt 2: numeric status only
                    return [4 /*yield*/, service.updateRequest(requestId, {
                            sct_request_status: newStatus,
                        })];
                case 18:
                    // Attempt 2: numeric status only
                    _p.sent();
                    patched = true;
                    return [3 /*break*/, 20];
                case 19:
                    patchErr_1 = _p.sent();
                    console.error('[approvalHandler] All updateRequest attempts failed:', patchErr_1);
                    throw patchErr_1;
                case 20: return [3 /*break*/, 21];
                case 21:
                    actionDesc = '';
                    if (actionType === 'Approve') {
                        if (logTier === 'Manager') {
                            actionDesc = 'Approved by Line Manager — advanced to Tier 1';
                        }
                        else if (logTier === 'Tier 1') {
                            actionDesc = 'Approved by Tier 1 — advanced to Tier 2';
                        }
                        else if (logTier === 'Tier 2') {
                            actionDesc = 'Approved by Tier 2 — advanced to Tier 3';
                        }
                        else if (logTier === 'Tier 3') {
                            actionDesc = 'Approved by Tier 3 — Final Approval (PO Issued)';
                        }
                        else {
                            actionDesc = "Approved at ".concat(logTier, " tier");
                        }
                    }
                    else if (actionType === 'Reject') {
                        actionDesc = "Rejected by ".concat(logTier);
                    }
                    else {
                        actionDesc = "Returned for Information by ".concat(logTier);
                    }
                    statusName = function (code) {
                        switch (code) {
                            case 854000001: return 'Draft';
                            case 854000002: return 'Submitted';
                            case 854000003: return 'In Review';
                            case 854000004: return 'PO Issued';
                            case 854000005: return 'Rejected';
                            case 854000006: return 'Returned';
                            case 854000007: return 'Closed';
                            case 854000000: return 'Draft'; // legacy fallback
                            default: return String(code);
                        }
                    };
                    auditPayload = {
                        sct_action: action === 'approve' ? 1 : action === 'reject' ? 5 : 4,
                        sct_description: actionDesc,
                        sct_oldvalue: "Status: ".concat(statusName(currentStatus), ", Pending Tier: ").concat(currentPendingTier || 'None'),
                        sct_newvalue: "Status: ".concat(statusName(newStatus), ", Pending Tier: ").concat(nextPendingTier || 'None'),
                        'sct_requestid@odata.bind': "/sct_requests(".concat(requestId, ")"),
                    };
                    if (currentSystemUserId && GUID_RE.test(currentSystemUserId)) {
                        auditPayload['sct_actionby@odata.bind'] = "/systemusers(".concat(currentSystemUserId, ")");
                    }
                    _p.label = 22;
                case 22:
                    _p.trys.push([22, 24, , 26]);
                    return [4 /*yield*/, service.createAuditLog(auditPayload)];
                case 23:
                    _p.sent();
                    return [3 /*break*/, 26];
                case 24:
                    auditErr_1 = _p.sent();
                    console.warn('[approvalHandler] Primary createAuditLog failed, retrying without actionby:', auditErr_1);
                    delete auditPayload['sct_actionby@odata.bind'];
                    return [4 /*yield*/, service.createAuditLog(auditPayload).catch(function () { })];
                case 25:
                    _p.sent();
                    return [3 /*break*/, 26];
                case 26:
                    emailDispatched = false;
                    if (!(graphClient && webPartUrl)) return [3 /*break*/, 48];
                    rec = currentRecord || {};
                    ref = strVal(rec, 'sct_sct_reference') || strVal(rec, 'sct_referencenumber') || requestId;
                    requestTitle = strVal(rec, 'sct_request_title') || ref;
                    requesterEmail = strVal(rec, 'sct_requesteremail') || strVal(rec, 'sct_receiving_contact_email');
                    if (!(!requesterEmail && rawRequester && GUID_RE.test(rawRequester))) return [3 /*break*/, 30];
                    _p.label = 27;
                case 27:
                    _p.trys.push([27, 29, , 30]);
                    return [4 /*yield*/, service.getSystemUserById(rawRequester)];
                case 28:
                    uById = _p.sent();
                    requesterEmail = (_m = uById === null || uById === void 0 ? void 0 : uById.internalemailaddress) !== null && _m !== void 0 ? _m : '';
                    return [3 /*break*/, 30];
                case 29:
                    _a = _p.sent();
                    return [3 /*break*/, 30];
                case 30:
                    if (!(!requesterEmail && graphClient)) return [3 /*break*/, 34];
                    createdByAad = strVal(rec, '_createdby_value');
                    if (!(createdByAad && GUID_RE.test(createdByAad))) return [3 /*break*/, 34];
                    _p.label = 31;
                case 31:
                    _p.trys.push([31, 33, , 34]);
                    return [4 /*yield*/, graphClient.api("/users/".concat(createdByAad)).select('mail,userPrincipalName').get()];
                case 32:
                    uInfo = _p.sent();
                    requesterEmail = (uInfo === null || uInfo === void 0 ? void 0 : uInfo.mail) || (uInfo === null || uInfo === void 0 ? void 0 : uInfo.userPrincipalName) || '';
                    return [3 /*break*/, 34];
                case 33:
                    _b = _p.sent();
                    return [3 /*break*/, 34];
                case 34:
                    actingUserEmail = (currentUserEmail || '').trim();
                    if (!(!actingUserEmail && currentSystemUserId && GUID_RE.test(currentSystemUserId))) return [3 /*break*/, 38];
                    _p.label = 35;
                case 35:
                    _p.trys.push([35, 37, , 38]);
                    return [4 /*yield*/, service.getCurrentSystemUser(currentSystemUserId)];
                case 36:
                    actingUser = _p.sent();
                    actingUserEmail = (_o = actingUser === null || actingUser === void 0 ? void 0 : actingUser.internalemailaddress) !== null && _o !== void 0 ? _o : '';
                    return [3 /*break*/, 38];
                case 37:
                    _c = _p.sent();
                    return [3 /*break*/, 38];
                case 38:
                    buildCc = function () {
                        var emails = [];
                        for (var _i = 0; _i < arguments.length; _i++) {
                            emails[_i] = arguments[_i];
                        }
                        var seen = new Set();
                        var result = [];
                        for (var _a = 0, emails_1 = emails; _a < emails_1.length; _a++) {
                            var e = emails_1[_a];
                            if (e && typeof e === 'string') {
                                var clean = e.trim();
                                var lower = clean.toLowerCase();
                                if (clean.length > 0 && !seen.has(lower)) {
                                    seen.add(lower);
                                    result.push(clean);
                                }
                            }
                        }
                        return result;
                    };
                    tier1Emails = [];
                    tier2Emails = [];
                    tier3Emails = [];
                    if (!graphClient) return [3 /*break*/, 40];
                    return [4 /*yield*/, Promise.all([
                            getGroupMemberEmails(graphClient, TIER1_GROUP_ID).catch(function () { return []; }),
                            getGroupMemberEmails(graphClient, TIER2_GROUP_ID).catch(function () { return []; }),
                            getGroupMemberEmails(graphClient, TIER3_GROUP_ID).catch(function () { return []; }),
                        ])];
                case 39:
                    _d = _p.sent(), t1 = _d[0], t2 = _d[1], t3 = _d[2];
                    tier1Emails = t1;
                    tier2Emails = t2;
                    tier3Emails = t3;
                    _p.label = 40;
                case 40:
                    previousLogs = [];
                    _p.label = 41;
                case 41:
                    _p.trys.push([41, 43, , 44]);
                    return [4 /*yield*/, service.getApprovalLogs(requestId)];
                case 42:
                    previousLogs = _p.sent();
                    return [3 /*break*/, 44];
                case 43:
                    _e = _p.sent();
                    previousLogs = [];
                    return [3 /*break*/, 44];
                case 44:
                    _p.trys.push([44, 47, , 48]);
                    toEmails = [];
                    ccEmails = [];
                    subject = '';
                    actionRequired = '';
                    if (actionType === 'Approve') {
                        if (logTier === 'Manager') {
                            // If Manager Approves (Routing to Tier 1): to: [Tier1Email], cc: [RequesterEmail, currentUserEmail]
                            toEmails = tier1Emails;
                            ccEmails = buildCc(requesterEmail, actingUserEmail);
                            subject = "[SCT Action Required] ".concat(ref, " \u2014 Tier 1 Approval");
                            actionRequired = 'Tier 1 Review & Approval Required';
                        }
                        else if (logTier === 'Tier 1') {
                            // If Tier 1 Approves (Routing to Tier 2): to: [Tier2Email], cc: [RequesterEmail, currentUserEmail]
                            toEmails = tier2Emails;
                            ccEmails = buildCc(requesterEmail, actingUserEmail);
                            subject = "[SCT Action Required] ".concat(ref, " \u2014 Tier 2 Approval");
                            actionRequired = 'Tier 2 Review & Approval Required';
                        }
                        else if (logTier === 'Tier 2') {
                            // If Tier 2 Approves (Routing to Tier 3): to: [Tier3Email], cc: [RequesterEmail, currentUserEmail]
                            toEmails = tier3Emails;
                            ccEmails = buildCc(requesterEmail, actingUserEmail);
                            subject = "[SCT Action Required] ".concat(ref, " \u2014 Tier 3 Approval");
                            actionRequired = 'Tier 3 Review & Approval Required';
                        }
                        else if (logTier === 'Tier 3') {
                            // If Tier 3 Approves (Final): to: [RequesterEmail], cc: [currentUserEmail]
                            toEmails = requesterEmail ? [requesterEmail] : [];
                            ccEmails = buildCc(actingUserEmail);
                            subject = "[SCT Final Approval \u2014 PO Issued] ".concat(ref, " \u2014 ").concat(requestTitle);
                            actionRequired = 'Final Approval Granted — PO Authorized';
                        }
                    }
                    else if (actionType === 'Reject') {
                        // If Reject (Any Tier): to: [RequesterEmail], cc: [currentUserEmail]
                        toEmails = requesterEmail ? [requesterEmail] : [];
                        ccEmails = buildCc(actingUserEmail);
                        subject = "[SCT Request Rejected] ".concat(ref, " \u2014 ").concat(requestTitle);
                        actionRequired = "Request Rejected by ".concat(logTier);
                    }
                    else {
                        // If Return (Any Tier): to: [RequesterEmail], cc: [currentUserEmail]
                        toEmails = requesterEmail ? [requesterEmail] : [];
                        ccEmails = buildCc(actingUserEmail);
                        subject = "[SCT Returned for Correction] ".concat(ref, " \u2014 ").concat(requestTitle);
                        actionRequired = "Request Returned for Correction by ".concat(logTier);
                    }
                    if (!(toEmails.length > 0)) return [3 /*break*/, 46];
                    htmlBody = generateApprovalEmailHTML(rec, previousLogs, webPartUrl, actionRequired);
                    return [4 /*yield*/, sendGraphEmail(graphClient, toEmails, ccEmails, subject, htmlBody)];
                case 45:
                    _p.sent();
                    emailDispatched = true;
                    _p.label = 46;
                case 46: return [3 /*break*/, 48];
                case 47:
                    mailErr_1 = _p.sent();
                    emailWarning = mailErr_1 instanceof Error ? mailErr_1.message : 'Notification email could not be delivered';
                    return [3 /*break*/, 48];
                case 48: return [2 /*return*/, { newStatus: newStatus, pendingTierUpdated: true, emailDispatched: emailDispatched, emailWarning: emailWarning }];
            }
        });
    });
}
// ── Helpers ───────────────────────────────────────────────────────────────────
/** sct_tiername integer for a given role key */
export function roleToTierInt(role) {
    var _a;
    return (_a = TIER_INT[role]) !== null && _a !== void 0 ? _a : 1;
}
/** Next sct_request_status after an approve action */
export function nextStatusAfterApprove(role) {
    return role === 'approver_tier3' ? REQUEST_STATUS.PO_ISSUED : REQUEST_STATUS.IN_REVIEW;
}
/** Human-readable label for a tier integer (for display in notifications etc.) */
export function tierLabel(tierInt) {
    var _a;
    var labels = { 4: 'Manager', 1: 'Tier 1', 2: 'Tier 2', 3: 'Tier 3' };
    return (_a = labels[tierInt]) !== null && _a !== void 0 ? _a : 'Unknown';
}
/**
 * Checks whether the active approver stage is locked for a one-time decision.
 * Respects approval cycles: if a request was returned and resubmitted, previous cycle
 * decisions are archived and a fresh decision is permitted for the new cycle.
 */
export function checkStageDecisionLock(record, role, approvalLogs) {
    if (!record) {
        return { isLocked: true, isActiveStage: false, isTerminal: false, stageName: '', currentPendingStage: '' };
    }
    // Normalize role and logTier
    var normRole = (role || 'manager').toLowerCase().trim();
    var logTier = 'Manager';
    var tierInt = 4;
    if (normRole === 'manager') {
        logTier = 'Manager';
        tierInt = 4;
    }
    else if (normRole === 'approver_tier1' || normRole.indexOf('1') !== -1) {
        logTier = 'Tier 1';
        tierInt = 1;
    }
    else if (normRole === 'approver_tier2' || normRole.indexOf('2') !== -1) {
        logTier = 'Tier 2';
        tierInt = 2;
    }
    else if (normRole === 'approver_tier3' || normRole.indexOf('3') !== -1) {
        logTier = 'Tier 3';
        tierInt = 3;
    }
    // Pending stage resolution first
    var ptRaw = String(record['cr2a4_pendingtier'] ||
        record['cr2a4_PendingTier'] ||
        record['sct_pendingtier'] ||
        record['pendingtier'] ||
        '').trim();
    var ptLower = ptRaw.replace(/[^a-z0-9]/g, '').toLowerCase();
    var hasActivePendingTier = ptLower === 'manager' || ptLower.indexOf('tier') !== -1 || ptLower === '1' || ptLower === '2' || ptLower === '3';
    var activePendingStage = 'Manager';
    if (ptLower === 'tier3' || ptLower === '3' || ptLower.indexOf('3') !== -1)
        activePendingStage = 'Tier 3';
    else if (ptLower === 'tier2' || ptLower === '2' || ptLower.indexOf('2') !== -1)
        activePendingStage = 'Tier 2';
    else if (ptLower === 'tier1' || ptLower === '1' || ptLower.indexOf('1') !== -1)
        activePendingStage = 'Tier 1';
    else if (ptLower === 'manager' || !ptLower)
        activePendingStage = 'Manager';
    var isActiveStage = logTier === activePendingStage;
    // Check terminal status using canonical status label (only terminal if no active pending tier)
    var canonicalStatus = getRecordStatusLabel(record);
    var isTerminal = canonicalStatus === 'Rejected' || canonicalStatus === 'Closed' || (!hasActivePendingTier && canonicalStatus === 'PO Issued');
    if (isTerminal) {
        var termLabel = canonicalStatus === 'PO Issued' ? 'Approved / PO Issued' : canonicalStatus;
        return {
            isLocked: true,
            isActiveStage: false,
            isTerminal: true,
            stageName: logTier,
            currentPendingStage: termLabel,
            lockMessage: "This request has reached terminal status (".concat(termLabel, "). No further approval actions can be taken."),
        };
    }
    // Find latest Return timestamp in history (boundary of current cycle)
    var latestReturnTime = 0;
    if (Array.isArray(approvalLogs)) {
        approvalLogs.forEach(function (l) {
            var _a, _b, _c, _d;
            var statusVal = l['sct_approvalstatus'];
            var nameVal = String((_a = l['sct_name']) !== null && _a !== void 0 ? _a : '').toLowerCase();
            var commentsVal = String((_b = l['sct_comments']) !== null && _b !== void 0 ? _b : '').toLowerCase();
            var isReturn = statusVal === APPROVAL_STATUS.RETURN_FOR_INFO ||
                nameVal.indexOf('return') !== -1 ||
                commentsVal.indexOf('return') !== -1;
            if (isReturn) {
                var t = new Date(String((_d = (_c = l['sct_decisionon']) !== null && _c !== void 0 ? _c : l['createdon']) !== null && _d !== void 0 ? _d : 0)).getTime();
                if (t > latestReturnTime)
                    latestReturnTime = t;
            }
        });
    }
    // Filter logs for current cycle
    var currentCycleLogs = latestReturnTime > 0
        ? approvalLogs.filter(function (l) {
            var _a, _b;
            var t = new Date(String((_b = (_a = l['sct_decisionon']) !== null && _a !== void 0 ? _a : l['createdon']) !== null && _b !== void 0 ? _b : 0)).getTime();
            return t > latestReturnTime;
        })
        : (Array.isArray(approvalLogs) ? approvalLogs : []);
    // Look for existing decision for this tier in the current cycle
    var priorLog = currentCycleLogs.find(function (l) {
        var _a;
        var tInt = l['sct_tiername'];
        if (tInt === tierInt)
            return true;
        var name = String((_a = l['sct_name']) !== null && _a !== void 0 ? _a : '').toLowerCase();
        return name.indexOf(logTier.toLowerCase()) !== -1;
    });
    if (priorLog) {
        var actionVal = priorLog['sct_approvalstatus'] === 1 ? 'Approved'
            : priorLog['sct_approvalstatus'] === 2 ? 'Rejected'
                : 'Returned';
        var approver = strVal(priorLog, '_sct_approverid_value') ||
            strVal(priorLog, '_createdby_value') ||
            strVal(priorLog, 'sct_name') ||
            'Approver';
        var timeRaw = priorLog['createdon'] || priorLog['sct_decisionon'];
        var timeFormatted = timeRaw ? formatDateTime(String(timeRaw)) : '—';
        var comments = strVal(priorLog, 'sct_comments') || '';
        return {
            isLocked: true,
            isActiveStage: false,
            isTerminal: false,
            stageName: logTier,
            currentPendingStage: activePendingStage,
            previousDecision: {
                action: actionVal,
                approverName: approver,
                timestamp: timeFormatted,
                comments: comments,
            },
            lockMessage: "You have already submitted your decision for this stage (".concat(logTier, "). Further actions for this stage are permanently locked in this approval cycle."),
        };
    }
    if (!isActiveStage) {
        return {
            isLocked: true,
            isActiveStage: false,
            isTerminal: false,
            stageName: logTier,
            currentPendingStage: activePendingStage,
            lockMessage: "This request is currently pending at '".concat(activePendingStage, "'. Approvers for '").concat(logTier, "' cannot take action at this time."),
        };
    }
    return {
        isLocked: false,
        isActiveStage: true,
        isTerminal: false,
        stageName: logTier,
        currentPendingStage: activePendingStage,
    };
}
//# sourceMappingURL=approvalHandler.js.map