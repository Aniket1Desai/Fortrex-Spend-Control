/**
 * ApprovalRoutingService — Schedule A approval routing engine.
 *
 * Implements the two-stage split-approval chain:
 *  Phase 1 — Pre-Request: which levels the requester must attach before submitting.
 *  Phase 2 — Post-Sourcing: does Procurement need FP&A / CFO review before issuing a PO?
 *  Role Resolution — dynamically resolves L1-L4 approver identities from sct_role,
 *                    including active-delegate substitution and Rule R10 self-approval escalation.
 */
import { __awaiter, __generator } from "tslib";
// ── Spend thresholds ──────────────────────────────────────────────────────────
var T_L1 = 2500; // L1 (first-line) ceiling
var T_L2 = 10000; // L2 (department head) ceiling
var T_LEGAL = 50000; // Legal Review required above this value
// ── Service ───────────────────────────────────────────────────────────────────
var ApprovalRoutingService = /** @class */ (function () {
    function ApprovalRoutingService(_svc) {
        this._svc = _svc;
    }
    // ── Phase 1: Pre-Request Approval Calculator ──────────────────────────────
    /**
     * Returns the approval level(s) the requester must collect BEFORE submission.
     *
     * Schedule A rules:
     *  New supplier (any value) → L1 + L2 + L3
     *  Known supplier, value ≤ $2,500                 → L1 only
     *  Known supplier, $2,500 < value ≤ $10,000        → L2 only
     *  Known supplier, value > $10,000                 → L3 only
     */
    ApprovalRoutingService.prototype.calculatePreApprovals = function (estimatedValue, isNewSupplier) {
        if (isNewSupplier) {
            return {
                levels: ['L1', 'L2', 'L3'],
                description: 'New supplier — pre-approval signatures required from L1 (Line Manager), L2 (Department Head) and L3 (CFO).',
            };
        }
        if (estimatedValue <= T_L1) {
            return {
                levels: ['L1'],
                description: "Value \u2264 $".concat(T_L1.toLocaleString(), " \u2014 L1 (Line Manager) sign-off required."),
            };
        }
        if (estimatedValue <= T_L2) {
            return {
                levels: ['L2'],
                description: "Value $".concat(T_L1.toLocaleString(), "\u2013$").concat(T_L2.toLocaleString(), " \u2014 L2 (Department Head) sign-off required."),
            };
        }
        return {
            levels: ['L3'],
            description: "Value > $".concat(T_L2.toLocaleString(), " \u2014 L3 (CFO) sign-off required."),
        };
    };
    // ── Phase 2: Post-Sourcing Approval Gate ──────────────────────────────────
    /**
     * Returns what additional post-sourcing approvals Procurement must trigger.
     *
     * Rules:
     *  value > $10k OR new supplier  → FP&A Review required first.
     *  After FP&A: always route to CFO if either condition is true.
     *  value > $50k                  → Legal Review flag raised additionally.
     */
    ApprovalRoutingService.prototype.calculatePostSourcingRequirements = function (estimatedValue, isNewSupplier) {
        var needsGate = estimatedValue > T_L2 || isNewSupplier;
        var needsLegal = estimatedValue > T_LEGAL;
        var descriptions = [];
        if (!needsGate) {
            descriptions.push('Within direct Procurement authority — PO may be issued immediately.');
        }
        else {
            if (isNewSupplier)
                descriptions.push('New supplier — FP&A review and CFO approval required before PO.');
            if (estimatedValue > T_L2)
                descriptions.push("Value > $".concat(T_L2.toLocaleString(), " \u2014 FP&A review and CFO approval required."));
            if (needsLegal)
                descriptions.push("Value > $".concat(T_LEGAL.toLocaleString(), " \u2014 Legal Review flag raised."));
        }
        return {
            needsFPAReview: needsGate,
            needsCFOApproval: needsGate,
            needsLegalReview: needsLegal,
            description: descriptions.join(' '),
        };
    };
    // ── Dynamic Role Resolution + Rule R10 ───────────────────────────────────
    /**
     * Resolves L1–L4 approvers from the sct_role table for the given site.
     * Applies:
     *  — Active-delegate substitution (today within effectivedate..expirydate)
     *  — Rule R10: if approver GUID = requesterSystemUserId, escalate to next level.
     *
     * @param siteId             GUID value from _sct_site_value (the site's sct_sitesmasterid)
     * @param requesterSystemUserId  Dataverse systemuserid GUID of the person submitting
     */
    ApprovalRoutingService.prototype.resolveApprovers = function (siteId, requesterSystemUserId) {
        return __awaiter(this, void 0, void 0, function () {
            var escalations, warnings, rawRoles, _a, siteRoles, byLevel, levels, resolved, _i, levels_1, level, roleRecords, role, now, effDate, expDate, delegateActive, personGuid, delegateGuid, activeGuid, activeName, nextLevel;
            var _b, _c, _d, _e, _f, _g, _h;
            return __generator(this, function (_j) {
                switch (_j.label) {
                    case 0:
                        escalations = [];
                        warnings = [];
                        rawRoles = [];
                        _j.label = 1;
                    case 1:
                        _j.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._svc.getRawRoles(true)];
                    case 2:
                        rawRoles = _j.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _j.sent();
                        warnings.push('Could not load roles from Dataverse — approvers not resolved.');
                        return [2 /*return*/, { approvers: [], escalations: escalations, warnings: warnings }];
                    case 4:
                        siteRoles = rawRoles.filter(function (r) {
                            var id = r['_sct_siteid_value'];
                            return id && String(id).toLowerCase() === siteId.toLowerCase();
                        });
                        if (!siteRoles.length) {
                            warnings.push("No roles found for site ".concat(siteId, ". Verify sct_role records in Dataverse."));
                        }
                        byLevel = {};
                        siteRoles.forEach(function (r) {
                            var _a, _b;
                            var level = String((_b = (_a = r['sct_approvallevel']) !== null && _a !== void 0 ? _a : r['sct_approval_level']) !== null && _b !== void 0 ? _b : '').toUpperCase();
                            if (!byLevel[level])
                                byLevel[level] = [];
                            byLevel[level].push(r);
                        });
                        levels = ['L1', 'L2', 'L3', 'L4'];
                        resolved = [];
                        for (_i = 0, levels_1 = levels; _i < levels_1.length; _i++) {
                            level = levels_1[_i];
                            roleRecords = (_b = byLevel[level]) !== null && _b !== void 0 ? _b : [];
                            if (!roleRecords.length)
                                continue;
                            role = roleRecords[0];
                            now = new Date();
                            effDate = role['sct_effectivedate'] ? new Date(String(role['sct_effectivedate'])) : null;
                            expDate = role['sct_expirydate'] ? new Date(String(role['sct_expirydate'])) : null;
                            delegateActive = role['_sct_delegateid_value'] != null &&
                                String(role['_sct_delegateid_value']).trim() !== '' &&
                                (effDate == null || now >= effDate) &&
                                (expDate == null || now <= expDate);
                            personGuid = String((_c = role['_sct_personid_value']) !== null && _c !== void 0 ? _c : '');
                            delegateGuid = delegateActive ? String((_d = role['_sct_delegateid_value']) !== null && _d !== void 0 ? _d : '') : '';
                            activeGuid = delegateActive ? delegateGuid : personGuid;
                            activeName = delegateActive
                                ? String((_e = role['_sct_delegateid_value@OData.Community.Display.V1.FormattedValue']) !== null && _e !== void 0 ? _e : delegateGuid)
                                : String((_f = role['_sct_personid_value@OData.Community.Display.V1.FormattedValue']) !== null && _f !== void 0 ? _f : personGuid);
                            if (!activeGuid) {
                                warnings.push("".concat(level, " role found but no approver GUID \u2014 check sct_role record ").concat(String((_g = role['sct_roleid']) !== null && _g !== void 0 ? _g : ''), "."));
                                continue;
                            }
                            // Rule R10: self-approval conflict — escalate to next level
                            if (activeGuid.toLowerCase() === requesterSystemUserId.toLowerCase()) {
                                nextLevel = this._nextLevel(level);
                                escalations.push("Rule R10: ".concat(level, " approver (").concat(activeName, ") is the requester \u2014 escalated to ").concat(nextLevel !== null && nextLevel !== void 0 ? nextLevel : 'Board (no higher level available)', "."));
                                if (!nextLevel) {
                                    warnings.push("".concat(level, " conflict cannot be escalated further \u2014 no higher level defined for this site."));
                                }
                                continue; // skip this level; the escalation will be picked up when processing the next level
                            }
                            resolved.push({
                                level: level,
                                systemUserId: activeGuid,
                                name: activeName,
                                isDelegate: delegateActive,
                                wasEscalated: false,
                                roleId: String((_h = role['sct_roleid']) !== null && _h !== void 0 ? _h : ''),
                            });
                        }
                        return [2 /*return*/, { approvers: resolved, escalations: escalations, warnings: warnings }];
                }
            });
        });
    };
    // ── Phase 8: Convenience matrix ──────────────────────────────────────────
    /**
     * Computes all three approval phases in one call and returns a combined matrix.
     * Useful for displaying the full approval picture on the Triage and FP&A panels.
     */
    ApprovalRoutingService.prototype.computeApprovalMatrix = function (siteId, requesterSystemUserId, estimatedValue, isNewSupplier) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, preApprovals, postSourcing, resolvedApprovers;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, Promise.all([
                            Promise.resolve(this.calculatePreApprovals(estimatedValue, isNewSupplier)),
                            Promise.resolve(this.calculatePostSourcingRequirements(estimatedValue, isNewSupplier)),
                            this.resolveApprovers(siteId, requesterSystemUserId),
                        ])];
                    case 1:
                        _a = _b.sent(), preApprovals = _a[0], postSourcing = _a[1], resolvedApprovers = _a[2];
                        return [2 /*return*/, { preApprovals: preApprovals, postSourcing: postSourcing, resolvedApprovers: resolvedApprovers }];
                }
            });
        });
    };
    // ── Private helpers ───────────────────────────────────────────────────────
    ApprovalRoutingService.prototype._nextLevel = function (current) {
        var order = ['L1', 'L2', 'L3', 'L4'];
        var idx = order.indexOf(current);
        return idx >= 0 && idx < order.length - 1 ? order[idx + 1] : null;
    };
    return ApprovalRoutingService;
}());
export { ApprovalRoutingService };
//# sourceMappingURL=ApprovalRoutingService.js.map