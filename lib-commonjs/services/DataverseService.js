"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DataverseService = exports.BASE_URL = void 0;
exports.scrubPayload = scrubPayload;
exports.strVal = strVal;
var tslib_1 = require("tslib");
var sp_http_1 = require("@microsoft/sp-http");
/** Fallback used only when the web part property has not been set */
var DEFAULT_DATAVERSE_HOST = 'https://org1bec48f0.crm8.dynamics.com';
/** @deprecated Use DataverseService.baseUrl instead — kept for any legacy direct imports */
exports.BASE_URL = "".concat(DEFAULT_DATAVERSE_HOST, "/api/data/v9.2/");
/**
 * Scrubs all undefined, null, or empty string navigation properties / fields
 * prior to dispatching HTTP POST or PATCH to Dataverse.
 */
function scrubPayload(payload) {
    var clean = tslib_1.__assign({}, payload);
    Object.keys(clean).forEach(function (key) {
        var val = clean[key];
        if (val === undefined || val === null || val === '') {
            if (key.indexOf('@odata.bind') !== -1 || val === undefined) {
                delete clean[key];
            }
        }
    });
    return clean;
}
var OData_HEADERS = {
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
    'Accept': 'application/json',
    'Content-Type': 'application/json; charset=utf-8',
    // Returns "fieldname@OData.Community.Display.V1.FormattedValue" alongside each raw value
    'Prefer': 'odata.include-annotations="*"',
    // Prevent browser/proxy caching of OData GET responses so post-approval reloads
    // always receive the latest committed Dataverse data, not a stale cached copy.
    'Cache-Control': 'no-cache',
    'Pragma': 'no-cache',
};
/** Pull GUID out of the OData-EntityId response header URL */
function extractGuidFromEntityId(entityIdHeader) {
    var match = entityIdHeader.match(/\(([0-9a-f-]{36})\)/i);
    if (!match)
        throw new Error("Cannot parse entity id from header: ".concat(entityIdHeader));
    return match[1];
}
/** Safe string cast for a raw OData field value */
function strVal(record, field) {
    // Prefer the OData formatted-value annotation when it exists
    var annotated = record["".concat(field, "@OData.Community.Display.V1.FormattedValue")];
    if (annotated !== undefined && annotated !== null)
        return String(annotated);
    var raw = record[field];
    return raw !== undefined && raw !== null ? String(raw) : '';
}
var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function assertGuid(value, label) {
    if (!GUID_RE.test(value)) {
        throw new Error("Invalid ".concat(label, ": expected a GUID, got \"").concat(value.slice(0, 40), "\""));
    }
}
function encodeODataString(val) {
    return val.replace(/'/g, "''");
}
var DataverseService = /** @class */ (function () {
    function DataverseService(_client, dataverseHost) {
        this._client = _client;
        var host = (dataverseHost || DEFAULT_DATAVERSE_HOST).replace(/\/$/, '');
        this.baseUrl = "".concat(host, "/api/data/v9.2/");
    }
    // ── Health check ────────────────────────────────────────────────────────────
    /**
     * Verifies the SCT tables are reachable in this environment.
     * Throws a descriptive error if the entity set is missing or access is denied.
     */
    DataverseService.prototype.testConnection = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, status_1;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._fetch("".concat(this.baseUrl, "sct_requests?$top=1&$select=sct_requestid"))];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            status_1 = response.status;
                            if (status_1 === 404) {
                                throw new Error("SCT tables not found in this Dataverse environment (HTTP 404). " +
                                    "Please ensure the Fortrex SCT solution has been imported and published in Power Apps before using this web part.");
                            }
                            if (status_1 === 401 || status_1 === 403) {
                                throw new Error("Access denied to Dataverse (HTTP ".concat(status_1, "). ") +
                                    "Ensure the AAD API permission (user_impersonation) has been approved in SharePoint Admin Centre \u2192 API Access.");
                            }
                            throw new Error("Dataverse connectivity check failed (HTTP ".concat(status_1, ")."));
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    // ── History logs ────────────────────────────────────────────────────────────
    DataverseService.prototype.getApprovalLogs = function (requestId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var select, logs;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        assertGuid(requestId, 'requestId');
                        select = [
                            'sct_approvallogid', 'sct_name',
                            'sct_approvalstatus', 'sct_tiername',
                            '_sct_approverid_value', '_createdby_value',
                            'sct_decisionon', 'sct_comments', 'createdon',
                        ].join(',');
                        return [4 /*yield*/, this._get("sct_approvallogs?$select=".concat(select, "&$filter=_sct_requestid_value eq ").concat(requestId, "&$orderby=createdon desc"))];
                    case 1:
                        logs = _a.sent();
                        // Explicit client-side sort: newest first by createdon
                        return [2 /*return*/, Array.isArray(logs)
                                ? tslib_1.__spreadArray([], logs, true).sort(function (a, b) {
                                    var _a, _b, _c, _d, _e, _f;
                                    var sA = String((_b = (_a = a['createdon']) !== null && _a !== void 0 ? _a : a['sct_decisionon']) !== null && _b !== void 0 ? _b : '').trim();
                                    var sB = String((_d = (_c = b['createdon']) !== null && _c !== void 0 ? _c : b['sct_decisionon']) !== null && _d !== void 0 ? _d : '').trim();
                                    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(sA))
                                        sA += 'Z';
                                    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(sB))
                                        sB += 'Z';
                                    var tA = new Date(sA).getTime() || 0;
                                    var tB = new Date(sB).getTime() || 0;
                                    if (tB !== tA)
                                        return tB - tA;
                                    return String((_e = b['sct_name']) !== null && _e !== void 0 ? _e : '').localeCompare(String((_f = a['sct_name']) !== null && _f !== void 0 ? _f : ''));
                                })
                                : []];
                }
            });
        });
    };
    DataverseService.prototype.getAuditLogs = function (requestId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var select, logs;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        assertGuid(requestId, 'requestId');
                        select = [
                            'sct_auditlogid', 'sct_name',
                            'sct_action', '_sct_actionby_value',
                            'sct_description', 'sct_oldvalue', 'sct_newvalue', 'createdon',
                        ].join(',');
                        return [4 /*yield*/, this._get("sct_auditlogs?$select=".concat(select, "&$filter=_sct_requestid_value eq ").concat(requestId, "&$orderby=createdon desc"))];
                    case 1:
                        logs = _a.sent();
                        // Explicit client-side sort: newest first by createdon
                        return [2 /*return*/, Array.isArray(logs)
                                ? tslib_1.__spreadArray([], logs, true).sort(function (a, b) {
                                    var _a, _b, _c, _d;
                                    var sA = String((_a = a['createdon']) !== null && _a !== void 0 ? _a : '').trim();
                                    var sB = String((_b = b['createdon']) !== null && _b !== void 0 ? _b : '').trim();
                                    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(sA))
                                        sA += 'Z';
                                    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(sB))
                                        sB += 'Z';
                                    var tA = new Date(sA).getTime() || 0;
                                    var tB = new Date(sB).getTime() || 0;
                                    if (tB !== tA)
                                        return tB - tA;
                                    return String((_c = b['sct_name']) !== null && _c !== void 0 ? _c : '').localeCompare(String((_d = a['sct_name']) !== null && _d !== void 0 ? _d : ''));
                                })
                                : []];
                }
            });
        });
    };
    DataverseService.prototype.createAuditLog = function (payload) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._post('sct_auditlogs', payload)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    DataverseService.prototype.createApprovalLog = function (payload) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var cleanPayload;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        cleanPayload = scrubPayload(payload);
                        return [4 /*yield*/, this._post('sct_approvallogs', cleanPayload)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Returns the GUIDs of "In Review" requests pending at the given tier,
     * using cr2a4_pendingtier (confirmed column, publisher prefix cr2a4_).
     * Single OData query — no N+1 log fetches.
     *
     * cr2a4_pendingtier values: 'Tier 1' | 'Tier 2' | 'Tier 3' | 'Manager' | null
     */
    DataverseService.prototype.getTierPendingRequestIds = function (tier) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var tierLabel, altLabel, filter, records, _a, all, matched, _b;
            return tslib_1.__generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        tierLabel = tier === 1 ? 'Tier 1' : tier === 2 ? 'Tier 2' : 'Tier 3';
                        altLabel = "Tier".concat(tier);
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 3, , 4]);
                        filter = encodeURIComponent("(sct_request_status eq 854000003 or sct_request_status eq 854000002) and (cr2a4_pendingtier eq '".concat(tierLabel, "' or cr2a4_pendingtier eq '").concat(altLabel, "' or cr2a4_pendingtier eq '").concat(tier, "')"));
                        return [4 /*yield*/, this._get("sct_requests?$select=sct_requestid&$filter=".concat(filter, "&$top=500"))];
                    case 2:
                        records = _c.sent();
                        if (records && records.length > 0) {
                            return [2 /*return*/, records.map(function (r) { var _a; return String((_a = r['sct_requestid']) !== null && _a !== void 0 ? _a : ''); }).filter(function (id) { return GUID_RE.test(id); })];
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _c.sent();
                        return [3 /*break*/, 4];
                    case 4:
                        _c.trys.push([4, 6, , 7]);
                        return [4 /*yield*/, this.getRawRequests().catch(function () { return []; })];
                    case 5:
                        all = _c.sent();
                        matched = all.filter(function (r) {
                            var _a, _b;
                            var rawStatus = typeof r['sct_request_status'] === 'number'
                                ? r['sct_request_status']
                                : parseInt(String((_b = (_a = r['sct_request_status']) !== null && _a !== void 0 ? _a : r['sct_requeststatus']) !== null && _b !== void 0 ? _b : ''), 10);
                            var statusLabel = (strVal(r, 'sct_request_status') || strVal(r, 'sct_requeststatus') || '').toLowerCase();
                            var isReview = rawStatus === 854000003 || rawStatus === 854000002 || statusLabel.indexOf('in review') !== -1 || statusLabel.indexOf('tier') !== -1;
                            if (!isReview)
                                return false;
                            var pt = String(r['cr2a4_pendingtier'] || r['sct_pendingtier'] || r['pendingtier'] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
                            return pt === "tier".concat(tier) || pt === "".concat(tier) || statusLabel.indexOf("tier ".concat(tier)) !== -1 || statusLabel.indexOf("tier".concat(tier)) !== -1;
                        });
                        return [2 /*return*/, matched.map(function (r) { var _a; return String((_a = r['sct_requestid']) !== null && _a !== void 0 ? _a : ''); }).filter(function (id) { return GUID_RE.test(id); })];
                    case 6:
                        _b = _c.sent();
                        return [2 /*return*/, []];
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    /** Convenience wrapper — badge count only */
    DataverseService.prototype.getTier1PendingCount = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                return [2 /*return*/, this.getTierPendingRequestIds(1).then(function (ids) { return ids.length; }).catch(function () { return 0; })];
            });
        });
    };
    /**
     * Returns the count of requests currently pending approval for this manager.
     */
    DataverseService.prototype.getManagerPendingCount = function (systemUserId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var directReportIds, safeUserId_1, reportsSet_1, managerFilter, reportsOr, filterB1, filterB2, _a, reqs1, reqs2, allRaw, approvedIdsArr, approvedIds_1, seen, candidateReqs, _i, _b, r, id, _c;
            var _d;
            return tslib_1.__generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        assertGuid(systemUserId, 'system user id');
                        _e.label = 1;
                    case 1:
                        _e.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, this.getDirectReportUserIds(systemUserId).catch(function () { return []; })];
                    case 2:
                        directReportIds = _e.sent();
                        safeUserId_1 = systemUserId.toLowerCase().replace(/[{}]/g, '');
                        reportsSet_1 = new Set(directReportIds.map(function (id) { return id.toLowerCase().replace(/[{}]/g, ''); }));
                        managerFilter = "_cr2a4_manager_value eq ".concat(systemUserId);
                        if (directReportIds.length > 0) {
                            reportsOr = directReportIds.slice(0, 15).map(function (id) { return "_sct_requester_value eq ".concat(id); }).join(' or ');
                            managerFilter = "(".concat(managerFilter, " or ").concat(reportsOr, ")");
                        }
                        filterB1 = encodeURIComponent("(sct_request_status eq 854000002 or sct_request_status eq 854000003 or sct_request_status eq 854000001) and ".concat(managerFilter));
                        filterB2 = encodeURIComponent("(sct_request_status eq 854000002 or sct_request_status eq 854000003 or sct_request_status eq 854000001) and _sct_manager_value eq ".concat(systemUserId));
                        return [4 /*yield*/, Promise.all([
                                this.getRawRequests(filterB1).catch(function () { return []; }),
                                this.getRawRequests(filterB2).catch(function () { return []; }),
                                this.getRawRequests().catch(function () { return []; }),
                                this.getMyApprovedRequestIds(systemUserId).catch(function () { return []; }),
                            ])];
                    case 3:
                        _a = _e.sent(), reqs1 = _a[0], reqs2 = _a[1], allRaw = _a[2], approvedIdsArr = _a[3];
                        approvedIds_1 = new Set(approvedIdsArr);
                        seen = new Set();
                        candidateReqs = [];
                        for (_i = 0, _b = tslib_1.__spreadArray(tslib_1.__spreadArray(tslib_1.__spreadArray([], reqs1, true), reqs2, true), allRaw, true); _i < _b.length; _i++) {
                            r = _b[_i];
                            id = String((_d = r['sct_requestid']) !== null && _d !== void 0 ? _d : '');
                            if (id && !seen.has(id)) {
                                seen.add(id);
                                candidateReqs.push(r);
                            }
                        }
                        return [2 /*return*/, candidateReqs.filter(function (r) {
                                var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k;
                                var id = String((_a = r['sct_requestid']) !== null && _a !== void 0 ? _a : '');
                                if (!id || approvedIds_1.has(id))
                                    return false;
                                var rawReq = (_d = (_c = (_b = r['_sct_requester_value']) !== null && _b !== void 0 ? _b : r['_createdby_value']) !== null && _c !== void 0 ? _c : r['sct_requester']) !== null && _d !== void 0 ? _d : '';
                                var reqId = rawReq ? String(rawReq).toLowerCase().replace(/[{}]/g, '') : '';
                                var createdById = r['_createdby_value'] ? String(r['_createdby_value']).toLowerCase().replace(/[{}]/g, '') : '';
                                if (reqId === safeUserId_1 || createdById === safeUserId_1)
                                    return false;
                                var rawMgr = (_h = (_g = (_f = (_e = r['_sct_manager_value']) !== null && _e !== void 0 ? _e : r['_cr2a4_manager_value']) !== null && _f !== void 0 ? _f : r['_manager_value']) !== null && _g !== void 0 ? _g : r['sct_manager']) !== null && _h !== void 0 ? _h : '';
                                var mgrId = rawMgr ? String(rawMgr).toLowerCase().replace(/[{}]/g, '') : '';
                                var mgrFormatted = (strVal(r, '_cr2a4_manager_value') || strVal(r, '_sct_manager_value') || strVal(r, 'cr2a4_manager') || strVal(r, 'manager') || '').toLowerCase();
                                var isDirectReport = Boolean(reportsSet_1.has(reqId) || (createdById && reportsSet_1.has(createdById)));
                                var isManagerOfReq = (mgrId === safeUserId_1) || isDirectReport;
                                if (!isManagerOfReq)
                                    return false;
                                var rawStatus = typeof r['sct_request_status'] === 'number'
                                    ? r['sct_request_status']
                                    : parseInt(String((_k = (_j = r['sct_request_status']) !== null && _j !== void 0 ? _j : r['sct_requeststatus']) !== null && _k !== void 0 ? _k : ''), 10);
                                var statusLabel = (strVal(r, 'sct_request_status') || strVal(r, 'sct_requeststatus') || '').toLowerCase();
                                var isPendingStatus = rawStatus === 854000002 || rawStatus === 854000003 || rawStatus === 854000001 ||
                                    statusLabel === 'submitted' || statusLabel === 'in review' ||
                                    statusLabel.indexOf('submitted') !== -1 || statusLabel.indexOf('in review') !== -1;
                                if (!isPendingStatus)
                                    return false;
                                var pt = String(r['cr2a4_pendingtier'] || r['sct_pendingtier'] || r['pendingtier'] || '').toLowerCase().trim();
                                var isPendingWithManager = pt === 'manager' || !pt;
                                return isPendingWithManager;
                            }).length];
                    case 4:
                        _c = _e.sent();
                        return [2 /*return*/, 0];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Returns the set of request IDs that the given system user has already approved
     * at the Manager tier. Used to distinguish Pending (not yet acted) from
     * already-forwarded requests in the manager's view.
     *
     * Queries approval logs where this user is either the assigned approver (_sct_approverid_value)
     * or the creator (_createdby_value), and filters for sct_tiername = 1 AND sct_approvalstatus = 1.
     */
    DataverseService.prototype.getMyApprovedRequestIds = function (systemUserId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var filter, logs;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        assertGuid(systemUserId, 'system user id');
                        filter = encodeURIComponent("(_sct_approverid_value eq ".concat(systemUserId, " or _createdby_value eq ").concat(systemUserId, ") and (sct_tiername eq 4 or sct_tiername eq 1) and sct_approvalstatus eq 1"));
                        return [4 /*yield*/, this._get("sct_approvallogs?$select=_sct_requestid_value,sct_name,sct_tiername,sct_approvalstatus&$filter=".concat(filter, "&$top=500")).catch(function () { return []; })];
                    case 1:
                        logs = _a.sent();
                        return [2 /*return*/, Array.from(new Set(logs
                                .filter(function (l) {
                                var _a;
                                var t = l['sct_tiername'];
                                var status = l['sct_approvalstatus'];
                                if (status !== 1)
                                    return false;
                                if (t === 4)
                                    return true;
                                if (t === 1) {
                                    var name_1 = String((_a = l['sct_name']) !== null && _a !== void 0 ? _a : '').toLowerCase();
                                    return name_1.indexOf('manager') !== -1;
                                }
                                return false;
                            })
                                .map(function (l) { var _a; return String((_a = l['_sct_requestid_value']) !== null && _a !== void 0 ? _a : ''); })
                                .filter(function (id) { return GUID_RE.test(id); })))];
                }
            });
        });
    };
    /** Returns IDs of requests this approver approved at the given tier. Supports new and legacy tier mappings. */
    DataverseService.prototype.getMyTierApprovedRequestIds = function (systemUserId, tierIndex) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var filter, logs;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        assertGuid(systemUserId, 'system user id');
                        filter = encodeURIComponent("(_sct_approverid_value eq ".concat(systemUserId, " or _createdby_value eq ").concat(systemUserId, ") and sct_tiername eq ").concat(tierIndex, " and sct_approvalstatus eq 1"));
                        return [4 /*yield*/, this._get("sct_approvallogs?$select=_sct_requestid_value,sct_name,sct_tiername,sct_approvalstatus&$filter=".concat(filter, "&$top=500")).catch(function () { return []; })];
                    case 1:
                        logs = _a.sent();
                        return [2 /*return*/, Array.from(new Set(logs
                                .filter(function (l) {
                                var _a;
                                var t = l['sct_tiername'];
                                var status = l['sct_approvalstatus'];
                                if (status !== 1)
                                    return false;
                                if (tierIndex === 1) {
                                    // Ensure this is not a Manager log stamped with legacy tiername 1
                                    var name_2 = String((_a = l['sct_name']) !== null && _a !== void 0 ? _a : '').toLowerCase();
                                    if (name_2.indexOf('manager') !== -1)
                                        return false;
                                }
                                return t === tierIndex;
                            })
                                .map(function (l) { var _a; return String((_a = l['_sct_requestid_value']) !== null && _a !== void 0 ? _a : ''); })
                                .filter(function (id) { return GUID_RE.test(id); })))];
                }
            });
        });
    };
    /**
     * Three-source data fetch for the Manager role, merged and deduplicated.
     *
     * Fetch A — authored: sct_requests where _sct_requester_value = userId
     * Fetch B — pending queue: sct_requests where status = In Review AND _cr2a4_manager_value = userId
     * Fetch C — approval history: sct_approvallogs where _createdby_value = userId
     *            → extract request GUIDs → fetch those requests
     *
     * Returns a single deduplicated array sorted newest-first.
     */
    DataverseService.prototype.getManagerMasterRequests = function (systemUserId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var allLogs, loggedIds, directReportIds, managerFilter, reportsOr, filterA, filterB1, filterB2, _a, fetchA, fetchB1, fetchB2, fetchC, allRaw, seen, merged, _i, _b, r, id;
            var _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        assertGuid(systemUserId, 'system user id');
                        return [4 /*yield*/, this._get("sct_approvallogs?$select=_sct_requestid_value" +
                                "&$filter=_createdby_value eq ".concat(systemUserId, "&$top=500")).catch(function () { return []; })];
                    case 1:
                        allLogs = _d.sent();
                        loggedIds = Array.from(new Set(allLogs
                            .map(function (l) { var _a; return String((_a = l['_sct_requestid_value']) !== null && _a !== void 0 ? _a : ''); })
                            .filter(function (id) { return GUID_RE.test(id); }))).slice(0, 50);
                        return [4 /*yield*/, this.getDirectReportUserIds(systemUserId).catch(function () { return []; })];
                    case 2:
                        directReportIds = _d.sent();
                        managerFilter = "_cr2a4_manager_value eq ".concat(systemUserId);
                        if (directReportIds.length > 0) {
                            reportsOr = directReportIds.slice(0, 15).map(function (id) { return "_sct_requester_value eq ".concat(id); }).join(' or ');
                            managerFilter = "(".concat(managerFilter, " or ").concat(reportsOr, ")");
                        }
                        filterA = encodeURIComponent("_sct_requester_value eq ".concat(systemUserId));
                        filterB1 = encodeURIComponent("(sct_request_status eq 854000002 or sct_request_status eq 854000003 or sct_request_status eq 854000001) and ".concat(managerFilter));
                        filterB2 = encodeURIComponent("(sct_request_status eq 854000002 or sct_request_status eq 854000003 or sct_request_status eq 854000001) and _sct_manager_value eq ".concat(systemUserId));
                        return [4 /*yield*/, Promise.all([
                                this.getRawRequests(filterA).catch(function () { return []; }),
                                this.getRawRequests(filterB1).catch(function () { return []; }),
                                this.getRawRequests(filterB2).catch(function () { return []; }),
                                loggedIds.length > 0
                                    ? this.getRawRequests(encodeURIComponent(loggedIds.map(function (id) { return "sct_requestid eq ".concat(id); }).join(' or '))).catch(function () { return []; })
                                    : Promise.resolve([]),
                                this.getRawRequests().catch(function () { return []; }),
                            ])];
                    case 3:
                        _a = _d.sent(), fetchA = _a[0], fetchB1 = _a[1], fetchB2 = _a[2], fetchC = _a[3], allRaw = _a[4];
                        seen = new Set();
                        merged = [];
                        for (_i = 0, _b = tslib_1.__spreadArray(tslib_1.__spreadArray(tslib_1.__spreadArray(tslib_1.__spreadArray(tslib_1.__spreadArray([], fetchA, true), fetchB1, true), fetchB2, true), fetchC, true), allRaw, true); _i < _b.length; _i++) {
                            r = _b[_i];
                            id = String((_c = r['sct_requestid']) !== null && _c !== void 0 ? _c : '');
                            if (id && !seen.has(id)) {
                                seen.add(id);
                                merged.push(r);
                            }
                        }
                        // Sort newest first (matches getRawRequests default ordering)
                        return [2 /*return*/, merged.sort(function (a, b) { var _a, _b; return String((_a = b['createdon']) !== null && _a !== void 0 ? _a : '').localeCompare(String((_b = a['createdon']) !== null && _b !== void 0 ? _b : '')); })];
                }
            });
        });
    };
    // ── Core helpers ────────────────────────────────────────────────────────────
    DataverseService.prototype._fetch = function (url, init) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, networkErr_1, msg;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._client.fetch(url, sp_http_1.AadHttpClient.configurations.v1, tslib_1.__assign(tslib_1.__assign({ method: 'GET' }, init), { headers: tslib_1.__assign(tslib_1.__assign({}, OData_HEADERS), init === null || init === void 0 ? void 0 : init.headers) }))];
                    case 1:
                        response = _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        networkErr_1 = _a.sent();
                        msg = networkErr_1 instanceof Error ? networkErr_1.message : String(networkErr_1);
                        throw new Error("Network error reaching Dataverse: ".concat(msg));
                    case 3: return [2 /*return*/, response];
                }
            });
        });
    };
    DataverseService.prototype._get = function (path) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, json;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._fetch("".concat(this.baseUrl).concat(path))];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, this._assertOk(response)];
                    case 2:
                        _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 3:
                        json = _a.sent();
                        return [2 /*return*/, (json.value !== undefined ? json.value : json)];
                }
            });
        });
    };
    DataverseService.prototype._post = function (path, body) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, networkErr_2, msg;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._client.fetch("".concat(this.baseUrl).concat(path), sp_http_1.AadHttpClient.configurations.v1, {
                                method: 'POST',
                                headers: OData_HEADERS,
                                body: JSON.stringify(body),
                            })];
                    case 1:
                        response = _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        networkErr_2 = _a.sent();
                        msg = networkErr_2 instanceof Error ? networkErr_2.message : String(networkErr_2);
                        throw new Error("Network error reaching Dataverse: ".concat(msg));
                    case 3: return [4 /*yield*/, this._assertOk(response, [201, 204])];
                    case 4:
                        _a.sent();
                        return [2 /*return*/, response];
                }
            });
        });
    };
    DataverseService.prototype._patch = function (path, body) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, networkErr_3, msg;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._client.fetch("".concat(this.baseUrl).concat(path), sp_http_1.AadHttpClient.configurations.v1, {
                                method: 'PATCH',
                                headers: tslib_1.__assign(tslib_1.__assign({}, OData_HEADERS), { 'If-Match': '*' }),
                                body: JSON.stringify(body),
                            })];
                    case 1:
                        response = _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        networkErr_3 = _a.sent();
                        msg = networkErr_3 instanceof Error ? networkErr_3.message : String(networkErr_3);
                        throw new Error("Network error reaching Dataverse: ".concat(msg));
                    case 3: return [4 /*yield*/, this._assertOk(response, [204])];
                    case 4:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    DataverseService.prototype._assertOk = function (response_1) {
        return tslib_1.__awaiter(this, arguments, void 0, function (response, expectedCodes) {
            var retryAfter, detail, errBody, _a;
            var _b, _c, _d;
            if (expectedCodes === void 0) { expectedCodes = [200]; }
            return tslib_1.__generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        if (expectedCodes.indexOf(response.status) !== -1)
                            return [2 /*return*/];
                        if (response.status === 401) {
                            throw new Error('Dataverse 401 Unauthorised — ensure the AAD API permission ' +
                                '(Dataverse / user_impersonation) has been granted in SharePoint Admin Centre → API Access.');
                        }
                        if (response.status === 403) {
                            throw new Error('Dataverse 403 Forbidden — the signed-in user lacks the required Dataverse security role.');
                        }
                        if (response.status === 429) {
                            retryAfter = (_b = response.headers.get('Retry-After')) !== null && _b !== void 0 ? _b : '60';
                            throw new Error("Dataverse 429 Too Many Requests \u2014 retry after ".concat(retryAfter, "s."));
                        }
                        detail = '';
                        _e.label = 1;
                    case 1:
                        _e.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, response.json()];
                    case 2:
                        errBody = _e.sent();
                        detail = (_d = (_c = errBody === null || errBody === void 0 ? void 0 : errBody.error) === null || _c === void 0 ? void 0 : _c.message) !== null && _d !== void 0 ? _d : '';
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _e.sent();
                        return [3 /*break*/, 4];
                    case 4: throw new Error("Dataverse error ".concat(response.status).concat(detail ? ': ' + detail : ''));
                }
            });
        });
    };
    // ── Metadata discovery ──────────────────────────────────────────────────────
    /**
     * Returns all readable attributes for the given entity logical name.
     * Use this to discover the real column names before building $select strings.
     */
    DataverseService.prototype.discoverEntityFields = function (entityLogicalName) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var path;
            return tslib_1.__generator(this, function (_a) {
                path = "EntityDefinitions(LogicalName='".concat(entityLogicalName, "')/Attributes") +
                    "?$select=LogicalName,AttributeType,IsCustomAttribute,IsValidForRead,DisplayName" +
                    "&$filter=IsValidForRead eq true";
                return [2 /*return*/, this._get(path)];
            });
        });
    };
    // ── System users ─────────────────────────────────────────────────────────────
    /**
     * Look up the Dataverse systemuser record by systemuserid primary key GUID.
     * Resolves internalemailaddress, fullname, and azureactivedirectoryobjectid.
     */
    DataverseService.prototype.getSystemUserById = function (systemUserId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var cleanId, user, _a, users, _b;
            return tslib_1.__generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        cleanId = (systemUserId || '').replace(/[{}]/g, '').trim();
                        if (!cleanId || !GUID_RE.test(cleanId))
                            return [2 /*return*/, null];
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 3, , 8]);
                        return [4 /*yield*/, this._get("systemusers(".concat(cleanId, ")?$select=systemuserid,fullname,internalemailaddress,azureactivedirectoryobjectid,_parentsystemuserid_value"))];
                    case 2:
                        user = _c.sent();
                        return [2 /*return*/, user || null];
                    case 3:
                        _a = _c.sent();
                        _c.label = 4;
                    case 4:
                        _c.trys.push([4, 6, , 7]);
                        return [4 /*yield*/, this._get("systemusers?$filter=systemuserid eq ".concat(cleanId, "&$select=systemuserid,fullname,internalemailaddress,azureactivedirectoryobjectid,_parentsystemuserid_value&$top=1"))];
                    case 5:
                        users = _c.sent();
                        return [2 /*return*/, Array.isArray(users) && users.length > 0 ? users[0] : null];
                    case 6:
                        _b = _c.sent();
                        return [2 /*return*/, null];
                    case 7: return [3 /*break*/, 8];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Look up the Dataverse systemuser record for the current AAD user.
     * The azureactivedirectoryobjectid field is the AAD Object ID (from SPFx pageContext.aadInfo.userId).
     * Falls back to null if the user is not found or permissions are insufficient.
     */
    DataverseService.prototype.getCurrentSystemUser = function (aadObjectId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var cleanAad, users, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        cleanAad = (aadObjectId || '').replace(/[{}]/g, '').trim();
                        if (!cleanAad)
                            return [2 /*return*/, null];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 5, , 6]);
                        return [4 /*yield*/, this._get("systemusers?$filter=azureactivedirectoryobjectid eq ".concat(cleanAad) +
                                "&$select=systemuserid,fullname,internalemailaddress,_parentsystemuserid_value&$top=1").catch(function () { return []; })];
                    case 2:
                        users = _b.sent();
                        if (!(!Array.isArray(users) || users.length === 0)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this._get("systemusers?$filter=azureactivedirectoryobjectid eq '".concat(cleanAad, "'") +
                                "&$select=systemuserid,fullname,internalemailaddress,_parentsystemuserid_value&$top=1").catch(function () { return []; })];
                    case 3:
                        // Fallback: try quoted filter if unquoted returned empty
                        users = _b.sent();
                        _b.label = 4;
                    case 4: return [2 /*return*/, Array.isArray(users) && users.length > 0 ? users[0] : null];
                    case 5:
                        _a = _b.sent();
                        return [2 /*return*/, null];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /** Look up systemuser by fullname */
    DataverseService.prototype.getSystemUserByName = function (fullName) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var clean, users, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        clean = (fullName || '').trim();
                        if (!clean)
                            return [2 /*return*/, null];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._get("systemusers?$filter=fullname eq '".concat(encodeODataString(clean), "'") +
                                "&$select=systemuserid,fullname,internalemailaddress,_parentsystemuserid_value&$top=1").catch(function () { return []; })];
                    case 2:
                        users = _b.sent();
                        return [2 /*return*/, Array.isArray(users) && users.length > 0 ? users[0] : null];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/, null];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /** Fallback: look up by email when AAD Object ID is unavailable */
    DataverseService.prototype.getSystemUserByEmail = function (email) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var cleanEmail, users, prefix, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        cleanEmail = (email || '').trim();
                        if (!cleanEmail)
                            return [2 /*return*/, null];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 5, , 6]);
                        return [4 /*yield*/, this._get("systemusers?$filter=internalemailaddress eq '".concat(encodeODataString(cleanEmail), "'") +
                                "&$select=systemuserid,fullname,internalemailaddress,_parentsystemuserid_value&$top=1").catch(function () { return []; })];
                    case 2:
                        users = _b.sent();
                        if (!(!Array.isArray(users) || users.length === 0)) return [3 /*break*/, 4];
                        prefix = cleanEmail.split('@')[0];
                        if (!(prefix && prefix.length > 2)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this._get("systemusers?$filter=startswith(internalemailaddress,'".concat(encodeODataString(prefix), "')") +
                                "&$select=systemuserid,fullname,internalemailaddress,_parentsystemuserid_value&$top=1").catch(function () { return []; })];
                    case 3:
                        users = _b.sent();
                        _b.label = 4;
                    case 4: return [2 /*return*/, Array.isArray(users) && users.length > 0 ? users[0] : null];
                    case 5:
                        _a = _b.sent();
                        return [2 /*return*/, null];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Resolves a Dataverse systemuserid string from the systemusers table by manager's email address.
     * Query: GET /api/data/v9.2/systemusers?$select=systemuserid,fullname&$filter=internalemailaddress eq '${managerEmail}'
     */
    DataverseService.prototype.getSystemUserIdByEmail = function (managerEmail) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var cleanEmail, users, prefix, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        cleanEmail = (managerEmail || '').trim();
                        if (!cleanEmail)
                            return [2 /*return*/, null];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 5, , 6]);
                        return [4 /*yield*/, this._get("systemusers?$select=systemuserid,fullname&$filter=internalemailaddress eq '".concat(encodeODataString(cleanEmail), "'&$top=1"))];
                    case 2:
                        users = _b.sent();
                        if (!((!Array.isArray(users) || users.length === 0) && cleanEmail.indexOf('@') !== -1)) return [3 /*break*/, 4];
                        prefix = cleanEmail.split('@')[0];
                        if (!(prefix && prefix.length > 2)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this._get("systemusers?$select=systemuserid,fullname&$filter=startswith(internalemailaddress,'".concat(encodeODataString(prefix), "')&$top=1"))];
                    case 3:
                        users = _b.sent();
                        _b.label = 4;
                    case 4:
                        if (Array.isArray(users) && users.length > 0 && users[0].systemuserid) {
                            return [2 /*return*/, users[0].systemuserid];
                        }
                        return [2 /*return*/, null];
                    case 5:
                        _a = _b.sent();
                        return [2 /*return*/, null];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    DataverseService.prototype.getManagerSystemUserId = function (requesterSystemUserId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var users, val, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        assertGuid(requesterSystemUserId, 'requester system user id');
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._get("systemusers?$filter=systemuserid eq ".concat(requesterSystemUserId) +
                                "&$select=_parentsystemuserid_value&$top=1")];
                    case 2:
                        users = _b.sent();
                        if (!Array.isArray(users) || users.length === 0)
                            return [2 /*return*/, null];
                        val = users[0]['_parentsystemuserid_value'];
                        return [2 /*return*/, typeof val === 'string' && GUID_RE.test(val) ? val : null];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/, null];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Returns true if any other systemuser record lists this user as their manager
     * (parentsystemuserid = systemUserId). Used to decide whether to show the
     * "Manager" role option — driven by Dataverse, not AAD org chart.
     * Returns false on any error so the app degrades gracefully.
     */
    DataverseService.prototype.hasDirectReportsInDataverse = function (systemUserId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var users, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        assertGuid(systemUserId, 'system user id');
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._get("systemusers?$filter=_parentsystemuserid_value eq ".concat(systemUserId) +
                                "&$select=systemuserid&$top=1")];
                    case 2:
                        users = _b.sent();
                        return [2 /*return*/, Array.isArray(users) && users.length > 0];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/, false];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Returns all systemuserid GUIDs for users whose manager is managerSystemUserId
     * (parentsystemuserid = managerSystemUserId in systemusers table).
     */
    DataverseService.prototype.getDirectReportUserIds = function (managerSystemUserId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var users, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        assertGuid(managerSystemUserId, 'manager system user id');
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._get("systemusers?$filter=_parentsystemuserid_value eq ".concat(managerSystemUserId) +
                                "&$select=systemuserid&$top=500")];
                    case 2:
                        users = _b.sent();
                        return [2 /*return*/, Array.isArray(users)
                                ? users.map(function (u) { var _a; return String((_a = u['systemuserid']) !== null && _a !== void 0 ? _a : ''); }).filter(function (id) { return GUID_RE.test(id); })
                                : []];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    // ── Relationship metadata ────────────────────────────────────────────────────
    /**
     * Returns the OData navigation property name to use in @odata.bind for a given
     * lookup attribute.
     *
     * Strategy (each step falls through to the next on failure or empty result):
     *   1. ManyToOneRelationships → ReferencingEntityNavigationPropertyName (most accurate)
     *   2. Attributes SchemaName  → e.g. 'cr2a4_Manager' (correct casing for @odata.bind)
     *   3. Caller's fallback or attributeLogicalName (last resort)
     *
     * The last-resort logical name (all-lowercase) triggers Dataverse 400
     * "undeclared property" when the actual nav prop uses mixed casing, which is why
     * step 2 (SchemaName) is critical.
     */
    DataverseService.prototype.getNavPropName = function (entityLogicalName, attributeLogicalName) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var rels, navProp, _a, attrs, schemaName, _b;
            return tslib_1.__generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _c.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._get("EntityDefinitions(LogicalName='".concat(entityLogicalName, "')/ManyToOneRelationships") +
                                "?$filter=ReferencingAttribute eq '".concat(attributeLogicalName, "'") +
                                "&$select=ReferencingEntityNavigationPropertyName")];
                    case 1:
                        rels = _c.sent();
                        navProp = Array.isArray(rels) && rels.length > 0
                            ? rels[0].ReferencingEntityNavigationPropertyName
                            : null;
                        // Note: Dataverse OData rejects 'sct_Manager' as undeclared on sct_requests.
                        // If the relationship metadata reports 'sct_Manager', prefer the confirmed schema name 'cr2a4_Manager'.
                        if (navProp && navProp.toLowerCase() === 'sct_manager') {
                            return [2 /*return*/, 'cr2a4_Manager'];
                        }
                        if (navProp && navProp !== attributeLogicalName)
                            return [2 /*return*/, navProp];
                        return [3 /*break*/, 3];
                    case 2:
                        _a = _c.sent();
                        return [3 /*break*/, 3];
                    case 3:
                        _c.trys.push([3, 5, , 6]);
                        return [4 /*yield*/, this._get("EntityDefinitions(LogicalName='".concat(entityLogicalName, "')/Attributes") +
                                "?$filter=LogicalName eq '".concat(attributeLogicalName, "'") +
                                "&$select=SchemaName")];
                    case 4:
                        attrs = _c.sent();
                        schemaName = Array.isArray(attrs) && attrs.length > 0 ? attrs[0].SchemaName : null;
                        if (schemaName)
                            return [2 /*return*/, schemaName];
                        return [3 /*break*/, 6];
                    case 5:
                        _b = _c.sent();
                        return [3 /*break*/, 6];
                    case 6: 
                    // Step 3: return the logical name as given; caller should supply a sane default via .catch()
                    return [2 /*return*/, attributeLogicalName];
                }
            });
        });
    };
    // ── Sites ───────────────────────────────────────────────────────────────────
    // No $select — let Dataverse return all columns so we never get a 400
    DataverseService.prototype.getSites = function () {
        return tslib_1.__awaiter(this, arguments, void 0, function (activeOnly) {
            var qs;
            if (activeOnly === void 0) { activeOnly = true; }
            return tslib_1.__generator(this, function (_a) {
                qs = "$orderby=createdon desc&$top=500";
                if (activeOnly)
                    qs += '&$filter=statecode eq 0';
                return [2 /*return*/, this._get("sct_sites_masters?".concat(qs))];
            });
        });
    };
    /** Raw ODataRecord[] for Sites — field names picked at runtime, avoids ISiteMaster mismatch */
    DataverseService.prototype.getSitesRaw = function () {
        return tslib_1.__awaiter(this, arguments, void 0, function (activeOnly) {
            var qs;
            if (activeOnly === void 0) { activeOnly = true; }
            return tslib_1.__generator(this, function (_a) {
                qs = "$orderby=createdon desc&$top=500";
                if (activeOnly)
                    qs += '&$filter=statecode eq 0';
                return [2 /*return*/, this._get("sct_sites_masters?".concat(qs))];
            });
        });
    };
    DataverseService.prototype.getSiteById = function (id) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._get("sct_sites_masters(".concat(id, ")"))];
                    case 1: return [2 /*return*/, _b.sent()];
                    case 2:
                        _a = _b.sent();
                        return [2 /*return*/, null];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    // ── Roles ───────────────────────────────────────────────────────────────────
    DataverseService.prototype.getRoles = function () {
        return tslib_1.__awaiter(this, arguments, void 0, function (activeOnly) {
            var qs;
            if (activeOnly === void 0) { activeOnly = true; }
            return tslib_1.__generator(this, function (_a) {
                qs = "$orderby=createdon desc&$top=500";
                if (activeOnly)
                    qs += '&$filter=statecode eq 0';
                return [2 /*return*/, this._get("sct_roles?".concat(qs))];
            });
        });
    };
    /** Raw ODataRecord[] for roles — used by ApprovalRoutingService for field-safe access */
    DataverseService.prototype.getRawRoles = function () {
        return tslib_1.__awaiter(this, arguments, void 0, function (activeOnly) {
            var qs;
            if (activeOnly === void 0) { activeOnly = true; }
            return tslib_1.__generator(this, function (_a) {
                qs = "$orderby=createdon desc&$top=500";
                if (activeOnly)
                    qs += '&$filter=statecode eq 0';
                return [2 /*return*/, this._get("sct_roles?".concat(qs))];
            });
        });
    };
    /**
     * Fetches all active role configuration records with the fields needed by
     * GroupRoleConfig (tier, role name, description/group, spend limit, site name).
     */
    DataverseService.prototype.getRoleConfigs = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var select;
            return tslib_1.__generator(this, function (_a) {
                select = [
                    'sct_roleid',
                    'sct_rolename',
                    'sct_roledescription',
                    'sct_approvallevel',
                    'sct_approvalspendlimit',
                    '_sct_siteid_value',
                    'statecode',
                ].join(',');
                return [2 /*return*/, this._get("sct_roles?$select=".concat(select, "&$filter=statecode eq 0&$orderby=sct_approvallevel asc,createdon asc&$top=500"))];
            });
        });
    };
    /**
     * PATCHes a single sct_role record — used by GroupRoleConfig to persist inline edits.
     */
    DataverseService.prototype.updateRoleConfig = function (roleId, payload) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        assertGuid(roleId, 'roleId');
                        return [4 /*yield*/, this._patch("sct_roles(".concat(roleId, ")"), payload)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    // ── Suppliers ───────────────────────────────────────────────────────────────
    DataverseService.prototype.getSuppliers = function () {
        return tslib_1.__awaiter(this, arguments, void 0, function (activeOnly) {
            var qs;
            if (activeOnly === void 0) { activeOnly = true; }
            return tslib_1.__generator(this, function (_a) {
                qs = "$orderby=createdon desc&$top=500";
                if (activeOnly)
                    qs += '&$filter=statecode eq 0';
                return [2 /*return*/, this._get("sct_suppliers?".concat(qs))];
            });
        });
    };
    // ── Requests ────────────────────────────────────────────────────────────────
    /**
     * Fetch requests with NO $select — returns every column the entity has.
     * This avoids 400 errors from guessed column names.
     * Optional OData $filter string, e.g. "statecode eq 0".
     */
    DataverseService.prototype.getRequests = function (filter) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var qs;
            return tslib_1.__generator(this, function (_a) {
                qs = "$orderby=createdon desc&$top=250";
                if (filter)
                    qs += "&$filter=".concat(filter);
                return [2 /*return*/, this._get("sct_requests?".concat(qs))];
            });
        });
    };
    /** Also exposes the raw ODataRecord[] for the Dashboard field inspector */
    DataverseService.prototype.getRawRequests = function (filter) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var qs;
            return tslib_1.__generator(this, function (_a) {
                qs = "$orderby=createdon desc&$top=250";
                if (filter)
                    qs += "&$filter=".concat(filter);
                return [2 /*return*/, this._get("sct_requests?".concat(qs))];
            });
        });
    };
    DataverseService.prototype.getRequestById = function (id) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        assertGuid(id, 'record id');
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._get("sct_requests(".concat(id, ")"))];
                    case 2: return [2 /*return*/, _b.sent()];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/, null];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    DataverseService.prototype.getRawRequestById = function (id) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        assertGuid(id, 'record id');
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._get("sct_requests(".concat(id, ")"))];
                    case 2: return [2 /*return*/, _b.sent()];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/, null];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /** Fetch all requests where the requester lookup matches the given systemuserid GUID.
     *  Tries sct_Requester first (schema name used in @odata.bind); falls back to sct_requestorid. */
    DataverseService.prototype.getMyRequests = function (systemUserId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var filter, qs, results, _a, filter, qs;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        assertGuid(systemUserId, 'system user id');
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        filter = "_sct_requester_value eq ".concat(systemUserId);
                        qs = "$orderby=createdon desc&$top=200&$filter=".concat(encodeURIComponent(filter));
                        return [4 /*yield*/, this._get("sct_requests?".concat(qs))];
                    case 2:
                        results = _b.sent();
                        return [2 /*return*/, results];
                    case 3:
                        _a = _b.sent();
                        filter = "_sct_requestorid_value eq ".concat(systemUserId);
                        qs = "$orderby=createdon desc&$top=200&$filter=".concat(encodeURIComponent(filter));
                        return [2 /*return*/, this._get("sct_requests?".concat(qs))];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    DataverseService.prototype.createRequest = function (payload) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var cleanPayload, response, entityId;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        cleanPayload = scrubPayload(payload);
                        return [4 /*yield*/, this._post('sct_requests', cleanPayload)];
                    case 1:
                        response = _b.sent();
                        entityId = (_a = response.headers.get('OData-EntityId')) !== null && _a !== void 0 ? _a : '';
                        return [2 /*return*/, extractGuidFromEntityId(entityId)];
                }
            });
        });
    };
    DataverseService.prototype.updateRequest = function (id, payload) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        assertGuid(id, 'record id');
                        return [4 /*yield*/, this._patch("sct_requests(".concat(id, ")"), payload)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /** Permanently delete a single sct_request record from Dataverse (HTTP DELETE → 204) */
    DataverseService.prototype.deleteRequest = function (id) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        assertGuid(id, 'record id');
                        return [4 /*yield*/, this._fetch("".concat(this.baseUrl, "sct_requests(").concat(id, ")"), { method: 'DELETE' })];
                    case 1:
                        response = _a.sent();
                        if (!(response.status !== 204)) return [3 /*break*/, 3];
                        return [4 /*yield*/, this._assertOk(response)];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    // ── Annotations (file attachments) ──────────────────────────────────────────
    DataverseService.prototype.uploadAnnotation = function (regardingObjectId, filename, mimeType, base64Body) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var body, response, entityId;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        assertGuid(regardingObjectId, 'request id');
                        body = {
                            'objectid_sct_request@odata.bind': "/sct_requests(".concat(regardingObjectId, ")"),
                            filename: filename,
                            mimetype: mimeType,
                            documentbody: base64Body,
                            subject: filename,
                        };
                        return [4 /*yield*/, this._post('annotations', body)];
                    case 1:
                        response = _b.sent();
                        entityId = (_a = response.headers.get('OData-EntityId')) !== null && _a !== void 0 ? _a : '';
                        return [2 /*return*/, extractGuidFromEntityId(entityId)];
                }
            });
        });
    };
    /** Fetch all annotations (file attachments) linked to a given sct_request record */
    DataverseService.prototype.getAnnotations = function (requestId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var filter;
            return tslib_1.__generator(this, function (_a) {
                assertGuid(requestId, 'request id');
                filter = encodeURIComponent("_objectid_value eq ".concat(requestId, " and objecttypecode eq 'sct_request'"));
                return [2 /*return*/, this._get("annotations?$filter=".concat(filter, "&$orderby=createdon desc&$top=50"))];
            });
        });
    };
    /** Delete a single annotation (removes the attachment) */
    DataverseService.prototype.deleteAnnotation = function (annotationId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        assertGuid(annotationId, 'annotation id');
                        return [4 /*yield*/, this._fetch("".concat(this.baseUrl, "annotations(").concat(annotationId, ")"), { method: 'DELETE' })];
                    case 1:
                        response = _a.sent();
                        if (!(response.status !== 204)) return [3 /*break*/, 3];
                        return [4 /*yield*/, this._assertOk(response)];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    return DataverseService;
}());
exports.DataverseService = DataverseService;
//# sourceMappingURL=DataverseService.js.map