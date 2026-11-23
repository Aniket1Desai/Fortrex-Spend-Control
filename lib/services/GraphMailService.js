import { __awaiter, __generator } from "tslib";
/**
 * Reusable Graph API Email Sender Utility
 * Dispatches an email via Microsoft Graph /v1.0/me/sendMail using the logged-in user's context.
 */
export var sendGraphEmail = function (graphClient, toEmails, ccEmails, subject, htmlBody) { return __awaiter(void 0, void 0, void 0, function () {
    var validTo, validCc, message, error_1;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                _a.trys.push([0, 2, , 3]);
                validTo = (toEmails || []).filter(function (e) { return Boolean(e && typeof e === 'string' && e.trim().length > 0); });
                validCc = (ccEmails || []).filter(function (e) { return Boolean(e && typeof e === 'string' && e.trim().length > 0); });
                if (validTo.length === 0) {
                    console.warn('[GraphMailService] sendGraphEmail aborted: toRecipients list is empty.');
                    return [2 /*return*/];
                }
                message = {
                    subject: subject,
                    body: { contentType: 'HTML', content: htmlBody },
                    toRecipients: validTo.map(function (email) { return ({ emailAddress: { address: email.trim() } }); }),
                };
                if (validCc.length > 0) {
                    message.ccRecipients = validCc.map(function (email) { return ({ emailAddress: { address: email.trim() } }); });
                }
                return [4 /*yield*/, graphClient.api('/me/sendMail').version('v1.0').post({ message: message, saveToSentItems: true })];
            case 1:
                _a.sent();
                return [3 /*break*/, 3];
            case 2:
                error_1 = _a.sent();
                console.error('Non-fatal error: Failed to send Graph email:', error_1);
                return [3 /*break*/, 3];
            case 3: return [2 /*return*/];
        }
    });
}); };
/**
 * Dynamically queries Microsoft Graph for all members of an Azure AD Group.
 * Returns clean email addresses (mail || userPrincipalName).
 * Zero hardcoded names or fallback lists.
 */
export function getGroupMemberEmails(graphClient, groupId) {
    return __awaiter(this, void 0, void 0, function () {
        var response, members, emails, seen, _i, members_1, m, email, lower, err_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, graphClient
                            .api("/groups/".concat(groupId, "/members"))
                            .select('id,displayName,mail,userPrincipalName')
                            .top(999)
                            .get()];
                case 1:
                    response = _a.sent();
                    members = (response === null || response === void 0 ? void 0 : response.value) || [];
                    emails = [];
                    seen = new Set();
                    for (_i = 0, members_1 = members; _i < members_1.length; _i++) {
                        m = members_1[_i];
                        email = String(m.mail || m.userPrincipalName || '').trim();
                        lower = email.toLowerCase();
                        if (email.length > 0 && !seen.has(lower) && email.indexOf('@') !== -1) {
                            seen.add(lower);
                            emails.push(email);
                        }
                    }
                    return [2 /*return*/, emails];
                case 2:
                    err_1 = _a.sent();
                    console.warn("[GraphMailService] Failed to dynamically fetch group members for ".concat(groupId, ":"), err_1);
                    return [2 /*return*/, []];
                case 3: return [2 /*return*/];
            }
        });
    });
}
var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function isGuid(val) {
    return typeof val === 'string' && GUID_RE.test(val.trim());
}
function cleanDisplayVal(val, fallback) {
    if (val === undefined || val === null)
        return fallback;
    var str = String(val).trim();
    if (!str || isGuid(str))
        return fallback;
    return str;
}
/**
 * Formats a date/timestamp to strict 'dd/MM/yyyy HH:mm tt' format in user's local timezone.
 * Example: 18/11/2026 09:31 AM
 */
function formatDateTime(val) {
    if (!val)
        return '—';
    try {
        var iso = val instanceof Date ? val.toISOString() : String(val).trim();
        if (!iso)
            return '—';
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
 * Dynamic HTML Template Generator with Cumulative Approval History
 * Injects prior approval logs and a deep-link action button into a responsive HTML email body.
 * Strictly maps Dataverse Formatted Values (@OData.Community.Display.V1.FormattedValue) to prevent GUID display.
 */
export function generateApprovalEmailHTML(request, previousLogs, siteUrl, actionRequired) {
    var _a, _b;
    var ref = String(request['sct_sct_reference@OData.Community.Display.V1.FormattedValue'] ||
        request['sct_sct_reference'] ||
        request['sct_referencenumber'] ||
        request['reference'] ||
        'Spend Request');
    var title = String(request['sct_request_title@OData.Community.Display.V1.FormattedValue'] ||
        request['sct_request_title'] ||
        request['title'] ||
        ref);
    // 1. Requester Name: Prioritize FormattedValue
    var requester = String(request['_sct_requester_value@OData.Community.Display.V1.FormattedValue'] ||
        request['sct_requester@OData.Community.Display.V1.FormattedValue'] ||
        request['_createdby_value@OData.Community.Display.V1.FormattedValue'] ||
        request['_sct_requestorid_value@OData.Community.Display.V1.FormattedValue'] ||
        cleanDisplayVal(request['requesterName'], '') ||
        cleanDisplayVal(request['_sct_requester_value'], '') ||
        'Unknown Requester');
    // 2. Business Unit: Prioritize FormattedValue
    var bu = String(request['sct_businessunit@OData.Community.Display.V1.FormattedValue'] ||
        request['sct_business_unit@OData.Community.Display.V1.FormattedValue'] ||
        cleanDisplayVal(request['sct_businessunit'], '') ||
        cleanDisplayVal(request['sct_business_unit'], '') ||
        cleanDisplayVal(request['businessUnit'], '') ||
        '—');
    // 3. Site / Location: Prioritize FormattedValue
    var site = String(request['_sct_site_value@OData.Community.Display.V1.FormattedValue'] ||
        request['sct_site@OData.Community.Display.V1.FormattedValue'] ||
        cleanDisplayVal(request['selectedSiteName'], '') ||
        cleanDisplayVal(request['_sct_site_value'], '') ||
        cleanDisplayVal(request['site'], '') ||
        '—');
    var amtRaw = parseFloat(String((_b = (_a = request['sct_estimated_total_value']) !== null && _a !== void 0 ? _a : request['estimatedAmount']) !== null && _b !== void 0 ? _b : '0').replace(/[^0-9.-]/g, ''));
    var amtFormatted = String(request['sct_estimated_total_value@OData.Community.Display.V1.FormattedValue'] ||
        (!isNaN(amtRaw) && amtRaw > 0
            ? "$".concat(amtRaw.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }))
            : '(not specified)'));
    var desc = String(request['sct_description@OData.Community.Display.V1.FormattedValue'] ||
        request['sct_description'] ||
        request['description'] ||
        '');
    // Deep link targeting the request reference
    var queryParam = ref !== 'Spend Request' && ref !== '(New Submission)' ? ref : String(request['id'] || request['sct_requestid'] || '');
    var deepLink = siteUrl.indexOf('?') !== -1
        ? "".concat(siteUrl, "&requestId=").concat(encodeURIComponent(queryParam))
        : "".concat(siteUrl, "?requestId=").concat(encodeURIComponent(queryParam));
    // Cumulative Prior Approval History HTML mapping
    var historyRowsHtml = '';
    if (Array.isArray(previousLogs) && previousLogs.length > 0) {
        historyRowsHtml = previousLogs.map(function (log) {
            var _a;
            var tierFormatted = log['sct_tiername@OData.Community.Display.V1.FormattedValue'];
            var tierInt = log['sct_tiername'];
            var tier = tierFormatted ? String(tierFormatted)
                : tierInt === 4 ? 'Line Manager'
                    : tierInt === 1 ? 'Tier 1 Approver'
                        : tierInt === 2 ? 'Tier 2 Approver'
                            : tierInt === 3 ? 'Tier 3 Approver'
                                : String((_a = log['sct_tiername']) !== null && _a !== void 0 ? _a : 'Approver');
            var statusFormatted = log['sct_approvalstatus@OData.Community.Display.V1.FormattedValue'];
            var statusInt = log['sct_approvalstatus'];
            var isApprove = statusFormatted
                ? String(statusFormatted).toLowerCase().indexOf('approved') !== -1
                : statusInt === 1;
            var isReject = statusFormatted
                ? String(statusFormatted).toLowerCase().indexOf('rejected') !== -1
                : statusInt === 2;
            var statusLabel = statusFormatted ? String(statusFormatted) : (isApprove ? 'Approved' : isReject ? 'Rejected' : 'Returned');
            var statusBg = isApprove ? '#DFF6DD' : isReject ? '#FDE7E9' : '#FFF4CE';
            var statusColor = isApprove ? '#107C10' : isReject ? '#A4262C' : '#7A4F01';
            // Approver Name: Prioritize FormattedValue
            var approverName = String(log['_sct_approver_value@OData.Community.Display.V1.FormattedValue'] ||
                log['_sct_approverid_value@OData.Community.Display.V1.FormattedValue'] ||
                log['sct_approverid@OData.Community.Display.V1.FormattedValue'] ||
                log['_createdby_value@OData.Community.Display.V1.FormattedValue'] ||
                cleanDisplayVal(log['sct_approvername'], '') ||
                cleanDisplayVal(log['sct_name'], '') ||
                'Unknown Approver');
            var comments = String(log['sct_comments@OData.Community.Display.V1.FormattedValue'] ||
                log['sct_comments'] ||
                '').trim();
            var rawIso = String(log['createdon'] ||
                log['sct_decisionon'] ||
                log['createdon@OData.Community.Display.V1.FormattedValue'] ||
                log['sct_decisionon@OData.Community.Display.V1.FormattedValue'] ||
                '');
            var dateStr = formatDateTime(rawIso);
            return "\n        <tr style=\"border-bottom: 1px solid #EDEBE9;\">\n          <td style=\"padding: 10px 12px; font-weight: 600; color: #323130; vertical-align: top; white-space: nowrap;\">\n            ".concat(tier, "\n          </td>\n          <td style=\"padding: 10px 12px; vertical-align: top; white-space: nowrap;\">\n            <span style=\"background: ").concat(statusBg, "; color: ").concat(statusColor, "; font-size: 11px; font-weight: 700; padding: 3px 8px; border-radius: 10px; display: inline-block;\">\n              ").concat(statusLabel, "\n            </span>\n          </td>\n          <td style=\"padding: 10px 12px; color: #605E5C; vertical-align: top;\">\n            <div style=\"font-weight: 600; color: #323130;\">").concat(approverName, "</div>\n            <div style=\"font-size: 11px; color: #8A8886; margin-top: 2px;\">").concat(dateStr, "</div>\n          </td>\n          <td style=\"padding: 10px 12px; color: #323130; vertical-align: top;\">\n            ").concat(comments ? "<div style=\"background: #F8F8F8; border-left: 3px solid #0078D4; padding: 6px 10px; border-radius: 2px; font-size: 12px; font-style: italic;\">\"".concat(comments, "\"</div>") : '<span style="color: #A19F9D; font-size: 12px;">(No comments provided)</span>', "\n          </td>\n        </tr>\n      ");
        }).join('');
    }
    else {
        historyRowsHtml = "\n      <tr>\n        <td colspan=\"4\" style=\"padding: 14px 12px; color: #8A8886; font-style: italic; text-align: center;\">\n          Initial submission \u2014 no prior approval history recorded.\n        </td>\n      </tr>\n    ";
    }
    return "\n<div style=\"font-family: 'Segoe UI', Tahoma, Arial, sans-serif; max-width: 650px; margin: 0 auto; color: #201F1E; background-color: #FAFAFA; padding: 20px;\">\n  <!-- Header Card -->\n  <div style=\"background: linear-gradient(135deg, #0078D4 0%, #004578 100%); padding: 24px 32px; border-radius: 8px 8px 0 0; color: #FFFFFF;\">\n    <div style=\"font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #B4D6FA; margin-bottom: 4px;\">\n      Fortrex Spend Control Tracker\n    </div>\n    <h1 style=\"margin: 0; font-size: 20px; font-weight: 700; color: #FFFFFF;\">\n      ".concat(actionRequired, "\n    </h1>\n    <p style=\"margin: 6px 0 0; font-size: 13px; color: #E1DFDD;\">\n      Reference: <strong style=\"color: #FFFFFF;\">").concat(ref, "</strong>\n    </p>\n  </div>\n\n  <!-- Body Container -->\n  <div style=\"background: #FFFFFF; padding: 28px 32px; border: 1px solid #EDEBE9; border-top: none; border-radius: 0 0 8px 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.04);\">\n    <p style=\"margin: 0 0 18px; font-size: 14px; line-height: 1.5; color: #323130;\">\n      A spend control request requires your attention. Below are the request details and complete cumulative approval trail:\n    </p>\n\n    <!-- Request Details Summary Card -->\n    <div style=\"background: #F8F9FA; border: 1px solid #E1DFDD; border-radius: 6px; padding: 18px 20px; margin-bottom: 24px;\">\n      <div style=\"font-size: 15px; font-weight: 700; color: #0078D4; margin-bottom: 12px; border-bottom: 1px solid #EDEBE9; padding-bottom: 8px;\">\n        ").concat(title, "\n      </div>\n      <table style=\"width: 100%; border-collapse: collapse; font-size: 13px;\">\n        <tr>\n          <td style=\"padding: 5px 0; color: #605E5C; width: 140px; font-weight: 600;\">Requested By:</td>\n          <td style=\"padding: 5px 0; color: #201F1E; font-weight: 600;\">").concat(requester, "</td>\n        </tr>\n        <tr>\n          <td style=\"padding: 5px 0; color: #605E5C; font-weight: 600;\">Estimated Value:</td>\n          <td style=\"padding: 5px 0; color: #107C10; font-weight: 700; font-size: 14px;\">").concat(amtFormatted, "</td>\n        </tr>\n        <tr>\n          <td style=\"padding: 5px 0; color: #605E5C; font-weight: 600;\">Business Unit:</td>\n          <td style=\"padding: 5px 0; color: #201F1E;\">").concat(bu, "</td>\n        </tr>\n        <tr>\n          <td style=\"padding: 5px 0; color: #605E5C; font-weight: 600;\">Site / Location:</td>\n          <td style=\"padding: 5px 0; color: #201F1E;\">").concat(site, "</td>\n        </tr>\n        ").concat(desc ? "\n        <tr>\n          <td style=\"padding: 8px 0 0; color: #605E5C; font-weight: 600; vertical-align: top;\">Description:</td>\n          <td style=\"padding: 8px 0 0; color: #323130; line-height: 1.4;\">".concat(desc, "</td>\n        </tr>") : '', "\n      </table>\n    </div>\n\n    <!-- Prior Approval History Section -->\n    <div style=\"margin-bottom: 26px;\">\n      <div style=\"font-size: 14px; font-weight: 700; color: #201F1E; margin-bottom: 10px;\">\n        Prior Approval History\n      </div>\n      <table style=\"width: 100%; border-collapse: collapse; font-size: 12px; border: 1px solid #EDEBE9; border-radius: 4px; overflow: hidden;\">\n        <thead>\n          <tr style=\"background: #F3F2F1; text-align: left; border-bottom: 1px solid #EDEBE9;\">\n            <th style=\"padding: 8px 12px; color: #605E5C; font-weight: 700; width: 110px;\">Stage</th>\n            <th style=\"padding: 8px 12px; color: #605E5C; font-weight: 700; width: 85px;\">Decision</th>\n            <th style=\"padding: 8px 12px; color: #605E5C; font-weight: 700; width: 140px;\">Approver &amp; Date</th>\n            <th style=\"padding: 8px 12px; color: #605E5C; font-weight: 700;\">Comments</th>\n          </tr>\n        </thead>\n        <tbody>\n          ").concat(historyRowsHtml, "\n        </tbody>\n      </table>\n    </div>\n\n    <!-- Action Deep Link CTA Button -->\n    <div style=\"text-align: center; margin: 30px 0 15px;\">\n      <a href=\"").concat(deepLink, "\" target=\"_blank\" style=\"display: inline-block; background-color: #0078D4; color: #FFFFFF; font-size: 14px; font-weight: 600; text-decoration: none; padding: 13px 28px; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,120,212,0.3);\">\n        View &amp; Act on Request\n      </a>\n    </div>\n\n    <!-- Footer Disclaimer -->\n    <div style=\"margin-top: 28px; padding-top: 14px; border-top: 1px solid #EDEBE9; font-size: 11px; color: #8A8886; text-align: center; line-height: 1.4;\">\n      This email was automatically generated and dispatched by the Fortrex Spend Control Tracker using Microsoft Graph API. Please do not reply directly to this email.\n    </div>\n  </div>\n</div>\n  ");
}
/** Fetches the current user's direct manager from MS Graph. Returns null if not found. */
export function getManager(graphClient) {
    return __awaiter(this, void 0, void 0, function () {
        var result, mail, _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _b.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, graphClient
                            .api('/me/manager')
                            .select('id,displayName,mail,userPrincipalName')
                            .get()];
                case 1:
                    result = _b.sent();
                    if (!result)
                        return [2 /*return*/, null];
                    mail = String(result.mail || result.userPrincipalName || '').trim();
                    if (!mail && !result.id)
                        return [2 /*return*/, null];
                    return [2 /*return*/, { id: result.id, displayName: result.displayName || '', mail: mail }];
                case 2:
                    _a = _b.sent();
                    return [2 /*return*/, null];
                case 3: return [2 /*return*/];
            }
        });
    });
}
/** Sends a notification email to a list of recipients (next-tier approvers). */
export function sendTierNotificationEmail(graphClient, recipients, tierName, requesterName, requestTitle, estimatedAmount, webPartUrl, decision) {
    return __awaiter(this, void 0, void 0, function () {
        var decisionLabel, decisionColor, amtFormatted, html;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!recipients.length)
                        return [2 /*return*/];
                    decisionLabel = decision === 'approved' ? 'Approved ✓' : decision === 'rejected' ? 'Rejected ✗' : 'Returned for Info ↩';
                    decisionColor = decision === 'approved' ? '#107C10' : decision === 'rejected' ? '#A4262C' : '#8764B8';
                    amtFormatted = estimatedAmount
                        ? "$".concat(parseFloat(estimatedAmount).toLocaleString('en-US', { minimumFractionDigits: 2 }))
                        : '(not specified)';
                    html = "\n<div style=\"font-family: Segoe UI, sans-serif; max-width: 600px; margin: 0 auto; color: #201F1E;\">\n  <div style=\"background: linear-gradient(135deg, #0078D4, #1BA1E2); padding: 24px 32px; border-radius: 8px 8px 0 0;\">\n    <h1 style=\"color: #fff; margin: 0; font-size: 20px; font-weight: 700;\">Fortrex Spend Control Tracker</h1>\n    <p style=\"color: rgba(255,255,255,0.85); margin: 6px 0 0; font-size: 14px;\">Request Requires Your Attention \u2014 ".concat(tierName, "</p>\n  </div>\n  <div style=\"background: #F3F2F1; padding: 28px 32px; border-radius: 0 0 8px 8px; border: 1px solid #EDEBE9;\">\n    <p style=\"margin: 0 0 16px; font-size: 15px;\">Hello,</p>\n    <p style=\"margin: 0 0 20px; font-size: 14px; line-height: 1.6; color: #605E5C;\">\n      A request from <strong>").concat(requesterName, "</strong> has been\n      <strong style=\"color: ").concat(decisionColor, ";\">").concat(decisionLabel, "</strong> at the previous stage\n      and is now awaiting your review as <strong>").concat(tierName, "</strong>.\n    </p>\n    <div style=\"background: #fff; border-radius: 6px; padding: 18px 20px; margin-bottom: 20px; border: 1px solid #EDEBE9;\">\n      <table style=\"width: 100%; border-collapse: collapse; font-size: 13px;\">\n        <tr><td style=\"padding: 6px 0; color: #A19F9D; width: 160px;\">Request Title</td><td style=\"font-weight: 600;\">").concat(requestTitle, "</td></tr>\n        <tr><td style=\"padding: 6px 0; color: #A19F9D;\">Requested By</td><td style=\"font-weight: 600;\">").concat(requesterName, "</td></tr>\n        <tr><td style=\"padding: 6px 0; color: #A19F9D;\">Estimated Value</td><td style=\"font-weight: 600;\">").concat(amtFormatted, "</td></tr>\n        <tr><td style=\"padding: 6px 0; color: #A19F9D;\">Current Stage</td><td style=\"font-weight: 600; color: ").concat(decisionColor, ";\">").concat(tierName, "</td></tr>\n      </table>\n    </div>\n    <a href=\"").concat(webPartUrl, "\" style=\"display: inline-block; background: #0078D4; color: #fff; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 14px;\">\n      Open SCT &amp; Take Action \u2192\n    </a>\n    <p style=\"margin: 24px 0 0; font-size: 12px; color: #A19F9D;\">\n      This notification was sent automatically by the Fortrex Spend Control Tracker. Do not reply to this email.\n    </p>\n  </div>\n</div>");
                    return [4 /*yield*/, graphClient.api('/me/sendMail').post({
                            message: {
                                subject: "[SCT \u2014 ".concat(tierName, " Action Required] ").concat(requestTitle),
                                body: { contentType: 'HTML', content: html },
                                toRecipients: recipients.map(function (r) { return ({
                                    emailAddress: { address: r.email, name: r.name },
                                }); }),
                            },
                            saveToSentItems: false,
                        })];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
/** Sends an approval-request email to the manager via Graph /me/sendMail. */
export function sendManagerApprovalEmail(graphClient, manager, requesterName, requestTitle, estimatedAmount, webPartUrl) {
    return __awaiter(this, void 0, void 0, function () {
        var amtFormatted, html;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    amtFormatted = estimatedAmount
                        ? "$".concat(parseFloat(estimatedAmount).toLocaleString('en-US', { minimumFractionDigits: 2 }))
                        : '(not specified)';
                    html = "\n<div style=\"font-family: Segoe UI, sans-serif; max-width: 600px; margin: 0 auto; color: #201F1E;\">\n  <div style=\"background: linear-gradient(135deg, #0078D4, #1BA1E2); padding: 24px 32px; border-radius: 8px 8px 0 0;\">\n    <h1 style=\"color: #fff; margin: 0; font-size: 20px; font-weight: 700;\">Fortrex Spend Control Tracker</h1>\n    <p style=\"color: rgba(255,255,255,0.85); margin: 6px 0 0; font-size: 14px;\">New Request Awaiting Your Approval</p>\n  </div>\n  <div style=\"background: #F3F2F1; padding: 28px 32px; border-radius: 0 0 8px 8px; border: 1px solid #EDEBE9;\">\n    <p style=\"margin: 0 0 16px; font-size: 15px;\">Hello <strong>".concat(manager.displayName, "</strong>,</p>\n    <p style=\"margin: 0 0 20px; font-size: 14px; line-height: 1.6; color: #605E5C;\">\n      A new spend control request has been submitted by <strong>").concat(requesterName, "</strong> and requires your approval.\n    </p>\n    <div style=\"background: #fff; border-radius: 6px; padding: 18px 20px; margin-bottom: 20px; border: 1px solid #EDEBE9;\">\n      <table style=\"width: 100%; border-collapse: collapse; font-size: 13px;\">\n        <tr><td style=\"padding: 6px 0; color: #A19F9D; width: 160px;\">Request Title</td><td style=\"font-weight: 600;\">").concat(requestTitle, "</td></tr>\n        <tr><td style=\"padding: 6px 0; color: #A19F9D;\">Requested By</td><td style=\"font-weight: 600;\">").concat(requesterName, "</td></tr>\n        <tr><td style=\"padding: 6px 0; color: #A19F9D;\">Estimated Value</td><td style=\"font-weight: 600;\">").concat(amtFormatted, "</td></tr>\n      </table>\n    </div>\n    <a href=\"").concat(webPartUrl, "\" style=\"display: inline-block; background: #0078D4; color: #fff; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 14px;\">\n      Review &amp; Approve Request \u2192\n    </a>\n    <p style=\"margin: 24px 0 0; font-size: 12px; color: #A19F9D;\">\n      This notification was sent automatically by the Fortrex Spend Control Tracker. Do not reply to this email.\n    </p>\n  </div>\n</div>");
                    return [4 /*yield*/, graphClient.api('/me/sendMail').post({
                            message: {
                                subject: "[SCT Approval Required] ".concat(requestTitle, " \u2014 ").concat(requesterName),
                                body: { contentType: 'HTML', content: html },
                                toRecipients: [{ emailAddress: { address: manager.mail, name: manager.displayName } }],
                            },
                            saveToSentItems: false,
                        })];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
/** Sends decision outcome email (Rejected or Returned for Info) to the requester. */
export function sendRequesterDecisionEmail(graphClient, requesterEmail, requesterName, requestTitle, estimatedAmount, webPartUrl, decision, approverName, tierName, comments) {
    return __awaiter(this, void 0, void 0, function () {
        var isRejected, decisionLabel, decisionColor, subjectPrefix, amtFormatted, html;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!requesterEmail)
                        return [2 /*return*/];
                    isRejected = decision === 'rejected';
                    decisionLabel = isRejected ? 'Rejected ✗' : 'Returned for Information ↩';
                    decisionColor = isRejected ? '#A4262C' : '#8764B8';
                    subjectPrefix = isRejected ? '[SCT Request Rejected]' : '[SCT Returned for Information]';
                    amtFormatted = estimatedAmount
                        ? "$".concat(parseFloat(estimatedAmount).toLocaleString('en-US', { minimumFractionDigits: 2 }))
                        : '(not specified)';
                    html = "\n<div style=\"font-family: Segoe UI, sans-serif; max-width: 600px; margin: 0 auto; color: #201F1E;\">\n  <div style=\"background: linear-gradient(135deg, ".concat(decisionColor, ", #201F1E); padding: 24px 32px; border-radius: 8px 8px 0 0;\">\n    <h1 style=\"color: #fff; margin: 0; font-size: 20px; font-weight: 700;\">Fortrex Spend Control Tracker</h1>\n    <p style=\"color: rgba(255,255,255,0.85); margin: 6px 0 0; font-size: 14px;\">Request Update: ").concat(decisionLabel, "</p>\n  </div>\n  <div style=\"background: #F3F2F1; padding: 28px 32px; border-radius: 0 0 8px 8px; border: 1px solid #EDEBE9;\">\n    <p style=\"margin: 0 0 16px; font-size: 15px;\">Hello <strong>").concat(requesterName, "</strong>,</p>\n    <p style=\"margin: 0 0 20px; font-size: 14px; line-height: 1.6; color: #605E5C;\">\n      Your spend request has been <strong style=\"color: ").concat(decisionColor, ";\">").concat(decisionLabel, "</strong> by <strong>").concat(approverName, "</strong> (").concat(tierName, ").\n    </p>\n    <div style=\"background: #fff; border-radius: 6px; padding: 18px 20px; margin-bottom: 20px; border: 1px solid #EDEBE9;\">\n      <table style=\"width: 100%; border-collapse: collapse; font-size: 13px;\">\n        <tr><td style=\"padding: 6px 0; color: #A19F9D; width: 140px;\">Request Title</td><td style=\"font-weight: 600;\">").concat(requestTitle, "</td></tr>\n        <tr><td style=\"padding: 6px 0; color: #A19F9D;\">Estimated Value</td><td style=\"font-weight: 600;\">").concat(amtFormatted, "</td></tr>\n        <tr><td style=\"padding: 6px 0; color: #A19F9D;\">Reviewed By</td><td style=\"font-weight: 600;\">").concat(approverName, " (").concat(tierName, ")</td></tr>\n        <tr><td style=\"padding: 6px 0; color: #A19F9D;\">Status</td><td style=\"font-weight: 600; color: ").concat(decisionColor, ";\">").concat(decisionLabel, "</td></tr>\n      </table>\n      ").concat(comments ? "\n      <div style=\"margin-top: 14px; padding-top: 12px; border-top: 1px solid #EDEBE9;\">\n        <div style=\"font-size: 11px; font-weight: 700; color: #A19F9D; text-transform: uppercase; margin-bottom: 4px;\">Approver Comments</div>\n        <div style=\"font-size: 13px; color: #201F1E; background: #F8F8F8; padding: 10px; border-radius: 4px; border: 1px solid #E1DFDD;\">".concat(comments, "</div>\n      </div>") : '', "\n    </div>\n    <a href=\"").concat(webPartUrl, "\" style=\"display: inline-block; background: #0078D4; color: #fff; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 14px;\">\n      ").concat(isRejected ? 'View Request in SCT →' : 'Edit &amp; Resubmit Request →', "\n    </a>\n    <p style=\"margin: 24px 0 0; font-size: 12px; color: #A19F9D;\">\n      This notification was sent automatically by the Fortrex Spend Control Tracker. Do not reply to this email.\n    </p>\n  </div>\n</div>");
                    return [4 /*yield*/, graphClient.api('/me/sendMail').post({
                            message: {
                                subject: "".concat(subjectPrefix, " ").concat(requestTitle),
                                body: { contentType: 'HTML', content: html },
                                toRecipients: [{ emailAddress: { address: requesterEmail, name: requesterName } }],
                            },
                            saveToSentItems: false,
                        })];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
/** Sends final approval (PO Issued) confirmation email to the requester and stakeholders. */
export function sendPOIssuedNotificationEmail(graphClient, requesterEmail, requesterName, requestTitle, estimatedAmount, webPartUrl, approverName, comments) {
    return __awaiter(this, void 0, void 0, function () {
        var amtFormatted, html;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!requesterEmail)
                        return [2 /*return*/];
                    amtFormatted = estimatedAmount
                        ? "$".concat(parseFloat(estimatedAmount).toLocaleString('en-US', { minimumFractionDigits: 2 }))
                        : '(not specified)';
                    html = "\n<div style=\"font-family: Segoe UI, sans-serif; max-width: 600px; margin: 0 auto; color: #201F1E;\">\n  <div style=\"background: linear-gradient(135deg, #107C10, #0B6A0B); padding: 24px 32px; border-radius: 8px 8px 0 0;\">\n    <h1 style=\"color: #fff; margin: 0; font-size: 20px; font-weight: 700;\">Fortrex Spend Control Tracker</h1>\n    <p style=\"color: rgba(255,255,255,0.85); margin: 6px 0 0; font-size: 14px;\">Final Approval Granted \u2014 PO Authorized</p>\n  </div>\n  <div style=\"background: #F3F2F1; padding: 28px 32px; border-radius: 0 0 8px 8px; border: 1px solid #EDEBE9;\">\n    <p style=\"margin: 0 0 16px; font-size: 15px;\">Hello <strong>".concat(requesterName, "</strong>,</p>\n    <p style=\"margin: 0 0 20px; font-size: 14px; line-height: 1.6; color: #605E5C;\">\n      Your spend request has successfully passed all approval stages and received final sign-off from <strong>").concat(approverName, "</strong> (Tier 3 / Executive).\n      Procurement has been authorized to issue a Purchase Order in Dynamics 365.\n    </p>\n    <div style=\"background: #fff; border-radius: 6px; padding: 18px 20px; margin-bottom: 20px; border: 1px solid #EDEBE9;\">\n      <table style=\"width: 100%; border-collapse: collapse; font-size: 13px;\">\n        <tr><td style=\"padding: 6px 0; color: #A19F9D; width: 140px;\">Request Title</td><td style=\"font-weight: 600;\">").concat(requestTitle, "</td></tr>\n        <tr><td style=\"padding: 6px 0; color: #A19F9D;\">Estimated Value</td><td style=\"font-weight: 600;\">").concat(amtFormatted, "</td></tr>\n        <tr><td style=\"padding: 6px 0; color: #A19F9D;\">Final Status</td><td style=\"font-weight: 600; color: #107C10;\">PO Issued (Approved)</td></tr>\n      </table>\n      ").concat(comments ? "\n      <div style=\"margin-top: 14px; padding-top: 12px; border-top: 1px solid #EDEBE9;\">\n        <div style=\"font-size: 11px; font-weight: 700; color: #A19F9D; text-transform: uppercase; margin-bottom: 4px;\">Sign-off Notes</div>\n        <div style=\"font-size: 13px; color: #201F1E; background: #F8F8F8; padding: 10px; border-radius: 4px; border: 1px solid #E1DFDD;\">".concat(comments, "</div>\n      </div>") : '', "\n    </div>\n    <a href=\"").concat(webPartUrl, "\" style=\"display: inline-block; background: #107C10; color: #fff; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 14px;\">\n      View Request in SCT \u2192\n    </a>\n    <p style=\"margin: 24px 0 0; font-size: 12px; color: #A19F9D;\">\n      This notification was sent automatically by the Fortrex Spend Control Tracker. Do not reply to this email.\n    </p>\n  </div>\n</div>");
                    return [4 /*yield*/, graphClient.api('/me/sendMail').post({
                            message: {
                                subject: "[SCT Approved \u2014 PO Authorized] ".concat(requestTitle),
                                body: { contentType: 'HTML', content: html },
                                toRecipients: [{ emailAddress: { address: requesterEmail, name: requesterName } }],
                            },
                            saveToSentItems: false,
                        })];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
//# sourceMappingURL=GraphMailService.js.map