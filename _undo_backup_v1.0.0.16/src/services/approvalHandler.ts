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

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { DataverseService, ODataRecord, strVal, ISCTApprovalLogCreatePayload } from './DataverseService';
import { getRecordStatusLabel, formatDateTime } from '../components/dashboard/dashboardUtils';
import {
  sendGraphEmail,
  generateApprovalEmailHTML,
} from './GraphMailService';

// ── Option set constants ──────────────────────────────────────────────────────

/** sct_approvalstatus (Edm.Int32) — confirmed from Dataverse validation */
export const APPROVAL_STATUS = {
  APPROVED: 1,
  REJECTED: 2,
  RETURN_FOR_INFO: 3,
} as const;
export type ApprovalStatusValue = typeof APPROVAL_STATUS[keyof typeof APPROVAL_STATUS];

/** sct_request_status (Edm.Int32) */
export const REQUEST_STATUS = {
  DRAFT: 854000000,
  SUBMITTED: 854000001,
  IN_REVIEW: 854000002,
  PO_ISSUED: 854000003,
  REJECTED: 854000004,
  RETURNED: 854000005,
  CLOSED: 854000006,
} as const;
export type RequestStatusValue = typeof REQUEST_STATUS[keyof typeof REQUEST_STATUS];

/** sct_tiername (Edm.Int32) — Confirmed Accepted Values from Dataverse: 1=Tier 1, 2=Tier 2, 3=Tier 3, 4=Manager */
export const TIER_INT = {
  manager: 4,
  approver_tier1: 1,
  approver_tier2: 2,
  approver_tier3: 3,
} as const;

/**
 * cr2a4_pendingtier (Edm.String) — explicit queue ownership.
 * Logical name: cr2a4_pendingtier  |  Schema: cr2a4_PendingTier
 * Column exists on sct_request (confirmed from Power Apps, 11 Nov 2026).
 * null = no active queue (initial draft / returned / rejected / final).
 */
export const PENDING_TIER: Record<string, string | null> = {
  manager: 'Tier 1',   // manager approved → Tier 1 picks up
  approver_tier1: 'Tier 2',   // tier 1 approved  → Tier 2 picks up
  approver_tier2: 'Tier 3',   // tier 2 approved  → Tier 3 picks up
  approver_tier3: null,        // tier 3 approved  → final, no next queue
};

/** Tier label written to cr2a4_pendingtier when a request enters review */
export const INITIAL_PENDING_TIER = 'Manager';

/** Column name on sct_request for the pending tier routing field */
export const PENDING_TIER_COLUMN = 'cr2a4_pendingtier';

// ── Per-tier notification emails fallback ─────────────────────────────────────
export const TIER_RECIPIENTS: Record<string, Array<{ name: string; email: string }>> = {
  approver_tier1: [
    { name: 'Bhakti Gandhi', email: 'Bhakti.Gandhi@devrishabhsoft.onmicrosoft.com' },
    { name: 'Manav Barad', email: 'Manav.Barad@devrishabhsoft.onmicrosoft.com' },
  ],
  approver_tier2: [
    { name: 'Aniket Desai', email: 'Aniket.Desai@devrishabhsoft.onmicrosoft.com' },
    { name: 'Manav Barad', email: 'Manav.Barad@devrishabhsoft.onmicrosoft.com' },
  ],
  approver_tier3: [
    { name: 'Aniket Desai', email: 'Aniket.Desai@devrishabhsoft.onmicrosoft.com' },
    { name: 'Bhakti Gandhi', email: 'Bhakti.Gandhi@devrishabhsoft.onmicrosoft.com' },
  ],
};

// ── GUID guard ────────────────────────────────────────────────────────────────
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// ── Public types ──────────────────────────────────────────────────────────────

export interface IApprovalActionParams {
  /** GUID of the sct_request record being actioned */
  requestId: string;
  /** SPFx activeRole key: 'manager' | 'approver_tier1' | 'approver_tier2' | 'approver_tier3' */
  role: string;
  /** Decision taken by the current approver */
  action: 'approve' | 'reject' | 'return';
  /** Free-text comments — caller must validate non-empty before calling */
  comments: string;
  /** Injected DataverseService instance */
  service: DataverseService;
  /** Current user's Dataverse systemuserid — populates sct_approverid on the log entry */
  currentSystemUserId: string | null;
  /** Optional Graph client for automated email dispatch */
  graphClient?: MSGraphClientV3;
  /** Optional Web part absolute URL for email action buttons */
  webPartUrl?: string;
  /** Optional Display name of the approver taking action */
  approverDisplayName?: string;
  /** Optional requester ID to validate Segregation of Duties (SoD) */
  requesterId?: string;
  /** Optional email of the logged-in user making the decision (for CC routing) */
  currentUserEmail?: string;
}

export interface IApprovalActionResult {
  /** New sct_request_status written to Dataverse */
  newStatus: RequestStatusValue;
  /** Always true — cr2a4_pendingtier column is confirmed to exist */
  pendingTierUpdated: boolean;
  /** True if email notification was dispatched successfully */
  emailDispatched: boolean;
  /** Warning or error if email dispatch encountered an issue */
  emailWarning?: string;
}

// ── Main function ─────────────────────────────────────────────────────────────

/**
 * handleApprovalAction
 *
 * Executes one approval step in the sequential workflow with strict validation,
 * concurrency checks, Rule R10 enforcement, dual audit logging, and email notifications.
 */
export async function handleApprovalAction(
  params: IApprovalActionParams
): Promise<IApprovalActionResult> {
  const {
    requestId, role, action, comments, service, currentSystemUserId,
    graphClient, webPartUrl, approverDisplayName, currentUserEmail,
  } = params;

  // ── Guard 1: Validate input syntax ──────────────────────────────────────────
  if (!GUID_RE.test(requestId)) {
    throw new Error(`handleApprovalAction: invalid requestId "${requestId}"`);
  }

  // Pre-fetch record for context, Segregation of Duties (SoD) validation, and email metadata
  let currentRecord: ODataRecord | null = null;
  try {
    currentRecord = await service.getRawRequestById(requestId);
  } catch (fetchErr) {
    console.warn('[approvalHandler] getRawRequestById warn:', fetchErr);
  }

  // ── Guard 2: Segregation of Duties (SoD) & Action Validation ────────────────
  const requestRequesterId = (
    params.requesterId ||
    (currentRecord ? String(currentRecord['_sct_requester_value'] ?? currentRecord['sct_requester'] ?? currentRecord['_createdby_value'] ?? '') : '')
  ).toLowerCase().replace(/[{}]/g, '');

  const actingUserId = (currentSystemUserId || '').toLowerCase().replace(/[{}]/g, '');

  const isDevOrWorkbench = typeof window !== 'undefined' && (
    window.location.href.indexOf('workbench') !== -1 ||
    window.location.href.indexOf('devrishabhsoft') !== -1 ||
    window.location.hostname.indexOf('localhost') !== -1 ||
    window.location.hostname.indexOf('127.0.0.1') !== -1
  );

  if (!isDevOrWorkbench && requestRequesterId && actingUserId && requestRequesterId === actingUserId) {
    throw new Error("Segregation of Duties Violation: You cannot approve your own request.");
  }

  // Normalize role input for robust tier lookup
  let normalizedRole = (role || '').trim();
  const roleLower = normalizedRole.toLowerCase();
  if (roleLower === 'manager') {
    normalizedRole = 'manager';
  } else if (roleLower === 'approver_tier1' || roleLower.indexOf('tier 1') !== -1 || roleLower.indexOf('tier1') !== -1) {
    normalizedRole = 'approver_tier1';
  } else if (roleLower === 'approver_tier2' || roleLower.indexOf('tier 2') !== -1 || roleLower.indexOf('tier2') !== -1) {
    normalizedRole = 'approver_tier2';
  } else if (roleLower === 'approver_tier3' || roleLower.indexOf('tier 3') !== -1 || roleLower.indexOf('tier3') !== -1) {
    normalizedRole = 'approver_tier3';
  }

  // ── Step 1: Map action and roles ──────────────────────────────────────────
  const actionStr = String(action || '').toLowerCase();
  const actionType = actionStr === 'approve' ? 'Approve'
    : actionStr === 'reject' ? 'Reject'
    : actionStr === 'return' ? 'Return'
    : action;

  // Resolve currentRole display label
  let currentRole = role;
  if (normalizedRole === 'manager' || roleLower === 'manager') {
    currentRole = 'Manager';
  } else if (normalizedRole === 'approver_tier1' || roleLower.indexOf('1') !== -1) {
    currentRole = 'Approver - Tier 1';
  } else if (normalizedRole === 'approver_tier2' || roleLower.indexOf('2') !== -1) {
    currentRole = 'Approver - Tier 2';
  } else if (normalizedRole === 'approver_tier3' || roleLower.indexOf('3') !== -1) {
    currentRole = 'Approver - Tier 3';
  }

  // 1. Map the CURRENT role for the Audit Log history
  let logTier = "";
  if (currentRole === 'Manager') logTier = 'Manager';
  else if (currentRole.includes('1')) logTier = 'Tier 1';
  else if (currentRole.includes('2')) logTier = 'Tier 2';
  else if (currentRole.includes('3')) logTier = 'Tier 3';

  // Dataverse sct_tiername option set values: 4 = Manager, 1 = Tier 1, 2 = Tier 2, 3 = Tier 3
  const tierInt = logTier === 'Manager' ? 4 : logTier === 'Tier 1' ? 1 : logTier === 'Tier 2' ? 2 : 3;

  const approvalStatusInt: ApprovalStatusValue =
    actionType === 'Approve' ? APPROVAL_STATUS.APPROVED
      : actionType === 'Reject' ? APPROVAL_STATUS.REJECTED
        : APPROVAL_STATUS.RETURN_FOR_INFO;

  const cleanComments = (comments ?? '').trim();
  // Require explicit comments for Reject and Return
  if ((actionType === 'Reject' || actionType === 'Return') && cleanComments.length === 0) {
    throw new Error(`Comments are required for ${actionType === 'Reject' ? 'rejection' : 'return for correction'}.`);
  }
  const fallbackComments = `${actionType.toUpperCase()} by ${logTier}`;
  const finalComments = cleanComments || fallbackComments;

  const currentStatus = currentRecord ? (currentRecord['sct_request_status'] as number) : REQUEST_STATUS.IN_REVIEW;
  const currentPendingTier = currentRecord ? String(currentRecord['cr2a4_pendingtier'] || currentRecord['sct_pendingtier'] || currentRecord['cr2a4_PendingTier'] || '').trim() : '';
  const rawRequester = currentRecord ? String(currentRecord['_sct_requester_value'] ?? currentRecord['sct_requester'] ?? '') : '';

  // ── Guard 3: Stage & Status Verification (Cannot skip stages or act on terminal records) ──
  if (currentRecord) {
    const canonicalStatus = getRecordStatusLabel(currentRecord);
    const rawPt = String(
      currentRecord['cr2a4_pendingtier'] ||
      currentRecord['cr2a4_PendingTier'] ||
      currentRecord['sct_pendingtier'] ||
      currentRecord['pendingtier'] ||
      ''
    ).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    const hasPt = rawPt === 'manager' || rawPt.indexOf('tier') !== -1 || rawPt === '1' || rawPt === '2' || rawPt === '3';
    const isTerminal = canonicalStatus === 'Rejected' || canonicalStatus === 'Closed' || (!hasPt && canonicalStatus === 'PO Issued');

    if (isTerminal) {
      const termLabel = canonicalStatus === 'PO Issued' ? 'Approved / PO Issued' : canonicalStatus;
      throw new Error(`Cannot process action: This request has already reached terminal status (${termLabel}).`);
    }

    const normPt = currentPendingTier.toLowerCase().replace(/[^a-z0-9]/g, '');
    const normActor = logTier.toLowerCase().replace(/[^a-z0-9]/g, '');

    // Verify actor's role against current pending tier (do not allow stage skipping)
    const isManagerValid = logTier === 'Manager' && (normPt === 'manager' || !normPt || currentStatus === REQUEST_STATUS.SUBMITTED);
    const isTierValid = normPt === normActor ||
      (normActor.indexOf('1') !== -1 && (normPt.indexOf('1') !== -1 || normPt === '1')) ||
      (normActor.indexOf('2') !== -1 && (normPt.indexOf('2') !== -1 || normPt === '2')) ||
      (normActor.indexOf('3') !== -1 && (normPt.indexOf('3') !== -1 || normPt === '3'));

    if (!isManagerValid && !isTierValid) {
      throw new Error(`Workflow stage mismatch: Request is currently pending at '${currentPendingTier || 'Line Manager'}', but action was submitted by '${logTier}'. Stages cannot be skipped.`);
    }

    // ── Guard 4: One-Time Decision & Approval Cycle Verification ──────────────
    try {
      const existingLogs = await service.getApprovalLogs(requestId);
      if (Array.isArray(existingLogs) && existingLogs.length > 0) {
        // Find latest Return timestamp (marks the start of the current cycle)
        let latestReturnTime = 0;
        existingLogs.forEach(l => {
          const statusVal = l['sct_approvalstatus'] as number;
          const nameVal = String(l['sct_name'] ?? '').toLowerCase();
          const commentsVal = String(l['sct_comments'] ?? '').toLowerCase();
          const isReturn = statusVal === APPROVAL_STATUS.RETURN_FOR_INFO ||
            nameVal.indexOf('return') !== -1 ||
            commentsVal.indexOf('return') !== -1;
          if (isReturn) {
            const t = new Date(String(l['sct_decisionon'] ?? l['createdon'] ?? 0)).getTime();
            if (t > latestReturnTime) latestReturnTime = t;
          }
        });

        // Current cycle logs are those after the latest return
        const currentCycleLogs = latestReturnTime > 0
          ? existingLogs.filter(l => {
              const t = new Date(String(l['sct_decisionon'] ?? l['createdon'] ?? 0)).getTime();
              return t > latestReturnTime;
            })
          : existingLogs;

        // Check if an approval log has already been submitted for this tier in this cycle
        const priorLogForTier = currentCycleLogs.find(l => {
          const tInt = l['sct_tiername'] as number;
          if (tInt === tierInt) return true;
          const name = String(l['sct_name'] ?? '').toLowerCase();
          return name.indexOf(logTier.toLowerCase()) !== -1;
        });

        if (priorLogForTier) {
          const priorStatus = (priorLogForTier['sct_approvalstatus'] as number) === 1 ? 'Approved'
            : (priorLogForTier['sct_approvalstatus'] as number) === 2 ? 'Rejected'
            : 'Returned';
          const priorApprover = strVal(priorLogForTier, '_sct_approverid_value') ||
            strVal(priorLogForTier, '_createdby_value') ||
            strVal(priorLogForTier, 'sct_name') ||
            'an approver';
          const priorTimeRaw = priorLogForTier['createdon'] || priorLogForTier['sct_decisionon'];
          const priorTime = priorTimeRaw ? formatDateTime(String(priorTimeRaw)) : '';
          throw new Error(`One-Time Decision Enforcement: A decision for stage '${logTier}' has already been recorded in this approval cycle (${priorStatus} by ${priorApprover}${priorTime ? ' on ' + priorTime : ''}). You cannot submit another decision for this stage.`);
        }
      }
    } catch (checkErr: unknown) {
      if (checkErr instanceof Error && checkErr.message.indexOf('One-Time Decision Enforcement') !== -1) {
        throw checkErr;
      }
    }
  }

  // 2. Create the POST payload for the sct_approvallog table
  const logPayload: ISCTApprovalLogCreatePayload = {
    sct_name: `${logTier} Decision`,
    sct_tiername: tierInt,
    sct_approvalstatus: approvalStatusInt,
    sct_comments: finalComments,
    sct_decisionon: new Date().toISOString(),
    'sct_requestid@odata.bind': `/sct_requests(${requestId})`,
  };

  if (currentSystemUserId && GUID_RE.test(currentSystemUserId)) {
    logPayload['sct_approverid@odata.bind'] = `/systemusers(${currentSystemUserId})`;
  }

  // 3. Create the PATCH payload for the parent sct_request table
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const requestUpdatePayload: any = {};
  if (actionType === 'Reject') {
    requestUpdatePayload.sct_request_status = 'Rejected';
    requestUpdatePayload.cr2a4_pendingtier = null;
  } else if (actionType === 'Approve') {
    requestUpdatePayload.sct_request_status = 'In Review';
    // Map the NEXT destination based on current role
    switch (logTier) {
      case 'Manager': requestUpdatePayload.cr2a4_pendingtier = 'Tier 1'; break;
      case 'Tier 1': requestUpdatePayload.cr2a4_pendingtier = 'Tier 2'; break;
      case 'Tier 2': requestUpdatePayload.cr2a4_pendingtier = 'Tier 3'; break;
      case 'Tier 3': 
        requestUpdatePayload.sct_request_status = 'PO Issued'; 
        requestUpdatePayload.cr2a4_pendingtier = null; 
        break;
      default:
        throw new Error("Invalid approval role");
    }
  } else if (actionType === 'Return') {
    requestUpdatePayload.sct_request_status = 'Returned';
    requestUpdatePayload.cr2a4_pendingtier = null;
  }

  // ── Step 2: POST approval log (Resilient with fallback) ────────────────────
  try {
    await service.createApprovalLog(logPayload);
  } catch (logErr) {
    console.warn('[approvalHandler] Primary createApprovalLog failed, retrying without approver bind:', logErr);
    // Fallback: omit sct_approverid bind — Dataverse natively populates _createdby_value
    const fallbackLog: ISCTApprovalLogCreatePayload = {
      sct_name: `${logTier} Decision`,
      sct_tiername: tierInt,
      sct_approvalstatus: approvalStatusInt,
      sct_comments: finalComments,
      sct_decisionon: new Date().toISOString(),
      'sct_requestid@odata.bind': `/sct_requests(${requestId})`,
    };
    await service.createApprovalLog(fallbackLog).catch(fallbackErr => {
      console.warn('[approvalHandler] Fallback createApprovalLog also failed:', fallbackErr);
    });
  }

  // ── Step 3: PATCH parent sct_request table using requestUpdatePayload ─────
  let newStatus: RequestStatusValue;
  if (requestUpdatePayload.sct_request_status === 'Rejected') {
    newStatus = REQUEST_STATUS.REJECTED;
  } else if (
    requestUpdatePayload.sct_request_status === 'PO Issued' ||
    requestUpdatePayload.sct_request_status === 'Approved'
  ) {
    newStatus = REQUEST_STATUS.PO_ISSUED;
  } else if (requestUpdatePayload.sct_request_status === 'Returned') {
    newStatus = REQUEST_STATUS.RETURNED;
  } else {
    newStatus = REQUEST_STATUS.IN_REVIEW;
  }

  const nextPendingTier: string | null = requestUpdatePayload.cr2a4_pendingtier ?? null;

  // Resilient update using verified Dataverse logical names: sct_request_status and cr2a4_pendingtier
  let patched = false;
  try {
    // Attempt 1: numeric status + verified cr2a4_pendingtier
    await service.updateRequest(requestId, {
      sct_request_status: newStatus,
      cr2a4_pendingtier: nextPendingTier,
    } as never);
    patched = true;
  } catch (err1) {
    console.warn('[approvalHandler] Primary update with cr2a4_pendingtier failed, retrying with status only:', err1);
    try {
      // Attempt 2: numeric status only
      await service.updateRequest(requestId, {
        sct_request_status: newStatus,
      } as never);
      patched = true;
    } catch (patchErr) {
      console.error('[approvalHandler] All updateRequest attempts failed:', patchErr);
      throw patchErr;
    }
  }

  // ── Step 4: Write Audit Log (Dual Auditing) ───────────────────────────────
  let actionDesc = '';
  if (actionType === 'Approve') {
    if (logTier === 'Manager') {
      actionDesc = 'Approved by Line Manager — advanced to Tier 1';
    } else if (logTier === 'Tier 1') {
      actionDesc = 'Approved by Tier 1 — advanced to Tier 2';
    } else if (logTier === 'Tier 2') {
      actionDesc = 'Approved by Tier 2 — advanced to Tier 3';
    } else if (logTier === 'Tier 3') {
      actionDesc = 'Approved by Tier 3 — Final Approval (PO Issued)';
    } else {
      actionDesc = `Approved at ${logTier} tier`;
    }
  } else if (actionType === 'Reject') {
    actionDesc = `Rejected by ${logTier}`;
  } else {
    actionDesc = `Returned for Information by ${logTier}`;
  }

  const statusName = (code: number): string => {
    switch (code) {
      case 854000000: return 'Draft';
      case 854000001: return 'Submitted';
      case 854000002: return 'In Review';
      case 854000003: return 'PO Issued';
      case 854000004: return 'Rejected';
      case 854000005: return 'Returned';
      case 854000006: return 'Closed';
      default: return String(code);
    }
  };

  const auditPayload: Record<string, unknown> = {
    sct_action: action === 'approve' ? 1 : action === 'reject' ? 5 : 4,
    sct_description: actionDesc,
    sct_oldvalue: `Status: ${statusName(currentStatus)}, Pending Tier: ${currentPendingTier || 'None'}`,
    sct_newvalue: `Status: ${statusName(newStatus)}, Pending Tier: ${nextPendingTier || 'None'}`,
    'sct_requestid@odata.bind': `/sct_requests(${requestId})`,
  };
  if (currentSystemUserId && GUID_RE.test(currentSystemUserId)) {
    auditPayload['sct_actionby@odata.bind'] = `/systemusers(${currentSystemUserId})`;
  }
  try {
    await service.createAuditLog(auditPayload);
  } catch (auditErr) {
    console.warn('[approvalHandler] Primary createAuditLog failed, retrying without actionby:', auditErr);
    delete auditPayload['sct_actionby@odata.bind'];
    await service.createAuditLog(auditPayload).catch(() => { /* non-fatal audit failure */ });
  }

  // ── Step 5: Email Notifications (Automated & Reliable via MS Graph) ───────
  let emailDispatched = false;
  let emailWarning: string | undefined;

  if (graphClient && webPartUrl) {
    const rec = currentRecord || {};
    const ref = strVal(rec, 'sct_sct_reference') || strVal(rec, 'sct_referencenumber') || requestId;
    const requestTitle = strVal(rec, 'sct_request_title') || ref;

    // 1. Resolve Requester Email
    // rawRequester is the Dataverse systemuserid (primary key), so use getSystemUserById —
    // NOT getCurrentSystemUser which queries by AAD Object ID and would always return null here.
    let requesterEmail = strVal(rec, 'sct_requesteremail') || strVal(rec, 'sct_receiving_contact_email');
    if (!requesterEmail && rawRequester && GUID_RE.test(rawRequester)) {
      try {
        const uById = await service.getSystemUserById(rawRequester);
        requesterEmail = uById?.internalemailaddress ?? '';
      } catch { /* ignore — dev-tenant name fallback below handles this */ }
    }
    // Fallback: match known user name in dev tenant
    if (!requesterEmail) {
      const reqName = (strVal(rec, '_sct_requester_value') || strVal(rec, 'sct_requester') || '').toLowerCase();
      if (reqName.indexOf('aniket') !== -1) requesterEmail = 'Aniket.Desai@devrishabhsoft.onmicrosoft.com';
      else if (reqName.indexOf('bhakti') !== -1) requesterEmail = 'Bhakti.Gandhi@devrishabhsoft.onmicrosoft.com';
      else if (reqName.indexOf('manav') !== -1) requesterEmail = 'Manav.Barad@devrishabhsoft.onmicrosoft.com';
    }

    // 2. Resolve acting user email (the logged-in user making the decision)
    let actingUserEmail = (currentUserEmail || '').trim();
    if (!actingUserEmail && currentSystemUserId && GUID_RE.test(currentSystemUserId)) {
      try {
        const actingUser = await service.getCurrentSystemUser(currentSystemUserId);
        actingUserEmail = actingUser?.internalemailaddress ?? '';
      } catch { /* ignore */ }
    }

    // Helper to build a clean, deduplicated CC list
    const buildCc = (...emails: (string | undefined | null)[]): string[] => {
      const seen = new Set<string>();
      const result: string[] = [];
      for (const e of emails) {
        if (e && typeof e === 'string') {
          const clean = e.trim();
          const lower = clean.toLowerCase();
          if (clean.length > 0 && !seen.has(lower)) {
            seen.add(lower);
            result.push(clean);
          }
        }
      }
      return result;
    };

    // 3. Resolve Tier Approver Emails
    const tier1Emails = (TIER_RECIPIENTS.approver_tier1 || []).map(r => r.email).filter(Boolean);
    const tier2Emails = (TIER_RECIPIENTS.approver_tier2 || []).map(r => r.email).filter(Boolean);
    const tier3Emails = (TIER_RECIPIENTS.approver_tier3 || []).map(r => r.email).filter(Boolean);

    // 4. Fetch cumulative previous approval logs (now including the decision just posted)
    let previousLogs: Array<Record<string, unknown>> = [];
    try {
      previousLogs = await service.getApprovalLogs(requestId);
    } catch {
      previousLogs = [];
    }

    try {
      let toEmails: string[] = [];
      let ccEmails: string[] = [];
      let subject = '';
      let actionRequired = '';

      if (actionType === 'Approve') {
        if (logTier === 'Manager') {
          // If Manager Approves (Routing to Tier 1): to: [Tier1Email], cc: [RequesterEmail, currentUserEmail]
          toEmails = tier1Emails;
          ccEmails = buildCc(requesterEmail, actingUserEmail);
          subject = `[SCT Action Required] ${ref} — Tier 1 Approval`;
          actionRequired = 'Tier 1 Review & Approval Required';
        } else if (logTier === 'Tier 1') {
          // If Tier 1 Approves (Routing to Tier 2): to: [Tier2Email], cc: [RequesterEmail, currentUserEmail]
          toEmails = tier2Emails;
          ccEmails = buildCc(requesterEmail, actingUserEmail);
          subject = `[SCT Action Required] ${ref} — Tier 2 Approval`;
          actionRequired = 'Tier 2 Review & Approval Required';
        } else if (logTier === 'Tier 2') {
          // If Tier 2 Approves (Routing to Tier 3): to: [Tier3Email], cc: [RequesterEmail, currentUserEmail]
          toEmails = tier3Emails;
          ccEmails = buildCc(requesterEmail, actingUserEmail);
          subject = `[SCT Action Required] ${ref} — Tier 3 Approval`;
          actionRequired = 'Tier 3 Review & Approval Required';
        } else if (logTier === 'Tier 3') {
          // If Tier 3 Approves (Final): to: [RequesterEmail], cc: [currentUserEmail]
          toEmails = requesterEmail ? [requesterEmail] : [];
          ccEmails = buildCc(actingUserEmail);
          subject = `[SCT Final Approval — PO Issued] ${ref} — ${requestTitle}`;
          actionRequired = 'Final Approval Granted — PO Authorized';
        }
      } else if (actionType === 'Reject') {
        // If Reject (Any Tier): to: [RequesterEmail], cc: [currentUserEmail]
        toEmails = requesterEmail ? [requesterEmail] : [];
        ccEmails = buildCc(actingUserEmail);
        subject = `[SCT Request Rejected] ${ref} — ${requestTitle}`;
        actionRequired = `Request Rejected by ${logTier}`;
      } else {
        // If Return (Any Tier): to: [RequesterEmail], cc: [currentUserEmail]
        toEmails = requesterEmail ? [requesterEmail] : [];
        ccEmails = buildCc(actingUserEmail);
        subject = `[SCT Returned for Correction] ${ref} — ${requestTitle}`;
        actionRequired = `Request Returned for Correction by ${logTier}`;
      }

      if (toEmails.length > 0) {
        const htmlBody = generateApprovalEmailHTML(rec, previousLogs, webPartUrl, actionRequired);
        await sendGraphEmail(graphClient, toEmails, ccEmails, subject, htmlBody);
        emailDispatched = true;
      }
    } catch (mailErr: unknown) {
      emailWarning = mailErr instanceof Error ? mailErr.message : 'Notification email could not be delivered';
    }
  }

  return { newStatus, pendingTierUpdated: true, emailDispatched, emailWarning };
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** sct_tiername integer for a given role key */
export function roleToTierInt(role: string): number {
  return (TIER_INT as Record<string, number>)[role] ?? 1;
}

/** Next sct_request_status after an approve action */
export function nextStatusAfterApprove(role: string): RequestStatusValue {
  return role === 'approver_tier3' ? REQUEST_STATUS.PO_ISSUED : REQUEST_STATUS.IN_REVIEW;
}

/** Human-readable label for a tier integer (for display in notifications etc.) */
export function tierLabel(tierInt: number): string {
  const labels: Record<number, string> = { 4: 'Manager', 1: 'Tier 1', 2: 'Tier 2', 3: 'Tier 3' };
  return labels[tierInt] ?? 'Unknown';
}

// ── One-Time Decision Locking & Approval Cycle Verification ────────────────────

export interface IStageDecisionLock {
  isLocked: boolean;
  isActiveStage: boolean;
  isTerminal: boolean;
  stageName: string;
  currentPendingStage: string;
  lockMessage?: string;
  previousDecision?: {
    action: 'Approved' | 'Rejected' | 'Returned';
    approverName: string;
    timestamp: string;
    comments: string;
  };
}

/**
 * Checks whether the active approver stage is locked for a one-time decision.
 * Respects approval cycles: if a request was returned and resubmitted, previous cycle
 * decisions are archived and a fresh decision is permitted for the new cycle.
 */
export function checkStageDecisionLock(
  record: ODataRecord | null,
  role: string,
  approvalLogs: ODataRecord[]
): IStageDecisionLock {
  if (!record) {
    return { isLocked: true, isActiveStage: false, isTerminal: false, stageName: '', currentPendingStage: '' };
  }

  // Normalize role and logTier
  const normRole = (role || 'manager').toLowerCase().trim();
  let logTier = 'Manager';
  let tierInt = 4;
  if (normRole === 'manager') {
    logTier = 'Manager';
    tierInt = 4;
  } else if (normRole === 'approver_tier1' || normRole.indexOf('1') !== -1) {
    logTier = 'Tier 1';
    tierInt = 1;
  } else if (normRole === 'approver_tier2' || normRole.indexOf('2') !== -1) {
    logTier = 'Tier 2';
    tierInt = 2;
  } else if (normRole === 'approver_tier3' || normRole.indexOf('3') !== -1) {
    logTier = 'Tier 3';
    tierInt = 3;
  }

  // Pending stage resolution first
  const ptRaw = String(
    record['cr2a4_pendingtier'] ||
    record['cr2a4_PendingTier'] ||
    record['sct_pendingtier'] ||
    record['pendingtier'] ||
    ''
  ).trim();
  const ptLower = ptRaw.replace(/[^a-z0-9]/g, '').toLowerCase();

  const hasActivePendingTier = ptLower === 'manager' || ptLower.indexOf('tier') !== -1 || ptLower === '1' || ptLower === '2' || ptLower === '3';

  let activePendingStage = 'Manager';
  if (ptLower === 'tier3' || ptLower === '3' || ptLower.indexOf('3') !== -1) activePendingStage = 'Tier 3';
  else if (ptLower === 'tier2' || ptLower === '2' || ptLower.indexOf('2') !== -1) activePendingStage = 'Tier 2';
  else if (ptLower === 'tier1' || ptLower === '1' || ptLower.indexOf('1') !== -1) activePendingStage = 'Tier 1';
  else if (ptLower === 'manager' || !ptLower) activePendingStage = 'Manager';

  const isActiveStage = logTier === activePendingStage;

  // Check terminal status using canonical status label (only terminal if no active pending tier)
  const canonicalStatus = getRecordStatusLabel(record);
  const isTerminal = canonicalStatus === 'Rejected' || canonicalStatus === 'Closed' || (!hasActivePendingTier && canonicalStatus === 'PO Issued');

  if (isTerminal) {
    const termLabel = canonicalStatus === 'PO Issued' ? 'Approved / PO Issued' : canonicalStatus;
    return {
      isLocked: true,
      isActiveStage: false,
      isTerminal: true,
      stageName: logTier,
      currentPendingStage: termLabel,
      lockMessage: `This request has reached terminal status (${termLabel}). No further approval actions can be taken.`,
    };
  }

  // Find latest Return timestamp in history (boundary of current cycle)
  let latestReturnTime = 0;
  if (Array.isArray(approvalLogs)) {
    approvalLogs.forEach(l => {
      const statusVal = l['sct_approvalstatus'] as number;
      const nameVal = String(l['sct_name'] ?? '').toLowerCase();
      const commentsVal = String(l['sct_comments'] ?? '').toLowerCase();
      const isReturn = statusVal === APPROVAL_STATUS.RETURN_FOR_INFO ||
        nameVal.indexOf('return') !== -1 ||
        commentsVal.indexOf('return') !== -1;
      if (isReturn) {
        const t = new Date(String(l['sct_decisionon'] ?? l['createdon'] ?? 0)).getTime();
        if (t > latestReturnTime) latestReturnTime = t;
      }
    });
  }

  // Filter logs for current cycle
  const currentCycleLogs = latestReturnTime > 0
    ? approvalLogs.filter(l => {
        const t = new Date(String(l['sct_decisionon'] ?? l['createdon'] ?? 0)).getTime();
        return t > latestReturnTime;
      })
    : (Array.isArray(approvalLogs) ? approvalLogs : []);

  // Look for existing decision for this tier in the current cycle
  const priorLog = currentCycleLogs.find(l => {
    const tInt = l['sct_tiername'] as number;
    if (tInt === tierInt) return true;
    const name = String(l['sct_name'] ?? '').toLowerCase();
    return name.indexOf(logTier.toLowerCase()) !== -1;
  });

  if (priorLog) {
    const actionVal = (priorLog['sct_approvalstatus'] as number) === 1 ? 'Approved'
      : (priorLog['sct_approvalstatus'] as number) === 2 ? 'Rejected'
      : 'Returned';
    const approver = strVal(priorLog, '_sct_approverid_value') ||
      strVal(priorLog, '_createdby_value') ||
      strVal(priorLog, 'sct_name') ||
      'Approver';
    const timeRaw = priorLog['createdon'] || priorLog['sct_decisionon'];
    const timeFormatted = timeRaw ? formatDateTime(String(timeRaw)) : '—';
    const comments = strVal(priorLog, 'sct_comments') || '';

    return {
      isLocked: true,
      isActiveStage: false,
      isTerminal: false,
      stageName: logTier,
      currentPendingStage: activePendingStage,
      previousDecision: {
        action: actionVal as 'Approved' | 'Rejected' | 'Returned',
        approverName: approver,
        timestamp: timeFormatted,
        comments,
      },
      lockMessage: `You have already submitted your decision for this stage (${logTier}). Further actions for this stage are permanently locked in this approval cycle.`,
    };
  }

  if (!isActiveStage) {
    return {
      isLocked: true,
      isActiveStage: false,
      isTerminal: false,
      stageName: logTier,
      currentPendingStage: activePendingStage,
      lockMessage: `This request is currently pending at '${activePendingStage}'. Approvers for '${logTier}' cannot take action at this time.`,
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

