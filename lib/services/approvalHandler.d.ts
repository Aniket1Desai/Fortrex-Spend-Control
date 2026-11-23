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
import { DataverseService, ODataRecord } from './DataverseService';
/** sct_approvalstatus (Edm.Int32) — confirmed from Dataverse validation */
export declare const APPROVAL_STATUS: {
    readonly APPROVED: 1;
    readonly REJECTED: 2;
    readonly RETURN_FOR_INFO: 3;
};
export type ApprovalStatusValue = typeof APPROVAL_STATUS[keyof typeof APPROVAL_STATUS];
/** sct_request_status (Edm.Int32) — Confirmed Dataverse choice values: 854000001 to 854000007 */
export declare const REQUEST_STATUS: {
    readonly DRAFT: 854000001;
    readonly SUBMITTED: 854000002;
    readonly IN_REVIEW: 854000003;
    readonly PO_ISSUED: 854000004;
    readonly REJECTED: 854000005;
    readonly RETURNED: 854000006;
    readonly CLOSED: 854000007;
};
export type RequestStatusValue = typeof REQUEST_STATUS[keyof typeof REQUEST_STATUS];
/** sct_tiername (Edm.Int32) — Confirmed Accepted Values from Dataverse: 1=Tier 1, 2=Tier 2, 3=Tier 3, 4=Manager */
export declare const TIER_INT: {
    readonly manager: 4;
    readonly approver_tier1: 1;
    readonly approver_tier2: 2;
    readonly approver_tier3: 3;
};
/**
 * cr2a4_pendingtier (Edm.String) — explicit queue ownership.
 * Logical name: cr2a4_pendingtier  |  Schema: cr2a4_PendingTier
 * Column exists on sct_request (confirmed from Power Apps, 11 Nov 2026).
 * null = no active queue (initial draft / returned / rejected / final).
 */
export declare const PENDING_TIER: Record<string, string | null>;
/** Tier label written to cr2a4_pendingtier when a request enters review */
export declare const INITIAL_PENDING_TIER = "Manager";
/** Column name on sct_request for the pending tier routing field */
export declare const PENDING_TIER_COLUMN = "cr2a4_pendingtier";
export declare const TIER1_GROUP_ID = "fdbf07a0-548b-4ccf-875b-900cb1b24a1c";
export declare const TIER2_GROUP_ID = "6a8a4077-c78f-4af2-ae5e-12428d1dcbc3";
export declare const TIER3_GROUP_ID = "e63dbc4f-6ad2-4969-a518-b229836756af";
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
    /** When true, allows self-approval for solo testing across tiers */
    soloTestMode?: boolean;
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
/**
 * handleApprovalAction
 *
 * Executes one approval step in the sequential workflow with strict validation,
 * concurrency checks, Rule R10 enforcement, dual audit logging, and email notifications.
 */
export declare function handleApprovalAction(params: IApprovalActionParams): Promise<IApprovalActionResult>;
/** sct_tiername integer for a given role key */
export declare function roleToTierInt(role: string): number;
/** Next sct_request_status after an approve action */
export declare function nextStatusAfterApprove(role: string): RequestStatusValue;
/** Human-readable label for a tier integer (for display in notifications etc.) */
export declare function tierLabel(tierInt: number): string;
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
export declare function checkStageDecisionLock(record: ODataRecord | null, role: string, approvalLogs: ODataRecord[]): IStageDecisionLock;
//# sourceMappingURL=approvalHandler.d.ts.map