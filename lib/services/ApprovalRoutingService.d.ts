/**
 * ApprovalRoutingService — Schedule A approval routing engine.
 *
 * Implements the two-stage split-approval chain:
 *  Phase 1 — Pre-Request: which levels the requester must attach before submitting.
 *  Phase 2 — Post-Sourcing: does Procurement need FP&A / CFO review before issuing a PO?
 *  Role Resolution — dynamically resolves L1-L4 approver identities from sct_role,
 *                    including active-delegate substitution and Rule R10 self-approval escalation.
 */
import { DataverseService } from './DataverseService';
import { ApprovalLevel } from '../models';
/** Which pre-approval level signatures must be collected before submit */
export interface IPreApprovalRequirements {
    levels: ApprovalLevel[];
    description: string;
}
/** Post-sourcing gate result — what additional reviews Procurement must trigger */
export interface IPostSourcingRequirements {
    needsFPAReview: boolean;
    needsCFOApproval: boolean;
    needsLegalReview: boolean;
    description: string;
}
/** A single resolved approver for one level */
export interface IResolvedApprover {
    level: ApprovalLevel;
    systemUserId: string;
    name: string;
    isDelegate: boolean;
    wasEscalated: boolean;
    roleId: string;
}
/** Full set of resolved approvers for a request */
export interface IResolvedApprovers {
    approvers: IResolvedApprover[];
    escalations: string[];
    warnings: string[];
}
/** Combined approval matrix — all three phases in one call */
export interface IApprovalMatrix {
    preApprovals: IPreApprovalRequirements;
    postSourcing: IPostSourcingRequirements;
    resolvedApprovers: IResolvedApprovers;
}
export declare class ApprovalRoutingService {
    private readonly _svc;
    constructor(_svc: DataverseService);
    /**
     * Returns the approval level(s) the requester must collect BEFORE submission.
     *
     * Schedule A rules:
     *  New supplier (any value) → L1 + L2 + L3
     *  Known supplier, value ≤ $2,500                 → L1 only
     *  Known supplier, $2,500 < value ≤ $10,000        → L2 only
     *  Known supplier, value > $10,000                 → L3 only
     */
    calculatePreApprovals(estimatedValue: number, isNewSupplier: boolean): IPreApprovalRequirements;
    /**
     * Returns what additional post-sourcing approvals Procurement must trigger.
     *
     * Rules:
     *  value > $10k OR new supplier  → FP&A Review required first.
     *  After FP&A: always route to CFO if either condition is true.
     *  value > $50k                  → Legal Review flag raised additionally.
     */
    calculatePostSourcingRequirements(estimatedValue: number, isNewSupplier: boolean): IPostSourcingRequirements;
    /**
     * Resolves L1–L4 approvers from the sct_role table for the given site.
     * Applies:
     *  — Active-delegate substitution (today within effectivedate..expirydate)
     *  — Rule R10: if approver GUID = requesterSystemUserId, escalate to next level.
     *
     * @param siteId             GUID value from _sct_site_value (the site's sct_sitesmasterid)
     * @param requesterSystemUserId  Dataverse systemuserid GUID of the person submitting
     */
    resolveApprovers(siteId: string, requesterSystemUserId: string): Promise<IResolvedApprovers>;
    /**
     * Computes all three approval phases in one call and returns a combined matrix.
     * Useful for displaying the full approval picture on the Triage and FP&A panels.
     */
    computeApprovalMatrix(siteId: string, requesterSystemUserId: string, estimatedValue: number, isNewSupplier: boolean): Promise<IApprovalMatrix>;
    private _nextLevel;
}
//# sourceMappingURL=ApprovalRoutingService.d.ts.map