/**
 * ApprovalRoutingService — Schedule A approval routing engine.
 *
 * Implements the two-stage split-approval chain:
 *  Phase 1 — Pre-Request: which levels the requester must attach before submitting.
 *  Phase 2 — Post-Sourcing: does Procurement need FP&A / CFO review before issuing a PO?
 *  Role Resolution — dynamically resolves L1-L4 approver identities from sct_role,
 *                    including active-delegate substitution and Rule R10 self-approval escalation.
 */

import { DataverseService, ODataRecord } from './DataverseService';
import { ApprovalLevel } from '../models';

// ── Spend thresholds ──────────────────────────────────────────────────────────

const T_L1     =  2500;   // L1 (first-line) ceiling
const T_L2     = 10000;   // L2 (department head) ceiling
const T_LEGAL  = 50000;   // Legal Review required above this value

// ── Public interfaces ─────────────────────────────────────────────────────────

/** Which pre-approval level signatures must be collected before submit */
export interface IPreApprovalRequirements {
  levels: ApprovalLevel[];
  description: string;
}

/** Post-sourcing gate result — what additional reviews Procurement must trigger */
export interface IPostSourcingRequirements {
  needsFPAReview:     boolean;
  needsCFOApproval:   boolean;
  needsLegalReview:   boolean;
  description:        string;
}

/** A single resolved approver for one level */
export interface IResolvedApprover {
  level:        ApprovalLevel;
  systemUserId: string;       // Dataverse systemuserid GUID
  name:         string;       // Display name (fullname)
  isDelegate:   boolean;      // true if the active delegate was substituted
  wasEscalated: boolean;      // true if Rule R10 caused escalation from a lower level
  roleId:       string;       // sct_roleid GUID
}

/** Full set of resolved approvers for a request */
export interface IResolvedApprovers {
  approvers:   IResolvedApprover[];
  escalations: string[];       // human-readable escalation notes (Rule R10)
  warnings:    string[];       // e.g. missing level, no active delegate
}

/** Combined approval matrix — all three phases in one call */
export interface IApprovalMatrix {
  preApprovals:     IPreApprovalRequirements;
  postSourcing:     IPostSourcingRequirements;
  resolvedApprovers: IResolvedApprovers;
}

// ── Service ───────────────────────────────────────────────────────────────────

export class ApprovalRoutingService {
  constructor(private readonly _svc: DataverseService) {}

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
  public calculatePreApprovals(
    estimatedValue: number,
    isNewSupplier: boolean
  ): IPreApprovalRequirements {
    if (isNewSupplier) {
      return {
        levels: ['L1', 'L2', 'L3'],
        description:
          'New supplier — pre-approval signatures required from L1 (Line Manager), L2 (Department Head) and L3 (CFO).',
      };
    }
    if (estimatedValue <= T_L1) {
      return {
        levels: ['L1'],
        description: `Value ≤ $${T_L1.toLocaleString()} — L1 (Line Manager) sign-off required.`,
      };
    }
    if (estimatedValue <= T_L2) {
      return {
        levels: ['L2'],
        description: `Value $${T_L1.toLocaleString()}–$${T_L2.toLocaleString()} — L2 (Department Head) sign-off required.`,
      };
    }
    return {
      levels: ['L3'],
      description: `Value > $${T_L2.toLocaleString()} — L3 (CFO) sign-off required.`,
    };
  }

  // ── Phase 2: Post-Sourcing Approval Gate ──────────────────────────────────

  /**
   * Returns what additional post-sourcing approvals Procurement must trigger.
   *
   * Rules:
   *  value > $10k OR new supplier  → FP&A Review required first.
   *  After FP&A: always route to CFO if either condition is true.
   *  value > $50k                  → Legal Review flag raised additionally.
   */
  public calculatePostSourcingRequirements(
    estimatedValue: number,
    isNewSupplier: boolean
  ): IPostSourcingRequirements {
    const needsGate     = estimatedValue > T_L2 || isNewSupplier;
    const needsLegal    = estimatedValue > T_LEGAL;
    const descriptions: string[] = [];

    if (!needsGate) {
      descriptions.push('Within direct Procurement authority — PO may be issued immediately.');
    } else {
      if (isNewSupplier) descriptions.push('New supplier — FP&A review and CFO approval required before PO.');
      if (estimatedValue > T_L2) descriptions.push(`Value > $${T_L2.toLocaleString()} — FP&A review and CFO approval required.`);
      if (needsLegal)    descriptions.push(`Value > $${T_LEGAL.toLocaleString()} — Legal Review flag raised.`);
    }

    return {
      needsFPAReview:   needsGate,
      needsCFOApproval: needsGate,
      needsLegalReview: needsLegal,
      description: descriptions.join(' '),
    };
  }

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
  public async resolveApprovers(
    siteId: string,
    requesterSystemUserId: string
  ): Promise<IResolvedApprovers> {
    const escalations: string[] = [];
    const warnings: string[]    = [];

    // Load all active roles as raw ODataRecord[] to avoid model field-name mismatch risk
    let rawRoles: ODataRecord[] = [];
    try {
      rawRoles = await this._svc.getRawRoles(true);
    } catch {
      warnings.push('Could not load roles from Dataverse — approvers not resolved.');
      return { approvers: [], escalations, warnings };
    }

    // Filter to this site
    const siteRoles = rawRoles.filter(r => {
      const id = r['_sct_siteid_value'];
      return id && String(id).toLowerCase() === siteId.toLowerCase();
    });

    if (!siteRoles.length) {
      warnings.push(`No roles found for site ${siteId}. Verify sct_role records in Dataverse.`);
    }

    // Group by approval level
    const byLevel: Record<string, ODataRecord[]> = {};
    siteRoles.forEach(r => {
      const level = String(r['sct_approvallevel'] ?? r['sct_approval_level'] ?? '').toUpperCase();
      if (!byLevel[level]) byLevel[level] = [];
      byLevel[level].push(r);
    });

    // Build a sorted list of levels to resolve
    const levels: ApprovalLevel[] = ['L1', 'L2', 'L3', 'L4'];
    const resolved: IResolvedApprover[] = [];

    for (const level of levels) {
      const roleRecords = byLevel[level] ?? [];
      if (!roleRecords.length) continue;

      const role = roleRecords[0]; // take first matching role

      // Delegate substitution — use delegate if today is within effective/expiry window
      const now = new Date();
      const effDate = role['sct_effectivedate'] ? new Date(String(role['sct_effectivedate'])) : null;
      const expDate = role['sct_expirydate']    ? new Date(String(role['sct_expirydate']))    : null;
      const delegateActive =
        role['_sct_delegateid_value'] != null &&
        String(role['_sct_delegateid_value']).trim() !== '' &&
        (effDate == null || now >= effDate) &&
        (expDate == null || now <= expDate);

      const personGuid  = String(role['_sct_personid_value']  ?? '');
      const delegateGuid = delegateActive ? String(role['_sct_delegateid_value'] ?? '') : '';
      const activeGuid  = delegateActive ? delegateGuid : personGuid;
      const activeName  = delegateActive
        ? String(role['_sct_delegateid_value@OData.Community.Display.V1.FormattedValue'] ?? delegateGuid)
        : String(role['_sct_personid_value@OData.Community.Display.V1.FormattedValue']   ?? personGuid);

      if (!activeGuid) {
        warnings.push(`${level} role found but no approver GUID — check sct_role record ${String(role['sct_roleid'] ?? '')}.`);
        continue;
      }

      // Rule R10: self-approval conflict — escalate to next level
      if (activeGuid.toLowerCase() === requesterSystemUserId.toLowerCase()) {
        const nextLevel = this._nextLevel(level);
        escalations.push(
          `Rule R10: ${level} approver (${activeName}) is the requester — escalated to ${nextLevel ?? 'Board (no higher level available)'}.`
        );
        if (!nextLevel) {
          warnings.push(`${level} conflict cannot be escalated further — no higher level defined for this site.`);
        }
        continue; // skip this level; the escalation will be picked up when processing the next level
      }

      resolved.push({
        level,
        systemUserId: activeGuid,
        name: activeName,
        isDelegate: delegateActive,
        wasEscalated: false,
        roleId: String(role['sct_roleid'] ?? ''),
      });
    }

    return { approvers: resolved, escalations, warnings };
  }

  // ── Phase 8: Convenience matrix ──────────────────────────────────────────

  /**
   * Computes all three approval phases in one call and returns a combined matrix.
   * Useful for displaying the full approval picture on the Triage and FP&A panels.
   */
  public async computeApprovalMatrix(
    siteId: string,
    requesterSystemUserId: string,
    estimatedValue: number,
    isNewSupplier: boolean
  ): Promise<IApprovalMatrix> {
    const [preApprovals, postSourcing, resolvedApprovers] = await Promise.all([
      Promise.resolve(this.calculatePreApprovals(estimatedValue, isNewSupplier)),
      Promise.resolve(this.calculatePostSourcingRequirements(estimatedValue, isNewSupplier)),
      this.resolveApprovers(siteId, requesterSystemUserId),
    ]);
    return { preApprovals, postSourcing, resolvedApprovers };
  }

  // ── Private helpers ───────────────────────────────────────────────────────

  private _nextLevel(current: ApprovalLevel): ApprovalLevel | null {
    const order: ApprovalLevel[] = ['L1', 'L2', 'L3', 'L4'];
    const idx = order.indexOf(current);
    return idx >= 0 && idx < order.length - 1 ? order[idx + 1] : null;
  }
}
