/**
 * Dataverse entity: sct_roles
 * Publisher prefix: sct_
 *
 * Stores the approval hierarchy (L1 – L4) per site.
 * Only approvers appear in this table — it does NOT contain all employees.
 * Run DataverseService.discoverEntityFields('sct_role') to verify logical names.
 */

export const enum RoleStateCode {
  Active   = 0,
  Inactive = 1,
}

/** Approval level — used to drive routing in Stage 3 Triage */
export type ApprovalLevel = 'L1' | 'L2' | 'L3' | 'L4';

export interface IRole {
  /** Primary key (GUID) */
  sct_roleid: string;

  /** Role display name (e.g. "Site Manager - Brisbane") */
  sct_rolename: string;

  /** Optional description of duties */
  sct_roledescription?: string;

  /**
   * Approval tier this role operates at.
   * L1 = first-line, L2 = department head, L3 = CFO, L4 = Board (if applicable).
   * Stored as a text choice to match Schedule A labels.
   */
  sct_approvallevel?: ApprovalLevel;

  /**
   * Maximum spend value (in home currency) this role can approve.
   * Determined by Schedule A. Requests above this threshold escalate to the next level.
   */
  sct_approvalspendlimit?: number;

  // ── Effective dates ──────────────────────────────────────────────────────────

  /** ISO date-time — when this role assignment became effective */
  sct_effectivedate?: string;

  /** ISO date-time — when this role assignment expires (null = no expiry) */
  sct_expirydate?: string;

  // ── Lookup: Site ─────────────────────────────────────────────────────────────

  /** GUID of the site this role covers */
  _sct_siteid_value?: string;
  'sct_siteid@odata.bind'?: string;
  sct_siteid?: {
    sct_sitesmasterid: string;
    sct_sitename: string;
    sct_sitecode: string;
  };

  // ── Lookup: Person (primary approver) ────────────────────────────────────────

  /** Azure AD Object ID of the approver (systemuser.azureactivedirectoryobjectid) */
  _sct_personid_value?: string;
  'sct_personid@odata.bind'?: string;
  sct_personid?: {
    systemuserid: string;
    fullname: string;
    internalemailaddress: string;
  };

  // ── Lookup: Delegate (backup approver) ───────────────────────────────────────

  /** Azure AD Object ID of the delegate who covers when the person is absent */
  _sct_delegateid_value?: string;
  'sct_delegateid@odata.bind'?: string;
  sct_delegateid?: {
    systemuserid: string;
    fullname: string;
    internalemailaddress: string;
  };

  // ── OData formatted value annotations ────────────────────────────────────────

  '_sct_siteid_value@OData.Community.Display.V1.FormattedValue'?:     string;
  '_sct_personid_value@OData.Community.Display.V1.FormattedValue'?:   string;
  '_sct_delegateid_value@OData.Community.Display.V1.FormattedValue'?: string;

  // ── Standard Dataverse system fields ─────────────────────────────────────────

  /** OData: state (0 = Active, 1 = Inactive) */
  statecode: RoleStateCode;

  /** OData: status reason */
  statuscode: number;

  /** ISO date-time — set by Dataverse */
  createdon?: string;

  /** ISO date-time — set by Dataverse */
  modifiedon?: string;
}
