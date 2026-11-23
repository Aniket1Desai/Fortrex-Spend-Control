/**
 * Dataverse entity: sct_suppliers
 * Publisher prefix: sct_
 *
 * Stores the approved supplier register. Only Active suppliers appear in
 * the Suggested Supplier dropdown on the SCT Request form.
 * Run DataverseService.discoverEntityFields('sct_supplier') to verify logical names.
 */

// ── Choice enumerations ────────────────────────────────────────────────────────

/** sct_supplierstatus — approval / trading status of the supplier */
export const enum SupplierStatus {
  Active   = 100000000,
  OnHold   = 100000001,
  Inactive = 100000002,
}

/**
 * sct_suppliercategory — primary spend category for this supplier.
 * Drives category-owner routing in Stage 4 (Source).
 */
export const enum SupplierCategory {
  Goods       = 100000000,
  Services    = 100000001,
  Works       = 100000002,
  Consulting  = 100000003,
  Software    = 100000004,
  Logistics   = 100000005,
}

export const enum SupplierStateCode {
  Active   = 0,
  Inactive = 1,
}

// ── Interface ─────────────────────────────────────────────────────────────────

export interface ISupplier {
  /** Primary key (GUID) */
  sct_supplierid: string;

  /** Supplier trading / display name — PrimaryNameAttribute confirmed as sct_supplier_name */
  sct_supplier_name: string;
  /** Alias kept for fallback compatibility */
  sct_suppliername?: string;

  /** Full registered legal name */
  sct_legalname?: string;

  /** Unique supplier code used in PO references (e.g. SUP-0042) */
  sct_suppliercode?: string;

  /**
   * Approval / trading status.
   * Only Active suppliers are shown in the requester's Suggested Supplier dropdown.
   * On Hold and Inactive suppliers are treated as "new supplier" for the requester.
   */
  sct_supplierstatus: SupplierStatus;

  /**
   * Whether this supplier is on the Preferred Supplier List (PSL).
   * Maintained by Procurement + Category Owner, endorsed by CFO.
   * Read-only for requesters.
   */
  sct_ispreferred?: boolean;

  /** Primary spend category for this supplier */
  sct_suppliercategory?: SupplierCategory;

  /** Primary contact full name */
  sct_contactname?: string;

  /** Primary contact email */
  sct_contactemail?: string;

  /** Primary contact phone */
  sct_contactphone?: string;

  /** Registered address */
  sct_address?: string;

  /** Whether a Non-Disclosure Agreement is in place */
  sct_hasnda?: boolean;

  /** Whether a Master Service Agreement is in place */
  sct_hasmsa?: boolean;

  /** ISO date — when compliance certification or insurance expires */
  sct_complianceexpiry?: string;

  /** Agreed payment terms (e.g. "Net 30") */
  sct_paymentterms?: string;

  /** Free-text notes from Procurement */
  sct_suppliernotes?: string;

  // ── OData formatted value annotations ────────────────────────────────────────

  'sct_supplierstatus@OData.Community.Display.V1.FormattedValue'?:  string;
  'sct_suppliercategory@OData.Community.Display.V1.FormattedValue'?: string;

  // ── Standard Dataverse system fields ─────────────────────────────────────────

  /** OData: state (0 = Active, 1 = Inactive) */
  statecode: SupplierStateCode;

  /** OData: status reason */
  statuscode: number;

  /** ISO date-time — set by Dataverse */
  createdon?: string;

  /** ISO date-time — set by Dataverse */
  modifiedon?: string;
}
