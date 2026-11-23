/**
 * Dataverse entity: sct_requests
 * Publisher prefix: sct_
 *
 * Covers every field from the Fortrex Spend Policy v2.0 (Schedule C).
 * Choice enum values must match the option set values in your Dataverse environment.
 * Run DataverseService.discoverEntityFields('sct_request') to verify logical names.
 */

// ── Choice enumerations ────────────────────────────────────────────────────────

/** sct_requesttype — type of procurement */
export const enum SCTRequestType {
  Goods             = 100000000,
  Services          = 100000001,
  Software          = 100000002,
  CapitalExpenditure= 100000003,
  Works             = 100000004,
}

/** sct_request_status — confirmed from Power Apps option set (sct_Request_Status) */
export const enum SCTRequestStatus {
  Draft      = 854000000,  // confirmed
  Submitted  = 854000001,  // confirmed
  InReview   = 854000002,  // confirmed
  POIssued   = 854000003,  // confirmed
  Rejected   = 854000004,  // confirmed
  Returned   = 854000005,  // confirmed
  Closed     = 854000006,  // confirmed
}

/** sct_urgency — urgency classification */
export const enum SCTUrgency {
  Normal    = 100000000,
  Urgent    = 100000001,
  Emergency = 100000002,
}

/** sct_pricingbasis — how the unit price is expressed */
export const enum SCTPricingBasis {
  Fixed     = 100000000,
  PerUnit   = 100000001,
  Monthly   = 100000002,
  Annual    = 100000003,
  Other     = 100000004,
}

/** statecode — standard Dataverse state */
export const enum SCTRequestStateCode {
  Active   = 0,
  Inactive = 1,
}

// ── Write payload (subset used for create / update) ───────────────────────────

/**
 * Fields you send when creating or updating a request.
 * Lookup fields use the OData bind syntax (e.g. sct_siteid@odata.bind).
 */
export interface ISCTRequestWrite {
  // ── Section B: Classification
  sct_businessunit?:               string;
  sct_costcenter?:                 string;
  sct_costisfor?:                  string;
  sct_function?:                   string;
  sct_requesttype?:                SCTRequestType;
  sct_category?:                   string;       // Set by SCT desk — read-only for requesters

  // ── Section C: Supplier
  sct_isnewsupplier?:              boolean;
  sct_newsuppliername?:            string;
  sct_newsupplierlegalname?:       string;
  sct_newsuppliercontactname?:     string;
  sct_newsuppliercontactemail?:    string;
  sct_newsuppliercontactphone?:    string;

  // ── Section D: Retrospective
  sct_isretrospective?:            boolean;

  // ── Section E: Requirement
  sct_requesttitle?:               string;
  sct_requestdescription?:         string;
  sct_justification?:              string;
  sct_quantity?:                   number;
  sct_unitprice?:                  number;
  sct_estimatedamount?:            number;       // Qty × UnitPrice — calculated client-side
  sct_pricingbasis?:               SCTPricingBasis;
  sct_existingagreementpo?:        string;       // Declare any existing order or contract

  // ── Section F: Delivery & Timeline (hidden when retrospective)
  sct_requiredbydate?:             string;       // ISO date-time
  sct_urgency?:                    SCTUrgency;
  sct_urgencyreason?:              string;
  sct_deliverto?:                  string;
  sct_equipmentassettag?:          string;

  // ── Section G: Compliance flags
  sct_costrechargedtocustomer?:    boolean;
  sct_customername?:               string;
  sct_customeragreementattached?:  boolean;
  sct_techordatainvolved?:         boolean;
  sct_supplieronpremises?:         boolean;
  sct_possiblecapitalpurchase?:    boolean;
  sct_isconfidential?:             boolean;

  // ── Section H: Term dates
  sct_termstartdate?:              string;       // ISO date-time
  sct_termenddate?:                string;       // ISO date-time

  // ── Section I: Declarations
  sct_declarationnopriorcommitment?:     boolean;
  sct_declarationnoconflictofinterest?:  boolean;

  // ── Lookup bindings (OData write syntax)
  'sct_siteid@odata.bind'?:             string;   // /sct_sitesmasters(GUID)
  'sct_supplierid@odata.bind'?:         string;   // /sct_suppliers(GUID)
  'sct_requestorid@odata.bind'?:        string;   // /systemusers(AAD-Object-ID)
  'sct_emergencydeclaredby@odata.bind'?: string;  // /systemusers(AAD-Object-ID)
  'sct_receivingcontact@odata.bind'?:   string;   // /systemusers(AAD-Object-ID)

  // ── Status / workflow (set by service layer)
  sct_status?:                     SCTRequestStatus;
  sct_requestdate?:                string;       // ISO date-time — set on submit
}

// ── Read model (returned by GET) ──────────────────────────────────────────────

export interface ISCTRequest extends ISCTRequestWrite {
  // ── Primary key
  sct_requestid:                   string;       // GUID

  // ── Auto-generated reference
  sct_requestnumber?:              string;       // SCT-{YYYY}-{SEQNUM:4} via Dataverse auto-number

  // ── Procurement-only fields (written by SCT desk / Procurement)
  sct_approvedamount?:             number;
  sct_ponumber?:                   string;
  sct_poissueddate?:               string;       // ISO date-time
  sct_approveddate?:               string;       // ISO date-time
  sct_closeddate?:                 string;       // ISO date-time
  sct_requestnotes?:               string;       // SCT desk return notes

  // ── Lookup value holders (read-only — Dataverse returns these on GET)
  _sct_siteid_value?:              string;       // GUID of linked site
  _sct_supplierid_value?:          string;       // GUID of linked supplier
  _sct_requestorid_value?:         string;       // GUID of requesting systemuser
  _sct_emergencydeclaredby_value?: string;
  _sct_receivingcontact_value?:    string;

  // ── Expanded navigation properties (populated when $expand is used)
  sct_siteid?:             { sct_sitesmasterid: string; sct_sitename: string; sct_sitecode: string; sct_costcenter?: string };
  sct_supplierid?:         { sct_supplierid: string; sct_suppliername: string; sct_ispreferred?: boolean };
  sct_requestorid?:        { systemuserid: string; fullname: string; internalemailaddress: string };
  sct_emergencydeclaredby?:{ systemuserid: string; fullname: string; internalemailaddress: string };
  sct_receivingcontact?:   { systemuserid: string; fullname: string; internalemailaddress: string };

  // ── OData formatted value annotations (returned when Prefer: odata.include-annotations="*")
  'sct_status@OData.Community.Display.V1.FormattedValue'?:       string;
  'sct_requesttype@OData.Community.Display.V1.FormattedValue'?:  string;
  'sct_urgency@OData.Community.Display.V1.FormattedValue'?:      string;
  'sct_pricingbasis@OData.Community.Display.V1.FormattedValue'?: string;
  '_sct_siteid_value@OData.Community.Display.V1.FormattedValue'?: string;
  '_sct_supplierid_value@OData.Community.Display.V1.FormattedValue'?: string;
  '_sct_requestorid_value@OData.Community.Display.V1.FormattedValue'?: string;

  // ── Standard Dataverse system fields
  statecode:               SCTRequestStateCode;
  statuscode:              number;
  createdon?:              string;
  modifiedon?:             string;
  '_createdby_value@OData.Community.Display.V1.FormattedValue'?: string;
}
