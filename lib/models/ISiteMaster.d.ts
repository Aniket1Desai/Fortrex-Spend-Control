/**
 * Dataverse entity: sct_sitesmaster
 * Publisher prefix: sct_
 *
 * Run DataverseService.discoverEntityFields('sct_sitesmaster') to verify logical names.
 */
export declare const enum SiteMasterStateCode {
    Active = 0,
    Inactive = 1
}
export interface ISiteMaster {
    /** Primary key (GUID) */
    sct_sitesmasterid: string;
    /** Short site code — unique identifier used in SCT Reference and PO */
    sct_sitecode: string;
    /** Full site / location name shown in the Site dropdown */
    sct_sitename: string;
    /** Physical address */
    sct_location?: string;
    /** Regional grouping (e.g. APAC, EMEA, Americas) */
    sct_region?: string;
    /** Business unit this site belongs to — auto-fills Business Unit on form */
    sct_businessunit?: string;
    /**
     * Cost centre code — auto-filled on the SCT Request form when the
     * requester selects this site. Requester cannot edit it.
     */
    sct_costcenter?: string;
    /**
     * General Ledger account code — written onto the PO by Procurement in Stage 5.
     * Stored here so Procurement can look it up from the site record.
     */
    sct_glaccountcode?: string;
    /** Whether the site is active for new requests */
    sct_isactive: boolean;
    /** Lookup → systemuser: the budget owner responsible for this site */
    _sct_budgetownerid_value?: string;
    'sct_budgetownerid@odata.bind'?: string;
    sct_budgetownerid?: {
        systemuserid: string;
        fullname: string;
        internalemailaddress: string;
    };
    /** OData: state (0 = Active, 1 = Inactive) */
    statecode: SiteMasterStateCode;
    /** OData: status reason */
    statuscode: number;
    /** ISO date-time — set by Dataverse */
    createdon?: string;
    /** ISO date-time — set by Dataverse */
    modifiedon?: string;
}
//# sourceMappingURL=ISiteMaster.d.ts.map