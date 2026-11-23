/**
 * Dataverse entity: sct_requests
 * Publisher prefix: sct_
 *
 * Covers every field from the Fortrex Spend Policy v2.0 (Schedule C).
 * Choice enum values must match the option set values in your Dataverse environment.
 * Run DataverseService.discoverEntityFields('sct_request') to verify logical names.
 */
/** sct_requesttype — type of procurement */
export declare const enum SCTRequestType {
    Goods = 100000000,
    Services = 100000001,
    Software = 100000002,
    CapitalExpenditure = 100000003,
    Works = 100000004
}
/** sct_request_status — confirmed from Power Apps option set (sct_Request_Status): 854000001 to 854000007 */
export declare const enum SCTRequestStatus {
    Draft = 854000001,// confirmed
    Submitted = 854000002,// confirmed
    InReview = 854000003,// confirmed
    POIssued = 854000004,// confirmed
    Rejected = 854000005,// confirmed
    Returned = 854000006,// confirmed
    Closed = 854000007
}
/** sct_urgency — urgency classification */
export declare const enum SCTUrgency {
    Normal = 100000000,
    Urgent = 100000001,
    Emergency = 100000002
}
/** sct_pricingbasis — how the unit price is expressed */
export declare const enum SCTPricingBasis {
    Fixed = 100000000,
    PerUnit = 100000001,
    Monthly = 100000002,
    Annual = 100000003,
    Other = 100000004
}
/** statecode — standard Dataverse state */
export declare const enum SCTRequestStateCode {
    Active = 0,
    Inactive = 1
}
/**
 * Fields you send when creating or updating a request.
 * Lookup fields use the OData bind syntax (e.g. sct_siteid@odata.bind).
 */
export interface ISCTRequestWrite {
    sct_businessunit?: string;
    sct_costcenter?: string;
    sct_costisfor?: string;
    sct_function?: string;
    sct_requesttype?: SCTRequestType;
    sct_category?: string;
    sct_isnewsupplier?: boolean;
    sct_newsuppliername?: string;
    sct_newsupplierlegalname?: string;
    sct_newsuppliercontactname?: string;
    sct_newsuppliercontactemail?: string;
    sct_newsuppliercontactphone?: string;
    sct_isretrospective?: boolean;
    sct_requesttitle?: string;
    sct_requestdescription?: string;
    sct_justification?: string;
    sct_quantity?: number;
    sct_unitprice?: number;
    sct_estimatedamount?: number;
    sct_pricingbasis?: SCTPricingBasis;
    sct_existingagreementpo?: string;
    sct_requiredbydate?: string;
    sct_urgency?: SCTUrgency;
    sct_urgencyreason?: string;
    sct_deliverto?: string;
    sct_equipmentassettag?: string;
    sct_costrechargedtocustomer?: boolean;
    sct_customername?: string;
    sct_customeragreementattached?: boolean;
    sct_techordatainvolved?: boolean;
    sct_supplieronpremises?: boolean;
    sct_possiblecapitalpurchase?: boolean;
    sct_isconfidential?: boolean;
    sct_termstartdate?: string;
    sct_termenddate?: string;
    sct_declarationnopriorcommitment?: boolean;
    sct_declarationnoconflictofinterest?: boolean;
    'sct_siteid@odata.bind'?: string;
    'sct_supplierid@odata.bind'?: string;
    'sct_requestorid@odata.bind'?: string;
    'sct_emergencydeclaredby@odata.bind'?: string;
    'sct_receivingcontact@odata.bind'?: string;
    sct_status?: SCTRequestStatus;
    sct_requestdate?: string;
}
export interface ISCTRequest extends ISCTRequestWrite {
    sct_requestid: string;
    sct_requestnumber?: string;
    sct_approvedamount?: number;
    sct_ponumber?: string;
    sct_poissueddate?: string;
    sct_approveddate?: string;
    sct_closeddate?: string;
    sct_requestnotes?: string;
    _sct_siteid_value?: string;
    _sct_supplierid_value?: string;
    _sct_requestorid_value?: string;
    _sct_emergencydeclaredby_value?: string;
    _sct_receivingcontact_value?: string;
    sct_siteid?: {
        sct_sitesmasterid: string;
        sct_sitename: string;
        sct_sitecode: string;
        sct_costcenter?: string;
    };
    sct_supplierid?: {
        sct_supplierid: string;
        sct_suppliername: string;
        sct_ispreferred?: boolean;
    };
    sct_requestorid?: {
        systemuserid: string;
        fullname: string;
        internalemailaddress: string;
    };
    sct_emergencydeclaredby?: {
        systemuserid: string;
        fullname: string;
        internalemailaddress: string;
    };
    sct_receivingcontact?: {
        systemuserid: string;
        fullname: string;
        internalemailaddress: string;
    };
    'sct_status@OData.Community.Display.V1.FormattedValue'?: string;
    'sct_requesttype@OData.Community.Display.V1.FormattedValue'?: string;
    'sct_urgency@OData.Community.Display.V1.FormattedValue'?: string;
    'sct_pricingbasis@OData.Community.Display.V1.FormattedValue'?: string;
    '_sct_siteid_value@OData.Community.Display.V1.FormattedValue'?: string;
    '_sct_supplierid_value@OData.Community.Display.V1.FormattedValue'?: string;
    '_sct_requestorid_value@OData.Community.Display.V1.FormattedValue'?: string;
    statecode: SCTRequestStateCode;
    statuscode: number;
    createdon?: string;
    modifiedon?: string;
    '_createdby_value@OData.Community.Display.V1.FormattedValue'?: string;
}
//# sourceMappingURL=ISCTRequest.d.ts.map