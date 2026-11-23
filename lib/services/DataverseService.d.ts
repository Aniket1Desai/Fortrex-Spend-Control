import { AadHttpClient } from '@microsoft/sp-http';
import { ISiteMaster, IRole, ISupplier, ISCTRequest } from '../models';
import { IDataverseService } from './IDataverseService';
/** @deprecated Use DataverseService.baseUrl instead — kept for any legacy direct imports */
export declare const BASE_URL = "https://org1bec48f0.crm8.dynamics.com/api/data/v9.2/";
/** Raw OData record — every field is accessible as string key */
export type ODataRecord = Record<string, unknown>;
/** Minimal systemuser record needed for @odata.bind lookups */
export interface ISystemUser {
    systemuserid: string;
    fullname: string;
    internalemailaddress: string;
    _parentsystemuserid_value?: string;
}
export interface IEntityAttribute {
    LogicalName: string;
    AttributeType: string;
    IsCustomAttribute: boolean;
    IsValidForRead: boolean;
    DisplayName: {
        UserLocalizedLabel?: {
            Label: string;
        };
    };
}
/**
 * Strict EDM schema for creating a new SCT Request record in Dataverse (sct_requests)
 */
export interface ISCTRequestCreatePayload {
    sct_description?: string;
    sct_function?: string;
    sct_urgency_reason?: string;
    sct_emergency_declared_by_email?: string;
    sct_delivery_address?: string;
    sct_equipment_asset_tag?: string;
    sct_customer_name?: string;
    sct_cost_center?: string;
    sct_business_unit?: number;
    sct_cost_is_for?: number;
    sct_purchase_type?: number;
    sct_pricing_basis?: number;
    sct_urgency?: number;
    sct_estimated_total_value?: number;
    sct_term_start_date?: string;
    sct_term_end_date?: string;
    sct_date_needed_by?: string;
    sct_is_new_supplier?: boolean;
    sct_cost_recharged_to_customer?: boolean;
    sct_customer_agreement_attached?: boolean;
    sct_tech_or_data_involved?: boolean;
    sct_supplier_on_premises?: boolean;
    sct_possible_capital_purchase?: boolean;
    sct_is_confidential?: boolean;
    sct_is_retrospective?: boolean;
    sct_declaration_no_prior_commitment?: boolean;
    sct_declaration_no_conflict_of_interest?: boolean;
    sct_request_status?: number;
    cr2a4_pendingtier?: string | null;
    'sct_Requester@odata.bind'?: string;
    'sct_Site@odata.bind'?: string;
    'sct_Suggested_Supplier@odata.bind'?: string;
    'cr2a4_Manager@odata.bind'?: string;
    'cr2a4_manager@odata.bind'?: string;
    [key: string]: unknown;
}
/**
 * Strict EDM schema for creating an Approval Log record in Dataverse (sct_approvallogs)
 */
export interface ISCTApprovalLogCreatePayload {
    sct_name: string;
    sct_tiername: number;
    sct_approvalstatus: number;
    sct_comments: string;
    sct_decisionon: string;
    'sct_requestid@odata.bind': string;
    'sct_approverid@odata.bind'?: string;
    [key: string]: unknown;
}
/**
 * Scrubs all undefined, null, or empty string navigation properties / fields
 * prior to dispatching HTTP POST or PATCH to Dataverse.
 */
export declare function scrubPayload<T extends Record<string, unknown>>(payload: T): T;
/** Safe string cast for a raw OData field value */
export declare function strVal(record: ODataRecord, field: string): string;
export declare class DataverseService implements IDataverseService {
    private readonly _client;
    /** Full OData root, e.g. https://org.crm8.dynamics.com/api/data/v9.2/ */
    readonly baseUrl: string;
    constructor(_client: AadHttpClient, dataverseHost?: string);
    /**
     * Verifies the SCT tables are reachable in this environment.
     * Throws a descriptive error if the entity set is missing or access is denied.
     */
    testConnection(): Promise<void>;
    getApprovalLogs(requestId: string): Promise<ODataRecord[]>;
    getAuditLogs(requestId: string): Promise<ODataRecord[]>;
    createAuditLog(payload: Record<string, unknown>): Promise<void>;
    createApprovalLog(payload: ISCTApprovalLogCreatePayload | Record<string, unknown>): Promise<void>;
    /**
     * Returns the GUIDs of "In Review" requests pending at the given tier,
     * using cr2a4_pendingtier (confirmed column, publisher prefix cr2a4_).
     * Single OData query — no N+1 log fetches.
     *
     * cr2a4_pendingtier values: 'Tier 1' | 'Tier 2' | 'Tier 3' | 'Manager' | null
     */
    getTierPendingRequestIds(tier: 1 | 2 | 3): Promise<string[]>;
    /** Convenience wrapper — badge count only */
    getTier1PendingCount(): Promise<number>;
    /**
     * Returns the count of requests currently pending approval for this manager.
     */
    getManagerPendingCount(systemUserId: string): Promise<number>;
    /**
     * Returns the set of request IDs that the given system user has already approved
     * at the Manager tier. Used to distinguish Pending (not yet acted) from
     * already-forwarded requests in the manager's view.
     *
     * Queries approval logs where this user is either the assigned approver (_sct_approverid_value)
     * or the creator (_createdby_value), and filters for sct_tiername = 1 AND sct_approvalstatus = 1.
     */
    getMyApprovedRequestIds(systemUserId: string): Promise<string[]>;
    /** Returns IDs of requests this approver approved at the given tier. Supports new and legacy tier mappings. */
    getMyTierApprovedRequestIds(systemUserId: string, tierIndex: 1 | 2 | 3 | 4): Promise<string[]>;
    /**
     * Three-source data fetch for the Manager role, merged and deduplicated.
     *
     * Fetch A — authored: sct_requests where _sct_requester_value = userId
     * Fetch B — pending queue: sct_requests where status = In Review AND _cr2a4_manager_value = userId
     * Fetch C — approval history: sct_approvallogs where _createdby_value = userId
     *            → extract request GUIDs → fetch those requests
     *
     * Returns a single deduplicated array sorted newest-first.
     */
    getManagerMasterRequests(systemUserId: string): Promise<ODataRecord[]>;
    private _fetch;
    private _get;
    private _post;
    private _patch;
    private _assertOk;
    /**
     * Returns all readable attributes for the given entity logical name.
     * Use this to discover the real column names before building $select strings.
     */
    discoverEntityFields(entityLogicalName: string): Promise<IEntityAttribute[]>;
    /**
     * Look up the Dataverse systemuser record by systemuserid primary key GUID.
     * Resolves internalemailaddress, fullname, and azureactivedirectoryobjectid.
     */
    getSystemUserById(systemUserId: string): Promise<ISystemUser | null>;
    /**
     * Look up the Dataverse systemuser record for the current AAD user.
     * The azureactivedirectoryobjectid field is the AAD Object ID (from SPFx pageContext.aadInfo.userId).
     * Falls back to null if the user is not found or permissions are insufficient.
     */
    getCurrentSystemUser(aadObjectId: string): Promise<ISystemUser | null>;
    /** Look up systemuser by fullname */
    getSystemUserByName(fullName: string): Promise<ISystemUser | null>;
    /** Fallback: look up by email when AAD Object ID is unavailable */
    getSystemUserByEmail(email: string): Promise<ISystemUser | null>;
    /**
     * Resolves a Dataverse systemuserid string from the systemusers table by manager's email address.
     * Query: GET /api/data/v9.2/systemusers?$select=systemuserid,fullname&$filter=internalemailaddress eq '${managerEmail}'
     */
    getSystemUserIdByEmail(managerEmail: string): Promise<string | null>;
    getManagerSystemUserId(requesterSystemUserId: string): Promise<string | null>;
    /**
     * Returns true if any other systemuser record lists this user as their manager
     * (parentsystemuserid = systemUserId). Used to decide whether to show the
     * "Manager" role option — driven by Dataverse, not AAD org chart.
     * Returns false on any error so the app degrades gracefully.
     */
    hasDirectReportsInDataverse(systemUserId: string): Promise<boolean>;
    /**
     * Returns all systemuserid GUIDs for users whose manager is managerSystemUserId
     * (parentsystemuserid = managerSystemUserId in systemusers table).
     */
    getDirectReportUserIds(managerSystemUserId: string): Promise<string[]>;
    /**
     * Returns the OData navigation property name to use in @odata.bind for a given
     * lookup attribute.
     *
     * Strategy (each step falls through to the next on failure or empty result):
     *   1. ManyToOneRelationships → ReferencingEntityNavigationPropertyName (most accurate)
     *   2. Attributes SchemaName  → e.g. 'cr2a4_Manager' (correct casing for @odata.bind)
     *   3. Caller's fallback or attributeLogicalName (last resort)
     *
     * The last-resort logical name (all-lowercase) triggers Dataverse 400
     * "undeclared property" when the actual nav prop uses mixed casing, which is why
     * step 2 (SchemaName) is critical.
     */
    getNavPropName(entityLogicalName: string, attributeLogicalName: string): Promise<string>;
    getSites(activeOnly?: boolean): Promise<ISiteMaster[]>;
    /** Raw ODataRecord[] for Sites — field names picked at runtime, avoids ISiteMaster mismatch */
    getSitesRaw(activeOnly?: boolean): Promise<ODataRecord[]>;
    getSiteById(id: string): Promise<ISiteMaster | null>;
    getRoles(activeOnly?: boolean): Promise<IRole[]>;
    /** Raw ODataRecord[] for roles — used by ApprovalRoutingService for field-safe access */
    getRawRoles(activeOnly?: boolean): Promise<ODataRecord[]>;
    /**
     * Fetches all active role configuration records with the fields needed by
     * GroupRoleConfig (tier, role name, description/group, spend limit, site name).
     */
    getRoleConfigs(): Promise<ODataRecord[]>;
    /**
     * PATCHes a single sct_role record — used by GroupRoleConfig to persist inline edits.
     */
    updateRoleConfig(roleId: string, payload: Record<string, unknown>): Promise<void>;
    getSuppliers(activeOnly?: boolean): Promise<ISupplier[]>;
    /**
     * Fetch requests with NO $select — returns every column the entity has.
     * This avoids 400 errors from guessed column names.
     * Optional OData $filter string, e.g. "statecode eq 0".
     */
    getRequests(filter?: string): Promise<ISCTRequest[]>;
    /** Also exposes the raw ODataRecord[] for the Dashboard field inspector */
    getRawRequests(filter?: string): Promise<ODataRecord[]>;
    getRequestById(id: string): Promise<ISCTRequest | null>;
    getRawRequestById(id: string): Promise<ODataRecord | null>;
    /** Fetch all requests where the requester lookup matches the given systemuserid GUID.
     *  Tries sct_Requester first (schema name used in @odata.bind); falls back to sct_requestorid. */
    getMyRequests(systemUserId: string): Promise<ODataRecord[]>;
    createRequest(payload: Partial<ISCTRequest> | ISCTRequestCreatePayload | Record<string, unknown>): Promise<string>;
    updateRequest(id: string, payload: Partial<ISCTRequest> | Record<string, unknown>): Promise<void>;
    /** Permanently delete a single sct_request record from Dataverse (HTTP DELETE → 204) */
    deleteRequest(id: string): Promise<void>;
    uploadAnnotation(regardingObjectId: string, filename: string, mimeType: string, base64Body: string): Promise<string>;
    /** Fetch all annotations (file attachments) linked to a given sct_request record */
    getAnnotations(requestId: string): Promise<ODataRecord[]>;
    /** Delete a single annotation (removes the attachment) */
    deleteAnnotation(annotationId: string): Promise<void>;
}
//# sourceMappingURL=DataverseService.d.ts.map