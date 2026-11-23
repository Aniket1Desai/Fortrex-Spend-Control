import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { ISiteMaster, IRole, ISupplier, ISCTRequest } from '../models';
import { IDataverseService } from './IDataverseService';

/** Fallback used only when the web part property has not been set */
const DEFAULT_DATAVERSE_HOST = 'https://org1bec48f0.crm8.dynamics.com';

/** @deprecated Use DataverseService.baseUrl instead — kept for any legacy direct imports */
export const BASE_URL = `${DEFAULT_DATAVERSE_HOST}/api/data/v9.2/`;

/** Raw OData record — every field is accessible as string key */
export type ODataRecord = Record<string, unknown>;

/** Minimal systemuser record needed for @odata.bind lookups */
export interface ISystemUser {
  systemuserid: string;
  fullname: string;
  internalemailaddress: string;
  _parentsystemuserid_value?: string; // Manager's systemuserid (from systemusers.parentsystemuserid)
}

export interface IEntityAttribute {
  LogicalName: string;
  AttributeType: string;
  IsCustomAttribute: boolean;
  IsValidForRead: boolean;
  DisplayName: { UserLocalizedLabel?: { Label: string } };
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
export function scrubPayload<T extends Record<string, unknown>>(payload: T): T {
  const clean = { ...payload };
  Object.keys(clean).forEach(key => {
    const val = clean[key];
    if (val === undefined || val === null || val === '') {
      if (key.indexOf('@odata.bind') !== -1 || val === undefined) {
        delete clean[key];
      }
    }
  });
  return clean;
}

const OData_HEADERS: Record<string, string> = {
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
  'Accept': 'application/json',
  'Content-Type': 'application/json; charset=utf-8',
  // Returns "fieldname@OData.Community.Display.V1.FormattedValue" alongside each raw value
  'Prefer': 'odata.include-annotations="*"',
  // Prevent browser/proxy caching of OData GET responses so post-approval reloads
  // always receive the latest committed Dataverse data, not a stale cached copy.
  'Cache-Control': 'no-cache',
  'Pragma': 'no-cache',
};

/** Pull GUID out of the OData-EntityId response header URL */
function extractGuidFromEntityId(entityIdHeader: string): string {
  const match = entityIdHeader.match(/\(([0-9a-f-]{36})\)/i);
  if (!match) throw new Error(`Cannot parse entity id from header: ${entityIdHeader}`);
  return match[1];
}

/** Safe string cast for a raw OData field value */
export function strVal(record: ODataRecord, field: string): string {
  // Prefer the OData formatted-value annotation when it exists
  const annotated = record[`${field}@OData.Community.Display.V1.FormattedValue`];
  if (annotated !== undefined && annotated !== null) return String(annotated);
  const raw = record[field];
  return raw !== undefined && raw !== null ? String(raw) : '';
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function assertGuid(value: string, label: string): void {
  if (!GUID_RE.test(value)) {
    throw new Error(`Invalid ${label}: expected a GUID, got "${value.slice(0, 40)}"`);
  }
}

function encodeODataString(val: string): string {
  return val.replace(/'/g, "''");
}

export class DataverseService implements IDataverseService {
  /** Full OData root, e.g. https://org.crm8.dynamics.com/api/data/v9.2/ */
  public readonly baseUrl: string;

  constructor(private readonly _client: AadHttpClient, dataverseHost?: string) {
    const host = (dataverseHost || DEFAULT_DATAVERSE_HOST).replace(/\/$/, '');
    this.baseUrl = `${host}/api/data/v9.2/`;
  }

  // ── Health check ────────────────────────────────────────────────────────────

  /**
   * Verifies the SCT tables are reachable in this environment.
   * Throws a descriptive error if the entity set is missing or access is denied.
   */
  public async testConnection(): Promise<void> {
    const response = await this._fetch(
      `${this.baseUrl}sct_requests?$top=1&$select=sct_requestid`
    );
    if (!response.ok) {
      const status = response.status;
      if (status === 404) {
        throw new Error(
          `SCT tables not found in this Dataverse environment (HTTP 404). ` +
          `Please ensure the Fortrex SCT solution has been imported and published in Power Apps before using this web part.`
        );
      }
      if (status === 401 || status === 403) {
        throw new Error(
          `Access denied to Dataverse (HTTP ${status}). ` +
          `Ensure the AAD API permission (user_impersonation) has been approved in SharePoint Admin Centre → API Access.`
        );
      }
      throw new Error(`Dataverse connectivity check failed (HTTP ${status}).`);
    }
  }

  // ── History logs ────────────────────────────────────────────────────────────

  public async getApprovalLogs(requestId: string): Promise<ODataRecord[]> {
    assertGuid(requestId, 'requestId');
    const select = [
      'sct_approvallogid', 'sct_name',
      'sct_approvalstatus', 'sct_tiername',
      '_sct_approverid_value', '_createdby_value',
      'sct_decisionon', 'sct_comments', 'createdon',
    ].join(',');
    const logs = await this._get<ODataRecord[]>(
      `sct_approvallogs?$select=${select}&$filter=_sct_requestid_value eq ${requestId}&$orderby=createdon desc`
    );
    // Explicit client-side sort: newest first by createdon
    return Array.isArray(logs)
      ? [...logs].sort((a, b) => {
          let sA = String(a['createdon'] ?? a['sct_decisionon'] ?? '').trim();
          let sB = String(b['createdon'] ?? b['sct_decisionon'] ?? '').trim();
          if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(sA)) sA += 'Z';
          if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(sB)) sB += 'Z';
          const tA = new Date(sA).getTime() || 0;
          const tB = new Date(sB).getTime() || 0;
          if (tB !== tA) return tB - tA;
          return String(b['sct_name'] ?? '').localeCompare(String(a['sct_name'] ?? ''));
        })
      : [];
  }

  public async getAuditLogs(requestId: string): Promise<ODataRecord[]> {
    assertGuid(requestId, 'requestId');
    const select = [
      'sct_auditlogid', 'sct_name',
      'sct_action', '_sct_actionby_value',
      'sct_description', 'sct_oldvalue', 'sct_newvalue', 'createdon',
    ].join(',');
    const logs = await this._get<ODataRecord[]>(
      `sct_auditlogs?$select=${select}&$filter=_sct_requestid_value eq ${requestId}&$orderby=createdon desc`
    );
    // Explicit client-side sort: newest first by createdon
    return Array.isArray(logs)
      ? [...logs].sort((a, b) => {
          let sA = String(a['createdon'] ?? '').trim();
          let sB = String(b['createdon'] ?? '').trim();
          if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(sA)) sA += 'Z';
          if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(sB)) sB += 'Z';
          const tA = new Date(sA).getTime() || 0;
          const tB = new Date(sB).getTime() || 0;
          if (tB !== tA) return tB - tA;
          return String(b['sct_name'] ?? '').localeCompare(String(a['sct_name'] ?? ''));
        })
      : [];
  }

  public async createAuditLog(payload: Record<string, unknown>): Promise<void> {
    await this._post('sct_auditlogs', payload);
  }

  public async createApprovalLog(payload: ISCTApprovalLogCreatePayload | Record<string, unknown>): Promise<void> {
    const cleanPayload = scrubPayload(payload);
    await this._post('sct_approvallogs', cleanPayload);
  }

  /**
   * Returns the GUIDs of "In Review" requests pending at the given tier,
   * using cr2a4_pendingtier (confirmed column, publisher prefix cr2a4_).
   * Single OData query — no N+1 log fetches.
   *
   * cr2a4_pendingtier values: 'Tier 1' | 'Tier 2' | 'Tier 3' | 'Manager' | null
   */
  public async getTierPendingRequestIds(tier: 1 | 2 | 3): Promise<string[]> {
    const tierLabel = tier === 1 ? 'Tier 1' : tier === 2 ? 'Tier 2' : 'Tier 3';
    const altLabel  = `Tier${tier}`;
    try {
      const filter = encodeURIComponent(
        `(sct_request_status eq 854000003 or sct_request_status eq 854000002) and (cr2a4_pendingtier eq '${tierLabel}' or cr2a4_pendingtier eq '${altLabel}' or cr2a4_pendingtier eq '${tier}')`
      );
      const records = await this._get<ODataRecord[]>(
        `sct_requests?$select=sct_requestid&$filter=${filter}&$top=500`
      );
      if (records && records.length > 0) {
        return records.map(r => String(r['sct_requestid'] ?? '')).filter(id => GUID_RE.test(id));
      }
    } catch { /* try in-memory fallback below */ }

    try {
      const all = await this.getRawRequests().catch(() => []);
      const matched = all.filter(r => {
        const rawStatus = typeof r['sct_request_status'] === 'number'
          ? r['sct_request_status']
          : parseInt(String(r['sct_request_status'] ?? r['sct_requeststatus'] ?? ''), 10);
        const statusLabel = (strVal(r, 'sct_request_status') || strVal(r, 'sct_requeststatus') || '').toLowerCase();
        const isReview = rawStatus === 854000003 || rawStatus === 854000002 || statusLabel.indexOf('in review') !== -1 || statusLabel.indexOf('tier') !== -1;
        if (!isReview) return false;
        const pt = String(r['cr2a4_pendingtier'] || r['sct_pendingtier'] || r['pendingtier'] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        return pt === `tier${tier}` || pt === `${tier}` || statusLabel.indexOf(`tier ${tier}`) !== -1 || statusLabel.indexOf(`tier${tier}`) !== -1;
      });
      return matched.map(r => String(r['sct_requestid'] ?? '')).filter(id => GUID_RE.test(id));
    } catch {
      return [];
    }
  }

  /** Convenience wrapper — badge count only */
  public async getTier1PendingCount(): Promise<number> {
    return this.getTierPendingRequestIds(1).then(ids => ids.length).catch(() => 0);
  }

  /**
   * Returns the count of requests currently pending approval for this manager.
   */
  public async getManagerPendingCount(systemUserId: string): Promise<number> {
    assertGuid(systemUserId, 'system user id');
    try {
      const directReportIds = await this.getDirectReportUserIds(systemUserId).catch((): string[] => []);
      const safeUserId = systemUserId.toLowerCase().replace(/[{}]/g, '');
      const reportsSet = new Set(directReportIds.map(id => id.toLowerCase().replace(/[{}]/g, '')));

      let managerFilter = `_cr2a4_manager_value eq ${systemUserId}`;
      if (directReportIds.length > 0) {
        const reportsOr = directReportIds.slice(0, 15).map(id => `_sct_requester_value eq ${id}`).join(' or ');
        managerFilter = `(${managerFilter} or ${reportsOr})`;
      }
      const filterB1 = encodeURIComponent(
        `(sct_request_status eq 854000002 or sct_request_status eq 854000003 or sct_request_status eq 854000001) and ${managerFilter}`
      );
      const filterB2 = encodeURIComponent(
        `(sct_request_status eq 854000002 or sct_request_status eq 854000003 or sct_request_status eq 854000001) and _sct_manager_value eq ${systemUserId}`
      );

      const [reqs1, reqs2, allRaw, approvedIdsArr] = await Promise.all([
        this.getRawRequests(filterB1).catch((): ODataRecord[] => []),
        this.getRawRequests(filterB2).catch((): ODataRecord[] => []),
        this.getRawRequests().catch((): ODataRecord[] => []),
        this.getMyApprovedRequestIds(systemUserId).catch((): string[] => []),
      ]);
      const approvedIds = new Set(approvedIdsArr);

      const seen = new Set<string>();
      const candidateReqs: ODataRecord[] = [];
      for (const r of [...reqs1, ...reqs2, ...allRaw]) {
        const id = String(r['sct_requestid'] ?? '');
        if (id && !seen.has(id)) {
          seen.add(id);
          candidateReqs.push(r);
        }
      }

      return candidateReqs.filter(r => {
        const id = String(r['sct_requestid'] ?? '');
        if (!id || approvedIds.has(id)) return false;

        const rawReq = r['_sct_requester_value'] ?? r['_createdby_value'] ?? r['sct_requester'] ?? '';
        const reqId = rawReq ? String(rawReq).toLowerCase().replace(/[{}]/g, '') : '';
        const createdById = r['_createdby_value'] ? String(r['_createdby_value']).toLowerCase().replace(/[{}]/g, '') : '';
        if (reqId === safeUserId || createdById === safeUserId) return false;

        const rawMgr = r['_sct_manager_value'] ?? r['_cr2a4_manager_value'] ?? r['_manager_value'] ?? r['sct_manager'] ?? '';
        const mgrId = rawMgr ? String(rawMgr).toLowerCase().replace(/[{}]/g, '') : '';
        const mgrFormatted = (strVal(r, '_cr2a4_manager_value') || strVal(r, '_sct_manager_value') || strVal(r, 'cr2a4_manager') || strVal(r, 'manager') || '').toLowerCase();
        const isDirectReport = Boolean(reportsSet.has(reqId) || (createdById && reportsSet.has(createdById)));
        const isManagerOfReq = (mgrId === safeUserId) || isDirectReport;
        if (!isManagerOfReq) return false;

        const rawStatus = typeof r['sct_request_status'] === 'number'
          ? r['sct_request_status']
          : parseInt(String(r['sct_request_status'] ?? r['sct_requeststatus'] ?? ''), 10);
        const statusLabel = (strVal(r, 'sct_request_status') || strVal(r, 'sct_requeststatus') || '').toLowerCase();
        const isPendingStatus = rawStatus === 854000002 || rawStatus === 854000003 || rawStatus === 854000001 ||
          statusLabel === 'submitted' || statusLabel === 'in review' ||
          statusLabel.indexOf('submitted') !== -1 || statusLabel.indexOf('in review') !== -1;
        if (!isPendingStatus) return false;

        const pt = String(r['cr2a4_pendingtier'] || r['sct_pendingtier'] || r['pendingtier'] || '').toLowerCase().trim();
        const isPendingWithManager = pt === 'manager' || !pt;
        return isPendingWithManager;
      }).length;
    } catch {
      return 0;
    }
  }

  /**
   * Returns the set of request IDs that the given system user has already approved
   * at the Manager tier. Used to distinguish Pending (not yet acted) from
   * already-forwarded requests in the manager's view.
   *
   * Queries approval logs where this user is either the assigned approver (_sct_approverid_value)
   * or the creator (_createdby_value), and filters for sct_tiername = 1 AND sct_approvalstatus = 1.
   */
  public async getMyApprovedRequestIds(systemUserId: string): Promise<string[]> {
    assertGuid(systemUserId, 'system user id');
    // Manager tier: sct_tiername = 4 (legacy fallback: 1 if sct_name indicates Manager); sct_approvalstatus = 1 (Approved)
    const filter = encodeURIComponent(
      `(_sct_approverid_value eq ${systemUserId} or _createdby_value eq ${systemUserId}) and (sct_tiername eq 4 or sct_tiername eq 1) and sct_approvalstatus eq 1`
    );
    const logs = await this._get<ODataRecord[]>(
      `sct_approvallogs?$select=_sct_requestid_value,sct_name,sct_tiername,sct_approvalstatus&$filter=${filter}&$top=500`
    ).catch((): ODataRecord[] => []);
    return Array.from(new Set(
      logs
        .filter(l => {
          const t = l['sct_tiername'] as number;
          const status = l['sct_approvalstatus'] as number;
          if (status !== 1) return false;
          if (t === 4) return true;
          if (t === 1) {
            const name = String(l['sct_name'] ?? '').toLowerCase();
            return name.indexOf('manager') !== -1;
          }
          return false;
        })
        .map(l => String(l['_sct_requestid_value'] ?? ''))
        .filter(id => GUID_RE.test(id))
    ));
  }

  /** Returns IDs of requests this approver approved at the given tier. Supports new and legacy tier mappings. */
  public async getMyTierApprovedRequestIds(systemUserId: string, tierIndex: 1 | 2 | 3 | 4): Promise<string[]> {
    assertGuid(systemUserId, 'system user id');
    // Tier 1: sct_tiername = 1
    // Tier 2: sct_tiername = 2
    // Tier 3: sct_tiername = 3
    const filter = encodeURIComponent(
      `(_sct_approverid_value eq ${systemUserId} or _createdby_value eq ${systemUserId}) and sct_tiername eq ${tierIndex} and sct_approvalstatus eq 1`
    );
    const logs = await this._get<ODataRecord[]>(
      `sct_approvallogs?$select=_sct_requestid_value,sct_name,sct_tiername,sct_approvalstatus&$filter=${filter}&$top=500`
    ).catch((): ODataRecord[] => []);
    return Array.from(new Set(
      logs
        .filter(l => {
          const t = l['sct_tiername'] as number;
          const status = l['sct_approvalstatus'] as number;
          if (status !== 1) return false;
          if (tierIndex === 1) {
            // Ensure this is not a Manager log stamped with legacy tiername 1
            const name = String(l['sct_name'] ?? '').toLowerCase();
            if (name.indexOf('manager') !== -1) return false;
          }
          return t === tierIndex;
        })
        .map(l => String(l['_sct_requestid_value'] ?? ''))
        .filter(id => GUID_RE.test(id))
    ));
  }

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
  public async getManagerMasterRequests(systemUserId: string): Promise<ODataRecord[]> {
    assertGuid(systemUserId, 'system user id');

    // Fetch C part 1: get all request IDs this user ever acted on (any tier/status)
    const allLogs = await this._get<ODataRecord[]>(
      `sct_approvallogs?$select=_sct_requestid_value` +
      `&$filter=_createdby_value eq ${systemUserId}&$top=500`
    ).catch((): ODataRecord[] => []);
    const loggedIds = Array.from(new Set(
      allLogs
        .map(l => String(l['_sct_requestid_value'] ?? ''))
        .filter(id => GUID_RE.test(id))
    )).slice(0, 50); // cap to avoid huge OR filter

    // Fetch direct report user IDs from Dataverse systemusers (parentsystemuserid = systemUserId)
    const directReportIds = await this.getDirectReportUserIds(systemUserId).catch((): string[] => []);
    let managerFilter = `_cr2a4_manager_value eq ${systemUserId}`;
    if (directReportIds.length > 0) {
      const reportsOr = directReportIds.slice(0, 15).map(id => `_sct_requester_value eq ${id}`).join(' or ');
      managerFilter = `(${managerFilter} or ${reportsOr})`;
    }

    // Fetch A, B1, B2, C2, and allRaw in parallel
    const filterA = encodeURIComponent(`_sct_requester_value eq ${systemUserId}`);
    // Fetch B: requests assigned to this manager or submitted by their direct reports
    const filterB1 = encodeURIComponent(
      `(sct_request_status eq 854000002 or sct_request_status eq 854000003 or sct_request_status eq 854000001) and ${managerFilter}`
    );
    const filterB2 = encodeURIComponent(
      `(sct_request_status eq 854000002 or sct_request_status eq 854000003 or sct_request_status eq 854000001) and _sct_manager_value eq ${systemUserId}`
    );

    const [fetchA, fetchB1, fetchB2, fetchC, allRaw] = await Promise.all([
      this.getRawRequests(filterA).catch((): ODataRecord[] => []),
      this.getRawRequests(filterB1).catch((): ODataRecord[] => []),
      this.getRawRequests(filterB2).catch((): ODataRecord[] => []),
      loggedIds.length > 0
        ? this.getRawRequests(
            encodeURIComponent(loggedIds.map(id => `sct_requestid eq ${id}`).join(' or '))
          ).catch((): ODataRecord[] => [])
        : Promise.resolve<ODataRecord[]>([]),
      this.getRawRequests().catch((): ODataRecord[] => []),
    ]);

    // Merge + deduplicate by sct_requestid
    const seen = new Set<string>();
    const merged: ODataRecord[] = [];
    for (const r of [...fetchA, ...fetchB1, ...fetchB2, ...fetchC, ...allRaw]) {
      const id = String(r['sct_requestid'] ?? '');
      if (id && !seen.has(id)) {
        seen.add(id);
        merged.push(r);
      }
    }
    // Sort newest first (matches getRawRequests default ordering)
    return merged.sort((a, b) =>
      String(b['createdon'] ?? '').localeCompare(String(a['createdon'] ?? ''))
    );
  }

  // ── Core helpers ────────────────────────────────────────────────────────────

  private async _fetch(url: string, init?: RequestInit): Promise<HttpClientResponse> {
    let response: HttpClientResponse;
    try {
      response = await this._client.fetch(url, AadHttpClient.configurations.v1, {
        method: 'GET',
        ...init,
        headers: { ...OData_HEADERS, ...(init?.headers as Record<string, string> | undefined) },
      });
    } catch (networkErr: unknown) {
      const msg = networkErr instanceof Error ? networkErr.message : String(networkErr);
      throw new Error(`Network error reaching Dataverse: ${msg}`);
    }
    return response;
  }

  private async _get<T>(path: string): Promise<T> {
    const response = await this._fetch(`${this.baseUrl}${path}`);
    await this._assertOk(response);
    const json = await response.json() as { value?: T; [key: string]: unknown };
    return (json.value !== undefined ? json.value : json) as T;
  }

  private async _post(path: string, body: object): Promise<HttpClientResponse> {
    let response: HttpClientResponse;
    try {
      response = await this._client.fetch(`${this.baseUrl}${path}`, AadHttpClient.configurations.v1, {
        method: 'POST',
        headers: OData_HEADERS,
        body: JSON.stringify(body),
      });
    } catch (networkErr: unknown) {
      const msg = networkErr instanceof Error ? networkErr.message : String(networkErr);
      throw new Error(`Network error reaching Dataverse: ${msg}`);
    }
    await this._assertOk(response, [201, 204]);
    return response;
  }

  private async _patch(path: string, body: object): Promise<void> {
    let response: HttpClientResponse;
    try {
      response = await this._client.fetch(`${this.baseUrl}${path}`, AadHttpClient.configurations.v1, {
        method: 'PATCH',
        headers: { ...OData_HEADERS, 'If-Match': '*' },
        body: JSON.stringify(body),
      });
    } catch (networkErr: unknown) {
      const msg = networkErr instanceof Error ? networkErr.message : String(networkErr);
      throw new Error(`Network error reaching Dataverse: ${msg}`);
    }
    await this._assertOk(response, [204]);
  }

  private async _assertOk(
    response: HttpClientResponse,
    expectedCodes: number[] = [200]
  ): Promise<void> {
    if (expectedCodes.indexOf(response.status) !== -1) return;

    if (response.status === 401) {
      throw new Error(
        'Dataverse 401 Unauthorised — ensure the AAD API permission ' +
        '(Dataverse / user_impersonation) has been granted in SharePoint Admin Centre → API Access.'
      );
    }
    if (response.status === 403) {
      throw new Error(
        'Dataverse 403 Forbidden — the signed-in user lacks the required Dataverse security role.'
      );
    }
    if (response.status === 429) {
      const retryAfter = response.headers.get('Retry-After') ?? '60';
      throw new Error(`Dataverse 429 Too Many Requests — retry after ${retryAfter}s.`);
    }

    let detail = '';
    try {
      const errBody = await response.json() as { error?: { message?: string } };
      detail = errBody?.error?.message ?? '';
    } catch { /* ignore */ }

    throw new Error(`Dataverse error ${response.status}${detail ? ': ' + detail : ''}`);
  }

  // ── Metadata discovery ──────────────────────────────────────────────────────

  /**
   * Returns all readable attributes for the given entity logical name.
   * Use this to discover the real column names before building $select strings.
   */
  public async discoverEntityFields(entityLogicalName: string): Promise<IEntityAttribute[]> {
    const path =
      `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes` +
      `?$select=LogicalName,AttributeType,IsCustomAttribute,IsValidForRead,DisplayName` +
      `&$filter=IsValidForRead eq true`;
    return this._get<IEntityAttribute[]>(path);
  }

  // ── System users ─────────────────────────────────────────────────────────────

  /**
   * Look up the Dataverse systemuser record by systemuserid primary key GUID.
   * Resolves internalemailaddress, fullname, and azureactivedirectoryobjectid.
   */
  public async getSystemUserById(systemUserId: string): Promise<ISystemUser | null> {
    const cleanId = (systemUserId || '').replace(/[{}]/g, '').trim();
    if (!cleanId || !GUID_RE.test(cleanId)) return null;
    try {
      const user = await this._get<ISystemUser>(
        `systemusers(${cleanId})?$select=systemuserid,fullname,internalemailaddress,azureactivedirectoryobjectid,_parentsystemuserid_value`
      );
      return user || null;
    } catch {
      try {
        const users = await this._get<ISystemUser[]>(
          `systemusers?$filter=systemuserid eq ${cleanId}&$select=systemuserid,fullname,internalemailaddress,azureactivedirectoryobjectid,_parentsystemuserid_value&$top=1`
        );
        return Array.isArray(users) && users.length > 0 ? users[0] : null;
      } catch {
        return null;
      }
    }
  }

  /**
   * Look up the Dataverse systemuser record for the current AAD user.
   * The azureactivedirectoryobjectid field is the AAD Object ID (from SPFx pageContext.aadInfo.userId).
   * Falls back to null if the user is not found or permissions are insufficient.
   */
  public async getCurrentSystemUser(aadObjectId: string): Promise<ISystemUser | null> {
    const cleanAad = (aadObjectId || '').replace(/[{}]/g, '').trim();
    if (!cleanAad) return null;
    try {
      // Dataverse OData v4: Edm.Guid properties must be unquoted
      let users = await this._get<ISystemUser[]>(
        `systemusers?$filter=azureactivedirectoryobjectid eq ${cleanAad}` +
        `&$select=systemuserid,fullname,internalemailaddress,_parentsystemuserid_value&$top=1`
      ).catch(() => []);
      if (!Array.isArray(users) || users.length === 0) {
        // Fallback: try quoted filter if unquoted returned empty
        users = await this._get<ISystemUser[]>(
          `systemusers?$filter=azureactivedirectoryobjectid eq '${cleanAad}'` +
          `&$select=systemuserid,fullname,internalemailaddress,_parentsystemuserid_value&$top=1`
        ).catch(() => []);
      }
      return Array.isArray(users) && users.length > 0 ? users[0] : null;
    } catch { return null; }
  }

  /** Look up systemuser by fullname */
  public async getSystemUserByName(fullName: string): Promise<ISystemUser | null> {
    const clean = (fullName || '').trim();
    if (!clean) return null;
    try {
      const users = await this._get<ISystemUser[]>(
        `systemusers?$filter=fullname eq '${encodeODataString(clean)}'` +
        `&$select=systemuserid,fullname,internalemailaddress,_parentsystemuserid_value&$top=1`
      ).catch(() => []);
      return Array.isArray(users) && users.length > 0 ? users[0] : null;
    } catch { return null; }
  }

  /** Fallback: look up by email when AAD Object ID is unavailable */
  public async getSystemUserByEmail(email: string): Promise<ISystemUser | null> {
    const cleanEmail = (email || '').trim();
    if (!cleanEmail) return null;
    try {
      let users = await this._get<ISystemUser[]>(
        `systemusers?$filter=internalemailaddress eq '${encodeODataString(cleanEmail)}'` +
        `&$select=systemuserid,fullname,internalemailaddress,_parentsystemuserid_value&$top=1`
      ).catch(() => []);
      if (!Array.isArray(users) || users.length === 0) {
        // Domain mismatch fallback: match by email prefix (e.g. Aniket.Desai)
        const prefix = cleanEmail.split('@')[0];
        if (prefix && prefix.length > 2) {
          users = await this._get<ISystemUser[]>(
            `systemusers?$filter=startswith(internalemailaddress,'${encodeODataString(prefix)}')` +
            `&$select=systemuserid,fullname,internalemailaddress,_parentsystemuserid_value&$top=1`
          ).catch(() => []);
        }
      }
      return Array.isArray(users) && users.length > 0 ? users[0] : null;
    } catch { return null; }
  }

  /**
   * Resolves a Dataverse systemuserid string from the systemusers table by manager's email address.
   * Query: GET /api/data/v9.2/systemusers?$select=systemuserid,fullname&$filter=internalemailaddress eq '${managerEmail}'
   */
  public async getSystemUserIdByEmail(managerEmail: string): Promise<string | null> {
    const cleanEmail = (managerEmail || '').trim();
    if (!cleanEmail) return null;
    try {
      let users = await this._get<Array<{ systemuserid: string; fullname: string }>>(
        `systemusers?$select=systemuserid,fullname&$filter=internalemailaddress eq '${encodeODataString(cleanEmail)}'&$top=1`
      );
      if ((!Array.isArray(users) || users.length === 0) && cleanEmail.indexOf('@') !== -1) {
        // Domain mismatch fallback: match by email prefix (e.g. Manav.Barad)
        const prefix = cleanEmail.split('@')[0];
        if (prefix && prefix.length > 2) {
          users = await this._get<Array<{ systemuserid: string; fullname: string }>>(
            `systemusers?$select=systemuserid,fullname&$filter=startswith(internalemailaddress,'${encodeODataString(prefix)}')&$top=1`
          );
        }
      }
      if (Array.isArray(users) && users.length > 0 && users[0].systemuserid) {
        return users[0].systemuserid;
      }
      return null;
    } catch { return null; }
  }


  public async getManagerSystemUserId(requesterSystemUserId: string): Promise<string | null> {
    assertGuid(requesterSystemUserId, 'requester system user id');
    try {
      const users = await this._get<ODataRecord[]>(
        `systemusers?$filter=systemuserid eq ${requesterSystemUserId}` +
        `&$select=_parentsystemuserid_value&$top=1`
      );
      if (!Array.isArray(users) || users.length === 0) return null;
      const val = users[0]['_parentsystemuserid_value'];
      return typeof val === 'string' && GUID_RE.test(val) ? val : null;
    } catch { return null; }
  }

  /**
   * Returns true if any other systemuser record lists this user as their manager
   * (parentsystemuserid = systemUserId). Used to decide whether to show the
   * "Manager" role option — driven by Dataverse, not AAD org chart.
   * Returns false on any error so the app degrades gracefully.
   */
  public async hasDirectReportsInDataverse(systemUserId: string): Promise<boolean> {
    assertGuid(systemUserId, 'system user id');
    try {
      const users = await this._get<ODataRecord[]>(
        `systemusers?$filter=_parentsystemuserid_value eq ${systemUserId}` +
        `&$select=systemuserid&$top=1`
      );
      return Array.isArray(users) && users.length > 0;
    } catch { return false; }
  }

  /**
   * Returns all systemuserid GUIDs for users whose manager is managerSystemUserId
   * (parentsystemuserid = managerSystemUserId in systemusers table).
   */
  public async getDirectReportUserIds(managerSystemUserId: string): Promise<string[]> {
    assertGuid(managerSystemUserId, 'manager system user id');
    try {
      const users = await this._get<ODataRecord[]>(
        `systemusers?$filter=_parentsystemuserid_value eq ${managerSystemUserId}` +
        `&$select=systemuserid&$top=500`
      );
      return Array.isArray(users)
        ? users.map(u => String(u['systemuserid'] ?? '')).filter(id => GUID_RE.test(id))
        : [];
    } catch {
      return [];
    }
  }

  // ── Relationship metadata ────────────────────────────────────────────────────

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
  public async getNavPropName(
    entityLogicalName: string,
    attributeLogicalName: string
  ): Promise<string> {
    // Step 1: relationship metadata — most accurate source
    try {
      const rels = await this._get<Array<{ ReferencingEntityNavigationPropertyName: string }>>(
        `EntityDefinitions(LogicalName='${entityLogicalName}')/ManyToOneRelationships` +
        `?$filter=ReferencingAttribute eq '${attributeLogicalName}'` +
        `&$select=ReferencingEntityNavigationPropertyName`
      );
      const navProp = Array.isArray(rels) && rels.length > 0
        ? rels[0].ReferencingEntityNavigationPropertyName
        : null;
      // Note: Dataverse OData rejects 'sct_Manager' as undeclared on sct_requests.
      // If the relationship metadata reports 'sct_Manager', prefer the confirmed schema name 'cr2a4_Manager'.
      if (navProp && navProp.toLowerCase() === 'sct_manager') {
        return 'cr2a4_Manager';
      }
      if (navProp && navProp !== attributeLogicalName) return navProp;
    } catch { /* fall through */ }

    // Step 2: attribute SchemaName — preserves the correct mixed casing for @odata.bind
    // e.g. logical 'cr2a4_manager' → schema 'cr2a4_Manager'
    try {
      const attrs = await this._get<Array<{ SchemaName: string }>>(
        `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes` +
        `?$filter=LogicalName eq '${attributeLogicalName}'` +
        `&$select=SchemaName`
      );
      const schemaName = Array.isArray(attrs) && attrs.length > 0 ? attrs[0].SchemaName : null;
      if (schemaName) return schemaName;
    } catch { /* fall through */ }

    // Step 3: return the logical name as given; caller should supply a sane default via .catch()
    return attributeLogicalName;
  }

  // ── Sites ───────────────────────────────────────────────────────────────────
  // No $select — let Dataverse return all columns so we never get a 400

  public async getSites(activeOnly = true): Promise<ISiteMaster[]> {
    let qs = `$orderby=createdon desc&$top=500`;
    if (activeOnly) qs += '&$filter=statecode eq 0';
    return this._get<ISiteMaster[]>(`sct_sites_masters?${qs}`);
  }

  /** Raw ODataRecord[] for Sites — field names picked at runtime, avoids ISiteMaster mismatch */
  public async getSitesRaw(activeOnly = true): Promise<ODataRecord[]> {
    let qs = `$orderby=createdon desc&$top=500`;
    if (activeOnly) qs += '&$filter=statecode eq 0';
    return this._get<ODataRecord[]>(`sct_sites_masters?${qs}`);
  }

  public async getSiteById(id: string): Promise<ISiteMaster | null> {
    try {
      return await this._get<ISiteMaster>(`sct_sites_masters(${id})`);
    } catch { return null; }
  }

  // ── Roles ───────────────────────────────────────────────────────────────────

  public async getRoles(activeOnly = true): Promise<IRole[]> {
    let qs = `$orderby=createdon desc&$top=500`;
    if (activeOnly) qs += '&$filter=statecode eq 0';
    return this._get<IRole[]>(`sct_roles?${qs}`);
  }

  /** Raw ODataRecord[] for roles — used by ApprovalRoutingService for field-safe access */
  public async getRawRoles(activeOnly = true): Promise<ODataRecord[]> {
    let qs = `$orderby=createdon desc&$top=500`;
    if (activeOnly) qs += '&$filter=statecode eq 0';
    return this._get<ODataRecord[]>(`sct_roles?${qs}`);
  }

  /**
   * Fetches all active role configuration records with the fields needed by
   * GroupRoleConfig (tier, role name, description/group, spend limit, site name).
   */
  public async getRoleConfigs(): Promise<ODataRecord[]> {
    const select = [
      'sct_roleid',
      'sct_rolename',
      'sct_roledescription',
      'sct_approvallevel',
      'sct_approvalspendlimit',
      '_sct_siteid_value',
      'statecode',
    ].join(',');
    return this._get<ODataRecord[]>(
      `sct_roles?$select=${select}&$filter=statecode eq 0&$orderby=sct_approvallevel asc,createdon asc&$top=500`
    );
  }

  /**
   * PATCHes a single sct_role record — used by GroupRoleConfig to persist inline edits.
   */
  public async updateRoleConfig(
    roleId: string,
    payload: Record<string, unknown>
  ): Promise<void> {
    assertGuid(roleId, 'roleId');
    await this._patch(`sct_roles(${roleId})`, payload);
  }

  // ── Suppliers ───────────────────────────────────────────────────────────────

  public async getSuppliers(activeOnly = true): Promise<ISupplier[]> {
    let qs = `$orderby=createdon desc&$top=500`;
    if (activeOnly) qs += '&$filter=statecode eq 0';
    return this._get<ISupplier[]>(`sct_suppliers?${qs}`);
  }

  // ── Requests ────────────────────────────────────────────────────────────────

  /**
   * Fetch requests with NO $select — returns every column the entity has.
   * This avoids 400 errors from guessed column names.
   * Optional OData $filter string, e.g. "statecode eq 0".
   */
  public async getRequests(filter?: string): Promise<ISCTRequest[]> {
    let qs = `$orderby=createdon desc&$top=250`;
    if (filter) qs += `&$filter=${filter}`;
    return this._get<ISCTRequest[]>(`sct_requests?${qs}`);
  }

  /** Also exposes the raw ODataRecord[] for the Dashboard field inspector */
  public async getRawRequests(filter?: string): Promise<ODataRecord[]> {
    let qs = `$orderby=createdon desc&$top=250`;
    if (filter) qs += `&$filter=${filter}`;
    return this._get<ODataRecord[]>(`sct_requests?${qs}`);
  }

  public async getRequestById(id: string): Promise<ISCTRequest | null> {
    assertGuid(id, 'record id');
    try {
      return await this._get<ISCTRequest>(`sct_requests(${id})`);
    } catch { return null; }
  }

  public async getRawRequestById(id: string): Promise<ODataRecord | null> {
    assertGuid(id, 'record id');
    try {
      return await this._get<ODataRecord>(`sct_requests(${id})`);
    } catch { return null; }
  }

  /** Fetch all requests where the requester lookup matches the given systemuserid GUID.
   *  Tries sct_Requester first (schema name used in @odata.bind); falls back to sct_requestorid. */
  public async getMyRequests(systemUserId: string): Promise<ODataRecord[]> {
    assertGuid(systemUserId, 'system user id');
    // Try primary filter — sct_Requester navigation property → _sct_requester_value lookup column
    try {
      const filter = `_sct_requester_value eq ${systemUserId}`;
      const qs = `$orderby=createdon desc&$top=200&$filter=${encodeURIComponent(filter)}`;
      const results = await this._get<ODataRecord[]>(`sct_requests?${qs}`);
      return results;
    } catch {
      // Fallback: some environments use sct_requestorid as the lookup column name
      const filter = `_sct_requestorid_value eq ${systemUserId}`;
      const qs = `$orderby=createdon desc&$top=200&$filter=${encodeURIComponent(filter)}`;
      return this._get<ODataRecord[]>(`sct_requests?${qs}`);
    }
  }

  public async createRequest(payload: Partial<ISCTRequest> | ISCTRequestCreatePayload | Record<string, unknown>): Promise<string> {
    const cleanPayload = scrubPayload(payload as Record<string, unknown>);
    const response = await this._post('sct_requests', cleanPayload);
    const entityId = response.headers.get('OData-EntityId') ?? '';
    return extractGuidFromEntityId(entityId);
  }

  public async updateRequest(id: string, payload: Partial<ISCTRequest> | Record<string, unknown>): Promise<void> {
    assertGuid(id, 'record id');
    await this._patch(`sct_requests(${id})`, payload);
  }

  /** Permanently delete a single sct_request record from Dataverse (HTTP DELETE → 204) */
  public async deleteRequest(id: string): Promise<void> {
    assertGuid(id, 'record id');
    const response = await this._fetch(`${this.baseUrl}sct_requests(${id})`, { method: 'DELETE' });
    if (response.status !== 204) await this._assertOk(response);
  }

  // ── Annotations (file attachments) ──────────────────────────────────────────

  public async uploadAnnotation(
    regardingObjectId: string,
    filename: string,
    mimeType: string,
    base64Body: string
  ): Promise<string> {
    assertGuid(regardingObjectId, 'request id');
    const body = {
      'objectid_sct_request@odata.bind': `/sct_requests(${regardingObjectId})`,
      filename,
      mimetype: mimeType,
      documentbody: base64Body,
      subject: filename,
    };
    const response = await this._post('annotations', body);
    const entityId = response.headers.get('OData-EntityId') ?? '';
    return extractGuidFromEntityId(entityId);
  }

  /** Fetch all annotations (file attachments) linked to a given sct_request record */
  public async getAnnotations(requestId: string): Promise<ODataRecord[]> {
    assertGuid(requestId, 'request id');
    const filter = encodeURIComponent(`_objectid_value eq ${requestId} and objecttypecode eq 'sct_request'`);
    return this._get<ODataRecord[]>(`annotations?$filter=${filter}&$orderby=createdon desc&$top=50`);
  }

  /** Delete a single annotation (removes the attachment) */
  public async deleteAnnotation(annotationId: string): Promise<void> {
    assertGuid(annotationId, 'annotation id');
    const response = await this._fetch(`${this.baseUrl}annotations(${annotationId})`, { method: 'DELETE' });
    if (response.status !== 204) await this._assertOk(response);
  }
}
