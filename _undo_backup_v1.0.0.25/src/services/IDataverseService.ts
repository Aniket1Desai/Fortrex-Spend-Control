import { ISiteMaster, IRole, ISupplier, ISCTRequest } from '../models';
import { ISystemUser } from './DataverseService';

export interface IDataverseService {
  // Health check
  testConnection(): Promise<void>;

  // History logs
  getApprovalLogs(requestId: string): Promise<import('./DataverseService').ODataRecord[]>;
  getAuditLogs(requestId: string): Promise<import('./DataverseService').ODataRecord[]>;
  createAuditLog(payload: Record<string, unknown>): Promise<void>;
  createApprovalLog(payload: Record<string, unknown>): Promise<void>;
  getTierPendingRequestIds(tier: 1 | 2 | 3): Promise<string[]>;
  getTier1PendingCount(): Promise<number>;
  getManagerPendingCount(systemUserId: string): Promise<number>;
  getMyApprovedRequestIds(systemUserId: string): Promise<string[]>;
  getMyTierApprovedRequestIds(systemUserId: string, tierName: 1 | 2 | 3 | 4): Promise<string[]>;
  getManagerMasterRequests(systemUserId: string): Promise<import('./DataverseService').ODataRecord[]>;

  // System users
  getSystemUserById(systemUserId: string): Promise<ISystemUser | null>;
  getCurrentSystemUser(aadObjectId: string): Promise<ISystemUser | null>;
  getSystemUserByEmail(email: string): Promise<ISystemUser | null>;
  getSystemUserIdByEmail(managerEmail: string): Promise<string | null>;
  getManagerSystemUserId(requesterSystemUserId: string): Promise<string | null>;
  getDirectReportUserIds(managerSystemUserId: string): Promise<string[]>;

  // Sites
  getSites(activeOnly?: boolean): Promise<ISiteMaster[]>;
  getSiteById(id: string): Promise<ISiteMaster | null>;

  // Roles
  getRoles(activeOnly?: boolean): Promise<IRole[]>;
  getRoleConfigs(): Promise<import('./DataverseService').ODataRecord[]>;
  updateRoleConfig(roleId: string, payload: Record<string, unknown>): Promise<void>;

  // Suppliers
  getSuppliers(activeOnly?: boolean): Promise<ISupplier[]>;

  // Requests
  getRequests(filter?: string): Promise<ISCTRequest[]>;
  getRequestById(id: string): Promise<ISCTRequest | null>;
  getRawRequestById(id: string): Promise<import('./DataverseService').ODataRecord | null>;
  createRequest(payload: Partial<ISCTRequest>): Promise<string>;   // returns new GUID
  updateRequest(id: string, payload: Partial<ISCTRequest> | Record<string, unknown>): Promise<void>;

  // Attachments (Dataverse annotations)
  uploadAnnotation(
    regardingObjectId: string,
    filename: string,
    mimeType: string,
    base64Body: string
  ): Promise<string>; // returns new annotation GUID
}
