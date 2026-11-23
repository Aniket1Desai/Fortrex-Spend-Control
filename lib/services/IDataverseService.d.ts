import { ISiteMaster, IRole, ISupplier, ISCTRequest } from '../models';
import { ISystemUser } from './DataverseService';
export interface IDataverseService {
    testConnection(): Promise<void>;
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
    getSystemUserById(systemUserId: string): Promise<ISystemUser | null>;
    getCurrentSystemUser(aadObjectId: string): Promise<ISystemUser | null>;
    getSystemUserByEmail(email: string): Promise<ISystemUser | null>;
    getSystemUserIdByEmail(managerEmail: string): Promise<string | null>;
    getManagerSystemUserId(requesterSystemUserId: string): Promise<string | null>;
    getDirectReportUserIds(managerSystemUserId: string): Promise<string[]>;
    getSites(activeOnly?: boolean): Promise<ISiteMaster[]>;
    getSiteById(id: string): Promise<ISiteMaster | null>;
    getRoles(activeOnly?: boolean): Promise<IRole[]>;
    getRoleConfigs(): Promise<import('./DataverseService').ODataRecord[]>;
    updateRoleConfig(roleId: string, payload: Record<string, unknown>): Promise<void>;
    getSuppliers(activeOnly?: boolean): Promise<ISupplier[]>;
    getRequests(filter?: string): Promise<ISCTRequest[]>;
    getRequestById(id: string): Promise<ISCTRequest | null>;
    getRawRequestById(id: string): Promise<import('./DataverseService').ODataRecord | null>;
    createRequest(payload: Partial<ISCTRequest>): Promise<string>;
    updateRequest(id: string, payload: Partial<ISCTRequest> | Record<string, unknown>): Promise<void>;
    uploadAnnotation(regardingObjectId: string, filename: string, mimeType: string, base64Body: string): Promise<string>;
}
//# sourceMappingURL=IDataverseService.d.ts.map