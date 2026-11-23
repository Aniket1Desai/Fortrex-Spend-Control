// models/index.ts — barrel export for all Dataverse entity models

export type { ISiteMaster }           from './ISiteMaster';
export      { SiteMasterStateCode }   from './ISiteMaster';

export type { IRole, ApprovalLevel }  from './IRole';
export      { RoleStateCode }         from './IRole';

export type { ISupplier }             from './ISupplier';
export      { SupplierStatus, SupplierCategory, SupplierStateCode } from './ISupplier';

export type { ISCTRequest, ISCTRequestWrite } from './ISCTRequest';
export      {
  SCTRequestType,
  SCTRequestStatus,
  SCTUrgency,
  SCTPricingBasis,
  SCTRequestStateCode,
} from './ISCTRequest';
