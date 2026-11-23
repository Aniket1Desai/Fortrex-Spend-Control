import { DataverseService } from '../services/DataverseService';
/**
 * Human-readable labels for edit-panel fields used in audit descriptions.
 * Extend this map as more fields become editable.
 */
export declare const AUDIT_FIELD_LABELS: Record<string, string>;
/**
 * Compares initialValues with currentValues and POSTs one sct_auditlog record per
 * changed field, in parallel via Promise.all.
 *
 * @param service       DataverseService instance
 * @param requestId     GUID of the sct_request being edited
 * @param systemUserId  GUID of the current systemuser (null = omit the @odata.bind lookup)
 * @param initialValues Field values as loaded when the edit panel was opened
 * @param currentValues Field values as they were when Save was clicked
 */
export declare function logAuditChanges(service: DataverseService, requestId: string, systemUserId: string | null, initialValues: Record<string, string>, currentValues: Record<string, string>): Promise<void>;
//# sourceMappingURL=auditLogger.d.ts.map