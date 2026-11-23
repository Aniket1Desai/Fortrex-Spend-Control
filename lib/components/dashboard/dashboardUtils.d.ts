import { SCTRequestStatus, SCTRequestType } from '../../models';
export declare const REQUEST_STATUS_LABELS: Partial<Record<SCTRequestStatus, string>>;
export declare const REQUEST_STATUS_COLORS: Partial<Record<SCTRequestStatus, string>>;
export declare const REQUEST_TYPE_LABELS: Partial<Record<SCTRequestType, string>>;
/** All status labels in workflow order — confirmed from Power Apps sct_Request_Status option set. */
export declare const ALL_REQUEST_STATUS_VALUES: string[];
export declare const BUSINESS_UNIT_LABELS: Record<number | string, string>;
export declare const PURCHASE_TYPE_LABELS: Record<number | string, string>;
export declare const URGENCY_LABELS: Record<number | string, string>;
export declare const PRICING_BASIS_LABELS: Record<number | string, string>;
export declare const COST_IS_FOR_LABELS: Record<number | string, string>;
/** Resolves the human-readable text label for any choice or lookup column */
export declare function getChoiceLabel(record: Record<string, unknown> | null | undefined, field: string, map?: Record<number | string, string>): string;
export declare function formatCurrency(amount: number): string;
export declare function formatDate(iso: string): string;
/**
 * Formats a date/timestamp to strict 'dd/MM/yyyy HH:mm tt' format in user's local timezone.
 * Example: 18/11/2026 09:31 AM
 *
 * Automatically handles:
 * - Dataverse UTC ISO-8601 strings (e.g. '2026-11-18T04:01:23Z') converted to local time
 * - Raw ISO strings missing timezone offsets (appends 'Z' so they parse as UTC)
 * - Date objects
 */
export declare function formatDateTime(val: string | Date | undefined | null): string;
/**
 * Canonical status label resolver for an SCT request record.
 * Accurately maps numeric choice integers, formatted values, and string columns
 * to the exact canonical status label: 'Draft' | 'Submitted' | 'In Review' | 'PO Issued' | 'Rejected' | 'Returned' | 'Closed'.
 *
 * Strict Workflow Rules:
 * 1. Explicit terminal rejections, returns, closures, or drafts always take priority.
 * 2. Active Pending Tier: If a record has an active pending tier (Manager, Tier 1, Tier 2, Tier 3),
 *    it is in-flight. It can NEVER be 'PO Issued'.
 *    - Pending Manager: 'Submitted'
 *    - Pending Tier 1, 2, or 3: 'In Review'
 * 3. Terminal 'PO Issued': Only valid when there is NO active pending tier AND status is PO Issued / Approved.
 */
export declare function getRecordStatusLabel(record: Record<string, unknown> | null | undefined): string;
//# sourceMappingURL=dashboardUtils.d.ts.map