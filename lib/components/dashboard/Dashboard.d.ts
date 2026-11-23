import * as React from 'react';
import { DataverseService, ODataRecord } from '../../services/DataverseService';
import { IResolvedApprovers } from '../../services/ApprovalRoutingService';
interface IDashboardProps {
    dataverseService: DataverseService;
    context: import('@microsoft/sp-webpart-base').WebPartContext;
    /** Active role from the AppShell Role Switcher — controls which records are shown */
    activeRole?: string;
    /** When true, allows self-approval for solo end-to-end testing */
    soloTestMode?: boolean;
    /** Callback to toggle solo test mode from within dashboard */
    onToggleSoloTestMode?: () => void;
    /** Deep-link request ID or reference to auto-open */
    deepLinkRequestId?: string | null;
}
interface IDashboardState {
    records: ODataRecord[];
    loading: boolean;
    saving: boolean;
    error: string | null;
    saveError: string | null;
    saveSuccess: string | null;
    viewRecord: ODataRecord | null;
    editRecord: ODataRecord | null;
    editValues: Record<string, string>;
    editInitialValues: Record<string, string>;
    poDate: Date | undefined;
    resolvedApprovers: IResolvedApprovers | null;
    loadingApprovers: boolean;
    deleteRecord: ODataRecord | null;
    historyRecord: ODataRecord | null;
    showFieldInspector: boolean;
    detectedFields: string[];
    allFields: string[];
    searchText: string;
    statusFilter: string;
    currentPage: number;
    pageSize: number;
    roleFilteredIds: Set<string> | null;
    roleFilterLoading: boolean;
    requesterTab: 'pending' | 'completed' | 'myRequests';
    managerApprovedIds: Set<string>;
    masterRequestData: ODataRecord[];
    approverApprovedIds: Set<string>;
    directReportIds: Set<string>;
    takeActionRecord: ODataRecord | null;
    takeActionComments: string;
    takeActionSaving: boolean;
    takeActionError: string | null;
    sortField: string | null;
    sortDesc: boolean;
    deepLinkProcessed: boolean;
    sodHiddenCount: number;
}
export default class Dashboard extends React.Component<IDashboardProps, IDashboardState> {
    private _service;
    private _primaryKey;
    private _router;
    private _currentSystemUserId;
    constructor(props: IDashboardProps);
    componentDidMount(): void;
    componentDidUpdate(prevProps: IDashboardProps, prevState: IDashboardState): void;
    /**
     * Builds `roleFilteredIds` based on the selected role.
     *
     * - 'requester'      → Set of IDs where _sct_requester_value === currentSystemUserId
     * - 'approver_tier1' → Set of In Review IDs pending at Tier 1 (async cross-reference)
     * - 'approver_tier2' → Set of In Review IDs pending at Tier 2 (async cross-reference)
     * - 'admin' / other  → null (no restriction — show all records)
     */
    private _applyRoleFilter;
    private _resolveCurrentUser;
    private _loadManagerApprovedIds;
    /**
     * Three-source data load for the Manager role.
     * Fetches authored requests, pending queue, and approval-log history,
     * merges and deduplicates into masterRequestData.
     * Also refreshes managerApprovedIds in the same call so both are in sync.
     */
    private _loadManagerData;
    /**
     * Loads both pending queue IDs and already-approved IDs for a tier approver role.
     * Tier approver 'My Requests' tab shows strictly what they approved at their tier.
     */
    private _loadApproverData;
    private _load;
    private _processDeepLink;
    private _getRecordId;
    /**
     * Pure role × tab filter — all 4 inputs explicit so every call site is self-documenting.
     *
     * Lifecycle contract (must never break):
     *   Requester submits → assigned to Manager via _cr2a4_manager_value
     *   → Manager sees it in Review Queue and acts
     *   → acted requests move to Manager's My Requests (processed)
     *   → Requester tracks current status in their My Requests tab
     *
     * managerApprovedIds / approverApprovedIds are read from this.state — they are
     * loaded async and are not caller-supplied data slices.
     *
     * Class-component equivalent of:
     *   function applyFilters(baseData, selectedRole, activeTab, currentUserId) { ... }
     * useEffect([selectedRole, activeTab]) equivalent: componentDidUpdate watches both.
     */
    /**
     * Strict Role Isolation Filter
     *
     * 1. REQUESTER ROLE: Show ONLY requests created by this user (reqId === safeUserId).
     * 2. MANAGER ROLE: Show ONLY requests managed by this user (managerId === safeUserId).
     *    STRICT ISOLATION: Hide the user's personal requests (if reqId === safeUserId return false).
     * 3. TIER APPROVER ROLES: Show requests pending at the tier or approved in logs.
     *    STRICT ISOLATION: Hide the user's personal requests (if reqId === safeUserId return false).
     */
    applyFilters(masterRequestData: ODataRecord[], selectedRole: string, selectedTab: string, currentUserId: string | null): ODataRecord[];
    private _applyFilters;
    private _filteredRecords;
    private _onSortColumn;
    private _exportToCsv;
    private _pagedRecords;
    private _counts;
    private _buildColumns;
    private _openEdit;
    /** Async — loads resolved approvers from sct_role for the record's site (Phase 8) */
    private _loadApprovers;
    /** Stage action — PATCH only the status (and optional extra fields) with audit and queue sync */
    private _stageAction;
    private _resubmitReturned;
    private _saveEdit;
    private _confirmDelete;
    private _renderViewPanel;
    private _renderEditPanel;
    private _renderDeleteDialog;
    private _renderFieldInspector;
    private _renderPagination;
    private _renderHistoryPanel;
    /** Display label for the current approver tier */
    private _tierName;
    private _renderTakeActionDialog;
    render(): React.ReactElement;
}
export {};
//# sourceMappingURL=Dashboard.d.ts.map