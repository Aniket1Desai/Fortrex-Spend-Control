import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from '@fluentui/react/lib/DetailsList';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { TextField } from '@fluentui/react/lib/TextField';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react/lib/Calendar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { sendManagerApprovalEmail, sendTierNotificationEmail, getManager } from '../../services/GraphMailService';

import { DataverseService, ODataRecord, strVal } from '../../services/DataverseService';
import { handleApprovalAction } from '../../services/approvalHandler';
import RequestLogsPanel from '../history/RequestLogsPanel';
import { logAuditChanges } from '../../utils/auditLogger';
import { ApprovalRoutingService, IResolvedApprovers } from '../../services/ApprovalRoutingService';
import { LoadingSpinner, EmptyState, ErrorBanner, TakeActionModal } from '../common';
import styles from '../shell/AppShell.module.scss';
import {
  formatCurrency, formatDate, ALL_REQUEST_STATUS_VALUES, getRecordStatusLabel,
  getChoiceLabel, BUSINESS_UNIT_LABELS, PURCHASE_TYPE_LABELS, URGENCY_LABELS,
} from './dashboardUtils';

// ── Types ─────────────────────────────────────────────────────────────────────

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
  // Panels / dialogs
  viewRecord: ODataRecord | null;
  editRecord: ODataRecord | null;
  editValues: Record<string, string>;
  editInitialValues: Record<string, string>;
  poDate: Date | undefined;           // DatePicker value for PO Issue Date in Sourcing panel
  resolvedApprovers: IResolvedApprovers | null;
  loadingApprovers: boolean;
  deleteRecord: ODataRecord | null;
  historyRecord: ODataRecord | null;
  showFieldInspector: boolean;
  // Column detection
  detectedFields: string[];
  allFields: string[];
  // Search / filter / pagination
  searchText: string;
  statusFilter: string;
  currentPage: number;
  pageSize: number;
  // Role-based filtering
  roleFilteredIds: Set<string> | null;  // null = no role restriction (admin/default)
  roleFilterLoading: boolean;
  // Requester / manager view tabs
  requesterTab: 'pending' | 'completed' | 'myRequests';
  // Manager: IDs already approved by this manager (to exclude from Pending, include in My Requests)
  managerApprovedIds: Set<string>;
  // Manager: merged 3-source master data (Fetch A authored + Fetch B pending queue + Fetch C approval history)
  masterRequestData: ODataRecord[];
  // Tier approver: IDs this user approved at their specific tier (for My Requests tab)
  approverApprovedIds: Set<string>;
  // Direct report systemuser IDs (from systemusers parentsystemuserid)
  directReportIds: Set<string>;
  // Take Action dialog (Phase 4 — approver actions)
  takeActionRecord: ODataRecord | null;
  takeActionComments: string;
  takeActionSaving: boolean;
  takeActionError: string | null;
  // Sorting & CSV & deep linking
  sortField: string | null;
  sortDesc: boolean;
  deepLinkProcessed: boolean;
  sodHiddenCount: number;
}

// ── Constants ─────────────────────────────────────────────────────────────────

const PAGE_SIZE_OPTIONS = [10, 15, 25, 50];

const ALWAYS_SKIP = new Set([
  '@odata.etag', '@odata.context', 'versionnumber',
  'timezoneruleversionnumber', 'utcconversiontimezonecode',
  'importsequencenumber', 'overriddencreatedon',
]);

// Fields to show as table columns — confirmed from actual Dataverse sct_request record.
// Boolean fields excluded per user request (they are stored but not shown in the table).
// Lookup fields use the _xxx_value OData pattern and show display text via FormattedValue annotation.
const PRIORITY_FIELDS = [
  // Reference & identity
  'sct_sct_reference',
  // Business status
  'sct_request_status',
  // Site lookup (_sct_site_value@associatednavigationproperty = "sct_Site")
  '_sct_site_value',
  // Requester lookup
  '_sct_requester_value',
  // Purchase classification
  'sct_purchase_type',
  'sct_urgency',
  'sct_business_unit',
  'sct_category',
  // Financials
  'sct_estimated_total_value',
  // Dates
  'sct_date_needed_by',
  'sct_submitted_date',
  'sct_term_start_date',
  'sct_term_end_date',
  // Supplier
  '_sct_suggested_supplier_value',
  // Description & function
  'sct_description',
  'sct_function',
  // Logistics
  'sct_cost_is_for',
  'sct_delivery_address',
  'sct_cost_center',
  'sct_customer_name',
  'sct_equipment_asset_tag',
  'sct_urgency_reason',
  // PO tracking
  'sct_po_number',
  'sct_po_issue_date',
  // Contact
  'sct_receiving_contact_email',
  '_sct_receiving_contact_value',
];

// ── Dashboard table columns — the 9 specific fields shown in the list view ────
// Ordered by business relevance; shown regardless of detection order.
const DASHBOARD_VISIBLE_COLUMNS: Array<{ field: string; minW: number; maxW: number }> = [
  { field: 'sct_sct_reference', minW: 110, maxW: 140 },
  { field: 'sct_request_status', minW: 130, maxW: 160 },
  { field: '_sct_requester_value', minW: 120, maxW: 160 },
  { field: 'sct_business_unit', minW: 90, maxW: 120 },
  { field: '_sct_site_value', minW: 120, maxW: 180 },
  { field: 'sct_purchase_type', minW: 100, maxW: 130 },
  { field: 'sct_estimated_total_value', minW: 110, maxW: 140 },
  { field: 'sct_urgency', minW: 90, maxW: 110 },
  { field: 'sct_submitted_date', minW: 110, maxW: 140 },
];

// ── Edit panel — fields editable by a requester ───────────────────────────────
// Read-only fields (reference, status, dates, choice ints, lookups, procurement)
// are shown in the View panel only — NOT as inputs in the Edit panel.
const EDIT_TEXT_FIELDS = new Set([
  'sct_description', 'sct_function', 'sct_urgency_reason',
  'sct_delivery_address', 'sct_equipment_asset_tag',
  'sct_customer_name', 'sct_receiving_contact_email',
  'sct_emergency_declared_by_email',
]);
const EDIT_NUMBER_FIELDS = new Set(['sct_estimated_total_value']);
const EDIT_DATE_FIELDS = new Set(['sct_date_needed_by', 'sct_term_start_date', 'sct_term_end_date']);
// Boolean fields are already defined in BOOLEAN_FIELDS — all are editable via Toggle.
// Fields NOT in any of the above sets are read-only: shown in View panel, not Edit panel.

// ── Helpers ───────────────────────────────────────────────────────────────────

/** True if the field only contains a bare GUID with no display-text annotation.
 *  Dataverse lookup fields (e.g. _sct_siteid_value) store the FK GUID as their raw value
 *  BUT also carry a FormattedValue annotation with the human-readable display text.
 *  If that annotation exists, strVal() will return the text — so we must NOT skip the field. */
function isGuidField(records: ODataRecord[], field: string): boolean {
  for (const r of records) {
    // If a formatted-value annotation exists, strVal() returns display text, not the GUID.
    const fmtKey = `${field}@OData.Community.Display.V1.FormattedValue`;
    if (r[fmtKey] != null && String(r[fmtKey]).trim() !== '') return false;
    const v = r[field];
    if (v == null) continue;
    return typeof v === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v);
  }
  return false;
}

/** All usable fields (not @annotations, not system noise) */
function allUsableFields(records: ODataRecord[]): string[] {
  if (!records.length) return [];
  return Object.keys(records[0]).filter(k => {
    if (ALWAYS_SKIP.has(k)) return false;
    if (k.startsWith('@')) return false;
    if (k.endsWith('@OData.Community.Display.V1.FormattedValue')) return false;
    return true;
  });
}

// Boolean fields confirmed from Dataverse record — excluded from table columns per spec.
const BOOLEAN_FIELDS = new Set([
  'sct_is_new_supplier', 'sct_cost_recharged_to_customer', 'sct_customer_agreement_attached',
  'sct_tech_or_data_involved', 'sct_supplier_on_premises', 'sct_possible_capital_purchase',
  'sct_is_confidential', 'sct_is_retrospective', 'sct_existing_agreement_or_po',
  'sct_declaration_no_prior_commitment', 'sct_declaration_no_conflict_of_interest',
]);

/** Fields to show as table columns — sct_ fields first, skip bare GUIDs and boolean fields */
function detectDisplayFields(records: ODataRecord[]): string[] {
  const usable = allUsableFields(records);
  const nonGuid = usable.filter(k => !isGuidField(records, k));

  // Separate custom (sct_) from system fields; exclude boolean columns
  const customFields = nonGuid.filter(k => k.startsWith('sct_') && !BOOLEAN_FIELDS.has(k));
  const systemDateFields = ['createdon', 'modifiedon'].filter(k => nonGuid.indexOf(k) !== -1);

  // Within custom fields: priority ones first, then the rest alpha
  const sorted: string[] = [];
  PRIORITY_FIELDS.forEach(p => { if (customFields.indexOf(p) !== -1) sorted.push(p); });
  customFields.forEach(k => { if (sorted.indexOf(k) === -1) sorted.push(k); });
  // Append key system dates at the end
  systemDateFields.forEach(k => { if (sorted.indexOf(k) === -1) sorted.push(k); });

  return sorted;
}

/** Safely convert any truthy string/value to a strict JavaScript boolean for Edm.Boolean fields.
 *  Dataverse rejects "Yes", "No", "true", "false" strings — must send actual true/false. */
function toBoolean(value: unknown): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    const lower = value.trim().toLowerCase();
    return lower === 'true' || lower === 'yes' || lower === '1';
  }
  return Boolean(value);
}

/**
 * Convert any date string to strict Edm.Date format (YYYY-MM-DD) required by Dataverse.
 * Handles: already-correct YYYY-MM-DD, localized M/D/YYYY, and JS Date strings.
 * Returns null when the input cannot be parsed — caller should skip the field.
 * Note: avoids padStart — not available in this project's TS lib target.
 */
function formatToEdmDate(dateInput: unknown): string | null {
  if (!dateInput) return null;
  const str = String(dateInput).trim();
  if (!str) return null;

  // Already in strict Edm.Date format
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;

  // Try JS Date parsing (handles ISO strings, long-form locale strings, etc.)
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const mo = d.getMonth() + 1;
    const da = d.getDate();
    const mm = mo < 10 ? ('0' + mo) : String(mo);
    const dd = da < 10 ? ('0' + da) : String(da);
    return y + '-' + mm + '-' + dd;
  }

  // Fallback: split on / or - for M/D/YYYY or M-D-YYYY patterns
  const parts = str.split(/[\/\-]/);
  if (parts.length === 3) {
    const rawMonth = parts[0].trim();
    const rawDay = parts[1].trim();
    let rawYear = parts[2].trim();
    if (rawYear.length === 2) rawYear = '20' + rawYear;
    const mm = rawMonth.length < 2 ? ('0' + rawMonth) : rawMonth;
    const dd = rawDay.length < 2 ? ('0' + rawDay) : rawDay;
    if (rawYear.length === 4) return rawYear + '-' + mm + '-' + dd;
  }

  return null;
}


/** Human-readable column header — field names confirmed from actual Dataverse record */
function fieldLabel(logicalName: string): string {
  const KNOWN_LABELS: Record<string, string> = {
    // Reference
    'sct_sct_reference': 'Reference',
    // Status
    'sct_request_status': 'Status',
    // Site lookup: _sct_site_value@associatednavigationproperty = "sct_Site"
    '_sct_site_value': 'Site',
    // Requester / contacts
    '_sct_requester_value': 'Requester',
    '_sct_receiving_contact_value': 'Receiving Contact',
    '_sct_emergency_declared_by_value': 'Emergency Declared By',
    // Supplier: _sct_suggested_supplier_value@associatednavigationproperty = "sct_Suggested_Supplier"
    '_sct_suggested_supplier_value': 'Suggested Supplier',
    // Purchase details
    'sct_purchase_type': 'Purchase Type',
    'sct_urgency': 'Urgency',
    'sct_business_unit': 'Business Unit',
    'sct_category': 'Category',
    'sct_pricing_basis': 'Pricing Basis',
    'sct_cost_is_for': 'Cost Is For',
    // Financials
    'sct_estimated_total_value': 'Estimated Value',
    // Dates
    'sct_date_needed_by': 'Date Needed By',
    'sct_submitted_date': 'Submitted Date',
    'sct_term_start_date': 'Term Start',
    'sct_term_end_date': 'Term End',
    'sct_po_issue_date': 'PO Issue Date',
    // Text
    'sct_description': 'Description',
    'sct_function': 'Function',
    'sct_delivery_address': 'Delivery Address',
    'sct_cost_center': 'Cost Centre',
    'sct_customer_name': 'Customer Name',
    'sct_equipment_asset_tag': 'Asset Tag',
    'sct_urgency_reason': 'Urgency Reason',
    'sct_po_number': 'PO Number',
    'sct_receiving_contact_email': 'Receiving Contact Email',
    'sct_emergency_declared_by_email': 'Emergency Email',
    'sct_sct_notes': 'Notes',
    // System lookups
    '_ownerid_value': 'Owner',
    '_createdby_value': 'Created By',
    '_modifiedby_value': 'Modified By',
  };
  if (KNOWN_LABELS[logicalName]) return KNOWN_LABELS[logicalName];

  return logicalName
    .replace(/^_/, '')
    .replace(/_value$/, '')
    .replace(/^sct_/, '')
    .replace(/_/g, ' ')
    .replace(/\b\w/g, c => c.toUpperCase());
}

/**
 * Detect live status values from records using canonical status label resolution.
 */
function detectStatusValues(records: ODataRecord[]): string[] {
  const vals = new Set<string>();
  records.forEach(r => {
    const label = getRecordStatusLabel(r);
    if (label && label.trim() !== '') vals.add(label.trim());
  });
  return Array.from(vals).sort();
}

/** Find the primary key GUID field for a set of records */
function findPrimaryKey(records: ODataRecord[]): string | null {
  if (!records.length) return null;
  const guidFields = Object.keys(records[0]).filter(k =>
    !k.startsWith('@') &&
    !k.endsWith('@OData.Community.Display.V1.FormattedValue') &&
    isGuidField(records, k) &&
    !k.startsWith('_')
  );
  // prefer sct_requestid
  return guidFields.find(k => k.endsWith('requestid')) ?? guidFields[0] ?? null;
}

// ── Status pill ───────────────────────────────────────────────────────────────

const STATUS_COLOURS: Record<string, { bg: string; fg: string }> = {
  // Core workflow states — mint-cyan palette
  draft: { bg: '#F3F2F1', fg: '#605E5C' },
  submitted: { bg: '#E0F5F2', fg: '#00897B' },
  'in review': { bg: '#E3F2FB', fg: '#0277BD' },
  'under review': { bg: '#E3F2FB', fg: '#0277BD' },
  pending: { bg: '#FFF8E1', fg: '#F57F17' },
  'awaiting': { bg: '#FFF8E1', fg: '#F57F17' },
  approved: { bg: '#E0F5F2', fg: '#00695C' },
  rejected: { bg: '#FDE7E9', fg: '#A4262C' },
  returned: { bg: '#FFF4CE', fg: '#7A4F01' },
  cancelled: { bg: '#F3F2F1', fg: '#A19F9D' },
  completed: { bg: '#E0F5F2', fg: '#00695C' },
  closed: { bg: '#E0F5F2', fg: '#00695C' },
  // PO-specific
  'po draft': { bg: '#FFF8E1', fg: '#F57F17' },
  'po issued': { bg: '#E0F5F2', fg: '#00695C' },
  'po raised': { bg: '#E0F5F2', fg: '#00695C' },
  'po pending': { bg: '#FFF8E1', fg: '#F57F17' },
  // Triage / sourcing
  triage: { bg: '#F3E5F5', fg: '#6A1B9A' },
  sourcing: { bg: '#F3E5F5', fg: '#6A1B9A' },
  'fp&a': { bg: '#F3E5F5', fg: '#6A1B9A' },
  'cfo': { bg: '#FDE7E9', fg: '#A4262C' },
  receiving: { bg: '#E0F5F2', fg: '#00695C' },
  // System state
  active: { bg: '#E0F5F2', fg: '#00695C' },
  inactive: { bg: '#F3F2F1', fg: '#605E5C' },
  // Choice / boolean
  'fixed price': { bg: '#E0F5F2', fg: '#00897B' },
  'time and material': { bg: '#FFF8E1', fg: '#F57F17' },
  yes: { bg: '#E0F5F2', fg: '#00695C' },
  no: { bg: '#F3F2F1', fg: '#605E5C' },
};

const INT_TO_STATUS: Record<string, string> = {
  '854000001': 'Draft',
  '854000002': 'Submitted',
  '854000003': 'In Review',
  '854000004': 'PO Issued',
  '854000005': 'Rejected',
  '854000006': 'Returned',
  '854000007': 'Closed',
  '854000000': 'Draft',
};

const StatusPill: React.FC<{ value: string }> = ({ value }) => {
  const displayVal = INT_TO_STATUS[value] ?? value;
  const key = Object.keys(STATUS_COLOURS).find(k => displayVal.toLowerCase().indexOf(k) !== -1);
  const c = key ? STATUS_COLOURS[key] : { bg: '#F3F2F1', fg: '#323130' };
  return (
    <span style={{
      background: c.bg, color: c.fg, fontSize: 12, fontWeight: 600,
      padding: '3px 10px', borderRadius: 12, whiteSpace: 'nowrap',
    }}>
      {displayVal || '—'}
    </span>
  );
};

// ── Main component ────────────────────────────────────────────────────────────

// ── Confirmed status integers (Power Apps > sct_Request_Status option set: 854000001 to 854000007) ────
const DSTATUS = {
  DRAFT: 854000001,
  SUBMITTED: 854000002,
  IN_REVIEW: 854000003,
  PO_ISSUED: 854000004,
  REJECTED: 854000005,
  RETURNED: 854000006,
  CLOSED: 854000007,
};

// CFO/FP&A gate threshold
const D_CFO_THRESHOLD = 10000;

// Date helper — no padStart (not in this project's TS lib target)
function _dateToEdm(d: Date): string {
  const y = d.getFullYear();
  const mo = d.getMonth() + 1;
  const da = d.getDate();
  return y + '-' + (mo < 10 ? '0' + mo : String(mo)) + '-' + (da < 10 ? '0' + da : String(da));
}

export default class Dashboard extends React.Component<IDashboardProps, IDashboardState> {
  private _service: DataverseService | null = null;
  private _primaryKey: string | null = null;
  private _router: ApprovalRoutingService | null = null;
  private _currentSystemUserId: string | null = null;

  constructor(props: IDashboardProps) {
    super(props);
    this.state = {
      records: [], loading: false, saving: false,
      error: null, saveError: null, saveSuccess: null,
      viewRecord: null, editRecord: null, editValues: {}, editInitialValues: {},
      poDate: undefined, resolvedApprovers: null, loadingApprovers: false,
      deleteRecord: null, historyRecord: null,
      showFieldInspector: false,
      detectedFields: [], allFields: [],
      searchText: '', statusFilter: 'all',
      currentPage: 1, pageSize: 10,
      roleFilteredIds: null, roleFilterLoading: false,
      requesterTab: 'pending',
      managerApprovedIds: new Set<string>(),
      masterRequestData: [],
      approverApprovedIds: new Set<string>(),
      directReportIds: new Set<string>(),
      takeActionRecord: null, takeActionComments: '', takeActionSaving: false, takeActionError: null,
      sortField: 'sct_submitted_date',
      sortDesc: true,
      deepLinkProcessed: false,
      sodHiddenCount: 0,
    };
  }

  public componentDidMount(): void {
    this._service = this.props.dataverseService;
    this._router = new ApprovalRoutingService(this.props.dataverseService);
    void this._load();
    void this._resolveCurrentUser();
    // Apply initial role filter after data loads — handled inside _load() callback
  }

  public componentDidUpdate(prevProps: IDashboardProps, prevState: IDashboardState): void {
    if (prevProps.activeRole !== this.props.activeRole || prevProps.soloTestMode !== this.props.soloTestMode) {
      // Immediately clear grid data and state to prevent data bleed across roles
      this.setState({
        records: [],
        masterRequestData: [],
        loading: true,
        requesterTab: 'pending',
        currentPage: 1,
        managerApprovedIds: new Set<string>(),
        approverApprovedIds: new Set<string>(),
        directReportIds: new Set<string>(),
        roleFilteredIds: null,
      }, () => {
        const role = this.props.activeRole || 'admin';
        if (role === 'manager') {
          void this._loadManagerData();
        } else if (role === 'approver_tier1' || role === 'approver_tier2' || role === 'approver_tier3') {
          void this._loadApproverData(role);
          void this._load();
        } else {
          void this._applyRoleFilter(role);
          void this._load();
        }
      });
    }

    // Check deep-link parameter once data is populated
    if (this.props.deepLinkRequestId && !this.state.deepLinkProcessed && (this.state.records.length > 0 || this.state.masterRequestData.length > 0)) {
      this._processDeepLink();
    }

    // useEffect([selectedRole, activeTab]) equivalent: reset to page 1 when tab changes
    if (prevState.requesterTab !== this.state.requesterTab) {
      this.setState({ currentPage: 1 });
    }
  }

  /**
   * Builds `roleFilteredIds` based on the selected role.
   *
   * - 'requester'      → Set of IDs where _sct_requester_value === currentSystemUserId
   * - 'approver_tier1' → Set of In Review IDs pending at Tier 1 (async cross-reference)
   * - 'approver_tier2' → Set of In Review IDs pending at Tier 2 (async cross-reference)
   * - 'admin' / other  → null (no restriction — show all records)
   */
  private async _applyRoleFilter(role: string): Promise<void> {
    // Admin sees everything — clear any restriction immediately
    if (!role || role === 'admin') {
      this.setState({ roleFilteredIds: null, roleFilterLoading: false, currentPage: 1 });
      return;
    }

    if (role === 'requester') {
      // Filter purely in-memory using the already-resolved current user ID
      const userId = this._currentSystemUserId;
      if (!userId) {
        // User not yet resolved — wait for _resolveCurrentUser to finish, then re-filter
        this.setState({ roleFilteredIds: new Set<string>(), currentPage: 1 });
        return;
      }
      const ids = new Set<string>();
      this.state.records.forEach(r => {
        const requesterId = String(r['_sct_requester_value'] ?? '');
        if (requesterId === userId) {
          const pk = this._getRecordId(r);
          if (pk) ids.add(pk);
        }
      });
      this.setState({ roleFilteredIds: ids, roleFilterLoading: false, currentPage: 1 });
      return;
    }

    if (role === 'manager') {
      // Manager tab filtering is handled entirely in _filteredRecords() per tab.
      // Set null so all records pass the role-filter step, then _filteredRecords() splits by tab.
      this.setState({ roleFilteredIds: null, roleFilterLoading: false, currentPage: 1 });
      return;
    }

    if (role === 'approver_tier1' || role === 'approver_tier2' || role === 'approver_tier3') {
      if (!this._service) return;
      this.setState({ roleFilterLoading: true });
      try {
        const tier = role === 'approver_tier1' ? 1 : role === 'approver_tier2' ? 2 : 3;
        const pendingIds = await this._service.getTierPendingRequestIds(tier);
        this.setState({
          roleFilteredIds: new Set<string>(pendingIds),
          roleFilterLoading: false,
          currentPage: 1,
        });
      } catch {
        // On error fall back to showing In Review records only (safe degradation)
        const ids = new Set<string>();
        this.state.records.forEach(r => {
          const statusLabel = String(
            r['sct_request_status@OData.Community.Display.V1.FormattedValue'] ?? ''
          );
          if (statusLabel === 'In Review') {
            const pk = this._getRecordId(r);
            if (pk) ids.add(pk);
          }
        });
        this.setState({ roleFilteredIds: ids, roleFilterLoading: false, currentPage: 1 });
      }
      return;
    }

    // Unknown role key — show all
    this.setState({ roleFilteredIds: null, roleFilterLoading: false, currentPage: 1 });
  }

  private async _resolveCurrentUser(): Promise<void> {
    if (!this._service) return;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = this.props.context as any;
      const aadId: string | undefined = ctx?.pageContext?.aadInfo?.userId?.toString();
      const email: string | undefined = ctx?.pageContext?.user?.email;

      let user: import('../../services/DataverseService').ISystemUser | null = null;
      if (aadId) {
        user = await this._service.getCurrentSystemUser(aadId);
      }
      if (!user && email) {
        user = await this._service.getSystemUserByEmail(email);
      }

      if (user) {
        this._currentSystemUserId = user.systemuserid;
        const role = this.props.activeRole;
        if (role === 'manager') {
          void this._loadManagerData();
        } else if (role === 'approver_tier1' || role === 'approver_tier2' || role === 'approver_tier3') {
          void this._loadApproverData(role);
        } else {
          this.setState({ currentPage: 1 });
        }
      }
    } catch {
      // Non-fatal — audit logs will simply omit the @odata.bind reference
    }
  }

  private async _loadManagerApprovedIds(): Promise<void> {
    if (!this._service || !this._currentSystemUserId) return;
    try {
      const ids = await this._service.getMyApprovedRequestIds(this._currentSystemUserId);
      this.setState({ managerApprovedIds: new Set<string>(ids) });
    } catch {
      // Non-fatal — degrade gracefully (Pending may show extra items)
    }
  }

  /**
   * Three-source data load for the Manager role.
   * Fetches authored requests, pending queue, and approval-log history,
   * merges and deduplicates into masterRequestData.
   * Also refreshes managerApprovedIds in the same call so both are in sync.
   */
  private async _loadManagerData(): Promise<void> {
    if (!this._service || !this._currentSystemUserId) return;
    this.setState({ loading: true });
    try {
      const [masterData, approvedIds, reports] = await Promise.all([
        this._service.getManagerMasterRequests(this._currentSystemUserId),
        this._service.getMyApprovedRequestIds(this._currentSystemUserId),
        this._service.getDirectReportUserIds(this._currentSystemUserId).catch(() => []),
      ]);
      // Update primary key detection from merged records if not yet set
      if (!this._primaryKey && masterData.length > 0) {
        this._primaryKey = findPrimaryKey(masterData);
      }
      this.setState({
        masterRequestData: masterData,
        records: masterData,
        managerApprovedIds: new Set<string>(approvedIds),
        directReportIds: new Set<string>(reports.map(id => id.toLowerCase().replace(/[{}]/g, ''))),
        // Ensure column detection covers master data fields
        detectedFields: detectDisplayFields(masterData.length > 0 ? masterData : this.state.records),
        allFields: allUsableFields(masterData.length > 0 ? masterData : this.state.records),
        loading: false,
        currentPage: 1,
      });
    } catch {
      this.setState({ loading: false });
    }
  }

  /**
   * Loads both pending queue IDs and already-approved IDs for a tier approver role.
   * Tier approver 'My Requests' tab shows strictly what they approved at their tier.
   */
  private async _loadApproverData(role: string): Promise<void> {
    if (!this._service) return;
    this.setState({ roleFilterLoading: true });

    let userId = this._currentSystemUserId;
    if (!userId) {
      try {
        const email = this.props.context.pageContext.user.email;
        if (email) {
          const user = await this._service.getSystemUserByEmail(email);
          if (user) {
            userId = user.systemuserid;
            this._currentSystemUserId = userId;
          }
        }
      } catch { /* ignore */ }
    }

    if (!userId) {
      this.setState({ roleFilterLoading: false });
      return;
    }

    // Review Queue is filtered in-memory by cr2a4_pendingtier (no server ID fetch needed).
    // Only load approvedIds for the "My Requests" tab.
    // approver_tier1 → 1, approver_tier2 → 2, approver_tier3 → 3
    const tierIndex = role === 'approver_tier1' ? 1 : role === 'approver_tier2' ? 2 : 3;
    try {
      const approvedIds = await this._service.getMyTierApprovedRequestIds(userId, tierIndex as 1 | 2 | 3);
      this.setState({
        approverApprovedIds: new Set<string>(approvedIds),
        roleFilteredIds: null, // not needed; Review Queue uses cr2a4_pendingtier filter
        roleFilterLoading: false,
        currentPage: 1,
      });
    } catch {
      this.setState({ roleFilterLoading: false });
    }
  }

  // ── Data ──────────────────────────────────────────────────────────────────

  private async _load(): Promise<void> {
    if (!this._service) return;
    this.setState({ loading: true, error: null, saveSuccess: null });
    try {
      const records = await this._service.getRawRequests();
      this._primaryKey = findPrimaryKey(records);
      this.setState({
        records,
        masterRequestData: records,
        detectedFields: detectDisplayFields(records),
        allFields: allUsableFields(records),
        loading: false,
        currentPage: 1,
      }, () => {
        if (this.props.deepLinkRequestId && !this.state.deepLinkProcessed) {
          this._processDeepLink();
        }
        const role = this.props.activeRole;
        if (role === 'manager') {
          if (this._currentSystemUserId) void this._loadManagerData();
        } else if (role === 'approver_tier1' || role === 'approver_tier2' || role === 'approver_tier3') {
          if (this._currentSystemUserId) void this._loadApproverData(role);
          else void this._applyRoleFilter(role); // pending IDs at minimum until user resolves
        } else if (role && role !== 'admin') {
          void this._applyRoleFilter(role);
        }
      });
    } catch (err: unknown) {
      this.setState({ error: err instanceof Error ? err.message : 'Load failed', loading: false });
    }
  }

  private _processDeepLink(): void {
    const target = (this.props.deepLinkRequestId || '').trim().toLowerCase();
    if (!target) return;
    const pool = this.state.masterRequestData.length > 0 ? this.state.masterRequestData : this.state.records;
    const found = pool.find(r => {
      const ref = String(r['sct_sct_reference'] || '').trim().toLowerCase();
      const id = String(this._getRecordId(r) || '').trim().toLowerCase();
      return ref === target || id === target;
    });
    if (found) {
      this.setState({ viewRecord: found, deepLinkProcessed: true });
    }
  }

  private _getRecordId(record: ODataRecord): string | null {
    if (!this._primaryKey) return null;
    const v = record[this._primaryKey];
    return typeof v === 'string' ? v : null;
  }

  // ── Filtering / pagination ────────────────────────────────────────────────

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
  public applyFilters(
    masterRequestData: ODataRecord[],
    selectedRole: string,
    selectedTab: string,
    currentUserId: string | null
  ): ODataRecord[] {
    if (!masterRequestData || masterRequestData.length === 0) {
      return [];
    }

    // Normalize IDs for safe comparison
    const normalizeId = (id: unknown): string => String(id || '').toLowerCase().replace(/[{}]/g, '').trim();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctxAny = this.props.context as any;
    const safeAadId = normalizeId(ctxAny?.pageContext?.aadInfo?.userId?.toString?.());
    const safeUserId = normalizeId(currentUserId) || safeAadId;

    // Normalize role string for robust comparisons (supports 'Requester', 'requester', 'Manager', 'manager', 'Tier 1 Approver', etc.)
    const normRole = (selectedRole || '').trim();
    const normRoleLower = normRole.toLowerCase();

    // Role classification
    const isRequesterRole = normRoleLower === 'requester';
    const isManagerRole = normRoleLower === 'manager';
    const isTierRole = normRoleLower.indexOf('tier') !== -1 || normRoleLower.indexOf('approver') !== -1 || normRole.indexOf('Tier') !== -1;

    // Normalize tab string for robust comparisons (supports 'Pending', 'pending', 'Review Queue', 'Completed', 'completed', 'My Requests', 'myRequests')
    const normTab = (selectedTab || '').trim();
    const normTabLower = normTab.toLowerCase();
    const isPendingTab = normTab === 'Review Queue' || normTab === 'Pending' || normTabLower === 'review queue' || normTabLower === 'pending' || normTabLower === 'reviewqueue' || !normTab;
    const isCompletedOrMyRequests = normTab === 'My Requests' || normTab === 'Completed' || normTabLower === 'my requests' || normTabLower === 'completed' || normTabLower === 'myrequests';

    let hiddenBySodCount = 0;

    const filtered = masterRequestData.filter(req => {
      // Normalize IDs for safe comparison
      const rawReq = req['_sct_requester_value'] ?? req['_createdby_value'] ?? req['sct_requester'] ?? req['_sct_requestorid_value'] ?? req['createdby'] ?? '';
      const reqId = rawReq ? String(rawReq).toLowerCase().replace(/[{}]/g, '') : '';
      const createdById = req['_createdby_value'] ? String(req['_createdby_value']).toLowerCase().replace(/[{}]/g, '') : '';

      const rawMgr = req['_sct_manager_value'] ?? req['_cr2a4_manager_value'] ?? req['_manager_value'] ?? req['sct_manager'] ?? req['cr2a4_manager'] ?? req['manager'] ?? '';
      let managerId = rawMgr ? String(rawMgr).toLowerCase().replace(/[{}]/g, '') : '';
      const isDirectReport = Boolean(this.state.directReportIds && (this.state.directReportIds.has(reqId) || (createdById && this.state.directReportIds.has(createdById))));
      if (!managerId && isDirectReport) {
        managerId = safeUserId;
      }
      const mgrFormatted = (strVal(req, '_cr2a4_manager_value') || strVal(req, '_sct_manager_value') || strVal(req, '_manager_value') || strVal(req, 'cr2a4_manager') || strVal(req, 'sct_manager') || strVal(req, 'manager') || '').toLowerCase();
      const currentUserName = (this.props.context.pageContext.user.displayName || '').toLowerCase().trim();
      const currentUserEmail = (this.props.context.pageContext.user.email || '').toLowerCase().trim();
      const isMgrByName = Boolean(currentUserName) && (
        mgrFormatted.indexOf(currentUserName) !== -1 ||
        (Boolean(currentUserEmail) && mgrFormatted.indexOf(currentUserEmail) !== -1)
      );
      const isManagerOfReq = (managerId === safeUserId) || isMgrByName || isDirectReport;

      // Resolve status
      const rawStatus = typeof req['sct_request_status'] === 'number'
        ? req['sct_request_status']
        : parseInt(String(req['sct_request_status'] ?? req['sct_requeststatus'] ?? ''), 10);
      const statusLabel = strVal(req, 'sct_request_status') || strVal(req, 'sct_requeststatus') || '';
      const statusLower = statusLabel.toLowerCase();
      const isInReview = statusLower === 'in review' ||
        statusLower.indexOf('in review') !== -1 ||
        rawStatus === 854000003 || rawStatus === 854000002 ||
        req['sct_requeststatus'] === 'In Review' ||
        String(req['sct_requeststatus'] ?? '').toLowerCase().indexOf('in review') !== -1;
      const isSubmitted = statusLower === 'submitted' ||
        statusLower.indexOf('submitted') !== -1 ||
        rawStatus === 854000002 || rawStatus === 854000001 ||
        req['sct_requeststatus'] === 'Submitted' ||
        String(req['sct_requeststatus'] ?? '').toLowerCase().indexOf('submitted') !== -1;
      const statusString = isInReview ? 'In Review' : (isSubmitted ? 'Submitted' : (statusLabel || String(rawStatus)));
      (req as Record<string, unknown>)['sct_requeststatus'] = statusString;

      // Resolve pending tier
      let rawPt = String(
        req['cr2a4_pendingtier'] ||
        req['cr2a4_PendingTier'] ||
        req['sct_pendingtier'] ||
        req['pendingtier'] ||
        ''
      ).trim();
      if (!rawPt && statusLower.indexOf('tier') !== -1) {
        const m = statusLower.match(/tier\s*([123])/i);
        if (m) rawPt = `Tier ${m[1]}`;
      }
      (req as Record<string, unknown>)['sct_pendingtier'] = rawPt;
      const ptLower = rawPt.toLowerCase();

      // Record ID for approval log check
      const pk = this._getRecordId(req);
      const isUserInApprovalLog = req['sct_approvallog_contains_user'] === true ||
        (pk !== null && this.state.approverApprovedIds.has(pk));

      // Resolve whether the current user is the requester of this record
      const reqFormatted = (strVal(req, '_sct_requester_value') || strVal(req, 'sct_requester') || '').toLowerCase().trim();
      const reqNameFromOdata = (strVal(req, '_sct_requester_value@OData.Community.Display.V1.FormattedValue') || strVal(req, 'requester') || '').toLowerCase().trim();
      const isOwnRequest = Boolean(
        (safeUserId && (reqId === safeUserId || createdById === safeUserId || reqId === safeAadId || createdById === safeAadId)) ||
        (currentUserName && (reqFormatted.indexOf(currentUserName) !== -1 || reqNameFromOdata.indexOf(currentUserName) !== -1)) ||
        (currentUserEmail && (reqFormatted.indexOf(currentUserEmail) !== -1 || reqId.indexOf(currentUserEmail) !== -1))
      );

      // ── 1. Global Segregation of Duties (SoD) Evaluation ───────────────────
      // When Solo Test Mode is OFF (Strict Mode), an approver cannot self-approve their own request.
      // We flag the record so the UI clearly shows the request at its active tier while locking self-approval.
      const isMyRequestsTab = normTabLower === 'my requests' || normTabLower === 'myrequests';
      const allowSelfApproval = this.props.soloTestMode !== false;
      const isSodLocked = !allowSelfApproval && !isRequesterRole && !isMyRequestsTab && isPendingTab && isOwnRequest;
      (req as Record<string, unknown>)['_is_sod_locked'] = isSodLocked;
      if (isSodLocked) {
        hiddenBySodCount++;
      }

      // ── 2. The Requester View / "My Requests" Tab ───────────────────────────
      // If activeRole === 'Requester' OR user is on the "My Requests" tab:
      // Return true for all requests authored by this user
      if (isRequesterRole || isMyRequestsTab) {
        return isOwnRequest;
      }

      // ── 3. The Manager View ────────────────────────────────────────────────
      // If activeRole === 'Manager':
      if (isManagerRole) {
        const rawMgr = req['_cr2a4_manager_value'] ?? req['_sct_manager_value'] ?? req['cr2a4_manager'] ?? req['sct_manager'] ?? req['_manager_value'] ?? '';
        const managerId = normalizeId(rawMgr);
        const isPendingManagerTier = ptLower === 'manager' || (!rawPt && (rawStatus === 854000002 || rawStatus === 854000001 || statusLower === 'submitted'));

        const isManagerAssigned = (managerId && (managerId === safeUserId || isManagerOfReq))
          || isMgrByName
          || isDirectReport
          || (!managerId && isPendingManagerTier);

        if (isPendingTab) {
          return isManagerAssigned && isPendingManagerTier && !(pk !== null && this.state.managerApprovedIds && this.state.managerApprovedIds.has(pk));
        }
        if (isCompletedOrMyRequests) {
          return (pk !== null && this.state.managerApprovedIds && this.state.managerApprovedIds.has(pk)) || (isManagerAssigned && !isPendingManagerTier);
        }
        return isManagerAssigned && isPendingManagerTier;
      }

      // ── 4. The Tier 1, Tier 2, and Tier 3 Views ────────────────────────────
      if (selectedRole.includes('Tier') || selectedRole.includes('approver') || selectedRole.includes('Approver') || normRoleLower.includes('tier') || normRoleLower.includes('approver')) {
        let expectedTierValue = '';
        if (selectedRole.includes('1') || normRoleLower.includes('1')) expectedTierValue = 'Tier 1';
        else if (selectedRole.includes('2') || normRoleLower.includes('2')) expectedTierValue = 'Tier 2';
        else if (selectedRole.includes('3') || normRoleLower.includes('3')) expectedTierValue = 'Tier 3';

        const tierNum = expectedTierValue.replace(/\D/g, '');
        const ptClean = ptLower.replace(/[^a-z0-9]/g, '');
        const expectedTierClean = expectedTierValue.toLowerCase().replace(/[^a-z0-9]/g, '');
        const matchesTier = rawPt === expectedTierValue || ptClean === expectedTierClean || ptClean === tierNum || ptClean === ('tier' + tierNum);
        const matchesStatus = (req as Record<string, unknown>)['sct_requeststatus'] === 'In Review' || isInReview || rawStatus === 854000003 || rawStatus === 854000002;

        if (isPendingTab) {
          return matchesTier && matchesStatus;
        }

        if (selectedTab === 'My Requests' || selectedTab === 'Completed' || isCompletedOrMyRequests) {
          return req['sct_approvallog_contains_user'] === true || isUserInApprovalLog;
        }
        return matchesTier && matchesStatus;
      }

      // Admin role fallback
      if (normRoleLower === 'admin') {
        return true;
      }

      return false; // Fallback
    });

    if (this.state.sodHiddenCount !== hiddenBySodCount) {
      setTimeout(() => {
        this.setState({ sodHiddenCount: hiddenBySodCount });
      }, 0);
    }

    return filtered;
  }

  private _applyFilters(
    baseData: ODataRecord[],
    selectedRole: string,
    activeTab: string,
    currentUserId: string | null
  ): ODataRecord[] {
    return this.applyFilters(baseData, selectedRole, activeTab, currentUserId);
  }

  private _filteredRecords(): ODataRecord[] {
    const {
      records, masterRequestData, detectedFields, searchText, statusFilter, requesterTab,
    } = this.state;
    const role = this.props.activeRole || 'admin';
    const userId = this._currentSystemUserId;

    // Use masterRequestData if available, otherwise records
    const baseData = (role === 'manager' && masterRequestData.length > 0)
      ? masterRequestData
      : (masterRequestData.length > 0 ? masterRequestData : records);

    // Strict role isolation: applyFilters always returns a fresh slice
    let result = this.applyFilters(baseData, role, requesterTab, userId);

    // ── Status dropdown filter ────────────────────────────────────────────────
    if (statusFilter !== 'all') {
      const normFilter = statusFilter.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
      result = result.filter(r => {
        const itemStatus = getRecordStatusLabel(r);
        const normItem = itemStatus.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
        return normItem === normFilter;
      });
    }

    // ── Full-text search ──────────────────────────────────────────────────────
    if (searchText.trim()) {
      const q = searchText.trim().toLowerCase();
      result = result.filter(r =>
        detectedFields.some(f => strVal(r, f).toLowerCase().indexOf(q) !== -1)
      );
    }

    // ── Sorting ───────────────────────────────────────────────────────────────
    const { sortField, sortDesc } = this.state;
    if (sortField) {
      result.sort((a, b) => {
        let valA: unknown = a[sortField];
        let valB: unknown = b[sortField];

        if (sortField.startsWith('_') && sortField.endsWith('_value')) {
          valA = a[`${sortField}@OData.Community.Display.V1.FormattedValue`] ?? valA;
          valB = b[`${sortField}@OData.Community.Display.V1.FormattedValue`] ?? valB;
        }

        if (sortField === 'sct_estimated_total_value' || sortField.indexOf('amount') !== -1 || sortField.indexOf('value') !== -1) {
          const numA = parseFloat(String(valA ?? 0)) || 0;
          const numB = parseFloat(String(valB ?? 0)) || 0;
          return sortDesc ? numB - numA : numA - numB;
        }

        if (sortField === 'sct_submitted_date' || sortField === 'createdon' || sortField === 'modifiedon' || sortField === 'sct_date_needed_by') {
          const dateA = valA ? new Date(String(valA)).getTime() : 0;
          const dateB = valB ? new Date(String(valB)).getTime() : 0;
          return sortDesc ? dateB - dateA : dateA - dateB;
        }

        const strA = String(valA ?? '').toLowerCase();
        const strB = String(valB ?? '').toLowerCase();
        const cmp = strA.localeCompare(strB);
        return sortDesc ? -cmp : cmp;
      });
    }

    return result;
  }

  private _onSortColumn = (columnKey: string): void => {
    this.setState(prevState => ({
      sortField: columnKey,
      sortDesc: prevState.sortField === columnKey ? !prevState.sortDesc : false,
      currentPage: 1,
    }));
  };

  private _exportToCsv(records: ODataRecord[]): void {
    if (!records || records.length === 0) return;
    const headers = [
      'Reference',
      'Status',
      'Pending Tier',
      'Requester',
      'Business Unit',
      'Site',
      'Purchase Type',
      'Estimated Value',
      'Urgency',
      'Date Needed By',
      'Submitted Date'
    ];

    const escapeCsv = (val: unknown): string => {
      const str = val === null || val === undefined ? '' : String(val);
      if (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1) {
        return `"${str.replace(/"/g, '""')}"`;
      }
      return str;
    };

    const rows = records.map(r => {
      const ref = strVal(r, 'sct_sct_reference') || strVal(r, 'sct_reference') || '';
      const status = getRecordStatusLabel(r);
      const pendingTier = strVal(r, 'cr2a4_pendingtier') || strVal(r, 'sct_pendingtier') || strVal(r, 'cr2a4_PendingTier') || '';
      const requester = strVal(r, '_sct_requester_value@OData.Community.Display.V1.FormattedValue')
        || strVal(r, 'sct_requester@OData.Community.Display.V1.FormattedValue')
        || strVal(r, 'requesterName')
        || strVal(r, 'sct_requester')
        || '';
      const bu = getChoiceLabel(r, 'sct_business_unit', BUSINESS_UNIT_LABELS);
      const site = strVal(r, '_sct_site_value@OData.Community.Display.V1.FormattedValue')
        || strVal(r, 'sct_site@OData.Community.Display.V1.FormattedValue')
        || strVal(r, 'site')
        || '';
      const ptype = getChoiceLabel(r, 'sct_purchase_type', PURCHASE_TYPE_LABELS);
      const rawAmt = r['sct_estimated_total_value'];
      const numAmt = parseFloat(String(rawAmt ?? 0));
      const val = !isNaN(numAmt) && numAmt > 0 ? formatCurrency(numAmt) : (rawAmt !== undefined && rawAmt !== null ? String(rawAmt) : '$0.00');
      const urgency = getChoiceLabel(r, 'sct_urgency', URGENCY_LABELS);
      const neededBy = r['sct_date_needed_by'] ? formatDate(String(r['sct_date_needed_by'])) : '';
      const submitted = r['sct_submitted_date'] ? formatDate(String(r['sct_submitted_date'])) : (r['createdon'] ? formatDate(String(r['createdon'])) : '');

      return [
        escapeCsv(ref),
        escapeCsv(status),
        escapeCsv(pendingTier),
        escapeCsv(requester),
        escapeCsv(bu),
        escapeCsv(site),
        escapeCsv(ptype),
        escapeCsv(val),
        escapeCsv(urgency),
        escapeCsv(neededBy),
        escapeCsv(submitted)
      ].join(',');
    });

    const csvContent = '\uFEFF' + [headers.join(','), ...rows].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `Spend_Control_Tracker_Export_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  private _pagedRecords(filtered: ODataRecord[]): ODataRecord[] {
    const { currentPage, pageSize } = this.state;
    const start = (currentPage - 1) * pageSize;
    return filtered.slice(start, start + pageSize);
  }

  // ── Stat counts ───────────────────────────────────────────────────────────

  private _counts(filtered: ODataRecord[]): { total: number; pending: number; approved: number; spend: number } {
    const role = this.props.activeRole || 'admin';
    const userId = this._currentSystemUserId;
    const baseData = role === 'manager' ? this.state.masterRequestData : this.state.records;

    // For approver/manager roles, compute stats across Pending and Completed views
    if (role === 'manager' || role === 'approver_tier1' || role === 'approver_tier2' || role === 'approver_tier3') {
      const pendingSet = this._applyFilters(baseData, role, 'pending', userId);
      const completedSet = this._applyFilters(baseData, role, 'completed', userId);
      const uniqueIds = new Set<string>();
      pendingSet.forEach(r => { const id = this._getRecordId(r); if (id) uniqueIds.add(id); });
      completedSet.forEach(r => { const id = this._getRecordId(r); if (id) uniqueIds.add(id); });
      let spend = 0;
      const targetSpendSet = this.state.requesterTab === 'completed' ? completedSet : (pendingSet.length > 0 ? pendingSet : completedSet);
      targetSpendSet.forEach(r => {
        const amt = parseFloat(String(r['sct_estimated_total_value'] ?? 0));
        if (!isNaN(amt)) spend += amt;
      });
      return {
        total: uniqueIds.size,
        pending: pendingSet.length,
        approved: completedSet.length,
        spend,
      };
    }

    let pending = 0, approved = 0, spend = 0;
    filtered.forEach(r => {
      const statusLabel = getRecordStatusLabel(r);
      if (statusLabel === 'Submitted' || statusLabel === 'In Review') pending++;
      if (statusLabel === 'PO Issued') approved++;
      const amt = parseFloat(String(r['sct_estimated_total_value'] ?? 0));
      if (!isNaN(amt)) spend += amt;
    });
    return { total: filtered.length, pending, approved, spend };
  }

  // ── Columns ───────────────────────────────────────────────────────────────

  private _buildColumns(_statusValues: string[]): IColumn[] {
    const { allFields } = this.state;

    // Only include columns that actually exist in the loaded data
    const visible = DASHBOARD_VISIBLE_COLUMNS.filter(
      col => allFields.indexOf(col.field) !== -1 ||
        allFields.some(f => f === col.field)
    );

    const dataCols: IColumn[] = visible.map(({ field, minW, maxW }) => ({
      key: field,
      name: fieldLabel(field),
      fieldName: field,
      minWidth: minW,
      maxWidth: maxW,
      isResizable: true,
      isSorted: this.state.sortField === field,
      isSortedDescending: this.state.sortField === field && this.state.sortDesc,
      onColumnClick: (_ev, col) => col && this._onSortColumn(col.key),
      onRender: (item: ODataRecord) => {
        // sct_submitted_date: fall back to createdon for Draft records (not yet submitted)
        const rawVal = strVal(item, field);
        let val = (field === 'sct_submitted_date' && !rawVal)
          ? strVal(item, 'createdon')
          : rawVal;

        if (field === 'sct_business_unit') {
          val = getChoiceLabel(item, 'sct_business_unit', BUSINESS_UNIT_LABELS);
        } else if (field === 'sct_purchase_type') {
          val = getChoiceLabel(item, 'sct_purchase_type', PURCHASE_TYPE_LABELS);
        } else if (field === 'sct_urgency') {
          val = getChoiceLabel(item, 'sct_urgency', URGENCY_LABELS);
        }

        if (!val) return <span style={{ color: '#C8C6C4' }}>—</span>;

        if (field === 'sct_estimated_total_value' || field.indexOf('amount') !== -1) {
          const num = parseFloat(val.replace(/[^0-9.-]/g, ''));
          return !isNaN(num)
            ? <span style={{ fontWeight: 700, color: '#00897B', fontFamily: 'monospace', fontSize: 13 }}>{formatCurrency(num)}</span>
            : <span>{val}</span>;
        }
        if (field === 'sct_request_status' || field === 'sct_requeststatus') {
          const statusText = getRecordStatusLabel(item);
          const pt = strVal(item, 'cr2a4_pendingtier') || strVal(item, 'sct_pendingtier') || strVal(item, 'cr2a4_PendingTier');
          const display = (statusText === 'In Review' && pt) ? `${statusText} (${pt})` : (statusText || val);

          // SLA Ageing calculation for pending items
          let slaBadge: React.ReactElement | null = null;
          if (statusText === 'Submitted' || statusText === 'In Review') {
            const dateStr = strVal(item, 'sct_submitted_date') || strVal(item, 'createdon');
            if (dateStr) {
              const diffMs = Date.now() - new Date(dateStr).getTime();
              const diffHours = diffMs / (1000 * 60 * 60);
              const diffDays = Math.floor(diffHours / 24);
              if (diffHours < 24) {
                slaBadge = (
                  <span style={{ fontSize: 10, padding: '1px 5px', borderRadius: 4, background: '#EFF6FF', color: '#1D4ED8', border: '1px solid #BFDBFE', fontWeight: 600, display: 'inline-block', marginTop: 3 }}>
                    ⚡ &lt;24h
                  </span>
                );
              } else if (diffDays <= 3) {
                slaBadge = (
                  <span style={{ fontSize: 10, padding: '1px 5px', borderRadius: 4, background: '#FEF3C7', color: '#B45309', border: '1px solid #FDE68A', fontWeight: 600, display: 'inline-block', marginTop: 3 }}>
                    ⏳ {diffDays}d
                  </span>
                );
              } else {
                slaBadge = (
                  <span style={{ fontSize: 10, padding: '1px 5px', borderRadius: 4, background: '#FEE2E2', color: '#B91C1C', border: '1px solid #FECACA', fontWeight: 600, display: 'inline-block', marginTop: 3 }}>
                    ⚠️ {diffDays}d overdue
                  </span>
                );
              }
            }
          }

          return (
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 2 }}>
              <StatusPill value={display} />
              {slaBadge}
            </div>
          );
        }
        if (field === 'sct_urgency' || field === 'sct_purchase_type' || field === 'sct_business_unit') {
          return <StatusPill value={val} />;
        }
        if (field === 'sct_submitted_date' || field === 'createdon' || field.indexOf('date') !== -1) {
          return <span style={{ color: '#605E5C', fontSize: 12 }}>{formatDate(val)}</span>;
        }
        return (
          <span style={{ color: '#201F1E', fontSize: 13 }}
            title={val.length > 36 ? val : undefined}>
            {val.length > 36 ? val.substring(0, 34) + '…' : val}
          </span>
        );
      },
    }));

    const currentRoleKey = (this.props.activeRole || '').toLowerCase().trim();
    const isApproverRole = (
      currentRoleKey === 'manager' ||
      currentRoleKey.indexOf('tier') !== -1 ||
      currentRoleKey.indexOf('approver') !== -1
    );

    const actionsCol: IColumn = {
      key: '_actions',
      name: 'Actions',
      minWidth: isApproverRole ? 180 : 110,
      maxWidth: isApproverRole ? 200 : 120,
      onRender: (item: ODataRecord) => {
        const statusRaw = typeof item['sct_request_status'] === 'number'
          ? item['sct_request_status'] as number
          : NaN;
        const statusLabel = strVal(item, 'sct_request_status');
        const isDraft = statusRaw === DSTATUS.DRAFT || statusLabel === 'Draft' || (!statusRaw && isNaN(statusRaw));
        const isReturned = statusRaw === DSTATUS.RETURNED || statusLabel === 'Returned';
        const canEditItem = isDraft || isReturned;
        return (
          <div style={{ display: 'flex', alignItems: 'center', gap: 4, flexWrap: 'nowrap' }}>
            {/* Icon-only utility buttons */}
            <button
              className={`${styles.actionBtn} ${styles.actionBtnView}`}
              title="View details"
              aria-label="View"
              onClick={e => { e.stopPropagation(); this.setState({ viewRecord: item }); }}
            >
              <Icon iconName="View" style={{ fontSize: 14 }} />
            </button>
            {!isApproverRole && (
              <>
                <button
                  className={`${styles.actionBtn} ${styles.actionBtnEdit}`}
                  title={canEditItem ? (isReturned ? 'Edit & resubmit returned request' : 'Edit record') : 'Edit only available for Draft or Returned records'}
                  aria-label="Edit"
                  disabled={!canEditItem}
                  onClick={e => { e.stopPropagation(); if (canEditItem) this._openEdit(item); }}
                  style={!canEditItem ? { opacity: 0.35, cursor: 'not-allowed' } : undefined}
                >
                  <Icon iconName="Edit" style={{ fontSize: 14 }} />
                </button>
                <button
                  className={`${styles.actionBtn} ${styles.actionBtnDelete}`}
                  title={isDraft ? 'Delete record' : 'Delete only available for Draft records'}
                  aria-label="Delete"
                  disabled={!isDraft}
                  onClick={e => { e.stopPropagation(); if (isDraft) this.setState({ deleteRecord: item }); }}
                  style={!isDraft ? { opacity: 0.35, cursor: 'not-allowed' } : undefined}
                >
                  <Icon iconName="Delete" style={{ fontSize: 14 }} />
                </button>
              </>
            )}
            <button
              className={styles.actionBtn}
              title="View request history"
              aria-label="History"
              style={{ color: '#8764B8' }}
              onClick={e => { e.stopPropagation(); this.setState({ historyRecord: item }); }}
            >
              <Icon iconName="History" style={{ fontSize: 14 }} />
            </button>
            {/* Take Action — prominent labelled button ONLY for actionable pending requests, NEVER on completed / PO Issued */}
            {(() => {
              if (!isApproverRole) return null;
              if (this.state.requesterTab === 'completed') return null; // Never on completed tab!

              const statusRawNum = typeof item['sct_request_status'] === 'number'
                ? item['sct_request_status']
                : parseInt(String(item['sct_request_status'] ?? ''), 10);
              const statusStr = getRecordStatusLabel(item);
              const isTerminal = (
                statusRawNum === DSTATUS.PO_ISSUED ||
                statusRawNum === DSTATUS.REJECTED ||
                statusRawNum === DSTATUS.CLOSED ||
                statusRawNum === DSTATUS.DRAFT ||
                statusRawNum === DSTATUS.RETURNED ||
                statusStr === 'PO Issued' ||
                statusStr === 'Rejected' ||
                statusStr === 'Closed' ||
                statusStr === 'Draft' ||
                statusStr === 'Returned'
              );
              if (isTerminal) return null;

              // Check if pending tier matches approver role
              const itemPt = (strVal(item, 'cr2a4_pendingtier') || strVal(item, 'sct_pendingtier') || '').toLowerCase();
              let isEligible = false;
              if (currentRoleKey === 'manager') {
                isEligible = itemPt === 'manager' || (!itemPt && (statusRawNum === DSTATUS.SUBMITTED || statusStr === 'Submitted'));
              } else if (currentRoleKey === 'approver_tier1') {
                isEligible = itemPt === 'tier 1' || itemPt === 'tier1' || itemPt === '1';
              } else if (currentRoleKey === 'approver_tier2') {
                isEligible = itemPt === 'tier 2' || itemPt === 'tier2' || itemPt === '2';
              } else if (currentRoleKey === 'approver_tier3') {
                isEligible = itemPt === 'tier 3' || itemPt === 'tier3' || itemPt === '3';
              } else {
                isEligible = true;
              }

              if (!isEligible) return null;

              if (item['_is_sod_locked'] === true) {
                return (
                  <button
                    title="Segregation of Duties Enforced: You submitted this request. Self-approval is locked in Strict Mode. Enable Solo Test Mode in header to test alone."
                    aria-label="SoD Locked"
                    onClick={e => { e.stopPropagation(); this.setState({ takeActionRecord: item, takeActionComments: '', takeActionError: null }); }}
                    style={{
                      display: 'inline-flex', alignItems: 'center', gap: 5,
                      background: '#FFFBEB', color: '#92400E', border: '1px solid #FCD34D',
                      borderRadius: 20, padding: '5px 12px', fontSize: 11, fontWeight: 600,
                      cursor: 'pointer', whiteSpace: 'nowrap',
                    }}
                  >
                    <Icon iconName="Lock" style={{ fontSize: 11, color: '#D97706' }} />
                    SoD Locked
                  </button>
                );
              }

              return (
                <button
                  title="Review and take action on this request"
                  aria-label="Take Action"
                  onClick={e => { e.stopPropagation(); this.setState({ takeActionRecord: item, takeActionComments: '', takeActionError: null }); }}
                  style={{
                    display: 'inline-flex', alignItems: 'center', gap: 5,
                    background: 'linear-gradient(135deg, #0F766E 0%, #115E59 100%)',
                    color: '#fff', border: 'none', borderRadius: 20,
                    padding: '5px 13px', fontSize: 12, fontWeight: 600,
                    cursor: 'pointer', letterSpacing: '0.2px',
                    boxShadow: '0 2px 6px rgba(15,118,110,0.35)',
                    whiteSpace: 'nowrap',
                  }}
                >
                  <Icon iconName="WorkFlow" style={{ fontSize: 11 }} />
                  Take Action
                </button>
              );
            })()}
          </div>
        );
      },
    };

    return [...dataCols, actionsCol];
  }

  // ── Edit helpers ──────────────────────────────────────────────────────────

  private _openEdit(record: ODataRecord): void {
    const editValues: Record<string, string> = {};

    // Text fields: FormattedValue annotation is fine — it IS the display text we want to edit
    EDIT_TEXT_FIELDS.forEach(field => {
      if (record[field] !== undefined) editValues[field] = strVal(record, field);
    });

    // Number fields: use the RAW numeric value, NOT the FormattedValue annotation.
    // strVal() for sct_estimated_total_value can return "7,500.00" or "$7,500.00" from the
    // OData annotation, which breaks parseFloat (returns 7 or NaN) in _saveEdit.
    EDIT_NUMBER_FIELDS.forEach(field => {
      const raw = record[field];
      if (raw !== undefined && raw !== null) editValues[field] = String(raw);
    });

    // Date fields: use the RAW OData Edm.Date string ("YYYY-MM-DD"), not the localized
    // FormattedValue annotation ("11/3/2026"), so formatToEdmDate always gets a clean input.
    EDIT_DATE_FIELDS.forEach(field => {
      const raw = record[field];
      if (raw !== undefined && raw !== null) editValues[field] = String(raw);
    });

    // Boolean fields: use actual JS boolean → "true"/"false" strings that toBoolean() and
    // tb() both handle. Avoids dependence on the "Yes"/"No" annotation locale.
    BOOLEAN_FIELDS.forEach(field => {
      const raw = record[field];
      if (raw !== undefined && raw !== null) {
        editValues[field] = typeof raw === 'boolean' ? (raw ? 'true' : 'false') : strVal(record, field);
      }
    });

    // Sourcing panel: pre-populate PO date if one exists
    // Use raw Edm.Date string for reliable parsing
    const poDateRaw = record['sct_po_issue_date'];
    const poDateStr = poDateRaw ? String(poDateRaw) : '';
    let poDate: Date | undefined;
    if (poDateStr) { const d = new Date(poDateStr); if (!isNaN(d.getTime())) poDate = d; }

    // Snapshot the values at open-time so _saveEdit can diff against them for audit logging
    const editInitialValues: Record<string, string> = {};
    Object.keys(editValues).forEach(k => { editInitialValues[k] = editValues[k]; });

    this.setState(
      { editRecord: record, editValues, editInitialValues, poDate, resolvedApprovers: null, loadingApprovers: false, saveError: null },
      () => { void this._loadApprovers(record); }
    );
  }

  /** Async — loads resolved approvers from sct_role for the record's site (Phase 8) */
  private async _loadApprovers(record: ODataRecord): Promise<void> {
    if (!this._router) return;
    const siteRaw = record['_sct_site_value'] ?? record['_sct_siteid_value'] ?? '';
    const siteId = typeof siteRaw === 'string' ? siteRaw : '';
    if (!siteId) return;
    const requesterGuid = String(record['_sct_requester_value'] ?? record['_sct_requestorid_value'] ?? '');
    this.setState({ loadingApprovers: true });
    try {
      const resolvedApprovers = await this._router.resolveApprovers(siteId, requesterGuid);
      this.setState({ resolvedApprovers, loadingApprovers: false });
    } catch {
      this.setState({ loadingApprovers: false });
    }
  }

  /** Stage action — PATCH only the status (and optional extra fields) with audit and queue sync */
  private async _stageAction(targetStatus: number, extra?: Record<string, unknown>): Promise<void> {
    const { editRecord } = this.state;
    if (!editRecord || !this._service) return;
    const id = this._getRecordId(editRecord);
    if (!id) return;
    this.setState({ saving: true, saveError: null });

    const payload: Record<string, unknown> = { sct_request_status: targetStatus, ...(extra ?? {}) };

    // Synchronize cr2a4_pendingtier: clear on terminal states, preserve or set to Manager on In Review
    if (
      targetStatus === DSTATUS.PO_ISSUED ||
      targetStatus === DSTATUS.REJECTED ||
      targetStatus === DSTATUS.RETURNED ||
      targetStatus === DSTATUS.CLOSED
    ) {
      payload['cr2a4_pendingtier'] = null;
    } else if (targetStatus === DSTATUS.IN_REVIEW) {
      const existingPt = String(
        editRecord['cr2a4_pendingtier'] || editRecord['sct_pendingtier'] || ''
      ).trim();
      payload['cr2a4_pendingtier'] = existingPt || 'Manager';
    }

    try {
      await this._service.updateRequest(id, payload);

      // Audit log the stage transition
      await this._service.createAuditLog({
        sct_action: targetStatus === DSTATUS.PO_ISSUED ? 1 : targetStatus === DSTATUS.REJECTED ? 5 : targetStatus === DSTATUS.RETURNED ? 4 : 2,
        sct_description: `Stage workbench updated status to ${targetStatus}`,
        sct_oldvalue: `Status: ${strVal(editRecord, 'sct_request_status')}`,
        sct_newvalue: `Status: ${targetStatus}, Pending Tier: ${payload['cr2a4_pendingtier'] ?? 'null'}`,
        'sct_requestid@odata.bind': `/sct_requests(${id})`,
        ...(this._currentSystemUserId ? { 'sct_actionby@odata.bind': `/systemusers(${this._currentSystemUserId})` } : {}),
      }).catch(() => { /* non-fatal */ });

      this.setState({ saving: false, editRecord: null, saveSuccess: 'Record updated.' });
      void this._load();
    } catch (err: unknown) {
      this.setState({ saving: false, saveError: err instanceof Error ? err.message : 'Update failed' });
    }
  }

  private async _resubmitReturned(): Promise<void> {
    const { editRecord, editValues } = this.state;
    if (!editRecord || !this._service) return;
    const id = this._getRecordId(editRecord);
    if (!id) return;
    this.setState({ saving: true, saveError: null });

    const payload: Record<string, unknown> = {
      sct_request_status: DSTATUS.SUBMITTED,
      cr2a4_pendingtier: 'Manager',
      sct_submitted_date: _dateToEdm(new Date()),
    };

    EDIT_TEXT_FIELDS.forEach(f => {
      const v = editValues[f];
      if (v !== undefined && v !== '') payload[f] = v;
    });
    EDIT_NUMBER_FIELDS.forEach(f => {
      const v = editValues[f];
      if (v !== undefined && v !== '') {
        const n = parseFloat(v.replace(/[^0-9.\-]/g, ''));
        if (!isNaN(n)) payload[f] = n;
      }
    });
    EDIT_DATE_FIELDS.forEach(f => {
      const v = editValues[f];
      if (v) {
        const edm = formatToEdmDate(v);
        if (edm) payload[f] = edm;
      }
    });
    BOOLEAN_FIELDS.forEach(f => {
      if (editValues[f] !== undefined) {
        payload[f] = toBoolean(editValues[f]);
      }
    });

    try {
      await this._service.updateRequest(id, payload);

      // Audit log the resubmission
      await this._service.createAuditLog({
        sct_action: 1,
        sct_description: 'Request resubmitted for approval by requester',
        sct_oldvalue: `Status: ${strVal(editRecord, 'sct_request_status')}`,
        sct_newvalue: `Status: Submitted, Pending Tier: Manager`,
        'sct_requestid@odata.bind': `/sct_requests(${id})`,
        ...(this._currentSystemUserId ? { 'sct_actionby@odata.bind': `/systemusers(${this._currentSystemUserId})` } : {}),
      }).catch(() => { /* non-fatal */ });

      // Notify manager via Graph
      try {
        const graphClient = await this.props.context.msGraphClientFactory.getClient('3');
        const manager = await getManager(graphClient);
        if (manager) {
          const title = strVal(editRecord, 'sct_request_title') || strVal(editRecord, 'sct_sct_reference');
          const amt = String(editRecord['sct_estimated_total_value'] ?? '');
          const reqName = this.props.context.pageContext.user.displayName;
          await sendManagerApprovalEmail(graphClient, manager, reqName, title, amt, this.props.context.pageContext.web.absoluteUrl);
        }
      } catch { /* non-fatal */ }

      this.setState({ saving: false, editRecord: null, saveSuccess: 'Request resubmitted for review successfully.' });
      void this._load();
    } catch (err: unknown) {
      this.setState({ saving: false, saveError: err instanceof Error ? err.message : 'Resubmission failed' });
    }
  }

  private async _saveEdit(): Promise<void> {
    const { editRecord, editValues, editInitialValues } = this.state;
    if (!editRecord || !this._service) return;
    const id = this._getRecordId(editRecord);
    if (!id) return;
    this.setState({ saving: true, saveError: null });

    // Build a type-safe PATCH payload. editValues stores everything as strings (from strVal/input),
    // but Dataverse requires strict types. Use the original record's raw value type as the guide.
    const payload: Record<string, unknown> = {};
    Object.keys(editValues).forEach(field => {
      const newStr = editValues[field];
      const originalRaw = editRecord[field];

      if (typeof originalRaw === 'boolean') {
        // Edm.Boolean — convert "Yes"/"No"/"true"/"false" to actual boolean
        payload[field] = toBoolean(newStr);
        return;
      }
      if (typeof originalRaw === 'number') {
        // Edm.Int32 / Edm.Decimal — strip any currency formatting characters before parsing.
        // _openEdit now loads raw numbers, but guard against formatted strings just in case.
        const stripped = newStr.replace(/[^0-9.\-]/g, '');
        const num = parseFloat(stripped);
        if (!isNaN(num)) payload[field] = num;
        // If still not parseable (e.g. "Standard" for a choice field display text), skip it —
        // choice field display text cannot be round-tripped; the edit panel makes them read-only.
        return;
      }
      // Date fields — must be strict YYYY-MM-DD for Edm.Date; localized strings like "2/1/2027" cause 400
      if (EDIT_DATE_FIELDS.has(field)) {
        const edm = formatToEdmDate(newStr);
        if (edm) payload[field] = edm;
        // If unparseable, skip the field rather than sending a bad value
        return;
      }
      // Plain string — send as-is if non-empty
      if (newStr !== '' && newStr !== undefined) payload[field] = newStr;
    });

    if (Object.keys(payload).length === 0) {
      // Nothing to patch — close without a round-trip
      this.setState({ saving: false, editRecord: null, saveSuccess: 'No changes to save.' });
      return;
    }

    try {
      await this._service.updateRequest(id, payload);
      // Close panel, show success, then reload the full list from Dataverse so the
      // table reflects the server's authoritative values (computed fields, triggers, etc.)
      this.setState({ saving: false, editRecord: null, saveSuccess: 'Record updated successfully.' });
      void this._load();
      // Log each changed field to sct_auditlog asynchronously (non-blocking)
      void logAuditChanges(
        this._service, id, this._currentSystemUserId, editInitialValues, editValues
      );
    } catch (err: unknown) {
      this.setState({
        saving: false,
        saveError: err instanceof Error ? err.message : 'Update failed',
      });
    }
  }

  private async _confirmDelete(): Promise<void> {
    const { deleteRecord } = this.state;
    if (!deleteRecord || !this._service) return;
    const id = this._getRecordId(deleteRecord);
    if (!id) return;
    this.setState({ saving: true });
    try {
      await this._service.deleteRequest(id);
      this.setState(prev => ({
        saving: false,
        deleteRecord: null,
        saveSuccess: 'Record deleted.',
        records: prev.records.filter(r => this._getRecordId(r) !== id),
      }));
      return;
    } catch (err: unknown) {
      this.setState({
        saving: false,
        saveError: err instanceof Error ? err.message : 'Delete failed',
        deleteRecord: null,
      });
      return;
    }
  }

  // ── Render helpers ────────────────────────────────────────────────────────

  private _renderViewPanel(): React.ReactElement | null {
    const { viewRecord, allFields } = this.state;
    if (!viewRecord) return null;

    // Build view field list: PRIORITY_FIELDS order first, then any remaining sct_ fields.
    // Exclude bare GUIDs, annotation-only keys, and system noise.
    // Booleans are shown here (unlike table columns) — this is a detail view.
    const usable = allFields.filter(f =>
      (f.startsWith('sct_') || f === 'createdon' || f === 'modifiedon') &&
      !isGuidField([viewRecord], f)
    );
    const prioritised: string[] = [];
    PRIORITY_FIELDS.forEach(p => { if (usable.indexOf(p) !== -1) prioritised.push(p); });
    usable.forEach(f => { if (prioritised.indexOf(f) === -1) prioritised.push(f); });
    const displayFields = prioritised;

    // Reference number as panel title; confirmed PrimaryNameAttribute = sct_sct_reference
    const panelTitle = strVal(viewRecord, 'sct_sct_reference') || 'Record Details';
    // Business status from sct_request_status (confirmed); fall back to statuscode if needed
    const statusLabel = strVal(viewRecord, 'sct_request_status') || strVal(viewRecord, 'statuscode');
    const isReturned = statusLabel === 'Returned';
    const canEdit = statusLabel === 'Draft' || isReturned;

    return (
      <Panel isOpen onDismiss={() => this.setState({ viewRecord: null })}
        type={PanelType.custom} customWidth="min(640px, 95vw)"
        headerText={panelTitle}
        closeButtonAriaLabel="Close"
        styles={{ content: { paddingBottom: 80 } }}
        isFooterAtBottom
        onRenderFooterContent={() => (
          <div style={{ display: 'flex', gap: 10, padding: '0 24px 24px' }}>
            <PrimaryButton
              text={isReturned ? 'Edit & Resubmit' : 'Edit Record'}
              iconProps={{ iconName: 'Edit' }}
              disabled={!canEdit}
              title={canEdit
                ? (isReturned ? 'Edit and resubmit this returned request' : 'Edit this draft request')
                : `Only Draft or Returned records can be edited — this record is "${statusLabel || 'unknown status'}"`}
              onClick={() => {
                const rec = viewRecord;
                this.setState({ viewRecord: null }, () => this._openEdit(rec));
              }}
              styles={{ root: { borderRadius: 6 } }}
            />
            <DefaultButton
              text="Close"
              onClick={() => this.setState({ viewRecord: null })}
              styles={{ root: { borderRadius: 6 } }}
            />
          </div>
        )}
      >
        {/* Status badge at top of panel */}
        {statusLabel && (
          <div style={{ marginBottom: 16 }}>
            <StatusPill value={statusLabel} />
          </div>
        )}

        {displayFields.map(field => {
          const val = strVal(viewRecord, field);
          return (
            <div key={field} style={{ borderBottom: '1px solid #F3F2F1', padding: '11px 0', display: 'flex', flexDirection: 'column', gap: 3 }}>
              <Text
                variant="small"
                style={{ color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.5px', fontWeight: 600 }}
              >
                {fieldLabel(field)}
              </Text>
              {field.indexOf('status') !== -1 || field.indexOf('type') !== -1 || field.indexOf('basis') !== -1
                ? <StatusPill value={val || '—'} />
                : <Text variant="medium" style={{ color: val ? '#201F1E' : '#C8C6C4', wordBreak: 'break-all', lineHeight: '1.5' }}>
                  {field.indexOf('amount') !== -1 && val
                    ? formatCurrency(parseFloat(val) || 0)
                    : (field === 'createdon' || field === 'modifiedon' || field.indexOf('date') !== -1) && val
                      ? formatDate(val)
                      : val || '—'}
                </Text>
              }
            </div>
          );
        })}
      </Panel>
    );
  }

  private _renderEditPanel(): React.ReactElement | null {
    const { editRecord, editValues, poDate, saving, saveError,
      resolvedApprovers, loadingApprovers } = this.state;
    if (!editRecord) return null;

    // ── Detect which stage panel to show ───────────────────────────────────
    const ref = strVal(editRecord, 'sct_sct_reference');
    const statusLabel = strVal(editRecord, 'sct_request_status');   // FormattedValue
    const statusRaw = editRecord['sct_request_status'];            // raw integer
    const site = strVal(editRecord, '_sct_site_value');
    const amt = parseFloat(String(editRecord['sct_estimated_total_value'] ?? 0));
    const isNew = editRecord['sct_is_new_supplier'] === true || editRecord['sct_is_new_supplier'] === 'Yes';
    const needsGate = (!isNaN(amt) && amt > D_CFO_THRESHOLD) || isNew;

    const isSubmitted = statusRaw === DSTATUS.SUBMITTED || statusLabel === 'Submitted';
    const isInReview = statusRaw === DSTATUS.IN_REVIEW || statusLabel === 'In Review';
    const isPOIssued = statusRaw === DSTATUS.PO_ISSUED || statusLabel === 'PO Issued';

    // Sourcing: In Review + standard (within threshold, known supplier)
    const isSourceStage = isInReview && !needsGate;
    // FP&A / CFO: In Review + gate required
    const isGateStage = isInReview && needsGate;

    // ── Helpers ───────────────────────────────────────────────────────────
    const setVal = (field: string) =>
      (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) =>
        this.setState(s => ({ editValues: { ...s.editValues, [field]: v ?? '' } }));
    const setToggle = (field: string) =>
      (_: React.MouseEvent<HTMLElement>, checked?: boolean) =>
        this.setState(s => ({ editValues: { ...s.editValues, [field]: checked ? 'true' : 'false' } }));
    const tv = (field: string): string => editValues[field] ?? '';
    const tb = (field: string): boolean => tv(field) === 'true' || tv(field) === 'Yes';

    const ss: React.CSSProperties = { marginBottom: 20, paddingBottom: 16, borderBottom: '1px solid #EDEBE9' };
    const tfS = { root: { marginBottom: 12 }, fieldGroup: { borderRadius: 6 } };
    const lbl: React.CSSProperties = { fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 2 };
    const roVal: React.CSSProperties = { fontSize: 13, color: '#201F1E', fontWeight: 500 };
    const close = (): void => this.setState({ editRecord: null, saveError: null });

    const SectionTitle: React.FC<{ label: string; color?: string }> = ({ label, color = '#0078D4' }) => (
      <div style={{ fontSize: 13, fontWeight: 700, color, marginBottom: 12 }}>{label}</div>
    );

    // ── Expanded read-only context strip (shown in all panels) ─────────────
    const ContextStrip = (): React.ReactElement => (
      <div style={{ background: '#F8F8F7', borderRadius: 8, padding: '12px 14px', marginBottom: 20 }}>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px 16px' }}>
          {([
            ['Reference', ref],
            ['Status', statusLabel],
            ['Site', site],
            ['Requester', strVal(editRecord, '_sct_requester_value')],
            ['Est. Value', isNaN(amt) ? '—' : formatCurrency(amt)],
            ['Purchase Type', strVal(editRecord, 'sct_purchase_type')],
            ['Urgency', strVal(editRecord, 'sct_urgency')],
            ['New Supplier', isNew ? 'Yes ⚑' : 'No'],
            ['Date Needed By', formatDate(strVal(editRecord, 'sct_date_needed_by'))],
            ['Submitted Date', formatDate(strVal(editRecord, 'sct_submitted_date'))],
          ] as Array<[string, string]>).map(([label, value]) => (
            <div key={label} style={{ marginBottom: 6 }}>
              <div style={lbl}>{label}</div>
              <div style={roVal}>{value || '—'}</div>
            </div>
          ))}
        </div>
        {strVal(editRecord, 'sct_description') && (
          <div style={{ marginTop: 10, paddingTop: 10, borderTop: '1px solid #EDEBE9' }}>
            <div style={lbl}>Description</div>
            <div style={{ fontSize: 12, color: '#605E5C', lineHeight: '1.5', maxHeight: 72, overflowY: 'auto' }}>
              {strVal(editRecord, 'sct_description')}
            </div>
          </div>
        )}
      </div>
    );

    // ── Resolved approvers block (Phase 8) ─────────────────────────────────
    const ApproversBlock = (): React.ReactElement => (
      <div style={{ marginTop: 16 }}>
        <SectionTitle label="Schedule A — Resolved Approvers" color="#605E5C" />
        {loadingApprovers && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '8px 0' }}>
            <Spinner size={SpinnerSize.small} />
            <Text variant="small" style={{ color: '#605E5C' }}>Resolving from sct_role…</Text>
          </div>
        )}
        {!loadingApprovers && resolvedApprovers && (
          <>
            {resolvedApprovers.approvers.length === 0 && (
              <Text variant="small" style={{ color: '#A19F9D', display: 'block', marginBottom: 8 }}>
                No approvers found for this site — verify sct_role records in Dataverse.
              </Text>
            )}
            {resolvedApprovers.approvers.map(a => (
              <div key={a.level + a.systemUserId} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', background: '#F8F8F8', borderRadius: 8, marginBottom: 8 }}>
                <div style={{ width: 34, height: 34, borderRadius: 8, background: '#0078D4', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontWeight: 700, fontSize: 11, flexShrink: 0 }}>
                  {a.level}
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: '#201F1E' }}>{a.name || a.systemUserId}</div>
                  <div style={{ display: 'flex', gap: 5, marginTop: 3 }}>
                    {a.isDelegate && <span style={{ fontSize: 10, background: '#FFF4CE', color: '#905400', padding: '1px 6px', borderRadius: 8, fontWeight: 700 }}>DELEGATE</span>}
                    {a.wasEscalated && <span style={{ fontSize: 10, background: '#FDE7E9', color: '#A4262C', padding: '1px 6px', borderRadius: 8, fontWeight: 700 }}>R10 ESCALATED</span>}
                  </div>
                </div>
              </div>
            ))}
            {resolvedApprovers.escalations.map((e, i) => (
              <MessageBar key={i} messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 6, marginBottom: 6 } }}>{e}</MessageBar>
            ))}
            {resolvedApprovers.warnings.map((w, i) => (
              <MessageBar key={i} messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 6, marginBottom: 6 } }}>{w}</MessageBar>
            ))}
          </>
        )}
      </div>
    );

    // ── Pre-approval requirements (Phase 8) ────────────────────────────────
    const preApproval = this._router?.calculatePreApprovals(isNaN(amt) ? 0 : amt, isNew);

    // ── Determine panel header text + footer buttons ────────────────────────
    let headerText = ref ? `Edit · ${ref}` : 'Edit Record';
    let footerContent: React.ReactElement;

    if (isSubmitted) {
      headerText = ref ? `Stage 3 — Triage · ${ref}` : 'Stage 3 — Triage';
      footerContent = (
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', padding: '0 24px 24px' }}>
          <PrimaryButton text={saving ? 'Saving…' : 'Submit to Sourcing'} disabled={saving}
            onClick={() => void this._stageAction(DSTATUS.IN_REVIEW, {
              sct_cost_center: tv('sct_cost_center') || undefined,
            })}
            styles={{ root: { borderRadius: 6 } }} />
          <DefaultButton text="Return for Information" disabled={saving}
            onClick={() => void this._stageAction(DSTATUS.RETURNED)}
            styles={{ root: { borderRadius: 6, color: '#905400', borderColor: '#C7A333' } }} />
          <DefaultButton text="Cancel" onClick={close} disabled={saving} styles={{ root: { borderRadius: 6 } }} />
        </div>
      );
    } else if (isSourceStage) {
      headerText = ref ? `Stage 4 — Sourcing · ${ref}` : 'Stage 4 — Sourcing';
      footerContent = (
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', padding: '0 24px 24px' }}>
          <PrimaryButton text={saving ? 'Saving…' : 'Issue PO'} disabled={saving}
            onClick={() => {
              const extra: Record<string, unknown> = {};
              const po = tv('sct_po_number').trim();
              if (po) extra['sct_po_number'] = po;
              if (poDate) extra['sct_po_issue_date'] = _dateToEdm(poDate);
              void this._stageAction(DSTATUS.PO_ISSUED, extra);
            }}
            styles={{ root: { borderRadius: 6 } }} />
          <DefaultButton text="Reject" disabled={saving}
            onClick={() => void this._stageAction(DSTATUS.REJECTED)}
            styles={{ root: { borderRadius: 6, color: '#A4262C', borderColor: '#A4262C' } }} />
          <DefaultButton text="Cancel" onClick={close} disabled={saving} styles={{ root: { borderRadius: 6 } }} />
        </div>
      );
    } else if (isGateStage) {
      headerText = ref ? `Stage 5 — FP&A / CFO · ${ref}` : 'Stage 5 — FP&A / CFO';
      footerContent = (
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', padding: '0 24px 24px' }}>
          <PrimaryButton text={saving ? 'Saving…' : 'Approve — Issue PO'} disabled={saving}
            onClick={() => void this._stageAction(DSTATUS.PO_ISSUED)}
            styles={{ root: { borderRadius: 6 } }} />
          <DefaultButton text="Reject" disabled={saving}
            onClick={() => void this._stageAction(DSTATUS.REJECTED)}
            styles={{ root: { borderRadius: 6, color: '#A4262C', borderColor: '#A4262C' } }} />
          <DefaultButton text="Return to Requester" disabled={saving}
            onClick={() => void this._stageAction(DSTATUS.RETURNED)}
            styles={{ root: { borderRadius: 6, color: '#905400', borderColor: '#C7A333' } }} />
          <DefaultButton text="Cancel" onClick={close} disabled={saving} styles={{ root: { borderRadius: 6 } }} />
        </div>
      );
    } else if (isPOIssued) {
      headerText = ref ? `Stage 6 — Receiving · ${ref}` : 'Stage 6 — Receiving';
      footerContent = (
        <div style={{ display: 'flex', gap: 10, padding: '0 24px 24px' }}>
          <PrimaryButton text={saving ? 'Saving…' : 'Confirm Received & Close'} disabled={saving}
            onClick={() => void this._stageAction(DSTATUS.CLOSED)}
            styles={{ root: { borderRadius: 6 } }} />
          <DefaultButton text="Cancel" onClick={close} disabled={saving} styles={{ root: { borderRadius: 6 } }} />
        </div>
      );
    } else {
      // Default — standard requester edit (Draft, Returned, or custom status)
      const isReturnedRecord = statusRaw === DSTATUS.RETURNED || statusLabel === 'Returned';
      footerContent = (
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', padding: '0 24px 24px' }}>
          <PrimaryButton
            text={saving ? 'Submitting…' : (isReturnedRecord ? 'Resubmit for Review' : 'Save Changes')}
            onClick={() => void (isReturnedRecord ? this._resubmitReturned() : this._saveEdit())}
            disabled={saving}
            styles={{ root: { borderRadius: 6 } }}
          />
          {!isReturnedRecord && (
            <DefaultButton
              text={saving ? 'Submitting…' : 'Submit for Review'}
              disabled={saving}
              onClick={() => void this._resubmitReturned()}
              styles={{ root: { borderRadius: 6, color: '#107C10', borderColor: '#107C10' } }}
            />
          )}
          <DefaultButton text="Cancel" onClick={close} disabled={saving} styles={{ root: { borderRadius: 6 } }} />
        </div>
      );
    }

    return (
      <Panel isOpen onDismiss={close}
        type={PanelType.custom} customWidth="min(820px, 95vw)"
        headerText={headerText}
        closeButtonAriaLabel="Close"
        styles={{ content: { paddingBottom: 88 } }}
        isFooterAtBottom
        onRenderFooterContent={() => footerContent}
      >
        {saveError && (
          <MessageBar messageBarType={MessageBarType.error} style={{ marginBottom: 16 }}>{saveError}</MessageBar>
        )}

        {/* Always-visible: extended read-only request details */}
        <ContextStrip />

        {/* ── Stage 3: Triage ───────────────────────────────────────────── */}
        {isSubmitted && (
          <>
            {preApproval && (
              <MessageBar
                messageBarType={preApproval.levels.length > 1 ? MessageBarType.warning : MessageBarType.info}
                styles={{ root: { borderRadius: 6, marginBottom: 14 } }}
              >
                <strong>Pre-Approval Required: {preApproval.levels.join(' + ')}</strong> — {preApproval.description}
              </MessageBar>
            )}
            <SectionTitle label="Stage 3 — Triage Assignment" color="#0078D4" />
            <TextField label="Cost Centre" value={tv('sct_cost_center')}
              onChange={setVal('sct_cost_center')} placeholder="e.g. CC-1042" styles={tfS} />
            <ApproversBlock />
          </>
        )}

        {/* ── Stage 4: Sourcing ─────────────────────────────────────────── */}
        {isSourceStage && (
          <>
            <SectionTitle label="Stage 4 — Procurement Sourcing" color="#107C10" />
            <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 6, marginBottom: 12 } }}>
              Within direct Procurement authority (≤ ${D_CFO_THRESHOLD.toLocaleString()}, known supplier). Issue PO directly.
            </MessageBar>
            <TextField label="D365 PO Number" value={tv('sct_po_number')}
              onChange={setVal('sct_po_number')} placeholder="e.g. PO-2026-00123" styles={tfS} />
            <DatePicker
              label="PO Issue Date"
              value={poDate}
              onSelectDate={d => this.setState({ poDate: d ?? undefined })}
              firstDayOfWeek={DayOfWeek.Monday}
              formatDate={d => d ? _dateToEdm(d) : ''}
              placeholder="Select PO issue date…"
              styles={{ root: { marginBottom: 14 } }}
              allowTextInput
            />
          </>
        )}

        {/* ── Stage 5: FP&A / CFO ───────────────────────────────────────── */}
        {isGateStage && (
          <>
            <SectionTitle label="Stage 5 — FP&A / CFO Review" color="#8764B8" />
            <MessageBar
              messageBarType={MessageBarType.warning}
              styles={{ root: { borderRadius: 6, marginBottom: 12 } }}
            >
              <strong>CFO approval gate triggered</strong> —
              {!isNaN(amt) && amt > D_CFO_THRESHOLD ? ` value ${formatCurrency(amt)} exceeds $${D_CFO_THRESHOLD.toLocaleString()}.` : ''}
              {isNew ? ' New supplier.' : ''}
              {' '}Approve to issue PO or Reject.
            </MessageBar>
            {/* Read-only sourcing info */}
            {(strVal(editRecord, 'sct_po_number') || strVal(editRecord, '_sct_suggested_supplier_value')) && (
              <div style={{ background: '#F8F8F7', borderRadius: 8, padding: '10px 14px', marginBottom: 14 }}>
                {strVal(editRecord, '_sct_suggested_supplier_value') && (
                  <div style={{ marginBottom: 6 }}>
                    <div style={lbl}>Supplier</div>
                    <div style={roVal}>{strVal(editRecord, '_sct_suggested_supplier_value')}</div>
                  </div>
                )}
                {strVal(editRecord, 'sct_po_number') && (
                  <div>
                    <div style={lbl}>PO Number</div>
                    <div style={roVal}>{strVal(editRecord, 'sct_po_number')}</div>
                  </div>
                )}
              </div>
            )}
            <ApproversBlock />
          </>
        )}

        {/* ── Stage 6: Receiving ────────────────────────────────────────── */}
        {isPOIssued && (
          <>
            <SectionTitle label="Stage 6 — Receiving" color="#00B294" />
            <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 6, marginBottom: 12 } }}>
              Confirm that goods or services have been received. This closes the request in the Spend Control Tower.
            </MessageBar>
            {strVal(editRecord, 'sct_po_number') && (
              <div style={{ background: '#F8F8F7', borderRadius: 8, padding: '10px 14px', marginBottom: 14 }}>
                <div style={lbl}>PO Number</div>
                <div style={roVal}>{strVal(editRecord, 'sct_po_number')}</div>
              </div>
            )}
          </>
        )}

        {/* ── Default: standard editable form (Draft or unknown status) ─── */}
        {!isSubmitted && !isInReview && !isPOIssued && (
          <>
            <div style={ss}>
              <SectionTitle label="Description & Scope" />
              <TextField label="Description" value={tv('sct_description')} multiline rows={4}
                onChange={setVal('sct_description')} styles={tfS} />
              <TextField label="Function / Department" value={tv('sct_function')}
                onChange={setVal('sct_function')} styles={tfS} />
            </div>
            <div style={ss}>
              <SectionTitle label="Urgency Details" />
              <TextField label="Urgency Reason" value={tv('sct_urgency_reason')} multiline rows={2}
                onChange={setVal('sct_urgency_reason')} styles={tfS} />
              <TextField label="Emergency Declared By (email)" value={tv('sct_emergency_declared_by_email')}
                onChange={setVal('sct_emergency_declared_by_email')} styles={tfS} />
            </div>
            <div style={ss}>
              <SectionTitle label="Financials & Dates" />
              <TextField label="Estimated Total Value (AUD)" value={tv('sct_estimated_total_value')}
                onChange={setVal('sct_estimated_total_value')} styles={tfS} />
              <TextField label="Date Needed By (YYYY-MM-DD)" value={tv('sct_date_needed_by')}
                onChange={setVal('sct_date_needed_by')} styles={tfS} />
              <TextField label="Term Start Date (YYYY-MM-DD)" value={tv('sct_term_start_date')}
                onChange={setVal('sct_term_start_date')} styles={tfS} />
              <TextField label="Term End Date (YYYY-MM-DD)" value={tv('sct_term_end_date')}
                onChange={setVal('sct_term_end_date')} styles={tfS} />
            </div>
            <div style={ss}>
              <SectionTitle label="Logistics" />
              <TextField label="Delivery Address" value={tv('sct_delivery_address')} multiline rows={2}
                onChange={setVal('sct_delivery_address')} styles={tfS} />
              <TextField label="Receiving Contact Email" value={tv('sct_receiving_contact_email')}
                onChange={setVal('sct_receiving_contact_email')} styles={tfS} />
              <TextField label="Equipment Asset Tag" value={tv('sct_equipment_asset_tag')}
                onChange={setVal('sct_equipment_asset_tag')} styles={tfS} />
              <TextField label="Customer Name (if recharged)" value={tv('sct_customer_name')}
                onChange={setVal('sct_customer_name')} styles={tfS} />
            </div>
            <div style={{ marginBottom: 8 }}>
              <SectionTitle label="Flags & Declarations" />
              <Toggle label="Is New Supplier" checked={tb('sct_is_new_supplier')} onChange={setToggle('sct_is_new_supplier')} styles={{ root: { marginBottom: 8 } }} />
              <Toggle label="Cost Recharged to Customer" checked={tb('sct_cost_recharged_to_customer')} onChange={setToggle('sct_cost_recharged_to_customer')} styles={{ root: { marginBottom: 8 } }} />
              <Toggle label="Customer Agreement Attached" checked={tb('sct_customer_agreement_attached')} onChange={setToggle('sct_customer_agreement_attached')} styles={{ root: { marginBottom: 8 } }} />
              <Toggle label="Tech or Data Involved" checked={tb('sct_tech_or_data_involved')} onChange={setToggle('sct_tech_or_data_involved')} styles={{ root: { marginBottom: 8 } }} />
              <Toggle label="Supplier on Premises" checked={tb('sct_supplier_on_premises')} onChange={setToggle('sct_supplier_on_premises')} styles={{ root: { marginBottom: 8 } }} />
              <Toggle label="Possible Capital Purchase" checked={tb('sct_possible_capital_purchase')} onChange={setToggle('sct_possible_capital_purchase')} styles={{ root: { marginBottom: 8 } }} />
              <Toggle label="Confidential" checked={tb('sct_is_confidential')} onChange={setToggle('sct_is_confidential')} styles={{ root: { marginBottom: 8 } }} />
              <Toggle label="Retrospective" checked={tb('sct_is_retrospective')} onChange={setToggle('sct_is_retrospective')} styles={{ root: { marginBottom: 8 } }} />
              <Toggle label="Existing Agreement or PO" checked={tb('sct_existing_agreement_or_po')} onChange={setToggle('sct_existing_agreement_or_po')} styles={{ root: { marginBottom: 8 } }} />
              <Toggle label="Declaration – No Prior Commitment" checked={tb('sct_declaration_no_prior_commitment')} onChange={setToggle('sct_declaration_no_prior_commitment')} styles={{ root: { marginBottom: 8 } }} />
              <Toggle label="Declaration – No Conflict of Interest" checked={tb('sct_declaration_no_conflict_of_interest')} onChange={setToggle('sct_declaration_no_conflict_of_interest')} styles={{ root: { marginBottom: 8 } }} />
            </div>
          </>
        )}
      </Panel>
    );
  }

  private _renderDeleteDialog(): React.ReactElement | null {
    const { deleteRecord, saving } = this.state;
    if (!deleteRecord) return null;
    const titleField = this.state.allFields.find(f => f.indexOf('name') !== -1 || f.indexOf('title') !== -1);
    const label = titleField ? strVal(deleteRecord, titleField) : 'this record';
    return (
      <Dialog isOpen hidden={false}
        onDismiss={() => this.setState({ deleteRecord: null })}
        dialogContentProps={{ type: DialogType.normal, title: 'Delete Record', subText: `Permanently delete "${label}"? This cannot be undone.` }}
        modalProps={{ isBlocking: true }}
      >
        <DialogFooter>
          <PrimaryButton text={saving ? 'Deleting…' : 'Delete'} onClick={() => void this._confirmDelete()} disabled={saving} styles={{ root: { background: '#A4262C', border: 'none', borderRadius: 6 } }} />
          <DefaultButton text="Cancel" onClick={() => this.setState({ deleteRecord: null })} disabled={saving} styles={{ root: { borderRadius: 6 } }} />
        </DialogFooter>
      </Dialog>
    );
  }

  private _renderFieldInspector(): React.ReactElement | null {
    const { showFieldInspector, records } = this.state;
    if (!showFieldInspector || !records.length) return null;
    const allKeys = Object.keys(records[0]).filter(k => !k.startsWith('@')).sort();
    return (
      <Panel isOpen onDismiss={() => this.setState({ showFieldInspector: false })}
        type={PanelType.custom} customWidth="min(640px, 95vw)"
        headerText={`Field Inspector — ${allKeys.length} columns from sct_request`}
        closeButtonAriaLabel="Close"
        styles={{ content: { paddingBottom: 40 } }}
      >
        <Text variant="small" style={{ color: '#605E5C', marginBottom: 12, display: 'block', lineHeight: '1.6' }}>
          Real logical names returned by your Dataverse <code style={{ background: '#F3F2F1', padding: '1px 5px', borderRadius: 4 }}>sct_request</code> entity.
          Use these to update <code style={{ background: '#F3F2F1', padding: '1px 5px', borderRadius: 4 }}>src/models/ISCTRequest.ts</code>.
        </Text>
        <div style={{ fontFamily: 'Consolas, monospace', fontSize: 12, background: '#F8F8F8', border: '1px solid #EDEBE9', borderRadius: 8, padding: 16, maxHeight: 520, overflowY: 'auto' }}>
          {allKeys.map(k => {
            const sample = records[0][k];
            const type = sample === null ? 'null' : typeof sample === 'number' ? 'number' : typeof sample === 'boolean' ? 'boolean' : 'string';
            const sampleStr = sample !== null && sample !== undefined ? String(sample).substring(0, 50) : '';
            return (
              <div key={k} style={{ marginBottom: 5, display: 'flex', gap: 8, alignItems: 'baseline' }}>
                <span style={{ color: '#0078D4', minWidth: 260 }}>{k}</span>
                <span style={{ color: '#107C10', minWidth: 60 }}>{type}</span>
                <span style={{ color: '#A19F9D', fontSize: 11 }}>{sampleStr}</span>
              </div>
            );
          })}
        </div>
      </Panel>
    );
  }

  private _renderPagination(totalFiltered: number): React.ReactElement {
    const { currentPage, pageSize } = this.state;
    const totalPages = Math.max(1, Math.ceil(totalFiltered / pageSize));
    const start = Math.min((currentPage - 1) * pageSize + 1, totalFiltered);
    const end = Math.min(currentPage * pageSize, totalFiltered);

    // Show at most 5 page numbers
    const pages: number[] = [];
    if (totalPages <= 5) {
      for (let i = 1; i <= totalPages; i++) pages.push(i);
    } else if (currentPage <= 3) {
      pages.push(1, 2, 3, 4, 5);
    } else if (currentPage >= totalPages - 2) {
      for (let i = totalPages - 4; i <= totalPages; i++) pages.push(i);
    } else {
      for (let i = currentPage - 2; i <= currentPage + 2; i++) pages.push(i);
    }

    return (
      <div className={styles.pagination}>
        <Text variant="small" className={styles.paginationInfo}>
          Showing {totalFiltered === 0 ? 0 : start}–{end} of {totalFiltered} records
        </Text>
        <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
          <div className={styles.paginationControls}>
            <button className={styles.pageBtn} disabled={currentPage === 1}
              onClick={() => this.setState({ currentPage: 1 })} aria-label="First page">«</button>
            <button className={styles.pageBtn} disabled={currentPage === 1}
              onClick={() => this.setState(s => ({ currentPage: s.currentPage - 1 }))} aria-label="Previous page">‹</button>
            {pages.map(p => (
              <button key={p}
                className={`${styles.pageBtn} ${p === currentPage ? styles.pageBtnActive : ''}`}
                onClick={() => this.setState({ currentPage: p })} aria-label={`Page ${p}`} aria-current={p === currentPage}>{p}</button>
            ))}
            <button className={styles.pageBtn} disabled={currentPage === totalPages}
              onClick={() => this.setState(s => ({ currentPage: s.currentPage + 1 }))} aria-label="Next page">›</button>
            <button className={styles.pageBtn} disabled={currentPage === totalPages}
              onClick={() => this.setState({ currentPage: totalPages })} aria-label="Last page">»</button>
          </div>
          <select className={styles.pageSizeSelect} value={pageSize}
            onChange={e => this.setState({ pageSize: Number(e.target.value), currentPage: 1 })}
            aria-label="Rows per page">
            {PAGE_SIZE_OPTIONS.map(n => <option key={n} value={n}>{n} / page</option>)}
          </select>
        </div>
      </div>
    );
  }

  // ── History panel ─────────────────────────────────────────────────────────

  private _renderHistoryPanel(): React.ReactElement | null {
    const { historyRecord } = this.state;
    if (!historyRecord || !this._service) return null;
    const id = this._getRecordId(historyRecord);
    if (!id) return null;
    const ref = strVal(historyRecord, 'sct_sct_reference');
    return (
      <RequestLogsPanel
        requestId={id}
        reference={ref}
        service={this._service}
        onDismiss={() => this.setState({ historyRecord: null })}
      />
    );
  }

  // ── Take Action (Phase 4) ─────────────────────────────────────────────────

  /** Display label for the current approver tier */
  private _tierName(): string {
    switch (this.props.activeRole) {
      case 'manager': return 'Manager';
      case 'approver_tier1': return 'Tier 1';
      case 'approver_tier2': return 'Tier 2';
      case 'approver_tier3': return 'Tier 3';
      default: return 'Approver';
    }
  }

  private _renderTakeActionDialog(): React.ReactElement | null {
    const { takeActionRecord } = this.state;
    if (!takeActionRecord || !this._service) return null;

    return (
      <TakeActionModal
        isOpen={takeActionRecord !== null}
        record={takeActionRecord}
        service={this._service}
        activeRole={this.props.activeRole || 'manager'}
        currentSystemUserId={this._currentSystemUserId}
        context={this.props.context}
        soloTestMode={this.props.soloTestMode}
        onDismiss={() => this.setState({ takeActionRecord: null })}
        onSuccess={(msg: string) => {
          this.setState({
            takeActionRecord: null,
            saveSuccess: msg,
          });
          setTimeout(() => {
            if (this.props.activeRole === 'manager') {
              void this._loadManagerData();
            } else {
              void this._load();
            }
          }, 500);
        }}
      />
    );
  }

  // ── Main render ───────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const {
      loading, error, saveSuccess, saveError,
      records,
      searchText, statusFilter,
    } = this.state;


    // Merge static known statuses (always shown) with any live values from Dataverse
    const liveStatusValues = detectStatusValues(records);
    const allStatusValues: string[] = ALL_REQUEST_STATUS_VALUES.slice();
    liveStatusValues.forEach(v => {
      if (allStatusValues.indexOf(v) === -1) allStatusValues.push(v);
    });

    const filtered = this._filteredRecords();
    const counts = this._counts(filtered);
    const paged = this._pagedRecords(filtered);

    return (
      <div>
        {/* ── Stat tiles ─────────────────────────────────────────────── */}
        <div className={styles.statRow}>
          <div className={`${styles.statCard} ${styles.statCardBlue}`}>
            <Icon iconName="DocumentSet" style={{ fontSize: 20, color: '#0078D4' }} aria-hidden="true" />
            <span className={styles.statValue}>{loading ? '…' : counts.total}</span>
            <span className={styles.statLabel}>Total Requests</span>
          </div>
          <div className={`${styles.statCard} ${styles.statCardOrange}`}>
            <Icon iconName="Clock" style={{ fontSize: 20, color: '#F7630C' }} aria-hidden="true" />
            <span className={styles.statValue}>{loading ? '…' : counts.pending}</span>
            <span className={styles.statLabel}>Pending Approval</span>
          </div>
          <div className={`${styles.statCard} ${styles.statCardGreen}`}>
            <Icon iconName="CheckMark" style={{ fontSize: 20, color: '#107C10' }} aria-hidden="true" />
            <span className={styles.statValue}>{loading ? '…' : counts.approved}</span>
            <span className={styles.statLabel}>Approved</span>
          </div>
          <div className={`${styles.statCard} ${styles.statCardTeal}`}>
            <Icon iconName="Money" style={{ fontSize: 20, color: '#00B294' }} aria-hidden="true" />
            <span className={styles.statValue} style={{ fontSize: 20 }}>
              {loading ? '…' : formatCurrency(counts.spend)}
            </span>
            <span className={styles.statLabel}>Est. Total Spend</span>
          </div>
        </div>

        {/* ── Success / error banners ─────────────────────────────────── */}
        {saveSuccess && (
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}
            onDismiss={() => this.setState({ saveSuccess: null })}
            styles={{ root: { borderRadius: 6, marginBottom: 12 } }}>
            {saveSuccess}
          </MessageBar>
        )}
        {saveError && (
          <ErrorBanner message={saveError} onDismiss={() => this.setState({ saveError: null })} />
        )}
        {error && (
          <ErrorBanner message={error} onRetry={() => void this._load()}
            onDismiss={() => this.setState({ error: null })} />
        )}

        {/* ── Requester single tab (My Requests) ──────────────────────── */}
        {this.props.activeRole === 'requester' && (
          <Pivot
            selectedKey="myRequests"
            styles={{ root: { marginBottom: 12 } }}
          >
            <PivotItem headerText="My Requests" itemKey="myRequests" itemIcon="Contact" />
          </Pivot>
        )}

        {/* ── Manager Pivot tabs (Pending / Completed / My Requests) ──── */}
        {this.props.activeRole === 'manager' && (
          <Pivot
            selectedKey={this.state.requesterTab}
            onLinkClick={item => item && this.setState({ requesterTab: item.props.itemKey as 'pending' | 'completed' | 'myRequests', currentPage: 1 })}
            styles={{ root: { marginBottom: 12 } }}
          >
            <PivotItem headerText="Pending My Approval" itemKey="pending" itemIcon="Clock" />
            <PivotItem headerText="Completed" itemKey="completed" itemIcon="CheckMark" />
            <PivotItem headerText="My Requests" itemKey="myRequests" itemIcon="Contact" />
          </Pivot>
        )}

        {/* ── Approver Tier Pivot tabs (Pending / Completed / My Requests) ── */}
        {(this.props.activeRole === 'approver_tier1' ||
          this.props.activeRole === 'approver_tier2' ||
          this.props.activeRole === 'approver_tier3') && (
            <Pivot
              selectedKey={this.state.requesterTab}
              onLinkClick={item => item && this.setState({ requesterTab: item.props.itemKey as 'pending' | 'completed' | 'myRequests', currentPage: 1 })}
              styles={{ root: { marginBottom: 12 } }}
            >
              <PivotItem headerText="Pending Approval" itemKey="pending" itemIcon="Clock" />
              <PivotItem headerText="Completed" itemKey="completed" itemIcon="CheckMark" />
              <PivotItem headerText="My Requests" itemKey="myRequests" itemIcon="Contact" />
            </Pivot>
          )}

        {/* ── Segregation of Duties (SoD) Alert Banner ── */}
        {this.state.sodHiddenCount > 0 && (
          <div style={{
            background: '#FFFBEB',
            border: '1px solid #FCD34D',
            borderRadius: 8,
            padding: '12px 18px',
            marginBottom: 14,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            gap: 12,
            boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <Icon iconName="ShieldAlert" style={{ fontSize: 20, color: '#D97706' }} />
              <div>
                <strong style={{ color: '#92400E', fontSize: 13 }}>Segregation of Duties (SoD) Active:</strong>{' '}
                <span style={{ color: '#78350F', fontSize: 13 }}>
                  {this.state.sodHiddenCount} pending request{this.state.sodHiddenCount > 1 ? 's' : ''} submitted by you {this.state.sodHiddenCount > 1 ? 'have' : 'has'} self-approval locked to maintain audit compliance.
                </span>
              </div>
            </div>
            {this.props.onToggleSoloTestMode && (
              <button
                type="button"
                onClick={this.props.onToggleSoloTestMode}
                style={{
                  background: '#D97706', color: '#ffffff', border: 'none', borderRadius: 6,
                  padding: '6px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                  boxShadow: '0 1px 2px rgba(0,0,0,0.1)',
                }}
              >
                <Icon iconName="Unlock" style={{ fontSize: 12 }} />
                Enable Solo Test Mode to Review
              </button>
            )}
          </div>
        )}

        {/* ── Table card ─────────────────────────────────────────────── */}
        <div className={styles.card}>

          {/* Toolbar */}
          <div className={styles.toolbar}>
            {/* Search */}
            <div className={styles.searchBox} role="search">
              <Icon iconName="Search" style={{ fontSize: 14, color: '#A19F9D', flexShrink: 0 }} aria-hidden="true" />
              <input
                type="search"
                placeholder="Search all fields…"
                value={searchText}
                aria-label="Search requests"
                onChange={e => this.setState({ searchText: e.target.value, currentPage: 1 })}
              />
              {searchText && (
                <button
                  style={{ border: 'none', background: 'none', cursor: 'pointer', padding: 0, color: '#A19F9D', lineHeight: 1 }}
                  onClick={() => this.setState({ searchText: '', currentPage: 1 })}
                  aria-label="Clear search"
                >
                  <Icon iconName="Cancel" style={{ fontSize: 12 }} />
                </button>
              )}
            </div>

            {/* Status filter */}
            <select
              className={styles.filterSelect}
              value={statusFilter}
              aria-label="Filter by status"
              onChange={e => this.setState({ statusFilter: e.target.value, currentPage: 1 })}
            >
              <option value="all">All Statuses</option>
              {allStatusValues.map(v => <option key={v} value={v}>{v}</option>)}
            </select>

            <div className={styles.toolbarSpacer} />

            {/* Card title inline — reflects the active role view */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <Icon iconName="DocumentSet" style={{ color: '#0078D4' }} aria-hidden="true" />
              <Text variant="mediumPlus" style={{ fontWeight: 600, color: '#201F1E' }}>
                {this.props.activeRole === 'requester' ? 'My Requests'
                  : this.props.activeRole === 'manager' ? (this.state.requesterTab === 'completed' ? 'Completed — Manager' : this.state.requesterTab === 'myRequests' ? 'My Submitted Requests' : 'Pending My Approval')
                    : this.props.activeRole === 'approver_tier1' ? (this.state.requesterTab === 'completed' ? 'Completed — Tier 1' : this.state.requesterTab === 'myRequests' ? 'My Submitted Requests' : 'Pending — Tier 1')
                      : this.props.activeRole === 'approver_tier2' ? (this.state.requesterTab === 'completed' ? 'Completed — Tier 2' : this.state.requesterTab === 'myRequests' ? 'My Submitted Requests' : 'Pending — Tier 2')
                        : this.props.activeRole === 'approver_tier3' ? (this.state.requesterTab === 'completed' ? 'Completed — Tier 3' : this.state.requesterTab === 'myRequests' ? 'My Submitted Requests' : 'Pending — Tier 3')
                          : 'All Requests'}
              </Text>
              {this.state.roleFilterLoading && (
                <span style={{ fontSize: 11, color: '#A19F9D', fontStyle: 'italic' }}>filtering…</span>
              )}
              {!loading && !this.state.roleFilterLoading && (
                <span style={{ background: '#EBF3FB', color: '#0078D4', fontSize: 12, fontWeight: 600, padding: '2px 10px', borderRadius: 12 }}>
                  {filtered.length}
                </span>
              )}
            </div>

            <div className={styles.toolbarSpacer} />

            {/* Export CSV button */}
            <DefaultButton
              text="Export CSV"
              iconProps={{ iconName: 'ExcelDocument' }}
              title="Download currently filtered requests as CSV"
              disabled={loading || filtered.length === 0}
              onClick={() => this._exportToCsv(filtered)}
              styles={{ root: { borderRadius: 6, height: 32, borderColor: '#EDEBE9' } }}
            />

            <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }}
              onClick={() => void this._load()} disabled={loading}
              styles={{ root: { borderRadius: 6, height: 32 } }}
            />
          </div>


          {/* Body */}
          <div style={{ padding: 0 }}>
            {loading ? (
              <LoadingSpinner label="Loading from Dataverse…" />
            ) : filtered.length === 0 && !error ? (
              <EmptyState
                message={searchText || statusFilter !== 'all' ? 'No matching records' : 'No requests found'}
                subMessage={searchText || statusFilter !== 'all' ? 'Try adjusting your search or filter.' : 'Submit a new spend request to get started.'}
                iconName="DocumentSearch"
                actionLabel={searchText || statusFilter !== 'all' ? 'Clear filters' : undefined}
                onAction={searchText || statusFilter !== 'all' ? () => this.setState({ searchText: '', statusFilter: 'all', currentPage: 1 }) : undefined}
              />
            ) : (
              <div style={{ overflowX: 'auto', width: '100%' }}>
                <DetailsList
                  items={paged}
                  columns={this._buildColumns(allStatusValues)}
                  layoutMode={DetailsListLayoutMode.justified}
                  selectionMode={SelectionMode.none}
                  styles={{ root: { minWidth: 600 } }}
                  ariaLabelForGrid="Spend control requests"
                />
              </div>
            )}
          </div>

          {/* Pagination */}
          {!loading && filtered.length > 0 && this._renderPagination(filtered.length)}
        </div>

        {/* ── Panels / dialogs ─────────────────────────────────────────── */}
        {this._renderViewPanel()}
        {this._renderEditPanel()}
        {this._renderDeleteDialog()}
        {this._renderFieldInspector()}
        {this._renderHistoryPanel()}
        {this._renderTakeActionDialog()}
      </div>
    );
  }
}
