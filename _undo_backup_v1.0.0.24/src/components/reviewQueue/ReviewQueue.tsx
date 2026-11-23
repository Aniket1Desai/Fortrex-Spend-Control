import * as React from 'react';
import { DataverseService, ODataRecord, strVal } from '../../services/DataverseService';
import { handleApprovalAction } from '../../services/approvalHandler';
import { ApprovalRoutingService, IResolvedApprovers, IPreApprovalRequirements, IPostSourcingRequirements } from '../../services/ApprovalRoutingService';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react/lib/Calendar';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

import { LoadingSpinner, EmptyState, ErrorBanner, TakeActionModal } from '../common';
import styles from '../shell/AppShell.module.scss';
import { formatCurrency, formatDate } from '../dashboard/dashboardUtils';

// ── Types ─────────────────────────────────────────────────────────────────────

interface IReviewQueueProps {
  dataverseService: DataverseService;
  context: import('@microsoft/sp-webpart-base').WebPartContext;
  activeRole?: string;
  soloTestMode?: boolean;
}

type StageId = 'triage' | 'sourcing' | 'fpa' | 'cfo' | 'receiving';

interface IPendingStageRow {
  stage: string;
  role: string;
}

interface IReviewQueueState {
  activeStage: StageId;
  records: ODataRecord[];
  loading: boolean;
  error: string | null;
  // Action panel
  actionRecord: ODataRecord | null;
  actionValues: Record<string, string>;
  poDate: Date | undefined;           // DatePicker for PO Issue Date
  saving: boolean;
  saveError: string | null;
  saveSuccess: string | null;
  // Approval routing
  preApproval: IPreApprovalRequirements | null;
  postSourcing: IPostSourcingRequirements | null;
  resolvedApprovers: IResolvedApprovers | null;
  loadingApprovers: boolean;
  // Pending stage tracker
  pendingStageRows: IPendingStageRow[];
  loadingPending: boolean;
  // Manager queue state
  systemUserId: string | null;
  managerApprovedIds: Set<string>;
  directReportIds: Set<string>;
  managerActionRecord: ODataRecord | null;
  managerComments: string;
  managerSaving: boolean;
  managerError: string | null;
}

// ── Status option-set integers — confirmed from Power Apps sct_Request_Status (854000001 to 854000007) ──
// Source: Power Apps > Fortrex Spend Control T. > Tables > SCT Request > Columns > Status
const STATUS = {
  DRAFT:      854000001,  // Draft
  SUBMITTED:  854000002,  // Submitted
  IN_REVIEW:  854000003,  // In Review   (covers Triage → Sourcing → FP&A → CFO)
  PO_ISSUED:  854000004,  // PO Issued   (Procurement complete, awaiting receipt)
  REJECTED:   854000005,  // Rejected
  RETURNED:   854000006,  // Returned    (returned to requester for information)
  CLOSED:     854000007,  // Closed      (received and complete)
};

const CFO_THRESHOLD = 10000;

// ── Stage definitions ─────────────────────────────────────────────────────────

interface IStageDef {
  id: StageId;
  stage: string;
  title: string;
  desc: string;
  iconName: string;
  color: string;
  bg: string;
  watchLabels: string[];
  primaryAction: string;
  primaryStatus: number;
  altAction?: string;
  altStatus?: number;
}

// Note: Dataverse has 7 statuses only (confirmed from Power Apps).
// FP&A and CFO are conceptual sub-stages that BOTH operate while status = "In Review".
// They are distinguished in the UI by filtering on estimated value / new-supplier flag.
// Both resolve by issuing a PO → status transitions to "PO Issued" (854000003).
const STAGES: IStageDef[] = [
  {
    id: 'triage',
    stage: '3',
    title: 'SCT Desk — Triage',
    desc: 'Assign cost centre and advance submitted requests into the review pipeline',
    iconName: 'Sort',
    color: '#0078D4',
    bg: '#EBF3FB',
    watchLabels: ['Submitted'],
    primaryAction: 'Submit to Sourcing',
    primaryStatus: STATUS.IN_REVIEW,
    altAction: 'Return for Information',
    altStatus: STATUS.RETURNED,
  },
  {
    id: 'sourcing',
    stage: '4',
    title: 'Procurement — Sourcing',
    desc: 'Collect quotes and raise PO for standard requests (value ≤ $10k, known supplier)',
    iconName: 'ShoppingCart',
    color: '#107C10',
    bg: '#DFF6DD',
    watchLabels: ['In Review'],
    primaryAction: 'Raise PO',
    primaryStatus: STATUS.PO_ISSUED,
    altAction: 'Reject',
    altStatus: STATUS.REJECTED,
  },
  {
    id: 'fpa',
    stage: '5a',
    title: 'FP&A Review',
    desc: 'Post-sourcing review for requests >$10k or involving a new supplier — still "In Review" in Dataverse',
    iconName: 'BarChart4',
    color: '#8764B8',
    bg: '#F4F0FF',
    watchLabels: ['In Review'],
    primaryAction: 'FP&A Approved — Issue PO',
    primaryStatus: STATUS.PO_ISSUED,
    altAction: 'Reject',
    altStatus: STATUS.REJECTED,
  },
  {
    id: 'cfo',
    stage: '5b',
    title: 'CFO Approval',
    desc: 'L4 CFO gate — requests >$10k or new supplier requiring CFO sign-off before PO can be issued',
    iconName: 'Permissions',
    color: '#A4262C',
    bg: '#FDE7E9',
    watchLabels: ['In Review'],
    primaryAction: 'CFO Approved — Issue PO',
    primaryStatus: STATUS.PO_ISSUED,
    altAction: 'Reject',
    altStatus: STATUS.REJECTED,
  },
  {
    id: 'receiving',
    stage: '6',
    title: 'Receiving',
    desc: 'Confirm goods or services received and close the spend request',
    iconName: 'Completed',
    color: '#00B294',
    bg: '#E6F7F4',
    watchLabels: ['PO Issued'],
    primaryAction: 'Confirm Received & Close',
    primaryStatus: STATUS.CLOSED,
  },
];

// ── Date helpers (no padStart — not in this project's TS lib target) ──────────

function dateToEdm(d: Date): string {
  const y  = d.getFullYear();
  const mo = d.getMonth() + 1;
  const da = d.getDate();
  const mm = mo < 10 ? ('0' + mo) : String(mo);
  const dd = da < 10 ? ('0' + da) : String(da);
  return y + '-' + mm + '-' + dd;
}

function edmToDate(str: string): Date | undefined {
  if (!str) return undefined;
  const d = new Date(str);
  return isNaN(d.getTime()) ? undefined : d;
}

// ── Status pill ───────────────────────────────────────────────────────────────

const StatusPill: React.FC<{ value: string; color: string; bg: string }> = ({ value, color, bg }) => (
  <span style={{ background: bg, color, fontSize: 11, fontWeight: 700, padding: '3px 10px', borderRadius: 12, whiteSpace: 'nowrap' }}>
    {value || '—'}
  </span>
);

// ── Approver card ─────────────────────────────────────────────────────────────

const ApproverCard: React.FC<{ level: string; name: string; isDelegate: boolean; wasEscalated: boolean }> = (
  { level, name, isDelegate, wasEscalated }
) => (
  <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', background: '#F8F8F8', borderRadius: 8, marginBottom: 8 }}>
    <div style={{ width: 36, height: 36, borderRadius: 8, background: '#0078D4', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontWeight: 700, fontSize: 12, flexShrink: 0 }}>
      {level}
    </div>
    <div style={{ flex: 1 }}>
      <div style={{ fontSize: 13, fontWeight: 600, color: '#201F1E' }}>{name}</div>
      <div style={{ display: 'flex', gap: 6, marginTop: 3 }}>
        {isDelegate && <span style={{ fontSize: 10, background: '#FFF4CE', color: '#905400', padding: '1px 7px', borderRadius: 8, fontWeight: 700 }}>DELEGATE</span>}
        {wasEscalated && <span style={{ fontSize: 10, background: '#FDE7E9', color: '#A4262C', padding: '1px 7px', borderRadius: 8, fontWeight: 700 }}>R10 ESCALATED</span>}
      </div>
    </div>
  </div>
);

// ── Main component ────────────────────────────────────────────────────────────

export default class ReviewQueue extends React.Component<IReviewQueueProps, IReviewQueueState> {
  private _service: DataverseService;
  private _router: ApprovalRoutingService;
  private _primaryKey: string | null = null;

  constructor(props: IReviewQueueProps) {
    super(props);
    this._service = props.dataverseService;
    this._router  = new ApprovalRoutingService(props.dataverseService);
    this.state = {
      activeStage: 'triage',
      records: [],
      loading: false,
      error: null,
      actionRecord: null,
      actionValues: {},
      poDate: undefined,
      saving: false,
      saveError: null,
      saveSuccess: null,
      preApproval: null,
      postSourcing: null,
      resolvedApprovers: null,
      loadingApprovers: false,
      pendingStageRows: [],
      loadingPending: false,
      systemUserId: null,
      managerApprovedIds: new Set<string>(),
      directReportIds: new Set<string>(),
      managerActionRecord: null,
      managerComments: '',
      managerSaving: false,
      managerError: null,
    };
  }

  public componentDidMount(): void {
    void this._load();
    void this._resolveUser();
  }

  public componentDidUpdate(prevProps: IReviewQueueProps): void {
    if (prevProps.activeRole !== this.props.activeRole) {
      this.setState({ managerActionRecord: null, managerComments: '', managerError: null, saveSuccess: null });
      void this._load();
      if (!this.state.systemUserId) {
        void this._resolveUser();
      } else if (this.props.activeRole === 'manager') {
        void this._service.getDirectReportUserIds(this.state.systemUserId).then(reports => {
          this.setState({
            directReportIds: new Set(reports.map(id => id.toLowerCase().replace(/[{}]/g, ''))),
          });
        }).catch(() => {});
        void this._loadManagerApprovedIds(this.state.systemUserId);
      }
    }
  }

  private async _resolveUser(): Promise<void> {
    try {
      const ctx = this.props.context;
      const aadId = ctx.pageContext?.aadInfo?.userId?.toString?.() ?? '';
      let sysUserId: string | null = null;
      if (aadId) {
        const user = await this._service.getCurrentSystemUser(aadId);
        sysUserId = user?.systemuserid ?? null;
      }
      if (!sysUserId && ctx.pageContext?.user?.email) {
        const user = await this._service.getSystemUserByEmail(ctx.pageContext.user.email);
        sysUserId = user?.systemuserid ?? null;
      }
      if (sysUserId) {
        let reports: string[] = [];
        if (this.props.activeRole === 'manager') {
          reports = await this._service.getDirectReportUserIds(sysUserId).catch(() => []);
        }
        this.setState({
          systemUserId: sysUserId,
          directReportIds: new Set(reports.map(id => id.toLowerCase().replace(/[{}]/g, ''))),
        });
        if (this.props.activeRole === 'manager') {
          void this._loadManagerApprovedIds(sysUserId);
        }
      }
    } catch { /* non-fatal */ }
  }

  private async _loadManagerApprovedIds(systemUserId?: string): Promise<void> {
    const userId = systemUserId ?? this.state.systemUserId;
    if (!userId) return;
    try {
      const ids = await this._service.getMyApprovedRequestIds(userId);
      this.setState({ managerApprovedIds: new Set<string>(ids) });
    } catch { /* non-fatal — Pending may show extra items */ }
  }

  // ── Data ──────────────────────────────────────────────────────────────────

  private async _load(): Promise<void> {
    this.setState({ loading: true, error: null, saveSuccess: null });
    try {
      const records = await this._service.getRawRequests();
      if (records.length && !this._primaryKey) {
        const guidFields = Object.keys(records[0]).filter(k =>
          !k.startsWith('@') && !k.startsWith('_') &&
          typeof records[0][k] === 'string' &&
          /^[0-9a-f]{8}-[0-9a-f]{4}-/.test(String(records[0][k]))
        );
        this._primaryKey = guidFields.find(k => k.endsWith('requestid')) ?? guidFields[0] ?? null;
      }
      this.setState({ records, loading: false });
    } catch (err: unknown) {
      this.setState({ error: err instanceof Error ? err.message : 'Load failed', loading: false });
    }
  }

  private _getId(r: ODataRecord): string | null {
    if (!this._primaryKey) return null;
    const v = r[this._primaryKey];
    return typeof v === 'string' ? v : null;
  }

  // ── Queue filtering ───────────────────────────────────────────────────────

  private _queueFor(stageDef: IStageDef): ODataRecord[] {
    return this.state.records.filter(r => {
      const label = String(r['sct_request_status@OData.Community.Display.V1.FormattedValue'] ?? '');
      const matchesStatus = stageDef.watchLabels.some(w => label.toLowerCase() === w.toLowerCase());
      if (!matchesStatus) return false;

      // FP&A and CFO both watch "In Review" but only show high-value / new-supplier requests.
      // Sourcing shows the remainder (standard, within-threshold requests).
      const amt    = parseFloat(String(r['sct_estimated_total_value'] ?? 0));
      const isNew  = r['sct_is_new_supplier'] === true || r['sct_is_new_supplier'] === 'Yes';
      const needsGate = (!isNaN(amt) && amt > CFO_THRESHOLD) || isNew;

      if (stageDef.id === 'sourcing') return !needsGate;   // standard only
      if (stageDef.id === 'fpa')      return needsGate;    // above threshold / new supplier
      if (stageDef.id === 'cfo')      return needsGate;    // same population as FP&A

      return true; // triage + receiving: no extra filter
    });
  }

  // ── Pending stage calculation ─────────────────────────────────────────────

  private _calculatePendingStage(logs: ODataRecord[]): IPendingStageRow[] {
    // Filter to approved logs only (sct_approvalstatus === 1)
    const approved = logs.filter(l => (l['sct_approvalstatus'] as number) === 1);

    if (approved.length === 0) {
      // No approvals recorded yet — request sits with the requester's manager
      return [
        { stage: 'Pending with Manager', role: "Requester's Line Manager" },
      ];
    }

    // Determine stage based on tiers approved in approval logs
    // Dataverse sct_tiername: 4 = Manager, 1 = Tier 1, 2 = Tier 2, 3 = Tier 3
    const tiersApproved = new Set(approved.map(l => l['sct_tiername'] as number));
    if (tiersApproved.has(3)) {
      return [{ stage: 'All Stages Approved', role: 'Awaiting PO Issuance' }];
    } else if (tiersApproved.has(2)) {
      return [{ stage: 'Pending with Tier 3 / CFO', role: 'Tier 3 / CFO Approver Group' }];
    } else if (tiersApproved.has(1)) {
      return [{ stage: 'Pending with Tier 2', role: 'Tier 2 AAD Approver Group' }];
    } else if (tiersApproved.has(4)) {
      return [{ stage: 'Pending with Tier 1', role: 'Tier 1 AAD Approver Group' }];
    } else {
      return [{ stage: 'Pending with Manager', role: "Requester's Line Manager" }];
    }
  }

  // ── Open action panel ─────────────────────────────────────────────────────

  private _openAction(record: ODataRecord, stageDef: IStageDef): void {
    const poDateStr = strVal(record, 'sct_po_issue_date');
    const statusLabel = String(
      record['sct_request_status@OData.Community.Display.V1.FormattedValue'] ?? ''
    ).toLowerCase();
    const isInReview = statusLabel === 'in review';

    this.setState({
      actionRecord: record,
      actionValues: { sct_cost_center: strVal(record, 'sct_cost_center'), sct_po_number: strVal(record, 'sct_po_number') },
      poDate: poDateStr ? edmToDate(poDateStr) : undefined,
      saveError: null,
      preApproval: null,
      postSourcing: null,
      resolvedApprovers: null,
      loadingApprovers: false,
      pendingStageRows: [],
      loadingPending: isInReview,
    }, () => {
      void this._loadRoutingData(record, stageDef);
      if (isInReview) void this._loadPendingStage(record);
    });
  }

  // ── Fetch approval logs and derive pending stage ───────────────────────────

  private async _loadPendingStage(record: ODataRecord): Promise<void> {
    const id = this._getId(record);
    if (!id) { this.setState({ loadingPending: false }); return; }
    try {
      const logs = await this._service.getApprovalLogs(id);
      this.setState({
        pendingStageRows: this._calculatePendingStage(Array.isArray(logs) ? logs : []),
        loadingPending: false,
      });
    } catch {
      this.setState({ loadingPending: false });
    }
  }

  // ── Load routing data for the opened action panel ─────────────────────────

  private async _loadRoutingData(record: ODataRecord, stageDef: IStageDef): Promise<void> {
    const estimatedValue = parseFloat(String(record['sct_estimated_total_value'] ?? 0));
    const isNewSupplier  = record['sct_is_new_supplier'] === true || record['sct_is_new_supplier'] === 'Yes';
    const siteId         = String(record['_sct_site_value'] ? record['_sct_site_value'] : record['_sct_siteid_value'] ?? '');

    // Always calculate deterministic routing (no async needed)
    const preApproval  = this._router.calculatePreApprovals(estimatedValue, isNewSupplier);
    const postSourcing = this._router.calculatePostSourcingRequirements(estimatedValue, isNewSupplier);

    this.setState({ preApproval, postSourcing });

    // Dynamic role resolution (async) — only for stages that show approver cards
    if (stageDef.id === 'triage' || stageDef.id === 'fpa' || stageDef.id === 'cfo') {
      const requesterGuid = String(record['_sct_requester_value'] ?? record['_sct_requestorid_value'] ?? '');
      if (siteId) {
        this.setState({ loadingApprovers: true });
        try {
          const resolvedApprovers = await this._router.resolveApprovers(siteId, requesterGuid);
          this.setState({ resolvedApprovers, loadingApprovers: false });
        } catch {
          this.setState({ loadingApprovers: false });
        }
      }
    }
  }

  // ── Save action ───────────────────────────────────────────────────────────

  private async _takeAction(stageDef: IStageDef, targetStatus: number): Promise<void> {
    const { actionRecord, actionValues, poDate, systemUserId } = this.state;
    if (!actionRecord) return;
    const id = this._getId(actionRecord);
    if (!id) return;

    this.setState({ saving: true, saveError: null });

    const payload: Record<string, unknown> = { sct_request_status: targetStatus };

    // Synchronize cr2a4_pendingtier: clear on termination states, preserve or set on In Review
    if (
      targetStatus === STATUS.PO_ISSUED ||
      targetStatus === STATUS.REJECTED ||
      targetStatus === STATUS.RETURNED ||
      targetStatus === STATUS.CLOSED
    ) {
      payload['cr2a4_pendingtier'] = null;
    } else if (targetStatus === STATUS.IN_REVIEW) {
      const existingPt = String(
        actionRecord['cr2a4_pendingtier'] || actionRecord['sct_pendingtier'] || ''
      ).trim();
      payload['cr2a4_pendingtier'] = existingPt || 'Manager';
    }

    // Triage: cost centre
    const cc = actionValues['sct_cost_center']?.trim();
    if (cc) payload['sct_cost_center'] = cc;

    // Sourcing: PO fields
    const po = actionValues['sct_po_number']?.trim();
    if (po) payload['sct_po_number'] = po;
    if (poDate) payload['sct_po_issue_date'] = dateToEdm(poDate);

    try {
      await this._service.updateRequest(id, payload);

      // Audit log the workbench stage action
      const currentStatusLabel = strVal(actionRecord, 'sct_request_status');
      const stageName = stageDef.title;
      await this._service.createAuditLog({
        sct_action: targetStatus === STATUS.PO_ISSUED ? 1 : targetStatus === STATUS.REJECTED ? 5 : targetStatus === STATUS.RETURNED ? 4 : 2,
        sct_description: `Stage ${stageDef.stage} (${stageName}) executed: ${stageDef.primaryAction}`,
        sct_oldvalue: `Status: ${currentStatusLabel}`,
        sct_newvalue: `Status: ${targetStatus}, Pending Tier: ${payload['cr2a4_pendingtier'] ?? 'null'}`,
        'sct_requestid@odata.bind': `/sct_requests(${id})`,
        ...(systemUserId ? { 'sct_actionby@odata.bind': `/systemusers(${systemUserId})` } : {}),
      }).catch(() => { /* non-fatal */ });

      this.setState({ saving: false, actionRecord: null, saveSuccess: 'Record updated successfully.' });
      void this._load();
    } catch (err: unknown) {
      this.setState({ saving: false, saveError: err instanceof Error ? err.message : 'Save failed' });
    }
  }

  // ── Helpers ───────────────────────────────────────────────────────────────

  private _estimatedValue(r: ODataRecord): number {
    return parseFloat(String(r['sct_estimated_total_value'] ?? 0));
  }

  private _isNewSupplier(r: ODataRecord): boolean {
    return r['sct_is_new_supplier'] === true || r['sct_is_new_supplier'] === 'Yes';
  }

  // ── Approver queue methods ─────────────────────────────────────────────────

  /** Records waiting for approval by the active approver role */
  private _approverQueue(): ODataRecord[] {
    const { records, systemUserId, directReportIds, managerApprovedIds } = this.state;
    const role = this.props.activeRole || 'manager';
    const myId = systemUserId ? systemUserId.toLowerCase().replace(/[{}]/g, '') : '';
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctxAny = this.props.context as any;
    const aadId: string = (ctxAny?.pageContext?.aadInfo?.userId?.toString?.() || '').toLowerCase().replace(/[{}]/g, '');
    const isDevOrWorkbench = typeof window !== 'undefined' && (
      window.location.href.indexOf('workbench') !== -1 ||
      window.location.href.indexOf('devrishabhsoft') !== -1 ||
      window.location.hostname.indexOf('localhost') !== -1 ||
      window.location.hostname.indexOf('127.0.0.1') !== -1
    );

    return records.filter(r => {
      const rawReq = String(r['_sct_requester_value'] ?? r['_createdby_value'] ?? r['sct_requester'] ?? '');
      const reqId  = rawReq.toLowerCase().replace(/[{}]/g, '');
      const createdById = r['_createdby_value'] ? String(r['_createdby_value']).toLowerCase().replace(/[{}]/g, '') : '';

      // 1. Global Segregation of Duties (SoD) Guard:
      const allowSelfApproval = Boolean(this.props.soloTestMode);
      if (!allowSelfApproval && myId && (reqId === myId || createdById === myId || String(r['_sct_requester_value'] || '').toLowerCase().replace(/[{}]/g, '') === myId)) {
        return false;
      }

      const rawStatus = typeof r['sct_request_status'] === 'number'
        ? r['sct_request_status']
        : parseInt(String(r['sct_request_status'] ?? ''), 10);
      const statusLabel = strVal(r, 'sct_request_status').toLowerCase();
      const isReview = rawStatus === 854000003 || rawStatus === 854000002 || statusLabel === 'in review' || statusLabel.indexOf('in review') !== -1;

      let ptRaw = String(
        r['cr2a4_pendingtier'] ||
        r['cr2a4_PendingTier'] ||
        r['cr2a4_pending_tier'] ||
        r['sct_pendingtier'] ||
        r['pendingtier'] ||
        ''
      ).trim();
      if (!ptRaw && statusLabel.indexOf('tier') !== -1) {
        const m = statusLabel.match(/tier\s*([123])/i);
        if (m) ptRaw = `Tier ${m[1]}`;
      }
      const pt = ptRaw.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();

      const normRole = (role || 'manager').toLowerCase().trim();
      const isManager = normRole === 'manager' || normRole.indexOf('manager') !== -1;
      const isTier1 = normRole === 'approver_tier1' || normRole.indexOf('tier 1') !== -1 || normRole.indexOf('tier1') !== -1;
      const isTier2 = normRole === 'approver_tier2' || normRole.indexOf('tier 2') !== -1 || normRole.indexOf('tier2') !== -1;
      const isTier3 = normRole === 'approver_tier3' || normRole.indexOf('tier 3') !== -1 || normRole.indexOf('tier3') !== -1;

      if (isManager) {
        const isPendingStatus = rawStatus === 854000002 || rawStatus === 854000003 || rawStatus === 854000001 || statusLabel === 'in review' || statusLabel === 'submitted';
        if (!isPendingStatus) return false;

        const isPendingWithManager = pt === 'manager' || ((rawStatus === 854000002 || rawStatus === 854000001) && !pt) || (statusLabel === 'submitted' && !pt);
        if (!isPendingWithManager) return false;

        const pk = this._getId(r);
        if (pk && managerApprovedIds.has(pk)) return false;

        const rawMgr = r['_sct_manager_value'] ?? r['_cr2a4_manager_value'] ?? r['_manager_value'] ?? r['sct_manager'] ?? r['cr2a4_manager'] ?? '';
        const mgrId = rawMgr ? String(rawMgr).toLowerCase().replace(/[{}]/g, '') : '';
        const mgrFormatted = (strVal(r, '_cr2a4_manager_value') || strVal(r, '_sct_manager_value') || strVal(r, '_manager_value') || strVal(r, 'cr2a4_manager') || strVal(r, 'sct_manager') || strVal(r, 'manager') || '').toLowerCase();
        const userDisplayName = (this.props.context.pageContext.user.displayName || '').toLowerCase().trim();
        const userEmail = (this.props.context.pageContext.user.email || '').toLowerCase().trim();
        const isMgrByName = Boolean(userDisplayName) && (
          mgrFormatted.indexOf(userDisplayName) !== -1 ||
          (Boolean(userEmail) && mgrFormatted.indexOf(userEmail) !== -1)
        );
        const isDirectReport = Boolean(directReportIds && (directReportIds.has(reqId) || (createdById && directReportIds.has(createdById))));
        const isManagerOfReq = (mgrId && (mgrId === myId || (aadId && mgrId === aadId.toLowerCase())))
          || isMgrByName
          || isDirectReport
          || (!mgrId && isPendingWithManager);
        if (!isManagerOfReq) return false;
        return true;
      }

      if (isTier1) {
        if (!isReview) return false;
        return pt === 'tier1' || pt.indexOf('tier1') !== -1 || statusLabel.indexOf('tier 1') !== -1 || statusLabel.indexOf('tier1') !== -1 || pt === '1';
      }

      if (isTier2) {
        if (!isReview) return false;
        return pt === 'tier2' || pt.indexOf('tier2') !== -1 || statusLabel.indexOf('tier 2') !== -1 || statusLabel.indexOf('tier2') !== -1 || pt === '2';
      }

      if (isTier3) {
        if (!isReview) return false;
        return pt === 'tier3' || pt.indexOf('tier3') !== -1 || statusLabel.indexOf('tier 3') !== -1 || statusLabel.indexOf('tier3') !== -1 || pt === '3';
      }

      return false;
    });
  }

  /** Take action (approve / reject / return) on a record from the approver queue */
  private async _handleManagerAction(decision: 'approve' | 'reject' | 'return'): Promise<void> {
    const { managerActionRecord, managerComments } = this.state;
    if (!managerActionRecord) return;

    const id = this._getId(managerActionRecord);
    if (!id) return;

    this.setState({ managerSaving: true, managerError: null });
    try {
      let graphClient: import('@microsoft/sp-http').MSGraphClientV3 | undefined;
      try {
        graphClient = await this.props.context.msGraphClientFactory.getClient('3');
      } catch { /* optional */ }

      const role = this.props.activeRole || 'manager';
      const rawReq = managerActionRecord
        ? String(managerActionRecord['_sct_requester_value'] ?? managerActionRecord['sct_requester'] ?? managerActionRecord['_createdby_value'] ?? '')
        : '';
      // Steps 1, 2, 3: write approval log + patch request status + audit log + notification emails
      const result = await handleApprovalAction({
        requestId:           id,
        role:                role,
        action:              decision,
        comments:            managerComments,
        service:             this._service,
        currentSystemUserId: this.state.systemUserId,
        requesterId:         rawReq,
        graphClient,
        webPartUrl:          this.props.context.pageContext.web.absoluteUrl,
        approverDisplayName: this.props.context.pageContext.user.displayName,
        currentUserEmail:    this.props.context.pageContext.user.email || undefined,
        soloTestMode:        Boolean(this.props.soloTestMode),
      });

      const successMsg = result.emailDispatched
        ? 'Action saved and notification email sent.'
        : 'Action saved successfully.';

      this.setState({ managerActionRecord: null, managerComments: '', managerSaving: false, saveSuccess: successMsg });
      // Wait 500ms for Dataverse to commit the PATCH before re-fetching
      await new Promise<void>(resolve => setTimeout(resolve, 500));
      void this._load();
      // Refresh approved IDs if manager acted so the record immediately disappears from Pending queue
      if (role === 'manager') {
        void this._loadManagerApprovedIds();
      }
    } catch (err: unknown) {
      this.setState({
        managerSaving: false,
        managerError: err instanceof Error ? err.message : 'Action failed. Please try again.',
      });
    }
  }

  private _renderManagerQueue(): React.ReactElement {
    const { loading, systemUserId, managerActionRecord, managerComments, managerSaving, managerError } = this.state;
    const role = this.props.activeRole || 'manager';
    const queue = this._approverQueue();

    const queueTitle = role === 'manager' ? 'Manager Review Queue'
      : role === 'approver_tier1' ? 'Tier 1 Review Queue'
      : role === 'approver_tier2' ? 'Tier 2 Review Queue'
      : role === 'approver_tier3' ? 'Tier 3 Review Queue'
      : 'Approver Review Queue';

    const queueDesc = role === 'manager' ? 'Requests assigned to you as line manager — pending your approval'
      : role === 'approver_tier1' ? 'Requests awaiting Tier 1 approval — status: In Review'
      : role === 'approver_tier2' ? 'Requests awaiting Tier 2 approval — status: In Review'
      : role === 'approver_tier3' ? 'Requests awaiting Tier 3 / Final approval — status: In Review'
      : 'Requests awaiting review';

    const headerIcon = role === 'manager' ? 'People' : 'WorkFlow';
    const themeColor = role === 'manager' ? '#0078D4' : role === 'approver_tier1' ? '#00897B' : role === 'approver_tier2' ? '#0277BD' : '#8764B8';

    const columns: IColumn[] = [
      {
        key: 'ref', name: 'Reference', minWidth: 110, maxWidth: 140, isResizable: true,
        onRender: (r: ODataRecord) => (
          <span style={{ fontWeight: 700, color: '#0078D4', fontFamily: 'monospace', fontSize: 12 }}>
            {strVal(r, 'sct_sct_reference') || '—'}
          </span>
        ),
      },
      {
        key: 'requester', name: 'Requester', minWidth: 120, maxWidth: 180, isResizable: true,
        onRender: (r: ODataRecord) => <span style={{ fontSize: 13 }}>{strVal(r, '_sct_requester_value') || '—'}</span>,
      },
      {
        key: 'site', name: 'Site', minWidth: 100, maxWidth: 150, isResizable: true,
        onRender: (r: ODataRecord) => <span style={{ fontSize: 13 }}>{strVal(r, '_sct_site_value') || '—'}</span>,
      },
      {
        key: 'type', name: 'Type', minWidth: 90, maxWidth: 130, isResizable: true,
        onRender: (r: ODataRecord) => <span style={{ fontSize: 13 }}>{strVal(r, 'sct_purchase_type') || '—'}</span>,
      },
      {
        key: 'value', name: 'Est. Value', minWidth: 100, maxWidth: 130, isResizable: true,
        onRender: (r: ODataRecord) => {
          const amt = this._estimatedValue(r);
          if (isNaN(amt)) return <span style={{ color: '#C8C6C4' }}>—</span>;
          return (
            <span style={{ fontWeight: 700, color: '#107C10', fontFamily: 'monospace', fontSize: 13 }}>
              {formatCurrency(amt)}
            </span>
          );
        },
      },
      {
        key: 'submitted', name: 'Submitted', minWidth: 100, maxWidth: 130, isResizable: true,
        onRender: (r: ODataRecord) => {
          const d = strVal(r, 'sct_submitted_date') || strVal(r, 'createdon');
          return <span style={{ fontSize: 12, color: '#605E5C' }}>{d ? formatDate(d) : '—'}</span>;
        },
      },
      {
        key: '_act', name: '', minWidth: 95, maxWidth: 100,
        onRender: (r: ODataRecord) => (
          <button
            title="Review and take action"
            onClick={e => { e.stopPropagation(); this.setState({ managerActionRecord: r, managerComments: '', managerError: null }); }}
            style={{
              display: 'inline-flex', alignItems: 'center', gap: 5,
              background: 'linear-gradient(135deg, #0078D4 0%, #005A9E 100%)',
              color: '#fff', border: 'none', borderRadius: 20,
              padding: '5px 13px', fontSize: 12, fontWeight: 600,
              cursor: 'pointer', letterSpacing: '0.2px',
              boxShadow: '0 1px 4px rgba(0,120,212,0.35)',
              whiteSpace: 'nowrap',
            }}
          >
            <Icon iconName="WorkFlow" style={{ fontSize: 11 }} />
            Review
          </button>
        ),
      },
    ];

    return (
      <div>
        {/* Header strip */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16, padding: '12px 16px', background: '#EBF3FB', borderRadius: 10 }}>
          <div style={{ width: 44, height: 44, borderRadius: 10, background: themeColor, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
            <Icon iconName={headerIcon} style={{ fontSize: 22, color: '#fff' }} />
          </div>
          <div style={{ flex: 1 }}>
            <Text variant="mediumPlus" style={{ color: '#201F1E', fontWeight: 700, display: 'block' }}>
              {queueTitle}
            </Text>
            <Text variant="small" style={{ color: '#605E5C' }}>
              {queueDesc}
            </Text>
          </div>
          <span style={{ background: themeColor, color: '#fff', fontSize: 18, fontWeight: 700, padding: '4px 14px', borderRadius: 12 }}>
            {queue.length}
          </span>
        </div>

        {!systemUserId && !loading && (
          <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 6, marginBottom: 12 } }}>
            Resolving your Dataverse user account…
          </MessageBar>
        )}

        {queue.length === 0 && !loading
          ? <EmptyState message="No pending reviews" subMessage="All requests assigned to you are up to date." iconName={headerIcon} />
          : <DetailsList items={queue} columns={columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              styles={{ root: { minWidth: 0 } }}
              ariaLabelForGrid={`${queueTitle} grid`} />
        }

        {/* Approver review dialog */}
        <TakeActionModal
          isOpen={managerActionRecord !== null}
          record={managerActionRecord}
          service={this._service}
          activeRole={this.props.activeRole || 'manager'}
          currentSystemUserId={this.state.systemUserId}
          context={this.props.context}
          onDismiss={() => this.setState({ managerActionRecord: null })}
          onSuccess={(msg: string) => {
            this.setState({
              managerActionRecord: null,
              saveSuccess: msg,
            });
            setTimeout(() => {
              void this._load();
              if ((this.props.activeRole || 'manager') === 'manager') {
                void this._loadManagerApprovedIds();
              }
            }, 500);
          }}
        />
      </div>
    );
  }
  // ── Render: Stage tabs ─────────────────────────────────────────────────────

  private _renderStageTabs(): React.ReactElement {
    const { activeStage } = this.state;
    return (
      <div style={{ display: 'flex', gap: 4, marginBottom: 20, flexWrap: 'wrap' }}>
        {STAGES.map(s => {
          const count    = this._queueFor(s).length;
          const isActive = activeStage === s.id;
          return (
            <button key={s.id}
              onClick={() => this.setState({ activeStage: s.id, saveSuccess: null })}
              style={{
                display: 'flex', alignItems: 'center', gap: 8,
                padding: '8px 16px', borderRadius: 8,
                border: isActive ? `2px solid ${s.color}` : '2px solid #EDEBE9',
                background: isActive ? s.bg : '#FFFFFF',
                color: isActive ? s.color : '#605E5C',
                fontWeight: isActive ? 700 : 500, fontSize: 13, cursor: 'pointer',
              }}
              aria-selected={isActive}
            >
              <Icon iconName={s.iconName} style={{ fontSize: 15, color: isActive ? s.color : '#A19F9D' }} />
              <span>Stage {s.stage}: {s.title.split('—')[0].trim()}</span>
              {!this.state.loading && count > 0 && (
                <span style={{ background: s.color, color: '#fff', fontSize: 11, fontWeight: 700, padding: '1px 7px', borderRadius: 10 }}>
                  {count}
                </span>
              )}
            </button>
          );
        })}
      </div>
    );
  }

  // ── Render: Workbench table ────────────────────────────────────────────────

  private _renderWorkbench(stageDef: IStageDef): React.ReactElement {
    const queue = this._queueFor(stageDef);
    const columns: IColumn[] = [
      {
        key: 'ref', name: 'Reference', minWidth: 110, maxWidth: 140, isResizable: true,
        onRender: (r: ODataRecord) => (
          <span style={{ fontWeight: 700, color: '#0078D4', fontFamily: 'monospace', fontSize: 12 }}>
            {strVal(r, 'sct_sct_reference') || '—'}
          </span>
        ),
      },
      {
        key: 'requester', name: 'Requester', minWidth: 120, maxWidth: 180, isResizable: true,
        onRender: (r: ODataRecord) => <span style={{ fontSize: 13 }}>{strVal(r, '_sct_requester_value') || '—'}</span>,
      },
      {
        key: 'site', name: 'Site', minWidth: 100, maxWidth: 160, isResizable: true,
        onRender: (r: ODataRecord) => <span style={{ fontSize: 13 }}>{strVal(r, '_sct_site_value') || '—'}</span>,
      },
      {
        key: 'type', name: 'Type', minWidth: 90, maxWidth: 130, isResizable: true,
        onRender: (r: ODataRecord) => <span style={{ fontSize: 13 }}>{strVal(r, 'sct_purchase_type') || '—'}</span>,
      },
      {
        key: 'value', name: 'Est. Value', minWidth: 100, maxWidth: 130, isResizable: true,
        onRender: (r: ODataRecord) => {
          const amt = this._estimatedValue(r);
          if (isNaN(amt)) return <span style={{ color: '#C8C6C4' }}>—</span>;
          const flag = amt > CFO_THRESHOLD;
          return (
            <span style={{ fontWeight: 700, color: flag ? '#A4262C' : '#107C10', fontFamily: 'monospace', fontSize: 13 }}>
              {formatCurrency(amt)}{flag && <span title="Exceeds $10k — CFO route"> ⚑</span>}
            </span>
          );
        },
      },
      {
        key: 'supplier', name: 'Supplier', minWidth: 110, maxWidth: 160, isResizable: true,
        onRender: (r: ODataRecord) =>
          this._isNewSupplier(r)
            ? <StatusPill value="New Supplier" color="#905400" bg="#FFF4CE" />
            : <span style={{ fontSize: 13 }}>{strVal(r, '_sct_suggested_supplier_value') || '—'}</span>,
      },
      {
        key: 'submitted', name: 'Submitted', minWidth: 100, maxWidth: 130, isResizable: true,
        onRender: (r: ODataRecord) => {
          const d = strVal(r, 'sct_submitted_date') || strVal(r, 'createdon');
          return <span style={{ fontSize: 12, color: '#605E5C' }}>{d ? formatDate(d) : '—'}</span>;
        },
      },
      {
        key: '_act', name: '', minWidth: 80, maxWidth: 80,
        onRender: (r: ODataRecord) => (
          <PrimaryButton text="Action"
            onClick={e => { e.stopPropagation(); this._openAction(r, stageDef); }}
            styles={{ root: { borderRadius: 6, height: 28, fontSize: 12, padding: '0 12px' } }}
          />
        ),
      },
    ];

    return (
      <div>
        {/* Stage strip */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 16, padding: '12px 16px', background: stageDef.bg, borderRadius: 10 }}>
          <div style={{ width: 44, height: 44, borderRadius: 10, background: stageDef.color, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
            <Icon iconName={stageDef.iconName} style={{ fontSize: 22, color: '#fff' }} />
          </div>
          <div style={{ flex: 1 }}>
            <Text variant="mediumPlus" style={{ color: '#201F1E', fontWeight: 700, display: 'block' }}>
              Stage {stageDef.stage}: {stageDef.title}
            </Text>
            <Text variant="small" style={{ color: '#605E5C' }}>{stageDef.desc}</Text>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <span style={{ background: stageDef.color, color: '#fff', fontSize: 18, fontWeight: 700, padding: '4px 14px', borderRadius: 12 }}>
              {queue.length}
            </span>
            <Text variant="small" style={{ color: stageDef.color, fontWeight: 600 }}>{queue.length === 1 ? 'item' : 'items'}</Text>
          </div>
        </div>

        <div style={{ fontSize: 11, color: '#A19F9D', marginBottom: 10 }}>
          Monitors status: {stageDef.watchLabels.map(l => `"${l}"`).join(', ')}
          {(stageDef.id === 'fpa' || stageDef.id === 'cfo') &&
            <span style={{ color: '#605E5C' }}>{' · '}Filtered to requests &gt; ${CFO_THRESHOLD.toLocaleString()} or new supplier</span>
          }
        </div>

        {queue.length === 0
          ? <EmptyState message="Queue is empty"
              subMessage={`No requests at the ${stageDef.watchLabels.join(' / ')} stage.`}
              iconName={stageDef.iconName} />
          : <DetailsList items={queue} columns={columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              styles={{ root: { minWidth: 0 } }}
              ariaLabelForGrid={`${stageDef.title} queue`} />
        }
      </div>
    );
  }

  // ── Render: Action panel ───────────────────────────────────────────────────

  // ── Render: Pending Stage Tracker ────────────────────────────────────────

  private _renderPendingTracker(): React.ReactElement | null {
    const { loadingPending, pendingStageRows, actionRecord } = this.state;
    if (!actionRecord) return null;

    const statusLabel = String(
      actionRecord['sct_request_status@OData.Community.Display.V1.FormattedValue'] ?? ''
    ).toLowerCase();
    if (statusLabel !== 'in review') return null;

    return (
      <div style={{ marginTop: 16, marginBottom: 4 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: '#0078D4', marginBottom: 10 }}>
          Pending Stage Tracker
        </div>
        {loadingPending ? (
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 0' }}>
            <Spinner size={SpinnerSize.small} />
            <span style={{ fontSize: 12, color: '#605E5C' }}>Calculating current stage…</span>
          </div>
        ) : (
          <table style={{
            width: '100%', borderCollapse: 'collapse',
            border: '1px solid #EDEBE9', borderRadius: 8, overflow: 'hidden',
            fontSize: 13,
          }}>
            <thead>
              <tr style={{ background: '#EBF3FB' }}>
                <th style={{
                  padding: '8px 14px', textAlign: 'left',
                  fontSize: 11, fontWeight: 700, color: '#0078D4',
                  textTransform: 'uppercase', letterSpacing: '0.5px',
                  borderBottom: '1px solid #EDEBE9',
                }}>
                  Pending Stage
                </th>
                <th style={{
                  padding: '8px 14px', textAlign: 'left',
                  fontSize: 11, fontWeight: 700, color: '#0078D4',
                  textTransform: 'uppercase', letterSpacing: '0.5px',
                  borderBottom: '1px solid #EDEBE9',
                }}>
                  Expected Role / Group
                </th>
              </tr>
            </thead>
            <tbody>
              {pendingStageRows.map((row, i) => (
                <tr key={i} style={{ background: i % 2 === 0 ? '#FFFFFF' : '#F8FFFE' }}>
                  <td style={{ padding: '10px 14px', color: '#201F1E', fontWeight: 600, borderBottom: '1px solid #EDEBE9' }}>
                    <span style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <Icon iconName="Clock" style={{ fontSize: 13, color: '#0078D4', flexShrink: 0 }} />
                      {row.stage}
                    </span>
                  </td>
                  <td style={{ padding: '10px 14px', color: '#605E5C', borderBottom: '1px solid #EDEBE9' }}>
                    {row.role}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    );
  }

  // ── Render: Action panel ───────────────────────────────────────────────────

  private _renderActionPanel(stageDef: IStageDef): React.ReactElement | null {
    const { actionRecord, actionValues, poDate, saving, saveError,
            preApproval, postSourcing, resolvedApprovers, loadingApprovers } = this.state;
    if (!actionRecord) return null;

    const ref    = strVal(actionRecord, 'sct_sct_reference');
    const status = strVal(actionRecord, 'sct_request_status');
    const site   = strVal(actionRecord, '_sct_site_value');
    const amt    = this._estimatedValue(actionRecord);
    const isNew  = this._isNewSupplier(actionRecord);
    const needsCFO = amt > CFO_THRESHOLD || isNew;

    const setVal = (field: string) =>
      (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) =>
        this.setState(s => ({ actionValues: { ...s.actionValues, [field]: v ?? '' } }));

    const tfStyles = { root: { marginBottom: 14 }, fieldGroup: { borderRadius: 6 } };
    const close    = (): void => this.setState({ actionRecord: null, saveError: null });

    const sectionHeader = (label: string, color: string): React.ReactElement => (
      <div style={{ fontSize: 13, fontWeight: 700, color, marginTop: 16, marginBottom: 10 }}>{label}</div>
    );

    return (
      <Panel isOpen onDismiss={close}
        type={PanelType.medium}
        headerText={ref ? `${stageDef.title} · ${ref}` : stageDef.title}
        closeButtonAriaLabel="Close"
        styles={{ content: { paddingBottom: 88 } }}
        isFooterAtBottom
        onRenderFooterContent={() => (
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', padding: '0 24px 24px' }}>
            <PrimaryButton
              text={saving ? 'Saving…' : stageDef.primaryAction}
              disabled={saving}
              onClick={() => void this._takeAction(stageDef, stageDef.primaryStatus)}
              styles={{ root: { borderRadius: 6 } }}
            />
            {stageDef.altAction && stageDef.altStatus !== undefined && (
              <DefaultButton
                text={stageDef.altAction}
                disabled={saving}
                onClick={() => void this._takeAction(stageDef, stageDef.altStatus!)}
                styles={{ root: {
                  borderRadius: 6,
                  color: stageDef.altAction === 'Reject' ? '#A4262C' : '#905400',
                  borderColor: stageDef.altAction === 'Reject' ? '#A4262C' : '#C7A333',
                }}}
              />
            )}
            <DefaultButton text="Cancel" onClick={close} disabled={saving} styles={{ root: { borderRadius: 6 } }} />
          </div>
        )}
      >
        {saveError && (
          <MessageBar messageBarType={MessageBarType.error} style={{ marginBottom: 14 }}>
            {saveError}
          </MessageBar>
        )}

        {/* ── Read-only context strip ─────────────────────────────────── */}
        <div style={{ background: '#F8F8F7', borderRadius: 8, padding: '12px 14px', marginBottom: 16 }}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px 16px' }}>
            {[
              ['Reference',      ref],
              ['Status',         status],
              ['Site',           site],
              ['Estimated Value', isNaN(amt) ? '—' : formatCurrency(amt)],
              ['Requester',      strVal(actionRecord, '_sct_requester_value')],
              ['Purchase Type',  strVal(actionRecord, 'sct_purchase_type')],
              ['Urgency',        strVal(actionRecord, 'sct_urgency')],
              ['New Supplier',   isNew ? 'Yes ⚑' : 'No'],
            ].map(([label, value]) => (
              <div key={label as string} style={{ marginBottom: 6 }}>
                <div style={{ fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 2 }}>{label}</div>
                <div style={{ fontSize: 13, color: '#201F1E', fontWeight: 500 }}>{value || '—'}</div>
              </div>
            ))}
          </div>
          {strVal(actionRecord, 'sct_description') && (
            <div style={{ marginTop: 10, paddingTop: 10, borderTop: '1px solid #EDEBE9' }}>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', marginBottom: 4 }}>Description</div>
              <div style={{ fontSize: 12, color: '#605E5C', lineHeight: '1.5', maxHeight: 72, overflowY: 'auto' }}>
                {strVal(actionRecord, 'sct_description')}
              </div>
            </div>
          )}
        </div>

        {/* ── Pre-Approval Requirements (Triage) ──────────────────────── */}
        {(stageDef.id === 'triage') && preApproval && (
          <>
            {sectionHeader('Pre-Approval Requirements', stageDef.color)}
            <MessageBar
              messageBarType={preApproval.levels.length > 1 ? MessageBarType.warning : MessageBarType.info}
              styles={{ root: { borderRadius: 6, marginBottom: 10 } }}
            >
              <strong>Required: {preApproval.levels.join(' + ')}</strong> — {preApproval.description}
            </MessageBar>
          </>
        )}

        {/* ── Stage 3: Cost Centre ─────────────────────────────────────── */}
        {stageDef.id === 'triage' && (
          <>
            {sectionHeader('Triage Assignment', stageDef.color)}
            <TextField label="Cost Centre" value={actionValues['sct_cost_center'] ?? ''}
              onChange={setVal('sct_cost_center')} placeholder="e.g. CC-1042"
              styles={tfStyles} />
          </>
        )}

        {/* ── Stage 4: PO Details (DatePicker for PO Issue Date) ───────── */}
        {stageDef.id === 'sourcing' && (
          <>
            {sectionHeader('PO Details', stageDef.color)}
            <TextField label="D365 PO Number" value={actionValues['sct_po_number'] ?? ''}
              onChange={setVal('sct_po_number')} placeholder="e.g. PO-2026-00123"
              styles={tfStyles} />
            <DatePicker
              label="PO Issue Date"
              value={poDate}
              onSelectDate={d => this.setState({ poDate: d ?? undefined })}
              firstDayOfWeek={DayOfWeek.Monday}
              formatDate={d => d ? dateToEdm(d) : ''}
              placeholder="Select PO issue date…"
              styles={{ root: { marginBottom: 14 } }}
              allowTextInput
            />

            {/* Post-Sourcing Gate */}
            {postSourcing && (
              <MessageBar
                messageBarType={postSourcing.needsFPAReview ? MessageBarType.warning : MessageBarType.info}
                styles={{ root: { borderRadius: 6, marginBottom: 10 } }}
              >
                <strong>{postSourcing.needsFPAReview ? 'FP&A + CFO approval required' : 'Direct PO authority'}</strong>
                {' — '}{postSourcing.description}
              </MessageBar>
            )}
            {postSourcing?.needsLegalReview && (
              <MessageBar messageBarType={MessageBarType.severeWarning} styles={{ root: { borderRadius: 6 } }}>
                ⚖ Legal Review required — value exceeds $50,000.
              </MessageBar>
            )}
          </>
        )}

        {/* ── Stage 5a: FP&A ──────────────────────────────────────────── */}
        {stageDef.id === 'fpa' && (
          <>
            {sectionHeader('FP&A Decision', stageDef.color)}
            {needsCFO
              ? <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 6, marginBottom: 10 } }}>
                  CFO approval is required — {isNaN(amt) ? '' : formatCurrency(amt)}{isNew ? ' / new supplier' : ''}.
                  Use <strong>Escalate to CFO</strong> unless you hold delegated authority.
                </MessageBar>
              : <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 6, marginBottom: 10 } }}>
                  Within FP&A direct approval authority. You may approve or escalate.
                </MessageBar>
            }
          </>
        )}

        {/* ── Stage 5b: CFO ───────────────────────────────────────────── */}
        {stageDef.id === 'cfo' && (
          <>
            {sectionHeader('CFO Decision', stageDef.color)}
            <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 6, marginBottom: 10 } }}>
              Approving sets status to "Approved". Rejecting sets status to "Rejected".
            </MessageBar>
          </>
        )}

        {/* ── Stage 6: Receiving ───────────────────────────────────────── */}
        {stageDef.id === 'receiving' && (
          <>
            {sectionHeader('Goods/Services Receipt', stageDef.color)}
            <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 6 } }}>
              Confirming receipt closes this request in the Spend Control Tower.
            </MessageBar>
          </>
        )}

        {/* ── Pending Stage Tracker (In Review records) ──────────────── */}
        {this._renderPendingTracker()}

        {/* ── Resolved Approvers (Triage / FP&A / CFO) ───────────────── */}
        {(stageDef.id === 'triage' || stageDef.id === 'fpa' || stageDef.id === 'cfo') && (
          <>
            {sectionHeader('Resolved Approvers', '#605E5C')}
            {loadingApprovers && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '8px 0' }}>
                <Spinner size={SpinnerSize.small} />
                <Text variant="small" style={{ color: '#605E5C' }}>Resolving approvers from sct_role…</Text>
              </div>
            )}
            {!loadingApprovers && resolvedApprovers && (
              <>
                {resolvedApprovers.approvers.length === 0 && (
                  <Text variant="small" style={{ color: '#A19F9D', display: 'block', marginBottom: 8 }}>
                    No approvers resolved — verify sct_role records for this site in Dataverse.
                  </Text>
                )}
                {resolvedApprovers.approvers.map(a => (
                  <ApproverCard key={a.level + a.systemUserId}
                    level={a.level} name={a.name}
                    isDelegate={a.isDelegate} wasEscalated={a.wasEscalated} />
                ))}
                {resolvedApprovers.escalations.map((e, i) => (
                  <MessageBar key={i} messageBarType={MessageBarType.warning}
                    styles={{ root: { borderRadius: 6, marginBottom: 6 } }}>
                    {e}
                  </MessageBar>
                ))}
                {resolvedApprovers.warnings.map((w, i) => (
                  <MessageBar key={i} messageBarType={MessageBarType.info}
                    styles={{ root: { borderRadius: 6, marginBottom: 6 } }}>
                    {w}
                  </MessageBar>
                ))}
              </>
            )}
          </>
        )}
      </Panel>
    );
  }

  // ── Main render ───────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const { activeStage, loading, error, saveSuccess } = this.state;
    const role = (this.props.activeRole || 'manager').toLowerCase().trim();
    const isApproverRole = (
      role === 'manager' || role.indexOf('manager') !== -1 ||
      role.indexOf('tier') !== -1 || role.indexOf('approver') !== -1
    );
    const stageDef = STAGES.find(s => s.id === activeStage)!;

    const queueTitle = (role === 'manager' || role.indexOf('manager') !== -1) ? 'Manager Review Queue'
      : (role === 'approver_tier1' || role.indexOf('tier 1') !== -1 || role.indexOf('tier1') !== -1) ? 'Tier 1 Review Queue'
      : (role === 'approver_tier2' || role.indexOf('tier 2') !== -1 || role.indexOf('tier2') !== -1) ? 'Tier 2 Review Queue'
      : (role === 'approver_tier3' || role.indexOf('tier 3') !== -1 || role.indexOf('tier3') !== -1) ? 'Tier 3 Review Queue'
      : 'Spend Control Tower';

    const queueSubtitle = (role === 'manager' || role.indexOf('manager') !== -1)
      ? 'Requests assigned to you as line manager, pending your review'
      : (role === 'approver_tier1' || role.indexOf('tier 1') !== -1 || role.indexOf('tier1') !== -1)
      ? 'Requests awaiting Tier 1 approval — status: In Review'
      : (role === 'approver_tier2' || role.indexOf('tier 2') !== -1 || role.indexOf('tier2') !== -1)
      ? 'Requests awaiting Tier 2 approval — status: In Review'
      : (role === 'approver_tier3' || role.indexOf('tier 3') !== -1 || role.indexOf('tier3') !== -1)
      ? 'Requests awaiting Tier 3 / Final approval — status: In Review'
      : 'Stage workbenches · Schedule A approval routing engine';

    const headerIcon = role === 'manager' ? 'People' : isApproverRole ? 'WorkFlow' : 'Checklist';

    return (
      <div>
        {/* Header */}
        <div className={styles.card} style={{ marginBottom: 20 }}>
          <div className={styles.cardBody} style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
            <div style={{
              width: 56, height: 56, borderRadius: 12,
              background: isApproverRole ? 'linear-gradient(135deg, #0078D4, #1BA1E2)' : 'linear-gradient(135deg, #8764B8, #0078D4)',
              display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0
            }}>
              <Icon iconName={headerIcon} style={{ fontSize: 26, color: '#ffffff' }} />
            </div>
            <div style={{ flex: 1 }}>
              <Text variant="xLarge" style={{ color: '#201F1E', fontWeight: 700, display: 'block' }}>
                {queueTitle}
              </Text>
              <Text variant="medium" style={{ color: '#605E5C' }}>
                {queueSubtitle}
              </Text>
            </div>
            <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }}
              onClick={() => void this._load()} disabled={loading}
              styles={{ root: { borderRadius: 6, height: 32 } }} />
          </div>
        </div>

        {saveSuccess && (
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}
            onDismiss={() => this.setState({ saveSuccess: null })}
            styles={{ root: { borderRadius: 6, marginBottom: 12 } }}>
            {saveSuccess}
          </MessageBar>
        )}
        {error && (
          <ErrorBanner message={error} onRetry={() => void this._load()}
            onDismiss={() => this.setState({ error: null })} />
        )}

        <div className={styles.card}>
          <div className={styles.cardBody}>
            {loading
              ? <LoadingSpinner label="Loading queue from Dataverse…" />
              : isApproverRole
                ? this._renderManagerQueue()
                : (
                  <>
                    {this._renderStageTabs()}
                    {this._renderWorkbench(stageDef)}
                  </>
                )
            }
          </div>
        </div>

        {!isApproverRole && this._renderActionPanel(stageDef)}
      </div>
    );
  }
}
