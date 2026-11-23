import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { TextField } from '@fluentui/react/lib/TextField';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { sendManagerApprovalEmail, getManager } from '../../services/GraphMailService';

import { DataverseService, ODataRecord, strVal } from '../../services/DataverseService';
import { LoadingSpinner } from '../common';
import styles from '../shell/AppShell.module.scss';
import { formatCurrency, formatDate, getRecordStatusLabel } from '../dashboard/dashboardUtils';
import LogsContent from '../history/LogsContent';

interface IMyRequestsProps {
  dataverseService: DataverseService;
  context: import('@microsoft/sp-webpart-base').WebPartContext;
  activeRole?: string;
}

interface IMyRequestsState {
  records: ODataRecord[];
  loading: boolean;
  error: string | null;
  systemUserId: string | null;
  activeFilter: string;   // 'All' | status label
  searchText: string;
  viewRecord: ODataRecord | null;
  resubmitRecord: ODataRecord | null;
  resubmitDescription: string;
  resubmitSaving: boolean;
  resubmitSuccess: string | null;
  resubmitError: string | null;
}

// ── Confirmed status constants (Power Apps sct_Request_Status) ─────────────────
const MY_STATUS_FILTERS = [
  { label: 'All', color: '#605E5C', bg: '#F3F2F1' },
  { label: 'Draft', color: '#605E5C', bg: '#EDEBE9' },
  { label: 'Submitted', color: '#0078D4', bg: '#DEECF9' },
  { label: 'In Review', color: '#FF8C00', bg: '#FFF4CE' },
  { label: 'PO Issued', color: '#00B294', bg: '#E1FAF4' },
  { label: 'Rejected', color: '#A4262C', bg: '#FDE7E9' },
  { label: 'Returned', color: '#8764B8', bg: '#F4F0FA' },
  { label: 'Closed', color: '#107C10', bg: '#DFF6DD' },
];

const STATUS_COLOR: Record<string, string> = {
  'Draft': '#605E5C',
  'Submitted': '#0078D4',
  'In Review': '#FF8C00',
  'PO Issued': '#00B294',
  'Rejected': '#A4262C',
  'Returned': '#8764B8',
  'Closed': '#107C10',
};

const STATUS_BG: Record<string, string> = {
  'Draft': '#EDEBE9',
  'Submitted': '#DEECF9',
  'In Review': '#FFF4CE',
  'PO Issued': '#E1FAF4',
  'Rejected': '#FDE7E9',
  'Returned': '#F4F0FA',
  'Closed': '#DFF6DD',
};

// Stage progress for visual indicator — maps status label to step index (0-based)
const STAGE_STEP: Record<string, number> = {
  'Draft': 0, 'Submitted': 1, 'In Review': 2, 'PO Issued': 3, 'Closed': 4,
  'Returned': -1, 'Rejected': -1,
};
const STAGE_LABELS = ['Draft', 'Submitted', 'In Review', 'PO Issued', 'Closed'];

export default class MyRequests extends React.Component<IMyRequestsProps, IMyRequestsState> {
  constructor(props: IMyRequestsProps) {
    super(props);
    this.state = {
      records: [], loading: true, error: null,
      systemUserId: null, activeFilter: 'All', searchText: '', viewRecord: null,
      resubmitRecord: null, resubmitDescription: '', resubmitSaving: false,
      resubmitSuccess: null, resubmitError: null,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadData();
  }

  public componentDidUpdate(prevProps: IMyRequestsProps): void {
    if (prevProps.activeRole !== this.props.activeRole) {
      this.setState({ records: [], loading: true }, () => {
        void this._loadData();
      });
    }
  }

  private async _loadData(): Promise<void> {
    this.setState({ loading: true, error: null });
    try {
      const ctx = this.props.context;
      const aadId = ctx.pageContext?.aadInfo?.userId?.toString?.() ?? '';
      let sysUserId: string | null = null;

      if (aadId) {
        const user = await this.props.dataverseService.getCurrentSystemUser(aadId);
        sysUserId = user?.systemuserid ?? null;
      }
      if (!sysUserId && ctx.pageContext?.user?.email) {
        const user = await this.props.dataverseService.getSystemUserByEmail(ctx.pageContext.user.email);
        sysUserId = user?.systemuserid ?? null;
      }

      if (!sysUserId) {
        this.setState({ loading: false, error: 'Could not resolve your Dataverse user account. Contact your administrator.' });
        return;
      }

      const role = (this.props.activeRole || 'requester').toLowerCase();
      const safeUserId = sysUserId.toLowerCase().replace(/[{}]/g, '');
      let records: ODataRecord[] = [];

      if (role === 'manager') {
        const [masterData, approvedIds, reports] = await Promise.all([
          this.props.dataverseService.getManagerMasterRequests(sysUserId),
          this.props.dataverseService.getMyApprovedRequestIds(sysUserId),
          this.props.dataverseService.getDirectReportUserIds(sysUserId).catch(() => []),
        ]);
        const directReportSet = new Set(reports.map(id => id.toLowerCase().replace(/[{}]/g, '')));
        const approvedSet = new Set(approvedIds);

        // Strict role isolation: Hide user's personal requests, show only requests managed/acted on by this manager
        records = masterData.filter(req => {
          const rawReq = req['_sct_requester_value'] ?? req['_createdby_value'] ?? req['sct_requester'] ?? '';
          const reqId = rawReq ? String(rawReq).toLowerCase().replace(/[{}]/g, '') : '';
          const createdById = req['_createdby_value'] ? String(req['_createdby_value']).toLowerCase().replace(/[{}]/g, '') : '';
          // STRICT ISOLATION: Hide the user's personal requests
          if (reqId === safeUserId || createdById === safeUserId) return false;

          const rawMgr = req['_sct_manager_value'] ?? req['_cr2a4_manager_value'] ?? req['_manager_value'] ?? req['sct_manager'] ?? req['cr2a4_manager'] ?? '';
          const mgrId = rawMgr ? String(rawMgr).toLowerCase().replace(/[{}]/g, '') : '';
          const mgrFormatted = (String(req['_cr2a4_manager_value@OData.Community.Display.V1.FormattedValue'] ?? req['_sct_manager_value@OData.Community.Display.V1.FormattedValue'] ?? '')).toLowerCase();
          const isMgrByName = mgrFormatted.indexOf('aniket') !== -1;
          const isManaged = (mgrId === safeUserId) || isMgrByName || directReportSet.has(reqId) || (createdById && directReportSet.has(createdById));
          const reqGuid = String(req['sct_requestid'] ?? '');

          return isManaged || (reqGuid ? approvedSet.has(reqGuid) : false);
        });
      } else if (role.indexOf('tier') !== -1 || role.indexOf('approver') !== -1) {
        const tierIndex = role.indexOf('1') !== -1 ? 1 : role.indexOf('2') !== -1 ? 2 : 3;
        const approvedIds = await this.props.dataverseService.getMyTierApprovedRequestIds(sysUserId, tierIndex as 1 | 2 | 3);
        const approvedSet = new Set(approvedIds);
        const allReqs = await this.props.dataverseService.getRawRequests();

        // Strict role isolation: Hide user's personal requests, show only approved requests
        records = allReqs.filter(req => {
          const rawReq = req['_sct_requester_value'] ?? req['sct_requester'] ?? '';
          const reqId = rawReq ? String(rawReq).toLowerCase().replace(/[{}]/g, '') : '';
          // STRICT ISOLATION: Hide the user's personal requests
          if (reqId === safeUserId) return false;

          const reqGuid = String(req['sct_requestid'] ?? '');
          return approvedSet.has(reqGuid);
        });
      } else {
        // Requester role: Show ONLY requests created by this user
        records = await this.props.dataverseService.getMyRequests(sysUserId);
      }

      this.setState({ records, systemUserId: sysUserId, loading: false });
    } catch (e: unknown) {
      this.setState({ loading: false, error: e instanceof Error ? e.message : String(e) });
    }
  }

  private _getStatusLabel(record: ODataRecord): string {
    return getRecordStatusLabel(record);
  }

  // sct_submitted_date confirmed from Power Apps (schema: sct_Submitted_Date, logical: sct_submitted_date)
  private _getSubmittedDate(record: ODataRecord): string {
    const submitted = record['sct_submitted_date'];
    if (submitted) return formatDate(String(submitted));
    const created = record['createdon'];
    if (created) return formatDate(String(created));
    return '—';
  }

  private _getFiltered(): ODataRecord[] {
    const { records, activeFilter, searchText } = this.state;
    const q = searchText.trim().toLowerCase();
    return records.filter(r => {
      const statusLabel = this._getStatusLabel(r);
      if (activeFilter !== 'All' && statusLabel !== activeFilter) return false;
      if (q) {
        const ref = strVal(r, 'sct_sct_reference').toLowerCase();
        const desc = strVal(r, 'sct_description').toLowerCase();
        const site = strVal(r, '_sct_site_value').toLowerCase();
        const type = strVal(r, 'sct_purchase_type').toLowerCase();
        if (!ref.includes(q) && !desc.includes(q) && !site.includes(q) && !type.includes(q)) return false;
      }
      return true;
    });
  }

  private _countFor(label: string): number {
    if (label === 'All') return this.state.records.length;
    return this.state.records.filter(r => this._getStatusLabel(r) === label).length;
  }

  private async _handleResubmit(): Promise<void> {
    const { resubmitRecord, resubmitDescription } = this.state;
    if (!resubmitRecord) return;
    const reqId = String(resubmitRecord['sct_requestid'] ?? resubmitRecord['id'] ?? '');
    if (!reqId) return;

    this.setState({ resubmitSaving: true, resubmitError: null });
    try {
      const payload: Record<string, unknown> = {
        sct_request_status: 854000001, // Submitted
        cr2a4_pendingtier: 'Manager',
        sct_description: resubmitDescription.trim() || strVal(resubmitRecord, 'sct_description'),
      };

      await this.props.dataverseService.updateRequest(reqId, payload);

      // Audit log the resubmission event
      await this.props.dataverseService.createAuditLog({
        sct_action: 1, // Submitted
        sct_description: 'Request resubmitted for Line Manager approval by requester',
        sct_oldvalue: 'Status: Returned, Pending Tier: None',
        sct_newvalue: 'Status: Submitted, Pending Tier: Manager',
        'sct_requestid@odata.bind': `/sct_requests(${reqId})`,
        ...(this.state.systemUserId ? { 'sct_actionby@odata.bind': `/systemusers(${this.state.systemUserId})` } : {}),
      }).catch(() => { /* non-fatal */ });

      // Notify manager via Graph
      try {
        const graphClient = await this.props.context.msGraphClientFactory.getClient('3');
        const manager = await getManager(graphClient);
        if (manager && manager.mail) {
          const title = strVal(resubmitRecord, 'sct_request_title') || strVal(resubmitRecord, 'sct_sct_reference');
          const amt = String(resubmitRecord['sct_estimated_total_value'] ?? '');
          const reqName = this.props.context.pageContext.user.displayName;
          await sendManagerApprovalEmail(graphClient, manager, reqName, title, amt, this.props.context.pageContext.web.absoluteUrl);
        }
      } catch { /* non-fatal email error */ }

      this.setState({
        resubmitSaving: false,
        resubmitRecord: null,
        viewRecord: null,
        resubmitSuccess: 'Request resubmitted successfully! It is now pending Line Manager review.',
      });
      await this._loadData();
    } catch (err: unknown) {
      const errMsg = err instanceof Error ? err.message : 'Resubmission failed. Please try again.';
      this.setState({ resubmitSaving: false, resubmitError: errMsg });
    }
  }

  public render(): React.ReactElement {
    const { loading, error, activeFilter, searchText, viewRecord } = this.state;
    const filtered = this._getFiltered();

    if (loading) return <LoadingSpinner label="Loading your requests…" />;

    return (
      <div>
        {this.state.resubmitSuccess && (
          <MessageBar
            messageBarType={MessageBarType.success}
            onDismiss={() => this.setState({ resubmitSuccess: null })}
            styles={{ root: { marginBottom: 16, borderRadius: 6 } }}
          >
            {this.state.resubmitSuccess}
          </MessageBar>
        )}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} style={{ marginBottom: 16, borderRadius: 6 }}>
            {error}
            <button
              style={{ marginLeft: 12, background: 'none', border: 'none', color: '#A4262C', cursor: 'pointer', textDecoration: 'underline', fontFamily: 'inherit', fontSize: 13 }}
              onClick={() => void this._loadData()}
            >
              Retry
            </button>
          </MessageBar>
        )}

        {/* ── Filter pills ─────────────────────────────────── */}
        <div className={styles.card} style={{ marginBottom: 16 }}>
          <div className={styles.cardBody} style={{ paddingTop: 14, paddingBottom: 14 }}>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
              <Text variant="small" style={{ color: '#605E5C', fontWeight: 600, marginRight: 4 }}>Filter:</Text>
              {MY_STATUS_FILTERS.map(f => {
                const count = this._countFor(f.label);
                const active = activeFilter === f.label;
                return (
                  <button
                    key={f.label}
                    onClick={() => this.setState({ activeFilter: f.label })}
                    style={{
                      background: active ? f.color : f.bg,
                      color: active ? '#fff' : f.color,
                      border: active ? `2px solid ${f.color}` : '2px solid transparent',
                      borderRadius: 16, padding: '4px 12px',
                      fontSize: 12, fontWeight: 600, cursor: 'pointer',
                      fontFamily: 'inherit', display: 'flex', alignItems: 'center', gap: 5,
                      transition: 'all 0.15s',
                    }}
                  >
                    {f.label}
                    <span style={{
                      background: active ? 'rgba(255,255,255,0.25)' : 'rgba(0,0,0,0.10)',
                      borderRadius: 10, padding: '0 5px', fontSize: 11,
                    }}>
                      {count}
                    </span>
                  </button>
                );
              })}
              <div style={{ flex: 1, minWidth: 200 }}>
                <SearchBox
                  placeholder="Search reference, description, site…"
                  value={searchText}
                  onChange={(_, v) => this.setState({ searchText: v ?? '' })}
                  styles={{ root: { borderRadius: 6 } }}
                />
              </div>
            </div>
          </div>
        </div>

        {/* ── Request cards ─────────────────────────────────── */}
        <div className={styles.card}>
          <div className={styles.cardHeader}>
            <div className={styles.cardTitle}>
              <div className={styles.cardTitleIcon} style={{ background: '#DFF6DD' }} aria-hidden="true">
                <Icon iconName="DocumentSearch" style={{ color: '#107C10', fontSize: 16 }} />
              </div>
              My Requests
              <span style={{ marginLeft: 8, fontSize: 12, color: '#A19F9D', fontWeight: 400 }}>
                {filtered.length} of {this.state.records.length}
              </span>
            </div>
            <button
              onClick={() => void this._loadData()}
              style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#0078D4', display: 'flex', alignItems: 'center', gap: 4, fontSize: 13, fontFamily: 'inherit' }}
              title="Refresh"
            >
              <Icon iconName="Refresh" style={{ fontSize: 14 }} /> Refresh
            </button>
          </div>

          <div className={styles.cardBody}>
            {filtered.length === 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '48px 0', gap: 12 }}>
                <Icon iconName="DocumentReply" style={{ fontSize: 48, color: '#C8C6C4' }} aria-hidden="true" />
                <Text variant="mediumPlus" style={{ color: '#A19F9D', fontWeight: 500, textAlign: 'center' }}>
                  {this.state.records.length === 0 ? 'You have no requests yet' : 'No requests match your filter'}
                </Text>
                <Text variant="small" style={{ color: '#C8C6C4', textAlign: 'center' }}>
                  {this.state.records.length === 0 ? 'Use the New Request tab to submit your first request.' : 'Try clearing the search or selecting All.'}
                </Text>
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                {filtered.map(r => this._renderCard(r))}
              </div>
            )}
          </div>
        </div>

        {/* ── View detail panel ─────────────────────────────── */}
        {viewRecord && this._renderDetailPanel(viewRecord)}

        {/* ── Revise & Resubmit Dialog ──────────────────────── */}
        <Dialog
          hidden={!this.state.resubmitRecord}
          onDismiss={() => { if (!this.state.resubmitSaving) this.setState({ resubmitRecord: null }); }}
          dialogContentProps={{
            type: DialogType.normal,
            title: `Revise & Resubmit — ${this.state.resubmitRecord ? strVal(this.state.resubmitRecord, 'sct_sct_reference') : ''}`,
            subText: 'Update the request description/justification based on approver feedback. Once resubmitted, the request will restart from the Line Manager approval stage.',
          }}
          modalProps={{ isBlocking: true }}
        >
          {this.state.resubmitError && (
            <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 12, borderRadius: 6 } }}>
              {this.state.resubmitError}
            </MessageBar>
          )}

          <TextField
            label="Updated Description / Justification"
            multiline
            rows={4}
            value={this.state.resubmitDescription}
            onChange={(_, val) => this.setState({ resubmitDescription: val || '' })}
            placeholder="Add clarifications or changes requested by approvers..."
            disabled={this.state.resubmitSaving}
            required
          />

          <DialogFooter>
            <PrimaryButton
              text={this.state.resubmitSaving ? 'Resubmitting…' : 'Resubmit Request'}
              iconProps={this.state.resubmitSaving ? undefined : { iconName: 'Send' }}
              onClick={() => void this._handleResubmit()}
              disabled={this.state.resubmitSaving || !this.state.resubmitDescription.trim()}
              styles={{ root: { borderRadius: 6, background: '#107C10', borderColor: '#107C10' } }}
            >
              {this.state.resubmitSaving && <Spinner size={SpinnerSize.xSmall} style={{ marginRight: 6 }} />}
            </PrimaryButton>
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ resubmitRecord: null })}
              disabled={this.state.resubmitSaving}
              styles={{ root: { borderRadius: 6 } }}
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  private _renderCard(r: ODataRecord): React.ReactElement {
    const ref = strVal(r, 'sct_sct_reference') || '(no reference)';
    const statusLabel = this._getStatusLabel(r);
    const site = strVal(r, '_sct_site_value') || strVal(r, 'sct_site');
    const desc = strVal(r, 'sct_description');
    const amtRaw = parseFloat(String(r['sct_estimated_total_value'] ?? ''));
    const amt = !isNaN(amtRaw) ? formatCurrency(amtRaw) : '—';
    const type = strVal(r, 'sct_purchase_type') || strVal(r, 'sct_requesttype');
    const submittedOn = this._getSubmittedDate(r);
    const poNum = strVal(r, 'sct_po_number');
    const isRejected = statusLabel === 'Rejected';
    const isReturned = statusLabel === 'Returned';
    const stepIndex = STAGE_STEP[statusLabel] ?? -1;

    const pt = strVal(r, 'cr2a4_pendingtier');
    const displayStatus = (statusLabel === 'In Review' && pt) ? `${statusLabel} (${pt})` : statusLabel;

    const pill = (
      <span style={{
        background: STATUS_BG[statusLabel] ?? '#EDEBE9',
        color: STATUS_COLOR[statusLabel] ?? '#605E5C',
        borderRadius: 12, padding: '2px 10px',
        fontSize: 11, fontWeight: 700, whiteSpace: 'nowrap',
      }}>
        {displayStatus}
      </span>
    );

    return (
      <div
        key={String(r['sct_requestid'] ?? ref)}
        onClick={() => this.setState({ viewRecord: r })}
        style={{
          border: `1px solid ${isRejected ? '#F3CFCF' : isReturned ? '#D9D0E8' : '#EDEBE9'}`,
          borderLeft: `4px solid ${STATUS_COLOR[statusLabel] ?? '#EDEBE9'}`,
          borderRadius: 8, padding: '14px 16px', cursor: 'pointer',
          background: '#FAFAFA', transition: 'box-shadow 0.15s',
        }}
        onMouseEnter={e => (e.currentTarget.style.boxShadow = '0 2px 8px rgba(0,0,0,0.10)')}
        onMouseLeave={e => (e.currentTarget.style.boxShadow = 'none')}
      >
        {/* Row 1: Reference + status pill + site */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6, flexWrap: 'wrap' }}>
          <Text variant="medium" style={{ fontWeight: 700, color: '#201F1E' }}>{ref}</Text>
          {pill}
          {site && (
            <span style={{ fontSize: 12, color: '#605E5C', display: 'flex', alignItems: 'center', gap: 4 }}>
              <Icon iconName="MapPin" style={{ fontSize: 11 }} />{site}
            </span>
          )}
          <span style={{ marginLeft: 'auto', fontSize: 11, color: '#A19F9D' }}>
            <Icon iconName="Clock" style={{ fontSize: 11, marginRight: 4 }} />
            Submitted: {submittedOn}
          </span>
        </div>

        {/* Row 2: Description */}
        {desc && (
          <div style={{ fontSize: 13, color: '#605E5C', marginBottom: 8, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '90%' }}>
            {desc}
          </div>
        )}

        {/* Row 3: Meta chips */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap', fontSize: 12, color: '#605E5C' }}>
          {amt !== '—' && (
            <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
              <Icon iconName="Money" style={{ fontSize: 12, color: '#0078D4' }} />
              <strong style={{ color: '#201F1E' }}>{amt}</strong>
            </span>
          )}
          {type && (
            <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
              <Icon iconName="Tag" style={{ fontSize: 11 }} />{type}
            </span>
          )}
          {poNum && (
            <span style={{ display: 'flex', alignItems: 'center', gap: 4, color: '#00B294', fontWeight: 600 }}>
              <Icon iconName="CheckMark" style={{ fontSize: 11 }} />PO: {poNum}
            </span>
          )}
          <span style={{ marginLeft: 'auto', color: '#0078D4', fontSize: 12, display: 'flex', alignItems: 'center', gap: 4 }}>
            View details <Icon iconName="ChevronRight" style={{ fontSize: 10 }} />
          </span>
        </div>

        {/* Row 4: Stage progress bar (not shown for terminal statuses) */}
        {stepIndex >= 0 && (
          <div style={{ marginTop: 10, display: 'flex', alignItems: 'center', gap: 4 }}>
            {STAGE_LABELS.map((s, i) => (
              <React.Fragment key={s}>
                <div style={{
                  flex: 1, height: 4, borderRadius: 2,
                  background: i <= stepIndex ? (STATUS_COLOR[statusLabel] ?? '#0078D4') : '#EDEBE9',
                  transition: 'background 0.3s',
                }} title={s} />
              </React.Fragment>
            ))}
          </div>
        )}
        {(isRejected || isReturned) && (
          <div style={{ marginTop: 8, fontSize: 12, color: STATUS_COLOR[statusLabel], fontWeight: 600, display: 'flex', alignItems: 'center', gap: 4 }}>
            <Icon iconName={isRejected ? 'StatusErrorFull' : 'ReturnToSession'} style={{ fontSize: 12 }} />
            {isRejected ? 'This request was rejected.' : 'This request was returned — please revise and resubmit.'}
          </div>
        )}
      </div>
    );
  }

  private _renderDetailPanel(r: ODataRecord): React.ReactElement {
    const ref = strVal(r, 'sct_sct_reference') || '(no reference)';
    const statusLabel = this._getStatusLabel(r);
    const site = strVal(r, '_sct_site_value');
    const desc = strVal(r, 'sct_description');
    const amtRaw = parseFloat(String(r['sct_estimated_total_value'] ?? ''));
    const amt = !isNaN(amtRaw) ? formatCurrency(amtRaw) : '—';
    const close = (): void => this.setState({ viewRecord: null });

    const lbl: React.CSSProperties = { fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 2 };
    const val: React.CSSProperties = { fontSize: 13, color: '#201F1E', fontWeight: 500, marginBottom: 10 };

    const Row: React.FC<{ label: string; value: string }> = ({ label, value }) => (
      <div>
        <div style={lbl}>{label}</div>
        <div style={val}>{value || '—'}</div>
      </div>
    );

    const requestId = String(r['sct_requestid'] ?? '');
    const canEdit = statusLabel === 'Draft';

    return (
      <Panel isOpen onDismiss={close} type={PanelType.custom} customWidth="min(720px, 96vw)"
        headerText={`Request: ${ref}`}
        closeButtonAriaLabel="Close"
        isFooterAtBottom
        onRenderFooterContent={() => (
          <div style={{ display: 'flex', gap: 10, padding: '0 24px 24px' }}>
            {statusLabel === 'Returned' ? (
              <PrimaryButton
                text="Revise & Resubmit"
                iconProps={{ iconName: 'ReturnToSession' }}
                onClick={() => {
                  this.setState({
                    resubmitRecord: r,
                    resubmitDescription: strVal(r, 'sct_description'),
                    resubmitError: null,
                  });
                }}
                styles={{ root: { borderRadius: 6, background: '#7A4F01', borderColor: '#C8961F' } }}
              />
            ) : (
              <PrimaryButton
                text="Edit Record"
                iconProps={{ iconName: 'Edit' }}
                disabled={!canEdit}
                title={canEdit ? 'Edit this draft request' : `Records with status "${statusLabel}" cannot be edited`}
                onClick={() => {
                  close();
                  // TODO: open edit panel for this record
                }}
                styles={{ root: { borderRadius: 6 } }}
              />
            )}
            <DefaultButton
              text="Close"
              onClick={close}
              styles={{ root: { borderRadius: 6 } }}
            />
          </div>
        )}
        styles={{ content: { paddingBottom: 60 } }}
      >
        {/* Status pill */}
        <div style={{ marginBottom: 16 }}>
          <span style={{ background: STATUS_BG[statusLabel] ?? '#EDEBE9', color: STATUS_COLOR[statusLabel] ?? '#605E5C', borderRadius: 12, padding: '4px 14px', fontSize: 13, fontWeight: 700 }}>
            {statusLabel}
          </span>
        </div>

        {/* Context */}
        <div style={{ background: '#F8F8F7', borderRadius: 8, padding: '12px 14px', marginBottom: 20 }}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '4px 16px' }}>
            <Row label="Reference" value={ref} />
            <Row label="Status" value={statusLabel} />
            <Row label="Site" value={site} />
            <Row label="Purchase Type" value={strVal(r, 'sct_purchase_type') || strVal(r, 'sct_requesttype')} />
            <Row label="Estimated Value" value={amt} />
            <Row label="Urgency" value={strVal(r, 'sct_urgency')} />
            <Row label="Submitted Date" value={this._getSubmittedDate(r)} />
            <Row label="Date Needed By" value={formatDate(strVal(r, 'sct_date_needed_by'))} />
            <Row label="PO Number" value={strVal(r, 'sct_po_number')} />
            <Row label="PO Issue Date" value={formatDate(strVal(r, 'sct_po_issue_date'))} />
          </div>
        </div>

        {desc && (
          <div style={{ marginBottom: 16 }}>
            <div style={lbl}>Description</div>
            <div style={{ fontSize: 13, color: '#201F1E', lineHeight: '1.6' }}>{desc}</div>
          </div>
        )}

        {strVal(r, 'sct_urgency_reason') && (
          <div style={{ marginBottom: 16 }}>
            <div style={lbl}>Urgency Reason</div>
            <div style={{ fontSize: 13, color: '#201F1E', lineHeight: '1.6' }}>{strVal(r, 'sct_urgency_reason')}</div>
          </div>
        )}

        {/* Flags */}
        <div style={{ borderTop: '1px solid #EDEBE9', paddingTop: 14 }}>
          <div style={lbl}>Flags</div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 6 }}>
            {([
              ['sct_is_new_supplier', 'New Supplier', '#905400', '#FFF4CE'],
              ['sct_is_retrospective', 'Retrospective', '#605E5C', '#EDEBE9'],
              ['sct_is_confidential', 'Confidential', '#A4262C', '#FDE7E9'],
              ['sct_tech_or_data_involved', 'Tech/Data', '#0078D4', '#DEECF9'],
              ['sct_possible_capital_purchase', 'Capital Purchase', '#8764B8', '#F4F0FA'],
              ['sct_cost_recharged_to_customer', 'Recharged to Customer', '#00B294', '#E1FAF4'],
            ] as Array<[string, string, string, string]>).map(([field, label, color, bg]) => {
              const v = r[field];
              const on = v === true || v === 'Yes' || v === '1';
              if (!on) return null;
              return (
                <span key={field} style={{ background: bg, color, borderRadius: 12, padding: '2px 10px', fontSize: 11, fontWeight: 700 }}>
                  {label}
                </span>
              );
            })}
          </div>
        </div>

        {/* ── Approval + Audit logs ──────────────────────── */}
        {requestId && (
          <div style={{ marginTop: 24 }}>
            <div style={{
              display: 'flex', alignItems: 'center', gap: 10,
              marginBottom: 14, paddingBottom: 10, borderBottom: '2px solid #EDEBE9',
            }}>
              <div style={{
                width: 28, height: 28, borderRadius: 6, flexShrink: 0,
                background: 'linear-gradient(135deg, #107C10, #8764B8)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
              }}>
                <Icon iconName="History" style={{ fontSize: 14, color: '#fff' }} />
              </div>
              <span style={{ fontSize: 15, fontWeight: 700, color: '#201F1E' }}>Request History</span>
            </div>
            <LogsContent requestId={requestId} service={this.props.dataverseService} />
          </div>
        )}
      </Panel>
    );
  }
}
