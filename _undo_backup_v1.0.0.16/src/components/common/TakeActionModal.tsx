import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { IconButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import { DataverseService, ODataRecord, strVal } from '../../services/DataverseService';
import { handleApprovalAction, checkStageDecisionLock } from '../../services/approvalHandler';
import { formatCurrency, formatDate, formatDateTime } from '../dashboard/dashboardUtils';

export interface ITakeActionModalProps {
  isOpen: boolean;
  record: ODataRecord | null;
  service: DataverseService;
  activeRole: string; // 'manager' | 'approver_tier1' | 'approver_tier2' | 'approver_tier3'
  currentSystemUserId: string | null;
  context: WebPartContext;
  onDismiss: () => void;
  onSuccess: (message: string) => void;
}

interface ITakeActionModalState {
  comments: string;
  commentsError: string | null;
  isSubmitting: boolean;
  submitError: string | null;
  loadingHistory: boolean;
  historyLogs: ODataRecord[];
  historyError: string | null;
  showHistory: boolean;
}

export default class TakeActionModal extends React.Component<ITakeActionModalProps, ITakeActionModalState> {
  constructor(props: ITakeActionModalProps) {
    super(props);
    this.state = {
      comments: '',
      commentsError: null,
      isSubmitting: false,
      submitError: null,
      loadingHistory: false,
      historyLogs: [],
      historyError: null,
      showHistory: true,
    };
  }

  public componentDidMount(): void {
    if (this.props.isOpen && this.props.record) {
      void this._loadHistory();
    }
  }

  public componentDidUpdate(prevProps: ITakeActionModalProps): void {
    if (this.props.isOpen && (!prevProps.isOpen || prevProps.record !== this.props.record)) {
      this.setState({
        comments: '',
        commentsError: null,
        isSubmitting: false,
        submitError: null,
      });
      void this._loadHistory();
    }
  }

  private _getRecordId(r: ODataRecord | null): string {
    if (!r) return '';
    return String(r['sct_requestid'] ?? r['sct_requests_id'] ?? r['id'] ?? '');
  }

  private async _loadHistory(): Promise<void> {
    const { record, service } = this.props;
    if (!record) return;
    const reqId = this._getRecordId(record);
    if (!reqId) return;

    this.setState({ loadingHistory: true, historyError: null });
    try {
      const logs = await service.getApprovalLogs(reqId);
      this.setState({ historyLogs: logs || [], loadingHistory: false });
    } catch (err: unknown) {
      const errMsg = err instanceof Error ? err.message : String(err);
      this.setState({ historyError: errMsg, loadingHistory: false });
    }
  }

  /** Role title for display */
  private _roleDisplay(): string {
    const r = (this.props.activeRole || '').toLowerCase();
    if (r === 'manager') return 'Line Manager';
    if (r.indexOf('1') !== -1) return 'Tier 1 Approver';
    if (r.indexOf('2') !== -1) return 'Tier 2 Approver';
    if (r.indexOf('3') !== -1) return 'Tier 3 Approver';
    return 'Approver';
  }

  /** Target destination text for the approve button */
  private _approveActionText(): string {
    const r = (this.props.activeRole || '').toLowerCase();
    if (r === 'manager') return 'Approve & Route to Tier 1';
    if (r.indexOf('1') !== -1) return 'Approve & Route to Tier 2';
    if (r.indexOf('2') !== -1) return 'Approve & Route to Tier 3';
    if (r.indexOf('3') !== -1) return 'Final Approval (Issue PO)';
    return 'Approve Request';
  }

  /** Current pending stage text */
  private _currentStageInfo(): { label: string; currentStep: number } {
    const { record } = this.props;
    if (!record) return { label: 'Line Manager Review', currentStep: 0 };
    const pt = String(record['cr2a4_pendingtier'] || record['cr2a4_PendingTier'] || record['sct_pendingtier'] || record['pendingtier'] || '').toLowerCase();
    if (pt.indexOf('3') !== -1) return { label: 'Tier 3 Approver Review', currentStep: 3 };
    if (pt.indexOf('2') !== -1) return { label: 'Tier 2 Approver Review', currentStep: 2 };
    if (pt.indexOf('1') !== -1) return { label: 'Tier 1 Approver Review', currentStep: 1 };
    return { label: 'Line Manager Review', currentStep: 0 };
  }

  private async _handleDecision(decision: 'approve' | 'reject' | 'return'): Promise<void> {
    const { record, service, currentSystemUserId, activeRole, context } = this.props;
    const { comments } = this.state;
    if (!record) return;

    const requestId = this._getRecordId(record);
    if (!requestId) {
      this.setState({ submitError: 'Invalid request: Missing record ID.' });
      return;
    }

    const trimmedComments = (comments || '').trim();

    // Enforce mandatory rationale for Reject and Return
    if ((decision === 'reject' || decision === 'return') && trimmedComments.length === 0) {
      this.setState({
        commentsError: `Comments are mandatory when taking a "${decision === 'reject' ? 'Reject' : 'Return for Correction'}" action. Please provide a clear reason for the requester.`,
      });
      return;
    }

    this.setState({ isSubmitting: true, submitError: null, commentsError: null });

    try {
      let graphClient: MSGraphClientV3 | undefined;
      try {
        graphClient = await context.msGraphClientFactory.getClient('3');
      } catch { /* optional Graph client */ }

      const rawReq = String(record['_sct_requester_value'] ?? record['sct_requester'] ?? record['_createdby_value'] ?? '');

      const result = await handleApprovalAction({
        requestId,
        role: activeRole || 'manager',
        action: decision,
        comments: trimmedComments,
        service,
        currentSystemUserId,
        requesterId: rawReq,
        graphClient,
        webPartUrl: context.pageContext.web.absoluteUrl,
        approverDisplayName: context.pageContext.user.displayName,
        currentUserEmail: context.pageContext.user.email || undefined,
      });

      const roleLabel = this._roleDisplay();
      let successMsg = '';
      if (decision === 'approve') {
        const next = (activeRole || '').toLowerCase().indexOf('3') !== -1 ? 'Approved and PO Issued' : 'advanced to the next approval tier';
        successMsg = `Request successfully approved by ${roleLabel} and ${next}.`;
      } else if (decision === 'reject') {
        successMsg = `Request successfully rejected by ${roleLabel}. Workflow terminated.`;
      } else {
        successMsg = `Request returned for correction by ${roleLabel}. The requester may now revise and resubmit.`;
      }

      if (result.emailDispatched) {
        successMsg += ' Automated notification email dispatched.';
      }

      this.setState({ isSubmitting: false });
      this.props.onSuccess(successMsg);
    } catch (err: unknown) {
      const errMsg = err instanceof Error ? err.message : 'Approval action failed. Please try again.';
      this.setState({ isSubmitting: false, submitError: errMsg });
    }
  }

  public render(): React.ReactElement | null {
    const { isOpen, record, onDismiss } = this.props;
    const { comments, commentsError, isSubmitting, submitError, loadingHistory, historyLogs, historyError, showHistory } = this.state;

    if (!isOpen || !record) return null;

    const ref = strVal(record, 'sct_sct_reference') || 'Request';
    const title = strVal(record, 'sct_request_title') || '—';
    const desc = strVal(record, 'sct_description');
    const reqName = strVal(record, '_sct_requester_value@OData.Community.Display.V1.FormattedValue')
      || strVal(record, 'sct_requester@OData.Community.Display.V1.FormattedValue')
      || strVal(record, 'requesterName')
      || strVal(record, 'sct_requester')
      || '—';
    const bu = strVal(record, 'sct_business_unit') || strVal(record, 'sct_businessunit') || '—';
    const costCenter = strVal(record, 'sct_cost_center') || strVal(record, 'sct_costcenter') || '—';
    const site = strVal(record, '_sct_site_value') || '—';
    const ptype = strVal(record, 'sct_purchase_type') || '—';
    const urgency = strVal(record, 'sct_urgency') || 'Standard';
    const dateNeeded = strVal(record, 'sct_date_needed_by');
    const amtRaw = parseFloat(String(record['sct_estimated_total_value'] ?? 0));
    const amountFormatted = !isNaN(amtRaw) && amtRaw > 0 ? formatCurrency(amtRaw) : '—';

    const stageInfo = this._currentStageInfo();
    const roleTitle = this._roleDisplay();
    const approveText = this._approveActionText();

    const decisionLock = checkStageDecisionLock(record, this.props.activeRole, historyLogs);
    const canAct = !isSubmitting && !decisionLock.isLocked;

    const STAGES = ['Manager', 'Tier 1', 'Tier 2', 'Tier 3', 'Approved'];

    const btnBase: React.CSSProperties = {
      display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 7,
      padding: '10px 18px', borderRadius: 6, border: 'none',
      fontSize: 13, fontWeight: 600, cursor: isSubmitting ? 'not-allowed' : 'pointer',
      opacity: isSubmitting ? 0.6 : 1, transition: 'all 0.15s ease-in-out',
      minWidth: 130, height: 38,
    };

    return (
      <Panel
        isOpen={this.props.isOpen && record !== null}
        onDismiss={() => !isSubmitting && onDismiss()}
        type={PanelType.custom}
        customWidth="min(800px, 96vw)"
        hasCloseButton={false}
        isFooterAtBottom={true}
        onRenderNavigationContent={() => (
          <div style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '16px 24px',
            borderBottom: '1px solid #EDEBE9',
            background: '#FFFFFF',
            width: '100%',
            boxSizing: 'border-box',
          }}>
            <div>
              <div style={{ fontSize: 18, fontWeight: 700, color: '#201F1E', display: 'flex', alignItems: 'center', gap: 8 }}>
                <Icon iconName="CheckList" style={{ color: '#0078D4', fontSize: 18 }} />
                <span>Take Action — {ref}</span>
              </div>
              <div style={{ fontSize: 12, color: '#605E5C', marginTop: 2 }}>
                Review and decide on this spend request as <strong>{roleTitle}</strong>
              </div>
            </div>
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              ariaLabel="Close"
              title="Close panel"
              disabled={isSubmitting}
              onClick={onDismiss}
              styles={{
                root: { color: '#605E5C', borderRadius: 4 },
                rootHovered: { color: '#201F1E', background: '#F3F2F1' },
              }}
            />
          </div>
        )}
        styles={{
          main: { boxShadow: '-4px 0 24px rgba(0,0,0,0.12)' },
          scrollableContent: { display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' },
          content: { flex: 1, overflowY: 'auto', paddingTop: 16, paddingBottom: 8, paddingLeft: 24, paddingRight: 24 },
          footer: { borderTop: 'none', padding: 0 },
        }}
        onRenderFooterContent={() => (
          <div style={{
            display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'flex-end',
            padding: '14px 24px', background: '#FAFAFA', borderTop: '1px solid #EDEBE9',
          }}>
            <button
              style={{
                ...btnBase,
                background: canAct ? '#107C10' : '#EDEBE9',
                color: canAct ? '#FFFFFF' : '#A19F9D',
                boxShadow: canAct && !isSubmitting ? '0 2px 6px rgba(16,124,16,0.25)' : 'none',
                cursor: canAct ? 'pointer' : 'not-allowed',
              }}
              disabled={!canAct}
              title={canAct ? approveText : (decisionLock.lockMessage || 'Actions are locked for this stage')}
              onClick={() => void this._handleDecision('approve')}
            >
              {isSubmitting ? <Spinner size={SpinnerSize.xSmall} styles={{ root: { display: 'inline-flex' } }} /> : <Icon iconName="CheckMark" style={{ fontSize: 12 }} />}
              {isSubmitting ? 'Saving…' : approveText}
            </button>

            <button
              style={{
                ...btnBase,
                background: canAct ? '#FFF4CE' : '#EDEBE9',
                color: canAct ? '#7A4F01' : '#A19F9D',
                border: canAct ? '1.5px solid #C8961F' : '1px solid #EDEBE9',
                cursor: canAct ? 'pointer' : 'not-allowed',
              }}
              disabled={!canAct}
              title={canAct ? 'Return for Correction' : (decisionLock.lockMessage || 'Actions are locked for this stage')}
              onClick={() => void this._handleDecision('return')}
            >
              {isSubmitting ? <Spinner size={SpinnerSize.xSmall} styles={{ root: { display: 'inline-flex' } }} /> : <Icon iconName="ReturnToSession" style={{ fontSize: 12 }} />}
              {isSubmitting ? 'Saving…' : 'Return for Correction'}
            </button>

            <button
              style={{
                ...btnBase,
                background: canAct ? '#FDE7E9' : '#EDEBE9',
                color: canAct ? '#A4262C' : '#A19F9D',
                border: canAct ? '1.5px solid #A4262C' : '1px solid #EDEBE9',
                cursor: canAct ? 'pointer' : 'not-allowed',
              }}
              disabled={!canAct}
              title={canAct ? 'Reject' : (decisionLock.lockMessage || 'Actions are locked for this stage')}
              onClick={() => void this._handleDecision('reject')}
            >
              {isSubmitting ? <Spinner size={SpinnerSize.xSmall} styles={{ root: { display: 'inline-flex' } }} /> : <Icon iconName="Cancel" style={{ fontSize: 12 }} />}
              {isSubmitting ? 'Saving…' : 'Reject'}
            </button>

            <button
              style={{ ...btnBase, background: '#F3F2F1', color: '#323130', border: '1px solid #C8C6C4', minWidth: 90 }}
              disabled={isSubmitting}
              onClick={onDismiss}
            >
              Cancel
            </button>
          </div>
        )}
      >
        {/* Top Error MessageBar if any submission error */}
        {submitError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline
            onDismiss={() => this.setState({ submitError: null })}
            styles={{ root: { borderRadius: 6, marginBottom: 14 } }}
          >
            {submitError}
          </MessageBar>
        )}

        {/* One-Time Decision Locking Banner */}
        {decisionLock.isLocked && (
          <div style={{
            background: decisionLock.previousDecision ? '#F0FDF4' : '#FFFBEB',
            border: `1.5px solid ${decisionLock.previousDecision ? '#86EFAC' : '#FDE68A'}`,
            borderRadius: 8,
            padding: '12px 16px',
            marginBottom: 16,
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
              <Icon
                iconName={decisionLock.previousDecision ? 'CompletedSolid' : 'LockSolid'}
                style={{ fontSize: 16, color: decisionLock.previousDecision ? '#16A34A' : '#D97706' }}
              />
              <span style={{ fontSize: 13, fontWeight: 700, color: decisionLock.previousDecision ? '#166534' : '#92400E' }}>
                {decisionLock.previousDecision
                  ? `You have already submitted your decision for this stage (${decisionLock.stageName}).`
                  : decisionLock.lockMessage}
              </span>
            </div>
            {decisionLock.previousDecision && (
              <div style={{ fontSize: 12, color: '#1E293B', marginLeft: 24, lineHeight: 1.6 }}>
                <div>
                  Decision: <strong style={{ color: decisionLock.previousDecision.action === 'Approved' ? '#16A34A' : decisionLock.previousDecision.action === 'Rejected' ? '#DC2626' : '#D97706' }}>{decisionLock.previousDecision.action}</strong> by <strong>{decisionLock.previousDecision.approverName}</strong> on <strong>{decisionLock.previousDecision.timestamp}</strong>.
                </div>
                {decisionLock.previousDecision.comments && (
                  <div style={{ fontStyle: 'italic', marginTop: 2 }}>
                    &ldquo;{decisionLock.previousDecision.comments}&rdquo;
                  </div>
                )}
                <div style={{ marginTop: 4, fontSize: 11, color: '#64748B' }}>
                  Further actions for this stage are permanently locked in this approval cycle.
                </div>
              </div>
            )}
          </div>
        )}

        {/* Role & Stage Banner */}
        <div style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          background: '#F0F6FF', border: '1px solid #C7E0F4', borderRadius: 8,
          padding: '10px 14px', marginBottom: 16,
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <Icon iconName="Shield" style={{ color: '#0078D4', fontSize: 16 }} />
            <Text variant="small" style={{ color: '#004578', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.4px' }}>
              Acting Role: {roleTitle}
            </Text>
          </div>
          <span style={{
            background: '#0078D4', color: '#fff', fontSize: 11, fontWeight: 700,
            padding: '3px 10px', borderRadius: 12, letterSpacing: '0.3px',
          }}>
            {stageInfo.label}
          </span>
        </div>

        {/* Workflow Progression Stepper */}
        <div style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          background: '#FAF9F8', border: '1px solid #EDEBE9', borderRadius: 8,
          padding: '10px 16px', marginBottom: 18,
        }}>
          {STAGES.map((s, idx) => {
            const isCurrent = idx === stageInfo.currentStep;
            const isDone = idx < stageInfo.currentStep;
            const color = isCurrent ? '#0078D4' : isDone ? '#107C10' : '#8A8886';
            const bg = isCurrent ? '#DEECF9' : isDone ? '#DFF6DD' : '#F3F2F1';
            return (
              <React.Fragment key={s}>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 3 }}>
                  <div style={{
                    width: 22, height: 22, borderRadius: '50%', background: bg, color,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    fontSize: 11, fontWeight: 700, border: `1.5px solid ${color}`,
                  }}>
                    {isDone ? <Icon iconName="CheckMark" style={{ fontSize: 11 }} /> : idx + 1}
                  </div>
                  <span style={{ fontSize: 11, fontWeight: isCurrent ? 700 : 500, color }}>
                    {s}
                  </span>
                </div>
                {idx < STAGES.length - 1 && (
                  <div style={{
                    flex: 1, height: 2, margin: '0 6px',
                    background: idx < stageInfo.currentStep ? '#107C10' : '#EDEBE9',
                  }} />
                )}
              </React.Fragment>
            );
          })}
        </div>

        {/* Request & Requester Summary Card */}
        <div style={{
          background: '#FFFFFF', border: '1px solid #E1DFDD', borderRadius: 8,
          padding: '14px 18px', marginBottom: 18, boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
        }}>
          <div style={{ fontSize: 15, fontWeight: 700, color: '#201F1E', marginBottom: 12 }}>
            {title || ref}
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px 16px' }}>
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Requester</div>
              <div style={{ fontSize: 13, fontWeight: 600, color: '#323130' }}>{reqName}</div>
            </div>
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Business Unit</div>
              <div style={{ fontSize: 13, fontWeight: 600, color: '#323130' }}>{bu}</div>
            </div>
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Estimated Value</div>
              <div style={{ fontSize: 14, fontWeight: 700, color: '#107C10' }}>{amountFormatted}</div>
            </div>
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Site / Location</div>
              <div style={{ fontSize: 13, color: '#323130' }}>{site}</div>
            </div>
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Purchase Type</div>
              <div style={{ fontSize: 13, color: '#323130' }}>{ptype}</div>
            </div>
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Urgency</div>
              <div style={{
                fontSize: 12, fontWeight: 700,
                color: urgency === 'Emergency' ? '#A4262C' : urgency === 'Urgent' ? '#C8961F' : '#107C10',
              }}>
                {urgency}
              </div>
            </div>
            {costCenter && costCenter !== '—' && (
              <div>
                <div style={{ fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Cost Center</div>
                <div style={{ fontSize: 13, color: '#323130' }}>{costCenter}</div>
              </div>
            )}
            {dateNeeded && (
              <div>
                <div style={{ fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Date Needed</div>
                <div style={{ fontSize: 13, color: '#323130' }}>{formatDate(dateNeeded)}</div>
              </div>
            )}
          </div>

          {desc && (
            <div style={{ marginTop: 12, paddingTop: 10, borderTop: '1px solid #F3F2F1' }}>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#8A8886', textTransform: 'uppercase', letterSpacing: '0.4px', marginBottom: 4 }}>
                Requirement Justification
              </div>
              <div style={{ fontSize: 12, color: '#494543', lineHeight: 1.5, background: '#F8F9FA', padding: '8px 10px', borderRadius: 4 }}>
                {desc}
              </div>
            </div>
          )}
        </div>

        {/* Previous Approval History Section */}
        <div style={{
          border: '1px solid #EDEBE9', borderRadius: 8, marginBottom: 18, overflow: 'hidden',
        }}>
          <button
            onClick={() => this.setState({ showHistory: !showHistory })}
            style={{
              width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'space-between',
              padding: '9px 14px', background: '#FAF9F8', border: 'none', cursor: 'pointer',
              textAlign: 'left',
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <Icon iconName="History" style={{ color: '#605E5C', fontSize: 13 }} />
              <span style={{ fontSize: 12, fontWeight: 700, color: '#323130' }}>
                Previous Approval History
              </span>
              <span style={{
                fontSize: 11, background: '#EDEBE9', color: '#605E5C', fontWeight: 600,
                padding: '1px 7px', borderRadius: 10, marginLeft: 4,
              }}>
                {historyLogs.length}
              </span>
            </div>
            <Icon iconName={showHistory ? 'ChevronUp' : 'ChevronDown'} style={{ fontSize: 11, color: '#605E5C' }} />
          </button>

          {showHistory && (
            <div style={{ padding: '10px 14px', background: '#FFFFFF', maxHeight: 260, overflowY: 'auto' }}>
              {loadingHistory && (
                <div style={{ padding: '8px 0', textAlign: 'center' }}>
                  <Spinner size={SpinnerSize.small} label="Loading approval history…" />
                </div>
              )}

              {historyError && (
                <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 4 } }}>
                  Could not load approval logs: {historyError}
                </MessageBar>
              )}

              {!loadingHistory && !historyError && historyLogs.length === 0 && (
                <div style={{ fontSize: 12, color: '#605E5C', fontStyle: 'italic', padding: '4px 0' }}>
                  No previous approvals recorded. This request is at the initial approval stage.
                </div>
              )}

              {!loadingHistory && historyLogs.map((log, idx) => {
                const logTierRaw = log['sct_tiername'];
                const logTierName = logTierRaw === 4 ? 'Manager'
                  : logTierRaw === 1 ? 'Tier 1'
                    : logTierRaw === 2 ? 'Tier 2'
                      : logTierRaw === 3 ? 'Tier 3'
                        : strVal(log, 'sct_tiername') || 'Approver';

                const statusRaw = log['sct_approvalstatus'];
                const isApproved = statusRaw === 1;
                const isRejected = statusRaw === 2;
                const statusBadgeBg = isApproved ? '#DFF6DD' : isRejected ? '#FDE7E9' : '#FFF4CE';
                const statusBadgeColor = isApproved ? '#107C10' : isRejected ? '#A4262C' : '#7A4F01';
                const statusText = isApproved ? 'Approved' : isRejected ? 'Rejected' : 'Returned';

                const approver = strVal(log, '_sct_approverid_value') || strVal(log, '_createdby_value') || 'Approver';
                const commentsText = strVal(log, 'sct_comments') || '—';
                const rawDate = (log['createdon'] ?? log['sct_decisionon'] ?? '') as string;
                const dateText = formatDateTime(rawDate);

                return (
                  <div key={idx} style={{
                    display: 'flex', flexDirection: 'column', gap: 4,
                    padding: '8px 10px', borderRadius: 6, background: '#F8F9FA',
                    border: '1px solid #EDEBE9', marginBottom: idx < historyLogs.length - 1 ? 8 : 0,
                  }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <span style={{ fontSize: 12, fontWeight: 700, color: '#201F1E' }}>{approver}</span>
                        <span style={{ fontSize: 10, background: '#E1DFDD', color: '#323130', padding: '1px 6px', borderRadius: 4, fontWeight: 600 }}>
                          {logTierName}
                        </span>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <span style={{
                          fontSize: 10, fontWeight: 700, padding: '2px 8px', borderRadius: 10,
                          background: statusBadgeBg, color: statusBadgeColor,
                        }}>
                          {statusText}
                        </span>
                        <span style={{ fontSize: 11, color: '#8A8886' }}>{dateText}</span>
                      </div>
                    </div>
                    {commentsText && commentsText !== '—' && (
                      <div style={{ fontSize: 12, color: '#494543', fontStyle: 'italic', paddingLeft: 4 }}>
                        &ldquo;{commentsText}&rdquo;
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* Decision Comments Field */}
        <TextField
          label="Decision Comments"
          required={canAct}
          multiline
          rows={3}
          placeholder={canAct ? "Enter your rationale, conditions, or instructions for this decision…" : "Decision comments are locked because action has already been taken or this stage is inactive."}
          value={comments}
          onChange={(_, v) => this.setState({ comments: v ?? '', commentsError: null })}
          errorMessage={commentsError || undefined}
          disabled={!canAct}
          description={canAct ? "Mandatory for Reject and Return for Correction decisions." : undefined}
          styles={{ root: { marginBottom: 12 }, fieldGroup: { borderRadius: 6 } }}
        />

        {/* Decision Action Buttons are rendered in Panel onRenderFooterContent */}
      </Panel>
    );
  }
}
