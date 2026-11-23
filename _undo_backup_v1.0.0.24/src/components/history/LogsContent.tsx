/**
 * LogsContent — accordion Approval + Audit log sections with no Panel wrapper.
 * Used by RequestLogsPanel (wraps in a Panel) and MyRequests detail view (inline).
 */
import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';

import { DataverseService, ODataRecord, strVal } from '../../services/DataverseService';
import { formatDateTime } from '../dashboard/dashboardUtils';

// ── Props ──────────────────────────────────────────────────────────────────────

export interface ILogsContentProps {
  requestId: string;
  service: DataverseService;
}

// ── Choice maps ───────────────────────────────────────────────────────────────

const APPROVAL_STATUS: Record<number, { label: string; color: string; bg: string }> = {
  1: { label: 'Approved', color: '#107C10', bg: '#DFF6DD' },
  2: { label: 'Rejected', color: '#A4262C', bg: '#FDE7E9' },
  3: { label: 'Return for Info', color: '#905400', bg: '#FFF4CE' },
};
// sct_tiername Edm.Int32 confirmed from Dataverse — Accepted Values: 1=Tier 1, 2=Tier 2, 3=Tier 3, 4=Manager
const APPROVAL_TIER: Record<number, string> = { 1: 'Tier 1', 2: 'Tier 2', 3: 'Tier 3', 4: 'Manager' };

const AUDIT_ACTION: Record<number, { label: string; color: string; bg: string }> = {
  1: { label: 'Submitted', color: '#004578', bg: '#DBEAFE' },
  2: { label: 'Edited', color: '#00897B', bg: '#D1FAE5' },
  3: { label: 'Draft', color: '#605E5C', bg: '#F3F2F1' },
  4: { label: 'Returned', color: '#905400', bg: '#FFF4CE' },
  5: { label: 'Rejected', color: '#A4262C', bg: '#FDE7E9' },
};

// ── Helpers ───────────────────────────────────────────────────────────────────

function fmtDateTime(record?: ODataRecord, fallback?: string): string {
  if (!record && !fallback) return '—';
  // CRITICAL: Prioritize raw 'createdon' (or 'sct_decisionon') EDM timestamp.
  // Never read @OData.Community.Display.V1.FormattedValue for dates because Dataverse
  // formats it on the server in UTC and in US MM/dd/yyyy format.
  const rawIso = record
    ? (record['createdon'] ?? record['sct_decisionon'] ?? fallback ?? '')
    : (fallback ?? '');
  return formatDateTime(String(rawIso));
}

function resolveChoice(item: ODataRecord, field: string, map: Record<number, string>): string {
  const raw = item[field];
  if (typeof raw === 'number' && map[raw]) return map[raw];
  if (typeof raw === 'string' && /^\d+$/.test(raw) && map[Number(raw)]) return map[Number(raw)];
  const fmt = strVal(item, field);
  if (fmt) return fmt;
  if (typeof raw === 'number') return map[raw] ?? String(raw);
  return '—';
}

// ── Micro-components ──────────────────────────────────────────────────────────

const Badge: React.FC<{ label: string; color: string; bg: string }> = ({ label, color, bg }) => (
  <span style={{
    display: 'inline-block', padding: '2px 10px', borderRadius: 10,
    fontSize: 11, fontWeight: 700, background: bg, color, whiteSpace: 'nowrap',
  }}>
    {label}
  </span>
);

const DetailRow: React.FC<{ label: string; value: React.ReactNode }> = ({ label, value }) => (
  <div style={{ display: 'flex', gap: 8, marginBottom: 6, alignItems: 'baseline', flexWrap: 'wrap' }}>
    <span style={{
      fontSize: 10, fontWeight: 700, color: '#A19F9D',
      textTransform: 'uppercase', letterSpacing: '0.6px', flexShrink: 0, minWidth: 88,
    }}>{label}</span>
    <span style={{ fontSize: 13, color: '#201F1E' }}>{value || '—'}</span>
  </div>
);

// ── Inner row accordion ───────────────────────────────────────────────────────

interface IRowAccordionProps {
  header: React.ReactNode;
  children: React.ReactNode;
  isOpen: boolean;
  onToggle: () => void;
  accentColor: string;
}

const RowAccordion: React.FC<IRowAccordionProps> = ({ header, children, isOpen, onToggle, accentColor }) => (
  <div style={{
    border: '1px solid',
    borderColor: isOpen ? accentColor : '#EDEBE9',
    borderRadius: 8, marginBottom: 6, overflow: 'hidden',
    boxShadow: isOpen ? '0 2px 8px rgba(0,0,0,0.07)' : '0 1px 2px rgba(0,0,0,0.04)',
  }}>
    <button
      onClick={onToggle}
      style={{
        width: '100%', display: 'flex', alignItems: 'center', gap: 10,
        padding: '9px 12px', background: isOpen ? '#F8FFFE' : '#FAFAFA',
        border: 'none', cursor: 'pointer', textAlign: 'left', fontFamily: 'inherit',
      }}
    >
      <Icon iconName={isOpen ? 'ChevronDown' : 'ChevronRight'} style={{ fontSize: 11, color: accentColor, flexShrink: 0 }} />
      <div style={{ flex: 1, minWidth: 0 }}>{header}</div>
    </button>
    {isOpen && (
      <div style={{ padding: '10px 14px 12px 36px', borderTop: '1px solid #EDEBE9', background: '#FFFFFF', borderLeft: `3px solid ${accentColor}` }}>
        {children}
      </div>
    )}
  </div>
);

// ── Section header (top-level collapsible) ────────────────────────────────────

interface ISectionHeaderProps {
  title: string;
  count: number | undefined;
  loading: boolean;
  isOpen: boolean;
  onToggle: () => void;
  color: string;
  bg: string;
  iconName: string;
}

const SectionHeader: React.FC<ISectionHeaderProps> = ({
  title, count, loading, isOpen, onToggle, color, bg, iconName,
}) => (
  <button
    onClick={onToggle}
    style={{
      width: '100%', display: 'flex', alignItems: 'center', gap: 12,
      padding: '12px 16px', background: isOpen ? bg : '#F3F2F1',
      border: '1px solid', borderColor: isOpen ? color : '#EDEBE9',
      borderRadius: isOpen ? '8px 8px 0 0' : 8,
      cursor: 'pointer', textAlign: 'left', fontFamily: 'inherit',
      transition: 'background 0.15s, border-color 0.15s',
      boxShadow: isOpen ? '0 2px 6px rgba(0,0,0,0.06)' : 'none',
    }}
  >
    <Icon iconName={iconName} style={{ fontSize: 18, color: isOpen ? color : '#605E5C', flexShrink: 0 }} />
    <span style={{ flex: 1, fontSize: 14, fontWeight: 700, color: isOpen ? color : '#201F1E' }}>{title}</span>
    {loading && <Spinner size={SpinnerSize.xSmall} />}
    {!loading && count !== undefined && (
      <span style={{
        background: isOpen ? color : '#C8C6C4', color: '#fff',
        fontSize: 11, fontWeight: 700, padding: '2px 8px', borderRadius: 10,
        minWidth: 20, textAlign: 'center', display: 'inline-block',
      }}>
        {count}
      </span>
    )}
    <Icon iconName={isOpen ? 'ChevronDown' : 'ChevronRight'} style={{ fontSize: 13, color: isOpen ? color : '#A19F9D', flexShrink: 0 }} />
  </button>
);

// ── Loading / empty ───────────────────────────────────────────────────────────

const Loader: React.FC = () => (
  <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '36px 16px' }}>
    <Spinner size={SpinnerSize.medium} />
    <span style={{ fontSize: 13, color: '#605E5C' }}>Loading…</span>
  </div>
);

const Empty: React.FC<{ message: string }> = ({ message }) => (
  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8, padding: '40px 0', textAlign: 'center' }}>
    <Icon iconName="EntitlementPolicy" style={{ fontSize: 36, color: '#C8C6C4' }} />
    <span style={{ fontSize: 13, color: '#605E5C' }}>{message}</span>
  </div>
);

// ── Approval section ──────────────────────────────────────────────────────────

interface IApprovalSectionProps {
  rows: ODataRecord[];
  loading: boolean;
  error: string | null;
  isOpen: boolean;
  onToggleSection: () => void;
}

const ApprovalSection: React.FC<IApprovalSectionProps> = ({ rows, loading, error, isOpen, onToggleSection }) => {
  const [expanded, setExpanded] = React.useState<Record<number, boolean>>({});

  const toggle = (i: number): void =>
    setExpanded(prev => {
      const next: Record<number, boolean> = {};
      Object.keys(prev).forEach(k => { next[Number(k)] = prev[Number(k)]; });
      next[i] = !prev[i];
      return next;
    });

  const body = (): React.ReactElement => {
    if (loading) return <Loader />;
    if (error) return <MessageBar messageBarType={MessageBarType.error} styles={{ root: { borderRadius: 6 } }}>{error}</MessageBar>;
    if (!rows.length) return <Empty message="No approval log entries for this request." />;
    return (
      <div>
        {rows.map((row, i) => {
          const logName = strVal(row, 'sct_name') || '—';
          const statusRaw = row['sct_approvalstatus'] as number;
          const statusDef = APPROVAL_STATUS[statusRaw] ?? { label: strVal(row, 'sct_approvalstatus') || 'Unknown', color: '#605E5C', bg: '#F3F2F1' };
          const tier = resolveChoice(row, 'sct_tiername', APPROVAL_TIER);
          const approver = strVal(row, '_sct_approverid_value') || '—';
          const date = fmtDateTime(row);
          const comments = strVal(row, 'sct_comments');

          const header = (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              <span style={{ width: 8, height: 8, borderRadius: '50%', background: statusDef.color, flexShrink: 0 }} />
              <span style={{ fontSize: 11, fontFamily: 'monospace', fontWeight: 700, color: '#0078D4', flexShrink: 0 }}>{logName}</span>
              <span style={{ fontSize: 13, fontWeight: 600, color: '#201F1E', flex: 1, minWidth: 0 }}>{approver}</span>
              <Badge label={tier} color="#0078D4" bg="#EBF3FB" />
              <span style={{ fontSize: 11, color: '#A19F9D', whiteSpace: 'nowrap' }}>{date}</span>
            </div>
          );

          return (
            <RowAccordion key={i} isOpen={!!expanded[i]} onToggle={() => toggle(i)} accentColor={statusDef.color} header={header}>
              <DetailRow label="Log No." value={<span style={{ fontFamily: 'monospace', fontWeight: 700, color: '#0078D4' }}>{logName}</span>} />
              <DetailRow label="Tier" value={<span style={{ fontWeight: 600, color: '#0078D4' }}>{tier}</span>} />
              <DetailRow label="Approver" value={approver} />
              <DetailRow label="Status" value={<Badge label={statusDef.label} color={statusDef.color} bg={statusDef.bg} />} />
              <DetailRow label="Date" value={date} />
              {comments ? (
                <div style={{ marginTop: 8 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', marginBottom: 5 }}>Comments</div>
                  <div style={{ fontSize: 12, color: '#201F1E', lineHeight: '1.6', background: '#F8F8F8', border: '1px solid #EDEBE9', borderRadius: 6, padding: '8px 12px' }}>
                    {comments}
                  </div>
                </div>
              ) : (
                <DetailRow label="Comments" value="—" />
              )}
            </RowAccordion>
          );
        })}
      </div>
    );
  };

  return (
    <div style={{ marginBottom: 12 }}>
      <SectionHeader title="Approval Logs" count={loading ? undefined : rows.length} loading={loading} isOpen={isOpen} onToggle={onToggleSection} color="#107C10" bg="#DFF6DD" iconName="Checklist" />
      {isOpen && (
        <div style={{ border: '1px solid #107C10', borderTop: 'none', borderRadius: '0 0 8px 8px', padding: 14, background: '#FAFFFE' }}>
          {body()}
        </div>
      )}
    </div>
  );
};

// ── Audit section ─────────────────────────────────────────────────────────────

interface IAuditSectionProps {
  rows: ODataRecord[];
  loading: boolean;
  error: string | null;
  isOpen: boolean;
  onToggleSection: () => void;
}

const AuditSection: React.FC<IAuditSectionProps> = ({ rows, loading, error, isOpen, onToggleSection }) => {
  const [expanded, setExpanded] = React.useState<Record<number, boolean>>({});

  const toggle = (i: number): void =>
    setExpanded(prev => {
      const next: Record<number, boolean> = {};
      Object.keys(prev).forEach(k => { next[Number(k)] = prev[Number(k)]; });
      next[i] = !prev[i];
      return next;
    });

  const body = (): React.ReactElement => {
    if (loading) return <Loader />;
    if (error) return <MessageBar messageBarType={MessageBarType.error} styles={{ root: { borderRadius: 6 } }}>{error}</MessageBar>;
    if (!rows.length) return <Empty message="No audit log entries for this request." />;
    return (
      <div>
        {rows.map((row, i) => {
          const logName = strVal(row, 'sct_name') || '—';
          const actionRaw = row['sct_action'] as number;
          const desc = strVal(row, 'sct_description');
          let actionDef = AUDIT_ACTION[actionRaw] ?? { label: strVal(row, 'sct_action') || 'Audit', color: '#605E5C', bg: '#F3F2F1' };
          if (desc && desc.toLowerCase().indexOf('approved') !== -1) {
            actionDef = { label: 'Approved', color: '#107C10', bg: '#DFF6DD' };
          } else if (desc && desc.toLowerCase().indexOf('submitted') !== -1) {
            actionDef = { label: 'Submitted', color: '#004578', bg: '#DBEAFE' };
          }
          const actionBy = strVal(row, '_sct_actionby_value') || '—';
          const date = fmtDateTime(row);
          const oldVal = strVal(row, 'sct_oldvalue');
          const newVal = strVal(row, 'sct_newvalue');

          const header = (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              <Badge label={actionDef.label} color={actionDef.color} bg={actionDef.bg} />
              <span style={{ fontSize: 11, fontFamily: 'monospace', fontWeight: 700, color: '#8764B8', flexShrink: 0 }}>{logName}</span>
              <span style={{ fontSize: 13, fontWeight: 600, color: '#201F1E', flex: 1, minWidth: 0 }}>{actionBy}</span>
              <span style={{ fontSize: 11, color: '#A19F9D', whiteSpace: 'nowrap' }}>{date}</span>
            </div>
          );

          return (
            <RowAccordion key={i} isOpen={!!expanded[i]} onToggle={() => toggle(i)} accentColor={actionDef.color} header={header}>
              <DetailRow label="Audit No." value={<span style={{ fontFamily: 'monospace', fontWeight: 700, color: '#8764B8' }}>{logName}</span>} />
              <DetailRow label="Action" value={<Badge label={actionDef.label} color={actionDef.color} bg={actionDef.bg} />} />
              <DetailRow label="Action By" value={actionBy} />
              <DetailRow label="Date" value={date} />
              <DetailRow label="Description" value={desc} />
              {(oldVal || newVal) && (
                <div style={{ marginTop: 10 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.6px', marginBottom: 6 }}>Changes</div>
                  <div style={{ display: 'flex', borderRadius: 6, border: '1px solid #EDEBE9', overflow: 'hidden', fontSize: 12 }}>
                    <div style={{ flex: 1, padding: '8px 12px', background: '#FEF3F2', borderRight: '1px solid #EDEBE9' }}>
                      <div style={{ fontSize: 10, fontWeight: 700, color: '#A4262C', textTransform: 'uppercase', marginBottom: 4 }}>Before</div>
                      <div style={{ color: '#605E5C', wordBreak: 'break-word', lineHeight: '1.5' }}>{oldVal || '(empty)'}</div>
                    </div>
                    <div style={{ flex: 1, padding: '8px 12px', background: '#F0FDF4' }}>
                      <div style={{ fontSize: 10, fontWeight: 700, color: '#107C10', textTransform: 'uppercase', marginBottom: 4 }}>After</div>
                      <div style={{ color: '#201F1E', fontWeight: 500, wordBreak: 'break-word', lineHeight: '1.5' }}>{newVal || '(empty)'}</div>
                    </div>
                  </div>
                </div>
              )}
            </RowAccordion>
          );
        })}
      </div>
    );
  };

  return (
    <div style={{ marginBottom: 12 }}>
      <SectionHeader title="Audit Logs" count={loading ? undefined : rows.length} loading={loading} isOpen={isOpen} onToggle={onToggleSection} color="#8764B8" bg="#F4F0FF" iconName="History" />
      {isOpen && (
        <div style={{ border: '1px solid #8764B8', borderTop: 'none', borderRadius: '0 0 8px 8px', padding: 14, background: '#FDFCFF' }}>
          {body()}
        </div>
      )}
    </div>
  );
};

// ── Workflow Stepper ──────────────────────────────────────────────────────────

const WorkflowStepper: React.FC<{ rows: ODataRecord[] }> = ({ rows }) => {
  const isApproved = (tier: number, nameMatch: string): boolean =>
    rows.some(r => (r['sct_tiername'] === tier || String(r['sct_name'] ?? '').toLowerCase().indexOf(nameMatch) !== -1) &&
      (r['sct_approvalstatus'] === 1 || r['sct_approvalstatus'] === 854000000));

  const isRejected = rows.some(r => r['sct_approvalstatus'] === 2 || r['sct_approvalstatus'] === 854000001);
  const isReturned = rows.some(r => r['sct_approvalstatus'] === 3 || r['sct_approvalstatus'] === 854000002);

  const mgrDone = isApproved(4, 'manager');
  const t1Done = isApproved(1, 'tier 1');
  const t2Done = isApproved(2, 'tier 2');
  const t3Done = isApproved(3, 'tier 3');

  let activeStage = 'Manager';
  if (mgrDone && !t1Done) activeStage = 'Tier 1';
  else if (t1Done && !t2Done) activeStage = 'Tier 2';
  else if (t2Done && !t3Done) activeStage = 'Tier 3';
  else if (t3Done) activeStage = 'PO Issued';

  const stages = [
    { label: 'Submitted', done: true, active: false },
    { label: 'Manager', done: mgrDone, active: activeStage === 'Manager' && !isRejected && !isReturned },
    { label: 'Tier 1', done: t1Done, active: activeStage === 'Tier 1' && !isRejected && !isReturned },
    { label: 'Tier 2', done: t2Done, active: activeStage === 'Tier 2' && !isRejected && !isReturned },
    { label: 'Tier 3', done: t3Done, active: activeStage === 'Tier 3' && !isRejected && !isReturned },
    { label: 'PO Issued', done: t3Done, active: activeStage === 'PO Issued' && !isRejected && !isReturned },
  ];

  return (
    <div style={{
      background: '#ffffff',
      border: '1px solid #E2E8F0',
      borderRadius: 10,
      padding: '12px 14px',
      marginBottom: 14,
      boxShadow: '0 1px 4px rgba(15, 23, 42, 0.04)',
    }}>
      <div style={{ fontSize: 11, fontWeight: 700, color: '#475569', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 10 }}>
        Workflow Pipeline Progress
      </div>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 4 }}>
        {stages.map((st, i) => {
          const isCurrentActive = st.active;
          const isDone = st.done;
          return (
            <React.Fragment key={i}>
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', flex: 1, minWidth: 0 }}>
                <div style={{
                  width: 24,
                  height: 24,
                  borderRadius: '50%',
                  background: isDone ? '#107C10' : isCurrentActive ? '#0078D4' : '#F1F5F9',
                  color: isDone || isCurrentActive ? '#ffffff' : '#94A3B8',
                  border: isCurrentActive ? '2px solid #60A5FA' : '1px solid #CBD5E1',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  fontSize: 10,
                  fontWeight: 700,
                  boxShadow: isCurrentActive ? '0 0 0 3px rgba(0, 120, 212, 0.2)' : 'none',
                }}>
                  {isDone ? <Icon iconName="CheckMark" style={{ fontSize: 10 }} /> : i + 1}
                </div>
                <div style={{
                  fontSize: 10.5,
                  fontWeight: isCurrentActive ? 700 : isDone ? 600 : 500,
                  color: isDone ? '#107C10' : isCurrentActive ? '#0078D4' : '#64748B',
                  marginTop: 4,
                  textAlign: 'center',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap',
                }}>
                  {st.label}
                </div>
              </div>
              {i < stages.length - 1 && (
                <div style={{
                  flex: 1,
                  height: 2,
                  minWidth: 8,
                  background: st.done ? '#107C10' : '#E2E8F0',
                  marginBottom: 16,
                }} />
              )}
            </React.Fragment>
          );
        })}
      </div>
    </div>
  );
};

// ── Root: fetches data + renders both sections ─────────────────────────────────

const LogsContent: React.FC<ILogsContentProps> = ({ requestId, service }) => {
  const [approvalRows, setApprovalRows] = React.useState<ODataRecord[]>([]);
  const [auditRows, setAuditRows] = React.useState<ODataRecord[]>([]);
  const [loadingApproval, setLoadingApproval] = React.useState<boolean>(true);
  const [loadingAudit, setLoadingAudit] = React.useState<boolean>(true);
  const [approvalError, setApprovalError] = React.useState<string | null>(null);
  const [auditError, setAuditError] = React.useState<string | null>(null);
  const [approvalOpen, setApprovalOpen] = React.useState<boolean>(true);
  const [auditOpen, setAuditOpen] = React.useState<boolean>(true);

  React.useEffect(() => {
    let cancelled = false;

    const parseEpoch = (rec: ODataRecord): number => {
      let val = String(rec['createdon'] ?? rec['sct_decisionon'] ?? '').trim();
      if (!val) return 0;
      if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(val)) {
        val += 'Z';
      }
      const t = new Date(val).getTime();
      return isNaN(t) ? 0 : t;
    };

    const sortByNewest = (records: ODataRecord[]): ODataRecord[] => {
      return [...records].sort((a, b) => {
        const timeA = parseEpoch(a);
        const timeB = parseEpoch(b);
        if (timeB !== timeA) return timeB - timeA;
        return String(b['sct_name'] ?? '').localeCompare(String(a['sct_name'] ?? ''));
      });
    };

    const fetchApproval = async (): Promise<void> => {
      setLoadingApproval(true);
      setApprovalError(null);
      try {
        const rows = await service.getApprovalLogs(requestId);
        if (!cancelled) setApprovalRows(sortByNewest(Array.isArray(rows) ? rows : []));
      } catch (err: unknown) {
        if (!cancelled) setApprovalError(err instanceof Error ? err.message : 'Failed to load approval logs.');
      } finally {
        if (!cancelled) setLoadingApproval(false);
      }
    };

    const fetchAudit = async (): Promise<void> => {
      setLoadingAudit(true);
      setAuditError(null);
      try {
        const rows = await service.getAuditLogs(requestId);
        if (!cancelled) setAuditRows(sortByNewest(Array.isArray(rows) ? rows : []));
      } catch (err: unknown) {
        if (!cancelled) setAuditError(err instanceof Error ? err.message : 'Failed to load audit logs.');
      } finally {
        if (!cancelled) setLoadingAudit(false);
      }
    };

    void fetchApproval();
    void fetchAudit();
    return () => { cancelled = true; };
  }, [requestId]);

  const totalLoaded = !loadingApproval && !loadingAudit;
  return (
    <div>
      {totalLoaded && (
        <>
          <WorkflowStepper rows={approvalRows} />
          <div style={{
            display: 'flex', gap: 20, padding: '10px 14px',
            background: '#F8F9FA', borderRadius: 8, marginBottom: 14,
            border: '1px solid #EDEBE9', flexWrap: 'wrap', alignItems: 'center',
          }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
            <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#107C10', flexShrink: 0, display: 'inline-block' }} />
            <span style={{ fontSize: 12, color: '#323130', fontWeight: 600 }}>
              {approvalRows.length} Approval {approvalRows.length === 1 ? 'Entry' : 'Entries'}
            </span>
          </div>
          <div style={{ width: 1, alignSelf: 'stretch', background: '#EDEBE9', flexShrink: 0 }} />
          <div style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
            <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#8764B8', flexShrink: 0, display: 'inline-block' }} />
            <span style={{ fontSize: 12, color: '#323130', fontWeight: 600 }}>
              {auditRows.length} Audit {auditRows.length === 1 ? 'Entry' : 'Entries'}
            </span>
          </div>
        </div>
      </>
      )}
      <ApprovalSection
        rows={approvalRows} loading={loadingApproval} error={approvalError}
        isOpen={approvalOpen} onToggleSection={() => setApprovalOpen(o => !o)}
      />
      <AuditSection
        rows={auditRows} loading={loadingAudit} error={auditError}
        isOpen={auditOpen} onToggleSection={() => setAuditOpen(o => !o)}
      />
    </div>
  );
};

export default LogsContent;
