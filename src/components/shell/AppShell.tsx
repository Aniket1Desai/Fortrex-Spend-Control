import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { IRenderFunction } from '@fluentui/react/lib/Utilities';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import { IAppShellProps } from './IAppShellProps';
import styles from './AppShell.module.scss';

import Dashboard from '../dashboard/Dashboard';
import NewRequestForm from '../newRequest/NewRequestForm';
import MyRequests from '../myRequests/MyRequests';
import ReviewQueue from '../reviewQueue/ReviewQueue';

type TabId = 'dashboard' | 'newRequest' | 'myRequests' | 'reviewQueue';

interface ITab {
  id: TabId;
  label: string;
  iconName: string;
  badge?: number;
  ariaLabel: string;
}

const TABS: ITab[] = [
  { id: 'dashboard',    label: 'Dashboard',     iconName: 'ViewDashboard',   ariaLabel: 'Dashboard tab' },
  { id: 'newRequest',   label: 'New Request',   iconName: 'Add',             ariaLabel: 'New Request tab' },
  { id: 'myRequests',   label: 'My Requests',   iconName: 'DocumentSearch',  ariaLabel: 'My Requests tab' },
  { id: 'reviewQueue',  label: 'Review Queue',  iconName: 'Checklist',       ariaLabel: 'Review Queue tab' },
];

const TIER1_KEY = 'approver_tier1';

// ── Azure AD Security Group IDs (confirmed from Entra admin center) ───────────
const TIER1_GROUP_ID = 'fdbf07a0-548b-4ccf-875b-900cb1b24a1c'; // SCT-Fortrex_Approvers-Tier1
const TIER2_GROUP_ID = '6a8a4077-c78f-4af2-ae5e-12428d1dcbc3'; // SCT-Fortrex_Approvers-Tier2
const TIER3_GROUP_ID = 'e63dbc4f-6ad2-4969-a518-b229836756af'; // SCT-Fortrex_Approvers-Tier3

interface IAppShellState {
  activeTab: TabId;
  activeRole: string;
  tier1PendingCount: number;
  rolePendingCounts: Record<string, number>;
  roleOptions: IDropdownOption[];
  rolesLoading: boolean;
  userRoles?: string[];
  soloTestMode: boolean;
  editingRequest: import('../../services/DataverseService').ODataRecord | null;
  deepLinkRequestId: string | null;
  isFormDirty: boolean;
  showLeaveConfirmDialog: boolean;
  pendingTabSwitch: TabId | null;
  pendingRoleSwitch: string | null;
}

export default class AppShell extends React.Component<IAppShellProps, IAppShellState> {
  constructor(props: IAppShellProps) {
    super(props);
    let defaultTestMode = true;
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        const stored = window.localStorage.getItem('sct_solo_test_mode');
        if (stored !== null) defaultTestMode = stored === 'true';
      }
    } catch { /* ignore */ }

    this.state = {
      activeTab: 'dashboard',
      activeRole: 'requester',
      tier1PendingCount: 0,
      rolePendingCounts: {},
      roleOptions: [{ key: 'requester', text: 'Requester' }],
      rolesLoading: true,
      userRoles: ['requester'],
      soloTestMode: defaultTestMode,
      editingRequest: null,
      deepLinkRequestId: null,
      isFormDirty: false,
      showLeaveConfirmDialog: false,
      pendingTabSwitch: null,
      pendingRoleSwitch: null,
    };
  }

  public componentDidMount(): void {
    void this._loadUserRoles();
    if (!this.props.serviceError) {
      void this._loadAllPendingCounts();
    }
    // Deep-link auto-routing (e.g. from Outlook approval email: ?requestId=SCT-2026-xxxx)
    try {
      if (typeof window !== 'undefined' && window.location && window.location.search) {
        const params = new URLSearchParams(window.location.search);
        const reqParam = params.get('requestId') || params.get('sctRef') || params.get('id');
        if (reqParam) {
          this.setState({ deepLinkRequestId: reqParam.trim() });
        }
      }
    } catch { /* ignore */ }
  }

  private async _loadUserRoles(): Promise<void> {
    const ctx = this.props.context;
    const aadOid = ctx.pageContext.aadInfo?.userId?.toString()?.toLowerCase() ?? '';
    const normAadOid = aadOid.replace(/[{}-]/g, '');
    const userEmail = (ctx.pageContext.user?.email ?? '').toLowerCase().trim();
    const userDisplayName = (ctx.pageContext.user?.displayName ?? '').toLowerCase().trim();

    try {
      // ── Step 1: Resolve Dataverse systemuser record ──────────────────────────
      const systemUser = await this.props.dataverseService.getCurrentSystemUser(aadOid)
        ?? (userEmail ? await this.props.dataverseService.getSystemUserByEmail(userEmail) : null);

      const systemUserId = systemUser?.systemuserid ?? null;

      // ── Step 2: Query Azure AD & Dataverse dynamically for Manager status ───
      let adIsManager = false;
      let graphClient: MSGraphClientV3 | undefined;
      try {
        graphClient = await ctx.msGraphClientFactory.getClient('3');
      } catch { /* optional Graph client */ }

      if (graphClient) {
        try {
          const reportsResponse = await graphClient
            .api('/me/directReports')
            .select('id')
            .top(1)
            .get() as { value: unknown[] };
          if ((reportsResponse?.value || []).length > 0) {
            adIsManager = true;
          }
        } catch { /* ignore */ }
      }

      let dataverseIsManager = false;
      if (systemUserId) {
        dataverseIsManager = await this.props.dataverseService.hasDirectReportsInDataverse(systemUserId);
        if (!dataverseIsManager) {
          const mgrPending = await this.props.dataverseService.getManagerPendingCount(systemUserId);
          if (mgrPending > 0) dataverseIsManager = true;
        }
      }
      const isManager = adIsManager || dataverseIsManager;

      // ── Step 3: Query Azure AD Security Groups dynamically ───────────────────
      const userGroupIds = new Set<string>();

      if (graphClient) {
        // Query 1: /me/transitiveMemberOf (recursively checks direct and nested security groups)
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const transResp: any = await graphClient
            .api('/me/transitiveMemberOf')
            .select('id,displayName')
            .top(999)
            .get();
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (transResp?.value || []).forEach((g: any) => {
            if (g?.id) userGroupIds.add(String(g.id).toLowerCase());
          });
        } catch (err) {
          console.warn('[AppShell] Graph transitiveMemberOf check warning:', err);
        }

        // Query 2: /me/checkMemberGroups (fast check against exact group IDs)
        try {
          const checkResp = await graphClient.api('/me/checkMemberGroups').post({
            groupIds: [TIER1_GROUP_ID, TIER2_GROUP_ID, TIER3_GROUP_ID],
          }) as { value: string[] };
          (checkResp?.value || []).forEach(id => {
            if (id) userGroupIds.add(String(id).toLowerCase());
          });
        } catch (err) {
          console.warn('[AppShell] Graph checkMemberGroups check warning:', err);
        }
      }

      // Query 3: Dataverse sct_roles table (if roles are assigned in Dataverse)
      let dvTier1 = false;
      let dvTier2 = false;
      let dvTier3 = false;
      if (systemUserId) {
        try {
          const roleConfigs = await this.props.dataverseService.getRoleConfigs();
          const sId = systemUserId.toLowerCase().replace(/[{}-]/g, '');
          const matchingRoles = roleConfigs.filter(r => {
            const pId = String(r['_sct_personid_value'] ?? '').toLowerCase().replace(/[{}-]/g, '');
            const dId = String(r['_sct_delegateid_value'] ?? '').toLowerCase().replace(/[{}-]/g, '');
            return pId === sId || dId === sId || pId === normAadOid || dId === normAadOid;
          });
          matchingRoles.forEach(r => {
            const lvl = String(r['sct_approvallevel'] ?? '').toUpperCase();
            if (lvl === 'L1') dvTier1 = true;
            if (lvl === 'L2') dvTier2 = true;
            if (lvl === 'L3' || lvl === 'L4') dvTier3 = true;
          });
        } catch { /* optional sct_roles query */ }
      }

      // Final Dynamic Tier Roles: true if user is member of AD Security Group or Dataverse role
      const isTier1 = userGroupIds.has(TIER1_GROUP_ID.toLowerCase()) || dvTier1;
      const isTier2 = userGroupIds.has(TIER2_GROUP_ID.toLowerCase()) || dvTier2;
      const isTier3 = userGroupIds.has(TIER3_GROUP_ID.toLowerCase()) || dvTier3;

      // ── Step 4: Construct Dynamic Role Options (100% derived from Azure AD & Dataverse) ──
      const opts: IDropdownOption[] = [{ key: 'requester', text: 'Requester' }];
      if (isManager) opts.push({ key: 'manager', text: 'Manager' });
      if (isTier1)   opts.push({ key: TIER1_KEY, text: 'Approver — Tier 1' });
      if (isTier2)   opts.push({ key: 'approver_tier2', text: 'Approver — Tier 2' });
      if (isTier3)   opts.push({ key: 'approver_tier3', text: 'Approver — Tier 3' });

      // Strictly deduplicate options by key to prevent any duplicate dropdown items
      const seenKeys = new Set<string>();
      const uniqueOpts: IDropdownOption[] = [];
      for (const opt of opts) {
        const k = String(opt.key);
        if (!seenKeys.has(k)) {
          seenKeys.add(k);
          uniqueOpts.push(opt);
        }
      }

      // If current activeRole is not in the allowed roles, reset to Requester
      const validKeys = new Set(uniqueOpts.map(o => o.key as string));
      const newActiveRole = validKeys.has(this.state.activeRole) ? this.state.activeRole : 'requester';

      this.setState({
        roleOptions: uniqueOpts,
        userRoles: uniqueOpts.map(o => o.key as string),
        activeRole: newActiveRole,
        rolesLoading: false,
      });

      // Refresh pending counts for all assigned roles
      void this._loadAllPendingCounts(systemUserId ?? undefined);
    } catch {
      this.setState({ rolesLoading: false });
    }
  }

  /** Loads pending counts across all tiers and manager to display notification badges */
  private async _loadAllPendingCounts(systemUserId?: string): Promise<void> {
    const counts: Record<string, number> = {};
    try {
      const [t1Count, t2Count, t3Count, mgrCount] = await Promise.all([
        this.props.dataverseService.getTierPendingRequestIds(1).then(ids => ids.length).catch(() => 0),
        this.props.dataverseService.getTierPendingRequestIds(2).then(ids => ids.length).catch(() => 0),
        this.props.dataverseService.getTierPendingRequestIds(3).then(ids => ids.length).catch(() => 0),
        systemUserId ? this.props.dataverseService.getManagerPendingCount(systemUserId) : 0,
      ]);
      counts['manager'] = mgrCount;
      counts[TIER1_KEY] = t1Count;
      counts['approver_tier2'] = t2Count;
      counts['approver_tier3'] = t3Count;

      this.setState({
        tier1PendingCount: t1Count,
        rolePendingCounts: counts,
      });
    } catch { /* silently ignore */ }
  }

  // ── Role Switcher badge renderer ─────────────────────────────────────────────
  private _onRenderRoleOption: IRenderFunction<IDropdownOption> = (option, defaultRender) => {
    if (!option) return null;
    const count = this.state.rolePendingCounts[option.key as string] || (option.key === TIER1_KEY ? this.state.tier1PendingCount : 0);

    if (count > 0) {
      return (
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', width: '100%', gap: 8 }}>
          <span style={{ flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
            {option.text}
          </span>
          <span style={{
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            minWidth: 20, height: 20, padding: '0 6px',
            borderRadius: 10,
            background: '#D13438',
            color: '#ffffff',
            fontSize: 11, fontWeight: 700, lineHeight: 1,
            flexShrink: 0,
          }}>
            {count > 99 ? '99+' : count}
          </span>
        </div>
      );
    }

    return defaultRender ? defaultRender(option) : <span>{option.text}</span>;
  };

  // ── Selected-item renderer — mirrors badge in the collapsed trigger ──────────
  private _onRenderRoleTitle: IRenderFunction<IDropdownOption[]> = (options, defaultRender) => {
    if (!options || !options.length) return null;
    const option = options[0];
    const count = this.state.rolePendingCounts[option.key as string] || (option.key === TIER1_KEY ? this.state.tier1PendingCount : 0);

    if (count > 0) {
      return (
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <span>{option.text}</span>
          <span style={{
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            minWidth: 18, height: 18, padding: '0 5px',
            borderRadius: 9,
            background: '#D13438', color: '#ffffff',
            fontSize: 10, fontWeight: 700, lineHeight: 1,
            flexShrink: 0,
          }}>
            {count > 99 ? '99+' : count}
          </span>
        </div>
      );
    }

    return defaultRender ? defaultRender(options) : <span>{option.text}</span>;
  };

  private _getInitials(displayName: string): string {
    return displayName
      .split(' ')
      .slice(0, 2)
      .map(part => part.charAt(0).toUpperCase())
      .join('');
  }

  private _handleTabClick = (targetTab: TabId): void => {
    if (this.state.activeTab === targetTab) return;

    // Check if user has entered data in the new request form ("on half way")
    if (this.state.activeTab === 'newRequest' && this.state.isFormDirty) {
      this.setState({
        showLeaveConfirmDialog: true,
        pendingTabSwitch: targetTab,
        pendingRoleSwitch: null,
      });
      return;
    }

    this.setState({
      activeTab: targetTab,
      editingRequest: targetTab === 'newRequest' ? null : this.state.editingRequest,
    });
  };

  private _handleRoleChange = (option?: IDropdownOption): void => {
    if (!option) return;
    const newRole = option.key as string;
    const currentTab = this.state.activeTab;
    let newTab: TabId = currentTab;
    // If switching to Requester while on Review Queue, redirect to dashboard.
    // (New Request is now available for all roles).
    if (newRole === 'requester' && currentTab === 'reviewQueue') newTab = 'dashboard';

    // If switching role while new request form is dirty
    if (currentTab === 'newRequest' && this.state.isFormDirty && newTab !== 'newRequest') {
      this.setState({
        showLeaveConfirmDialog: true,
        pendingTabSwitch: newTab,
        pendingRoleSwitch: newRole,
      });
      return;
    }

    this.setState({ activeRole: newRole, activeTab: newTab });
  };

  private _confirmLeaveTab = (): void => {
    const { pendingTabSwitch, pendingRoleSwitch } = this.state;
    this.setState({
      showLeaveConfirmDialog: false,
      isFormDirty: false,
      activeTab: pendingTabSwitch || 'dashboard',
      activeRole: pendingRoleSwitch || this.state.activeRole,
      pendingTabSwitch: null,
      pendingRoleSwitch: null,
      editingRequest: null,
    });
  };

  private _cancelLeaveConfirm = (): void => {
    this.setState({
      showLeaveConfirmDialog: false,
      pendingTabSwitch: null,
      pendingRoleSwitch: null,
    });
  };

  private _renderTabContent(): React.ReactElement {
    // When the Dataverse tables are unreachable, render nothing — no data fetches
    if (this.props.serviceError) {
      return (
        <div style={{
          display: 'flex', flexDirection: 'column', alignItems: 'center',
          justifyContent: 'center', height: '100%', gap: 14, padding: 40,
          color: '#605E5C', textAlign: 'center',
        }}>
          <Icon iconName="PlugDisconnected" style={{ fontSize: 48, color: '#C8C6C4' }} />
          <div style={{ fontSize: 16, fontWeight: 600, color: '#323130' }}>
            No data available
          </div>
          <div style={{ fontSize: 13, maxWidth: 420, lineHeight: 1.6 }}>
            This web part cannot connect to the Dataverse environment.
            Open the property pane, verify the URL, and ensure the Fortrex SCT solution is imported and published.
          </div>
        </div>
      );
    }

    const { dataverseService, context } = this.props;
    switch (this.state.activeTab) {
      case 'dashboard':
        return (
          <Dashboard
            dataverseService={dataverseService}
            context={context}
            activeRole={this.state.activeRole}
            soloTestMode={this.state.soloTestMode}
            onToggleSoloTestMode={() => {
              const nextVal = !this.state.soloTestMode;
              try {
                if (typeof window !== 'undefined' && window.localStorage) {
                  window.localStorage.setItem('sct_solo_test_mode', String(nextVal));
                }
              } catch { /* ignore */ }
              this.setState({ soloTestMode: nextVal });
            }}
            deepLinkRequestId={this.state.deepLinkRequestId}
          />
        );
      case 'newRequest':
        return (
          <NewRequestForm
            dataverseService={dataverseService}
            context={context}
            activeRole={this.state.activeRole}
            initialRecord={this.state.editingRequest}
            onDirtyChange={(isDirty) => this.setState({ isFormDirty: isDirty })}
            onFinished={() => this.setState({ activeTab: 'myRequests', editingRequest: null, isFormDirty: false })}
          />
        );
      case 'myRequests':
        return (
          <MyRequests
            dataverseService={dataverseService}
            context={context}
            activeRole={this.state.activeRole}
            onEditRequest={(rec) => this.setState({ activeTab: 'newRequest', editingRequest: rec })}
          />
        );
      case 'reviewQueue':
        return <ReviewQueue dataverseService={dataverseService} context={context} activeRole={this.state.activeRole} soloTestMode={this.state.soloTestMode} />;
      default:
        return (
          <Dashboard
            dataverseService={dataverseService}
            context={context}
            activeRole={this.state.activeRole}
            soloTestMode={this.state.soloTestMode}
            onToggleSoloTestMode={() => {
              const nextVal = !this.state.soloTestMode;
              try {
                if (typeof window !== 'undefined' && window.localStorage) {
                  window.localStorage.setItem('sct_solo_test_mode', String(nextVal));
                }
              } catch { /* ignore */ }
              this.setState({ soloTestMode: nextVal });
            }}
          />
        );
    }
  }

  public render(): React.ReactElement {
    const { userDisplayName, serviceError } = this.props;
    const { activeTab } = this.state;
    const initials = this._getInitials(userDisplayName || 'User');
    const isDisabled = !!serviceError;

    // 1. Use the deduplicated dynamic role options from state (derived from Azure AD & Dataverse)
    const dynamicRoleOptions = (this.state.roleOptions && this.state.roleOptions.length > 0)
      ? this.state.roleOptions
      : [{ key: 'requester', text: 'Requester' }];

    return (
      <div className={styles.root} role="main">
        {/* ── Service-unavailable banner ───────────────── */}
        {isDisabled && (
          <div style={{
            display: 'flex', alignItems: 'flex-start', gap: 12,
            background: '#FFF4CE', borderBottom: '2px solid #F7C948',
            padding: '12px 20px', flexShrink: 0,
          }}>
            <Icon iconName="Warning" style={{ fontSize: 18, color: '#6B3A00', marginTop: 2, flexShrink: 0 }} />
            <div>
              <div style={{ fontWeight: 700, color: '#6B3A00', fontSize: 13, marginBottom: 2 }}>
                Dataverse tables not available
              </div>
              <div style={{ color: '#6B3A00', fontSize: 12, lineHeight: 1.5 }}>{serviceError}</div>
            </div>
          </div>
        )}

        {/* ── Header ──────────────────────────────────── */}
        <header className={styles.header}>
          <div className={styles.headerTop}>
            {/* Brand */}
            <div className={styles.headerBrand}>
              <div className={styles.headerIcon} aria-hidden="true">
                <Icon iconName="PaymentCard" />
              </div>
              <div>
                <div className={styles.headerTitle}>Fortrex Spend Control</div>
                <div style={{ fontSize: 11, color: '#94A3B8', fontWeight: 500, marginTop: 1, letterSpacing: '0.2px' }}>
                  Procurement & Budget Governance Platform
                </div>
              </div>
            </div>

            {/* Role display — text when only Requester, Dropdown when multiple roles */}
            {this.state.rolesLoading ? (
              <span style={{ color: 'rgba(255,255,255,0.6)', fontSize: 12, fontStyle: 'italic' }}>
                Loading roles…
              </span>
            ) : dynamicRoleOptions.length <= 1 ? (
              <span style={{
                color: '#fff', fontSize: 13, padding: '5px 12px',
                background: 'rgba(255,255,255,0.12)',
                border: '1px solid rgba(255,255,255,0.3)',
                borderRadius: 6,
              }}>
                Role: Requester
              </span>
            ) : (
              <Dropdown
                options={dynamicRoleOptions}
                selectedKey={this.state.activeRole}
                disabled={isDisabled}
                onRenderOption={this._onRenderRoleOption}
                onRenderTitle={this._onRenderRoleTitle}
                onChange={(_, option) => this._handleRoleChange(option)}
                ariaLabel="Switch active role"
                styles={{
                  root:   { minWidth: 180 },
                  title:  { borderRadius: 6, fontSize: 13, background: 'rgba(255,255,255,0.12)', color: '#fff', border: '1px solid rgba(255,255,255,0.3)', height: 32, lineHeight: '30px' },
                  caretDownWrapper: { lineHeight: '32px', color: '#fff' },
                  dropdownItem: { fontSize: 13 },
                  dropdownItemSelected: { fontSize: 13 },
                }}
              />
            )}

            {/* Solo Test Mode Toggle (Enables self-approval across tiers for solo testing) */}
            <button
              type="button"
              onClick={() => {
                const nextVal = !this.state.soloTestMode;
                try {
                  if (typeof window !== 'undefined' && window.localStorage) {
                    window.localStorage.setItem('sct_solo_test_mode', String(nextVal));
                  }
                } catch { /* ignore */ }
                this.setState({ soloTestMode: nextVal });
              }}
              title={this.state.soloTestMode
                ? 'Solo Test Mode is ON: Self-approval is enabled so you can test all stages alone. Click to enforce strict SoD.'
                : 'Strict Segregation of Duties: Self-approval is blocked. Click to enable Solo Test Mode.'}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: 6,
                padding: '0 12px',
                height: 32,
                borderRadius: 16,
                fontSize: 12,
                fontWeight: 600,
                cursor: 'pointer',
                border: this.state.soloTestMode ? '1px solid #107C10' : '1px solid rgba(255,255,255,0.4)',
                background: this.state.soloTestMode ? '#DFF6DD' : 'rgba(255,255,255,0.12)',
                color: this.state.soloTestMode ? '#107C10' : '#ffffff',
                transition: 'all 0.2s ease',
              }}
            >
              <Icon iconName={this.state.soloTestMode ? 'Unlock' : 'Lock'} style={{ fontSize: 13 }} />
              <span>{this.state.soloTestMode ? 'Solo Test: ON' : 'SoD: Strict'}</span>
            </button>

            {/* User chip */}
            <div className={styles.headerUser} aria-label={`Signed in as ${userDisplayName}`}>
              <div className={styles.headerUserAvatar} aria-hidden="true">{initials}</div>
              <span>{userDisplayName}</span>
            </div>
          </div>

          {/* ── Tab Bar ─────────────────────────────── */}
          <nav className={styles.tabBar} role="tablist" aria-label="Main navigation">
            {TABS.filter(tab => {
              const role = this.state.activeRole;
              // "Review Queue" not shown for Requester (they have no review responsibilities)
              if (tab.id === 'reviewQueue' && role === 'requester') return false;
              return true;
            }).map(tab => (
              <button
                key={tab.id}
                role="tab"
                aria-selected={!isDisabled && activeTab === tab.id}
                aria-label={tab.ariaLabel}
                aria-controls={`tabpanel-${tab.id}`}
                id={`tab-${tab.id}`}
                disabled={isDisabled}
                className={[
                  styles.tabItem,
                  !isDisabled && activeTab === tab.id ? styles.tabItemActive : '',
                ].join(' ')}
                style={isDisabled ? { opacity: 0.4, cursor: 'not-allowed', pointerEvents: 'none' } : undefined}
                onClick={() => !isDisabled && this._handleTabClick(tab.id)}
              >
                <Icon iconName={tab.iconName} style={{ fontSize: 15 }} aria-hidden="true" />
                {tab.label}
                {tab.badge !== undefined && tab.badge > 0 && (
                  <span className={styles.tabBadge}>{tab.badge}</span>
                )}
              </button>
            ))}
          </nav>
        </header>

        {/* ── Content ─────────────────────────────────── */}
        <main
          className={styles.content}
          id={`tabpanel-${activeTab}`}
          role="tabpanel"
          aria-labelledby={`tab-${activeTab}`}
          aria-hidden={isDisabled}
        >
          {this._renderTabContent()}
        </main>

        {/* ── Unsaved Changes Confirmation Dialog ───────── */}
        <Dialog
          hidden={!this.state.showLeaveConfirmDialog}
          onDismiss={this._cancelLeaveConfirm}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Unsaved Request in Progress',
            subText: 'You are currently filling out a new request. If you leave this tab now, your entered information will be discarded. Are you sure you want to leave?',
          }}
          modalProps={{
            isBlocking: true,
            styles: { main: { maxWidth: 440, borderRadius: 8 } },
          }}
        >
          <DialogFooter>
            <PrimaryButton
              text="Leave & Discard"
              onClick={this._confirmLeaveTab}
              styles={{ root: { background: '#D13438', borderColor: '#D13438', borderRadius: 6 } }}
            />
            <DefaultButton
              text="Stay & Keep Editing"
              onClick={this._cancelLeaveConfirm}
              styles={{ root: { borderRadius: 6 } }}
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
