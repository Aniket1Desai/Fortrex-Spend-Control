import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
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

// ── Verified fallback user → roles map (AAD Object IDs from Entra admin center) ───────
// Enforces strict authorized roles for known users:
// Aniket Desai: Manager, Approver Tier 2, Tier 3 (NOT Tier 1)
// Bhakti Gandhi: Approver Tier 1, Tier 3 (NOT Tier 2)
// Manav Barad: Manager, Approver Tier 1, Tier 2 (NOT Tier 3)
const VERIFIED_AD_ROLES: Record<string, string[]> = {
  '16f8a13b-0a91-4150-bf9a-255f06146a4f': ['manager', 'approver_tier2', 'approver_tier3'], // Aniket Desai
  'e83c5b84-1eba-4c01-ba17-39149888cb9e': ['approver_tier1', 'approver_tier3'],            // Bhakti Gandhi
  '1849ec99-9bb8-4478-bb11-e7f3a27ed7d4': ['manager', 'approver_tier1', 'approver_tier2'], // Manav Barad
};

// Users who are managers (have direct reports in Dataverse systemusers table)
const VERIFIED_MANAGERS = new Set([
  '16f8a13b-0a91-4150-bf9a-255f06146a4f', // Aniket Desai
  '1849ec99-9bb8-4478-bb11-e7f3a27ed7d4', // Manav Barad
]);

// ── Per-tier email notification lists (from Entra group member screenshots) ───
export const TIER_NOTIFICATION_EMAILS: Record<string, Array<{name: string; email: string}>> = {
  approver_tier1: [
    { name: 'Bhakti Gandhi', email: 'Bhakti.Gandhi@devrishabhsoft.onmicrosoft.com' },
    { name: 'Manav Barad',   email: 'Manav.Barad@devrishabhsoft.onmicrosoft.com'  },
  ],
  approver_tier2: [
    { name: 'Aniket Desai',  email: 'Aniket.Desai@devrishabhsoft.onmicrosoft.com' },
    { name: 'Manav Barad',   email: 'Manav.Barad@devrishabhsoft.onmicrosoft.com'  },
  ],
  approver_tier3: [
    { name: 'Aniket Desai',  email: 'Aniket.Desai@devrishabhsoft.onmicrosoft.com' },
    { name: 'Bhakti Gandhi', email: 'Bhakti.Gandhi@devrishabhsoft.onmicrosoft.com' },
  ],
};

interface IAppShellState {
  activeTab: TabId;
  activeRole: string;
  tier1PendingCount: number;
  rolePendingCounts: Record<string, number>;
  roleOptions: IDropdownOption[];
  rolesLoading: boolean;
  userRoles?: string[];
  soloTestMode: boolean;
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
    };
  }

  public componentDidMount(): void {
    void this._loadUserRoles();
    if (!this.props.serviceError) {
      void this._loadAllPendingCounts();
    }
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

      // ── Step 2: Query Dataverse systemusers table for Manager status ─────────
      // A user is a Manager if:
      // (a) Other users list this user as their manager in Dataverse (_parentsystemuserid_value), or
      // (b) User has requests pending manager approval in Dataverse, or
      // (c) User is in VERIFIED_MANAGERS
      let dataverseIsManager = false;
      if (systemUserId) {
        dataverseIsManager = await this.props.dataverseService.hasDirectReportsInDataverse(systemUserId);
        if (!dataverseIsManager) {
          const mgrPending = await this.props.dataverseService.getManagerPendingCount(systemUserId);
          if (mgrPending > 0) dataverseIsManager = true;
        }
      }

      // ── Step 3: Query Dataverse sct_roles table for Tier assignments ─────────
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

      // ── Step 4: Query Azure AD via Graph API for Security Groups & Manager ───
      let graphClient: MSGraphClientV3 | undefined;
      try {
        graphClient = await ctx.msGraphClientFactory.getClient('3');
      } catch { /* optional Graph client */ }

      let adTier1 = false;
      let adTier2 = false;
      let adTier3 = false;
      let adIsManager = false;

      if (graphClient) {
        // Step 4a: Check AD Security Groups via checkMemberGroups
        try {
          const groupResponse = await graphClient.api('/me/checkMemberGroups').post({
            groupIds: [TIER1_GROUP_ID, TIER2_GROUP_ID, TIER3_GROUP_ID],
          }) as { value: string[] };

          if (Array.isArray(groupResponse?.value)) {
            const memberOf = new Set<string>((groupResponse.value || []).map(g => g.toLowerCase()));
            adTier1 = memberOf.has(TIER1_GROUP_ID.toLowerCase());
            adTier2 = memberOf.has(TIER2_GROUP_ID.toLowerCase());
            adTier3 = memberOf.has(TIER3_GROUP_ID.toLowerCase());
          }
        } catch {
          // Fallback Graph check via memberOf
          try {
            const memberResponse = await graphClient.api('/me/memberOf').select('id').get() as { value: Array<{ id: string }> };
            if (Array.isArray(memberResponse?.value)) {
              const memberOf = new Set<string>(memberResponse.value.map(g => (g.id || '').toLowerCase()));
              adTier1 = memberOf.has(TIER1_GROUP_ID.toLowerCase());
              adTier2 = memberOf.has(TIER2_GROUP_ID.toLowerCase());
              adTier3 = memberOf.has(TIER3_GROUP_ID.toLowerCase());
            }
          } catch { /* graph memberOf failed */ }
        }

        // Step 4b: Graph /me/directReports check
        try {
          const reportsResponse = await graphClient
            .api('/me/directReports')
            .select('id')
            .top(1)
            .get() as { value: unknown[] };
          if ((reportsResponse?.value || []).length > 0) {
            adIsManager = true;
          }
        } catch { /* 403 / ignore */ }
      }

      // ── Step 5: Strict Verified Authority Mapping ───────────────────────────
      // Enforces strict role isolation for known users regardless of AD group overlap.
      // E.g., Aniket Desai is Tier 2 & Tier 3 ONLY; Bhakti Gandhi is Tier 1 & Tier 3 ONLY.
      let isTier1 = dvTier1 || adTier1;
      let isTier2 = dvTier2 || adTier2;
      let isTier3 = dvTier3 || adTier3;

      const knownRoles = VERIFIED_AD_ROLES[aadOid] ?? VERIFIED_AD_ROLES[normAadOid] ??
        (userEmail.indexOf('aniket.desai') !== -1 || userDisplayName.indexOf('aniket desai') !== -1 ? ['manager', 'approver_tier2', 'approver_tier3'] :
         userEmail.indexOf('bhakti.gandhi') !== -1 || userDisplayName.indexOf('bhakti gandhi') !== -1 ? ['approver_tier1', 'approver_tier3'] :
         userEmail.indexOf('manav.barad') !== -1 || userDisplayName.indexOf('manav barad') !== -1 ? ['manager', 'approver_tier1', 'approver_tier2'] : null);

      if (knownRoles) {
        isTier1 = knownRoles.indexOf('approver_tier1') !== -1;
        isTier2 = knownRoles.indexOf('approver_tier2') !== -1;
        isTier3 = knownRoles.indexOf('approver_tier3') !== -1;
      }

      // Determine final Manager status
      const isKnownManager = VERIFIED_MANAGERS.has(aadOid) || VERIFIED_MANAGERS.has(normAadOid) ||
        userEmail.indexOf('aniket.desai') !== -1 || userEmail.indexOf('manav.barad') !== -1;
      const isManager = dataverseIsManager || adIsManager || isKnownManager || (knownRoles && knownRoles.indexOf('manager') !== -1);

      // ── Step 6: Construct Strict Role Options (ONLY user's actual roles) ──────
      const opts: IDropdownOption[] = [{ key: 'requester', text: 'Requester' }];
      if (isManager) opts.push({ key: 'manager', text: 'Manager' });
      if (isTier1)   opts.push({ key: TIER1_KEY, text: 'Approver — Tier 1' });
      if (isTier2)   opts.push({ key: 'approver_tier2', text: 'Approver — Tier 2' });
      if (isTier3)   opts.push({ key: 'approver_tier3', text: 'Approver — Tier 3' });

      // If current activeRole is not in the allowed roles, reset to Requester
      const validKeys = new Set(opts.map(o => o.key as string));
      const newActiveRole = validKeys.has(this.state.activeRole) ? this.state.activeRole : 'requester';

      this.setState({
        roleOptions: opts,
        userRoles: opts.map(o => o.key as string),
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
        return <Dashboard dataverseService={dataverseService} context={context} activeRole={this.state.activeRole} soloTestMode={this.state.soloTestMode} />;
      case 'newRequest':
        return <NewRequestForm dataverseService={dataverseService} context={context} />;
      case 'myRequests':
        return <MyRequests dataverseService={dataverseService} context={context} activeRole={this.state.activeRole} />;
      case 'reviewQueue':
        return <ReviewQueue dataverseService={dataverseService} context={context} activeRole={this.state.activeRole} soloTestMode={this.state.soloTestMode} />;
      default:
        return <Dashboard dataverseService={dataverseService} context={context} soloTestMode={this.state.soloTestMode} />;
    }
  }

  public render(): React.ReactElement {
    const { userDisplayName, serviceError } = this.props;
    const { activeTab } = this.state;
    const initials = this._getInitials(userDisplayName || 'User');
    const isDisabled = !!serviceError;

    // 1. Map the authorized roles array (e.g., this.state.userRoles or props.userRoles)
    const dynamicRoleOptions = (this.state.userRoles || ['requester']).map((roleKey: string) => {
      let displayName = 'Requester';
      if (roleKey === 'manager') displayName = 'Manager';
      else if (roleKey === 'approver_tier1') displayName = 'Approver — Tier 1';
      else if (roleKey === 'approver_tier2') displayName = 'Approver — Tier 2';
      else if (roleKey === 'approver_tier3') displayName = 'Approver — Tier 3';

      return { key: roleKey, text: displayName };
    });

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
                <Icon iconName="Money" />
              </div>
              <div>
                <div className={styles.headerTitle}>Fortrex Spend Control</div>
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
                onChange={(_, option) => {
                  if (!option) return;
                  const newRole = option.key as string;
                  // If switching to Requester while on Review Queue, redirect to dashboard.
                  // If switching away from Requester while on New Request tab, redirect to dashboard.
                  const currentTab = this.state.activeTab;
                  let newTab: TabId = currentTab;
                  if (newRole !== 'requester' && currentTab === 'newRequest') newTab = 'dashboard';
                  if (newRole === 'requester' && currentTab === 'reviewQueue') newTab = 'dashboard';
                  this.setState({ activeRole: newRole, activeTab: newTab });
                }}
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
              // "New Request" only for Requester
              if (tab.id === 'newRequest' && role !== 'requester') return false;
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
                onClick={() => !isDisabled && this.setState({ activeTab: tab.id })}
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
      </div>
    );
  }
}
