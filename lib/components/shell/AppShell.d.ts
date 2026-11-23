import * as React from 'react';
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IAppShellProps } from './IAppShellProps';
type TabId = 'dashboard' | 'newRequest' | 'myRequests' | 'reviewQueue';
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
    constructor(props: IAppShellProps);
    componentDidMount(): void;
    private _loadUserRoles;
    /** Loads pending counts across all tiers and manager to display notification badges */
    private _loadAllPendingCounts;
    private _onRenderRoleOption;
    private _onRenderRoleTitle;
    private _getInitials;
    private _handleTabClick;
    private _handleRoleChange;
    private _confirmLeaveTab;
    private _cancelLeaveConfirm;
    private _renderTabContent;
    render(): React.ReactElement;
}
export {};
//# sourceMappingURL=AppShell.d.ts.map