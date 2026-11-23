import * as React from 'react';
import { DataverseService, ODataRecord } from '../../services/DataverseService';
import { IResolvedApprovers, IPreApprovalRequirements, IPostSourcingRequirements } from '../../services/ApprovalRoutingService';
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
    actionRecord: ODataRecord | null;
    actionValues: Record<string, string>;
    poDate: Date | undefined;
    saving: boolean;
    saveError: string | null;
    saveSuccess: string | null;
    preApproval: IPreApprovalRequirements | null;
    postSourcing: IPostSourcingRequirements | null;
    resolvedApprovers: IResolvedApprovers | null;
    loadingApprovers: boolean;
    pendingStageRows: IPendingStageRow[];
    loadingPending: boolean;
    systemUserId: string | null;
    managerApprovedIds: Set<string>;
    directReportIds: Set<string>;
    managerActionRecord: ODataRecord | null;
    managerComments: string;
    managerSaving: boolean;
    managerError: string | null;
}
export default class ReviewQueue extends React.Component<IReviewQueueProps, IReviewQueueState> {
    private _service;
    private _router;
    private _primaryKey;
    constructor(props: IReviewQueueProps);
    componentDidMount(): void;
    componentDidUpdate(prevProps: IReviewQueueProps): void;
    private _resolveUser;
    private _loadManagerApprovedIds;
    private _load;
    private _getId;
    private _queueFor;
    private _calculatePendingStage;
    private _openAction;
    private _loadPendingStage;
    private _loadRoutingData;
    private _takeAction;
    private _estimatedValue;
    private _isNewSupplier;
    /** Records waiting for approval by the active approver role */
    private _approverQueue;
    /** Take action (approve / reject / return) on a record from the approver queue */
    private _handleManagerAction;
    private _renderManagerQueue;
    private _renderStageTabs;
    private _renderWorkbench;
    private _renderPendingTracker;
    private _renderActionPanel;
    render(): React.ReactElement;
}
export {};
//# sourceMappingURL=ReviewQueue.d.ts.map