import * as React from 'react';
import { DataverseService, ODataRecord } from '../../services/DataverseService';
interface IMyRequestsProps {
    dataverseService: DataverseService;
    context: import('@microsoft/sp-webpart-base').WebPartContext;
    activeRole?: string;
    onEditRequest?: (record: ODataRecord) => void;
}
interface IMyRequestsState {
    records: ODataRecord[];
    loading: boolean;
    error: string | null;
    systemUserId: string | null;
    activeFilter: string;
    searchText: string;
    viewRecord: ODataRecord | null;
    resubmitRecord: ODataRecord | null;
    resubmitDescription: string;
    resubmitSaving: boolean;
    resubmitSuccess: string | null;
    resubmitError: string | null;
}
export default class MyRequests extends React.Component<IMyRequestsProps, IMyRequestsState> {
    constructor(props: IMyRequestsProps);
    componentDidMount(): Promise<void>;
    componentDidUpdate(prevProps: IMyRequestsProps): void;
    private _loadData;
    private _getStatusLabel;
    private _getSubmittedDate;
    private _getFiltered;
    private _countFor;
    private _handleResubmit;
    render(): React.ReactElement;
    private _renderCard;
    private _cancelRequest;
    private _renderDetailPanel;
}
export {};
//# sourceMappingURL=MyRequests.d.ts.map