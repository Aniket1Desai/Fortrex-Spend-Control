import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DataverseService, ODataRecord } from '../../services/DataverseService';
export interface ITakeActionModalProps {
    isOpen: boolean;
    record: ODataRecord | null;
    service: DataverseService;
    activeRole: string;
    currentSystemUserId: string | null;
    context: WebPartContext;
    onDismiss: () => void;
    onSuccess: (message: string) => void;
    /** When true, allows self-approval for solo testing */
    soloTestMode?: boolean;
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
    constructor(props: ITakeActionModalProps);
    componentDidMount(): void;
    componentDidUpdate(prevProps: ITakeActionModalProps): void;
    private _getRecordId;
    private _loadHistory;
    /** Role title for display */
    private _roleDisplay;
    /** Target destination text for the approve button */
    private _approveActionText;
    /** Current pending stage text */
    private _currentStageInfo;
    private _handleDecision;
    render(): React.ReactElement | null;
}
export {};
//# sourceMappingURL=TakeActionModal.d.ts.map