import * as React from 'react';
import { DataverseService } from '../../services/DataverseService';
interface IRequestLogsPanelProps {
    requestId: string;
    reference: string;
    service: DataverseService;
    onDismiss: () => void;
}
declare const RequestLogsPanel: React.FC<IRequestLogsPanelProps>;
export default RequestLogsPanel;
//# sourceMappingURL=RequestLogsPanel.d.ts.map