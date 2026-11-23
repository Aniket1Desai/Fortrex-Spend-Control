/**
 * LogsContent — accordion Approval + Audit log sections with no Panel wrapper.
 * Used by RequestLogsPanel (wraps in a Panel) and MyRequests detail view (inline).
 */
import * as React from 'react';
import { DataverseService } from '../../services/DataverseService';
export interface ILogsContentProps {
    requestId: string;
    service: DataverseService;
}
declare const LogsContent: React.FC<ILogsContentProps>;
export default LogsContent;
//# sourceMappingURL=LogsContent.d.ts.map