import { MSGraphClientV3 } from '@microsoft/sp-http';
export interface IManagerInfo {
    displayName: string;
    mail: string;
    id: string;
}
/**
 * Reusable Graph API Email Sender Utility
 * Dispatches an email via Microsoft Graph /v1.0/me/sendMail using the logged-in user's context.
 */
export declare const sendGraphEmail: (graphClient: MSGraphClientV3, toEmails: string[], ccEmails: string[], subject: string, htmlBody: string) => Promise<void>;
/**
 * Dynamically queries Microsoft Graph for all members of an Azure AD Group.
 * Returns clean email addresses (mail || userPrincipalName).
 * Zero hardcoded names or fallback lists.
 */
export declare function getGroupMemberEmails(graphClient: MSGraphClientV3, groupId: string): Promise<string[]>;
/**
 * Dynamic HTML Template Generator with Cumulative Approval History
 * Injects prior approval logs and a deep-link action button into a responsive HTML email body.
 * Strictly maps Dataverse Formatted Values (@OData.Community.Display.V1.FormattedValue) to prevent GUID display.
 */
export declare function generateApprovalEmailHTML(request: Record<string, unknown>, previousLogs: Array<Record<string, unknown>>, siteUrl: string, actionRequired: string): string;
/** Fetches the current user's direct manager from MS Graph. Returns null if not found. */
export declare function getManager(graphClient: MSGraphClientV3): Promise<IManagerInfo | null>;
/** Sends a notification email to a list of recipients (next-tier approvers). */
export declare function sendTierNotificationEmail(graphClient: MSGraphClientV3, recipients: Array<{
    name: string;
    email: string;
}>, tierName: string, requesterName: string, requestTitle: string, estimatedAmount: string, webPartUrl: string, decision: 'approved' | 'rejected' | 'returned'): Promise<void>;
/** Sends an approval-request email to the manager via Graph /me/sendMail. */
export declare function sendManagerApprovalEmail(graphClient: MSGraphClientV3, manager: IManagerInfo, requesterName: string, requestTitle: string, estimatedAmount: string, webPartUrl: string): Promise<void>;
/** Sends decision outcome email (Rejected or Returned for Info) to the requester. */
export declare function sendRequesterDecisionEmail(graphClient: MSGraphClientV3, requesterEmail: string, requesterName: string, requestTitle: string, estimatedAmount: string, webPartUrl: string, decision: 'rejected' | 'returned', approverName: string, tierName: string, comments: string): Promise<void>;
/** Sends final approval (PO Issued) confirmation email to the requester and stakeholders. */
export declare function sendPOIssuedNotificationEmail(graphClient: MSGraphClientV3, requesterEmail: string, requesterName: string, requestTitle: string, estimatedAmount: string, webPartUrl: string, approverName: string, comments?: string): Promise<void>;
//# sourceMappingURL=GraphMailService.d.ts.map