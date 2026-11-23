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
export const sendGraphEmail = async (
  graphClient: MSGraphClientV3,
  toEmails: string[],
  ccEmails: string[],
  subject: string,
  htmlBody: string
): Promise<void> => {
  try {
    const validTo = (toEmails || []).filter(e => Boolean(e && typeof e === 'string' && e.trim().length > 0));
    const validCc = (ccEmails || []).filter(e => Boolean(e && typeof e === 'string' && e.trim().length > 0));

    if (validTo.length === 0) {
      console.warn('[GraphMailService] sendGraphEmail aborted: toRecipients list is empty.');
      return;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const message: any = {
      subject,
      body: { contentType: 'HTML', content: htmlBody },
      toRecipients: validTo.map(email => ({ emailAddress: { address: email.trim() } })),
    };
    if (validCc.length > 0) {
      message.ccRecipients = validCc.map(email => ({ emailAddress: { address: email.trim() } }));
    }

    await graphClient.api('/me/sendMail').version('v1.0').post({ message, saveToSentItems: true });
  } catch (error) {
    console.error('Non-fatal error: Failed to send Graph email:', error);
  }
};

/**
 * Dynamically queries Microsoft Graph for all members of an Azure AD Group.
 * Returns clean email addresses (mail || userPrincipalName).
 * Zero hardcoded names or fallback lists.
 */
export async function getGroupMemberEmails(
  graphClient: MSGraphClientV3,
  groupId: string
): Promise<string[]> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const response: any = await graphClient
      .api(`/groups/${groupId}/members`)
      .select('id,displayName,mail,userPrincipalName')
      .top(999)
      .get();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const members: any[] = response?.value || [];
    const emails: string[] = [];
    const seen = new Set<string>();

    for (const m of members) {
      const email = String(m.mail || m.userPrincipalName || '').trim();
      const lower = email.toLowerCase();
      if (email.length > 0 && !seen.has(lower) && email.indexOf('@') !== -1) {
        seen.add(lower);
        emails.push(email);
      }
    }
    return emails;
  } catch (err) {
    console.warn(`[GraphMailService] Failed to dynamically fetch group members for ${groupId}:`, err);
    return [];
  }
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function isGuid(val: unknown): boolean {
  return typeof val === 'string' && GUID_RE.test(val.trim());
}

function cleanDisplayVal(val: unknown, fallback: string): string {
  if (val === undefined || val === null) return fallback;
  const str = String(val).trim();
  if (!str || isGuid(str)) return fallback;
  return str;
}

/**
 * Formats a date/timestamp to strict 'dd/MM/yyyy HH:mm tt' format in user's local timezone.
 * Example: 18/11/2026 09:31 AM
 */
function formatDateTime(val: string | Date | undefined | null): string {
  if (!val) return '—';
  try {
    let iso = val instanceof Date ? val.toISOString() : String(val).trim();
    if (!iso) return '—';
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?$/.test(iso)) {
      iso += 'Z';
    }
    const d = new Date(iso);
    if (isNaN(d.getTime())) return typeof val === 'string' ? val : '—';

    const pad2 = (n: number): string => (n < 10 ? '0' + n : String(n));
    const day = pad2(d.getDate());
    const month = pad2(d.getMonth() + 1);
    const year = d.getFullYear();
    const h = d.getHours();
    const tt = h >= 12 ? 'PM' : 'AM';
    const h12 = h % 12 || 12;
    const hh = pad2(h12);
    const mm = pad2(d.getMinutes());

    return `${day}/${month}/${year} ${hh}:${mm} ${tt}`;
  } catch {
    return typeof val === 'string' ? val : '—';
  }
}

/**
 * Dynamic HTML Template Generator with Cumulative Approval History
 * Injects prior approval logs and a deep-link action button into a responsive HTML email body.
 * Strictly maps Dataverse Formatted Values (@OData.Community.Display.V1.FormattedValue) to prevent GUID display.
 */
export function generateApprovalEmailHTML(
  request: Record<string, unknown>,
  previousLogs: Array<Record<string, unknown>>,
  siteUrl: string,
  actionRequired: string
): string {
  const ref = String(
    request['sct_sct_reference@OData.Community.Display.V1.FormattedValue'] ||
    request['sct_sct_reference'] ||
    request['sct_referencenumber'] ||
    request['reference'] ||
    'Spend Request'
  );
  const title = String(
    request['sct_request_title@OData.Community.Display.V1.FormattedValue'] ||
    request['sct_request_title'] ||
    request['title'] ||
    ref
  );

  // 1. Requester Name: Prioritize FormattedValue
  const requester = String(
    request['_sct_requester_value@OData.Community.Display.V1.FormattedValue'] ||
    request['sct_requester@OData.Community.Display.V1.FormattedValue'] ||
    request['_createdby_value@OData.Community.Display.V1.FormattedValue'] ||
    request['_sct_requestorid_value@OData.Community.Display.V1.FormattedValue'] ||
    cleanDisplayVal(request['requesterName'], '') ||
    cleanDisplayVal(request['_sct_requester_value'], '') ||
    'Unknown Requester'
  );

  // 2. Business Unit: Prioritize FormattedValue
  const bu = String(
    request['sct_businessunit@OData.Community.Display.V1.FormattedValue'] ||
    request['sct_business_unit@OData.Community.Display.V1.FormattedValue'] ||
    cleanDisplayVal(request['sct_businessunit'], '') ||
    cleanDisplayVal(request['sct_business_unit'], '') ||
    cleanDisplayVal(request['businessUnit'], '') ||
    '—'
  );

  // 3. Site / Location: Prioritize FormattedValue
  const site = String(
    request['_sct_site_value@OData.Community.Display.V1.FormattedValue'] ||
    request['sct_site@OData.Community.Display.V1.FormattedValue'] ||
    cleanDisplayVal(request['selectedSiteName'], '') ||
    cleanDisplayVal(request['_sct_site_value'], '') ||
    cleanDisplayVal(request['site'], '') ||
    '—'
  );

  const amtRaw = parseFloat(
    String(request['sct_estimated_total_value'] ?? request['estimatedAmount'] ?? '0').replace(/[^0-9.-]/g, '')
  );
  const amtFormatted = String(
    request['sct_estimated_total_value@OData.Community.Display.V1.FormattedValue'] ||
    (!isNaN(amtRaw) && amtRaw > 0
      ? `$${amtRaw.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
      : '(not specified)')
  );

  const desc = String(
    request['sct_description@OData.Community.Display.V1.FormattedValue'] ||
    request['sct_description'] ||
    request['description'] ||
    ''
  );

  // Deep link targeting the request reference
  const queryParam = ref !== 'Spend Request' && ref !== '(New Submission)' ? ref : String(request['id'] || request['sct_requestid'] || '');
  const deepLink = siteUrl.indexOf('?') !== -1
    ? `${siteUrl}&requestId=${encodeURIComponent(queryParam)}`
    : `${siteUrl}?requestId=${encodeURIComponent(queryParam)}`;

  // Cumulative Prior Approval History HTML mapping
  let historyRowsHtml = '';
  if (Array.isArray(previousLogs) && previousLogs.length > 0) {
    historyRowsHtml = previousLogs.map(log => {
      const tierFormatted = log['sct_tiername@OData.Community.Display.V1.FormattedValue'];
      const tierInt = log['sct_tiername'];
      const tier = tierFormatted ? String(tierFormatted)
        : tierInt === 4 ? 'Line Manager'
          : tierInt === 1 ? 'Tier 1 Approver'
            : tierInt === 2 ? 'Tier 2 Approver'
              : tierInt === 3 ? 'Tier 3 Approver'
                : String(log['sct_tiername'] ?? 'Approver');

      const statusFormatted = log['sct_approvalstatus@OData.Community.Display.V1.FormattedValue'];
      const statusInt = log['sct_approvalstatus'];
      const isApprove = statusFormatted
        ? String(statusFormatted).toLowerCase().indexOf('approved') !== -1
        : statusInt === 1;
      const isReject = statusFormatted
        ? String(statusFormatted).toLowerCase().indexOf('rejected') !== -1
        : statusInt === 2;
      const statusLabel = statusFormatted ? String(statusFormatted) : (isApprove ? 'Approved' : isReject ? 'Rejected' : 'Returned');
      const statusBg = isApprove ? '#DFF6DD' : isReject ? '#FDE7E9' : '#FFF4CE';
      const statusColor = isApprove ? '#107C10' : isReject ? '#A4262C' : '#7A4F01';

      // Approver Name: Prioritize FormattedValue
      const approverName = String(
        log['_sct_approver_value@OData.Community.Display.V1.FormattedValue'] ||
        log['_sct_approverid_value@OData.Community.Display.V1.FormattedValue'] ||
        log['sct_approverid@OData.Community.Display.V1.FormattedValue'] ||
        log['_createdby_value@OData.Community.Display.V1.FormattedValue'] ||
        cleanDisplayVal(log['sct_approvername'], '') ||
        cleanDisplayVal(log['sct_name'], '') ||
        'Unknown Approver'
      );

      const comments = String(
        log['sct_comments@OData.Community.Display.V1.FormattedValue'] ||
        log['sct_comments'] ||
        ''
      ).trim();

      const rawIso = String(
        log['createdon'] ||
        log['sct_decisionon'] ||
        log['createdon@OData.Community.Display.V1.FormattedValue'] ||
        log['sct_decisionon@OData.Community.Display.V1.FormattedValue'] ||
        ''
      );
      const dateStr = formatDateTime(rawIso);

      return `
        <tr style="border-bottom: 1px solid #EDEBE9;">
          <td style="padding: 10px 12px; font-weight: 600; color: #323130; vertical-align: top; white-space: nowrap;">
            ${tier}
          </td>
          <td style="padding: 10px 12px; vertical-align: top; white-space: nowrap;">
            <span style="background: ${statusBg}; color: ${statusColor}; font-size: 11px; font-weight: 700; padding: 3px 8px; border-radius: 10px; display: inline-block;">
              ${statusLabel}
            </span>
          </td>
          <td style="padding: 10px 12px; color: #605E5C; vertical-align: top;">
            <div style="font-weight: 600; color: #323130;">${approverName}</div>
            <div style="font-size: 11px; color: #8A8886; margin-top: 2px;">${dateStr}</div>
          </td>
          <td style="padding: 10px 12px; color: #323130; vertical-align: top;">
            ${comments ? `<div style="background: #F8F8F8; border-left: 3px solid #0078D4; padding: 6px 10px; border-radius: 2px; font-size: 12px; font-style: italic;">"${comments}"</div>` : '<span style="color: #A19F9D; font-size: 12px;">(No comments provided)</span>'}
          </td>
        </tr>
      `;
    }).join('');
  } else {
    historyRowsHtml = `
      <tr>
        <td colspan="4" style="padding: 14px 12px; color: #8A8886; font-style: italic; text-align: center;">
          Initial submission — no prior approval history recorded.
        </td>
      </tr>
    `;
  }

  return `
<div style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; max-width: 650px; margin: 0 auto; color: #201F1E; background-color: #FAFAFA; padding: 20px;">
  <!-- Header Card -->
  <div style="background: linear-gradient(135deg, #0078D4 0%, #004578 100%); padding: 24px 32px; border-radius: 8px 8px 0 0; color: #FFFFFF;">
    <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #B4D6FA; margin-bottom: 4px;">
      Fortrex Spend Control Tracker
    </div>
    <h1 style="margin: 0; font-size: 20px; font-weight: 700; color: #FFFFFF;">
      ${actionRequired}
    </h1>
    <p style="margin: 6px 0 0; font-size: 13px; color: #E1DFDD;">
      Reference: <strong style="color: #FFFFFF;">${ref}</strong>
    </p>
  </div>

  <!-- Body Container -->
  <div style="background: #FFFFFF; padding: 28px 32px; border: 1px solid #EDEBE9; border-top: none; border-radius: 0 0 8px 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.04);">
    <p style="margin: 0 0 18px; font-size: 14px; line-height: 1.5; color: #323130;">
      A spend control request requires your attention. Below are the request details and complete cumulative approval trail:
    </p>

    <!-- Request Details Summary Card -->
    <div style="background: #F8F9FA; border: 1px solid #E1DFDD; border-radius: 6px; padding: 18px 20px; margin-bottom: 24px;">
      <div style="font-size: 15px; font-weight: 700; color: #0078D4; margin-bottom: 12px; border-bottom: 1px solid #EDEBE9; padding-bottom: 8px;">
        ${title}
      </div>
      <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
        <tr>
          <td style="padding: 5px 0; color: #605E5C; width: 140px; font-weight: 600;">Requested By:</td>
          <td style="padding: 5px 0; color: #201F1E; font-weight: 600;">${requester}</td>
        </tr>
        <tr>
          <td style="padding: 5px 0; color: #605E5C; font-weight: 600;">Estimated Value:</td>
          <td style="padding: 5px 0; color: #107C10; font-weight: 700; font-size: 14px;">${amtFormatted}</td>
        </tr>
        <tr>
          <td style="padding: 5px 0; color: #605E5C; font-weight: 600;">Business Unit:</td>
          <td style="padding: 5px 0; color: #201F1E;">${bu}</td>
        </tr>
        <tr>
          <td style="padding: 5px 0; color: #605E5C; font-weight: 600;">Site / Location:</td>
          <td style="padding: 5px 0; color: #201F1E;">${site}</td>
        </tr>
        ${desc ? `
        <tr>
          <td style="padding: 8px 0 0; color: #605E5C; font-weight: 600; vertical-align: top;">Description:</td>
          <td style="padding: 8px 0 0; color: #323130; line-height: 1.4;">${desc}</td>
        </tr>` : ''}
      </table>
    </div>

    <!-- Prior Approval History Section -->
    <div style="margin-bottom: 26px;">
      <div style="font-size: 14px; font-weight: 700; color: #201F1E; margin-bottom: 10px;">
        Prior Approval History
      </div>
      <table style="width: 100%; border-collapse: collapse; font-size: 12px; border: 1px solid #EDEBE9; border-radius: 4px; overflow: hidden;">
        <thead>
          <tr style="background: #F3F2F1; text-align: left; border-bottom: 1px solid #EDEBE9;">
            <th style="padding: 8px 12px; color: #605E5C; font-weight: 700; width: 110px;">Stage</th>
            <th style="padding: 8px 12px; color: #605E5C; font-weight: 700; width: 85px;">Decision</th>
            <th style="padding: 8px 12px; color: #605E5C; font-weight: 700; width: 140px;">Approver &amp; Date</th>
            <th style="padding: 8px 12px; color: #605E5C; font-weight: 700;">Comments</th>
          </tr>
        </thead>
        <tbody>
          ${historyRowsHtml}
        </tbody>
      </table>
    </div>

    <!-- Action Deep Link CTA Button -->
    <div style="text-align: center; margin: 30px 0 15px;">
      <a href="${deepLink}" target="_blank" style="display: inline-block; background-color: #0078D4; color: #FFFFFF; font-size: 14px; font-weight: 600; text-decoration: none; padding: 13px 28px; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,120,212,0.3);">
        View &amp; Act on Request
      </a>
    </div>

    <!-- Footer Disclaimer -->
    <div style="margin-top: 28px; padding-top: 14px; border-top: 1px solid #EDEBE9; font-size: 11px; color: #8A8886; text-align: center; line-height: 1.4;">
      This email was automatically generated and dispatched by the Fortrex Spend Control Tracker using Microsoft Graph API. Please do not reply directly to this email.
    </div>
  </div>
</div>
  `;
}

/** Fetches the current user's direct manager from MS Graph. Returns null if not found. */
export async function getManager(graphClient: MSGraphClientV3): Promise<IManagerInfo | null> {
  try {
    const result = await graphClient
      .api('/me/manager')
      .select('id,displayName,mail,userPrincipalName')
      .get();
    if (!result) return null;
    const mail = String(result.mail || result.userPrincipalName || '').trim();
    if (!mail && !result.id) return null;
    return { id: result.id, displayName: result.displayName || '', mail };
  } catch {
    return null;
  }
}

/** Sends a notification email to a list of recipients (next-tier approvers). */
export async function sendTierNotificationEmail(
  graphClient: MSGraphClientV3,
  recipients: Array<{name: string; email: string}>,
  tierName: string,
  requesterName: string,
  requestTitle: string,
  estimatedAmount: string,
  webPartUrl: string,
  decision: 'approved' | 'rejected' | 'returned'
): Promise<void> {
  if (!recipients.length) return;

  const decisionLabel = decision === 'approved' ? 'Approved ✓' : decision === 'rejected' ? 'Rejected ✗' : 'Returned for Info ↩';
  const decisionColor = decision === 'approved' ? '#107C10' : decision === 'rejected' ? '#A4262C' : '#8764B8';

  const amtFormatted = estimatedAmount
    ? `$${parseFloat(estimatedAmount).toLocaleString('en-US', { minimumFractionDigits: 2 })}`
    : '(not specified)';

  const html = `
<div style="font-family: Segoe UI, sans-serif; max-width: 600px; margin: 0 auto; color: #201F1E;">
  <div style="background: linear-gradient(135deg, #0078D4, #1BA1E2); padding: 24px 32px; border-radius: 8px 8px 0 0;">
    <h1 style="color: #fff; margin: 0; font-size: 20px; font-weight: 700;">Fortrex Spend Control Tracker</h1>
    <p style="color: rgba(255,255,255,0.85); margin: 6px 0 0; font-size: 14px;">Request Requires Your Attention — ${tierName}</p>
  </div>
  <div style="background: #F3F2F1; padding: 28px 32px; border-radius: 0 0 8px 8px; border: 1px solid #EDEBE9;">
    <p style="margin: 0 0 16px; font-size: 15px;">Hello,</p>
    <p style="margin: 0 0 20px; font-size: 14px; line-height: 1.6; color: #605E5C;">
      A request from <strong>${requesterName}</strong> has been
      <strong style="color: ${decisionColor};">${decisionLabel}</strong> at the previous stage
      and is now awaiting your review as <strong>${tierName}</strong>.
    </p>
    <div style="background: #fff; border-radius: 6px; padding: 18px 20px; margin-bottom: 20px; border: 1px solid #EDEBE9;">
      <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
        <tr><td style="padding: 6px 0; color: #A19F9D; width: 160px;">Request Title</td><td style="font-weight: 600;">${requestTitle}</td></tr>
        <tr><td style="padding: 6px 0; color: #A19F9D;">Requested By</td><td style="font-weight: 600;">${requesterName}</td></tr>
        <tr><td style="padding: 6px 0; color: #A19F9D;">Estimated Value</td><td style="font-weight: 600;">${amtFormatted}</td></tr>
        <tr><td style="padding: 6px 0; color: #A19F9D;">Current Stage</td><td style="font-weight: 600; color: ${decisionColor};">${tierName}</td></tr>
      </table>
    </div>
    <a href="${webPartUrl}" style="display: inline-block; background: #0078D4; color: #fff; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 14px;">
      Open SCT &amp; Take Action →
    </a>
    <p style="margin: 24px 0 0; font-size: 12px; color: #A19F9D;">
      This notification was sent automatically by the Fortrex Spend Control Tracker. Do not reply to this email.
    </p>
  </div>
</div>`;

  await graphClient.api('/me/sendMail').post({
    message: {
      subject: `[SCT — ${tierName} Action Required] ${requestTitle}`,
      body: { contentType: 'HTML', content: html },
      toRecipients: recipients.map(r => ({
        emailAddress: { address: r.email, name: r.name },
      })),
    },
    saveToSentItems: false,
  });
}

/** Sends an approval-request email to the manager via Graph /me/sendMail. */
export async function sendManagerApprovalEmail(
  graphClient: MSGraphClientV3,
  manager: IManagerInfo,
  requesterName: string,
  requestTitle: string,
  estimatedAmount: string,
  webPartUrl: string
): Promise<void> {
  const amtFormatted = estimatedAmount
    ? `$${parseFloat(estimatedAmount).toLocaleString('en-US', { minimumFractionDigits: 2 })}`
    : '(not specified)';

  const html = `
<div style="font-family: Segoe UI, sans-serif; max-width: 600px; margin: 0 auto; color: #201F1E;">
  <div style="background: linear-gradient(135deg, #0078D4, #1BA1E2); padding: 24px 32px; border-radius: 8px 8px 0 0;">
    <h1 style="color: #fff; margin: 0; font-size: 20px; font-weight: 700;">Fortrex Spend Control Tracker</h1>
    <p style="color: rgba(255,255,255,0.85); margin: 6px 0 0; font-size: 14px;">New Request Awaiting Your Approval</p>
  </div>
  <div style="background: #F3F2F1; padding: 28px 32px; border-radius: 0 0 8px 8px; border: 1px solid #EDEBE9;">
    <p style="margin: 0 0 16px; font-size: 15px;">Hello <strong>${manager.displayName}</strong>,</p>
    <p style="margin: 0 0 20px; font-size: 14px; line-height: 1.6; color: #605E5C;">
      A new spend control request has been submitted by <strong>${requesterName}</strong> and requires your approval.
    </p>
    <div style="background: #fff; border-radius: 6px; padding: 18px 20px; margin-bottom: 20px; border: 1px solid #EDEBE9;">
      <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
        <tr><td style="padding: 6px 0; color: #A19F9D; width: 160px;">Request Title</td><td style="font-weight: 600;">${requestTitle}</td></tr>
        <tr><td style="padding: 6px 0; color: #A19F9D;">Requested By</td><td style="font-weight: 600;">${requesterName}</td></tr>
        <tr><td style="padding: 6px 0; color: #A19F9D;">Estimated Value</td><td style="font-weight: 600;">${amtFormatted}</td></tr>
      </table>
    </div>
    <a href="${webPartUrl}" style="display: inline-block; background: #0078D4; color: #fff; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 14px;">
      Review &amp; Approve Request →
    </a>
    <p style="margin: 24px 0 0; font-size: 12px; color: #A19F9D;">
      This notification was sent automatically by the Fortrex Spend Control Tracker. Do not reply to this email.
    </p>
  </div>
</div>`;

  await graphClient.api('/me/sendMail').post({
    message: {
      subject: `[SCT Approval Required] ${requestTitle} — ${requesterName}`,
      body: { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: manager.mail, name: manager.displayName } }],
    },
    saveToSentItems: false,
  });
}

/** Sends decision outcome email (Rejected or Returned for Info) to the requester. */
export async function sendRequesterDecisionEmail(
  graphClient: MSGraphClientV3,
  requesterEmail: string,
  requesterName: string,
  requestTitle: string,
  estimatedAmount: string,
  webPartUrl: string,
  decision: 'rejected' | 'returned',
  approverName: string,
  tierName: string,
  comments: string
): Promise<void> {
  if (!requesterEmail) return;

  const isRejected = decision === 'rejected';
  const decisionLabel = isRejected ? 'Rejected ✗' : 'Returned for Information ↩';
  const decisionColor = isRejected ? '#A4262C' : '#8764B8';
  const subjectPrefix = isRejected ? '[SCT Request Rejected]' : '[SCT Returned for Information]';

  const amtFormatted = estimatedAmount
    ? `$${parseFloat(estimatedAmount).toLocaleString('en-US', { minimumFractionDigits: 2 })}`
    : '(not specified)';

  const html = `
<div style="font-family: Segoe UI, sans-serif; max-width: 600px; margin: 0 auto; color: #201F1E;">
  <div style="background: linear-gradient(135deg, ${decisionColor}, #201F1E); padding: 24px 32px; border-radius: 8px 8px 0 0;">
    <h1 style="color: #fff; margin: 0; font-size: 20px; font-weight: 700;">Fortrex Spend Control Tracker</h1>
    <p style="color: rgba(255,255,255,0.85); margin: 6px 0 0; font-size: 14px;">Request Update: ${decisionLabel}</p>
  </div>
  <div style="background: #F3F2F1; padding: 28px 32px; border-radius: 0 0 8px 8px; border: 1px solid #EDEBE9;">
    <p style="margin: 0 0 16px; font-size: 15px;">Hello <strong>${requesterName}</strong>,</p>
    <p style="margin: 0 0 20px; font-size: 14px; line-height: 1.6; color: #605E5C;">
      Your spend request has been <strong style="color: ${decisionColor};">${decisionLabel}</strong> by <strong>${approverName}</strong> (${tierName}).
    </p>
    <div style="background: #fff; border-radius: 6px; padding: 18px 20px; margin-bottom: 20px; border: 1px solid #EDEBE9;">
      <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
        <tr><td style="padding: 6px 0; color: #A19F9D; width: 140px;">Request Title</td><td style="font-weight: 600;">${requestTitle}</td></tr>
        <tr><td style="padding: 6px 0; color: #A19F9D;">Estimated Value</td><td style="font-weight: 600;">${amtFormatted}</td></tr>
        <tr><td style="padding: 6px 0; color: #A19F9D;">Reviewed By</td><td style="font-weight: 600;">${approverName} (${tierName})</td></tr>
        <tr><td style="padding: 6px 0; color: #A19F9D;">Status</td><td style="font-weight: 600; color: ${decisionColor};">${decisionLabel}</td></tr>
      </table>
      ${comments ? `
      <div style="margin-top: 14px; padding-top: 12px; border-top: 1px solid #EDEBE9;">
        <div style="font-size: 11px; font-weight: 700; color: #A19F9D; text-transform: uppercase; margin-bottom: 4px;">Approver Comments</div>
        <div style="font-size: 13px; color: #201F1E; background: #F8F8F8; padding: 10px; border-radius: 4px; border: 1px solid #E1DFDD;">${comments}</div>
      </div>` : ''}
    </div>
    <a href="${webPartUrl}" style="display: inline-block; background: #0078D4; color: #fff; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 14px;">
      ${isRejected ? 'View Request in SCT →' : 'Edit &amp; Resubmit Request →'}
    </a>
    <p style="margin: 24px 0 0; font-size: 12px; color: #A19F9D;">
      This notification was sent automatically by the Fortrex Spend Control Tracker. Do not reply to this email.
    </p>
  </div>
</div>`;

  await graphClient.api('/me/sendMail').post({
    message: {
      subject: `${subjectPrefix} ${requestTitle}`,
      body: { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: requesterEmail, name: requesterName } }],
    },
    saveToSentItems: false,
  });
}

/** Sends final approval (PO Issued) confirmation email to the requester and stakeholders. */
export async function sendPOIssuedNotificationEmail(
  graphClient: MSGraphClientV3,
  requesterEmail: string,
  requesterName: string,
  requestTitle: string,
  estimatedAmount: string,
  webPartUrl: string,
  approverName: string,
  comments?: string
): Promise<void> {
  if (!requesterEmail) return;

  const amtFormatted = estimatedAmount
    ? `$${parseFloat(estimatedAmount).toLocaleString('en-US', { minimumFractionDigits: 2 })}`
    : '(not specified)';

  const html = `
<div style="font-family: Segoe UI, sans-serif; max-width: 600px; margin: 0 auto; color: #201F1E;">
  <div style="background: linear-gradient(135deg, #107C10, #0B6A0B); padding: 24px 32px; border-radius: 8px 8px 0 0;">
    <h1 style="color: #fff; margin: 0; font-size: 20px; font-weight: 700;">Fortrex Spend Control Tracker</h1>
    <p style="color: rgba(255,255,255,0.85); margin: 6px 0 0; font-size: 14px;">Final Approval Granted — PO Authorized</p>
  </div>
  <div style="background: #F3F2F1; padding: 28px 32px; border-radius: 0 0 8px 8px; border: 1px solid #EDEBE9;">
    <p style="margin: 0 0 16px; font-size: 15px;">Hello <strong>${requesterName}</strong>,</p>
    <p style="margin: 0 0 20px; font-size: 14px; line-height: 1.6; color: #605E5C;">
      Your spend request has successfully passed all approval stages and received final sign-off from <strong>${approverName}</strong> (Tier 3 / Executive).
      Procurement has been authorized to issue a Purchase Order in Dynamics 365.
    </p>
    <div style="background: #fff; border-radius: 6px; padding: 18px 20px; margin-bottom: 20px; border: 1px solid #EDEBE9;">
      <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
        <tr><td style="padding: 6px 0; color: #A19F9D; width: 140px;">Request Title</td><td style="font-weight: 600;">${requestTitle}</td></tr>
        <tr><td style="padding: 6px 0; color: #A19F9D;">Estimated Value</td><td style="font-weight: 600;">${amtFormatted}</td></tr>
        <tr><td style="padding: 6px 0; color: #A19F9D;">Final Status</td><td style="font-weight: 600; color: #107C10;">PO Issued (Approved)</td></tr>
      </table>
      ${comments ? `
      <div style="margin-top: 14px; padding-top: 12px; border-top: 1px solid #EDEBE9;">
        <div style="font-size: 11px; font-weight: 700; color: #A19F9D; text-transform: uppercase; margin-bottom: 4px;">Sign-off Notes</div>
        <div style="font-size: 13px; color: #201F1E; background: #F8F8F8; padding: 10px; border-radius: 4px; border: 1px solid #E1DFDD;">${comments}</div>
      </div>` : ''}
    </div>
    <a href="${webPartUrl}" style="display: inline-block; background: #107C10; color: #fff; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 14px;">
      View Request in SCT →
    </a>
    <p style="margin: 24px 0 0; font-size: 12px; color: #A19F9D;">
      This notification was sent automatically by the Fortrex Spend Control Tracker. Do not reply to this email.
    </p>
  </div>
</div>`;

  await graphClient.api('/me/sendMail').post({
    message: {
      subject: `[SCT Approved — PO Authorized] ${requestTitle}`,
      body: { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: requesterEmail, name: requesterName } }],
    },
    saveToSentItems: false,
  });
}
