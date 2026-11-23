# Fortrex Spend Control Tracker — Master Architectural Compliance Report
**Document ID:** `Working Fortex plan 18-11-2026 10AM`  
**Date & Timestamp:** 18-Nov-2026 10:00 AM IST  
**Role:** Senior SPFx & Microsoft Dataverse Technical Architect  
**Project:** Fortrex Spend Control Tracker (SCT) Web Part  
**SPFx Package Version:** `1.0.0.9` (`sharepoint/solution/ad-dataverse-fortrex.sppkg`)  

---

## Executive Summary
A comprehensive end-to-end architectural code audit was performed across the entire SPFx codebase, cross-referencing all Dataverse OData v4 operations, approval state machines, segregation of duties (SoD) filters, Microsoft Graph v1.0 email delivery pipelines, GUID-sanitization routines, and UI locking states.

**Audit Verdict:** **100% COMPLIANT WITH ZERO DEVIATIONS OR ARCHITECTURAL LOOPHOLES.**  
All 5 core architectural requirements are strictly enforced in production code, backed by successful builds, runtime verification on the SharePoint Workbench (Sidebar Flyout), and live Microsoft Graph email deliveries in Microsoft Outlook.

---

## 1. State Machine Validation
**Specification:** Verify that `approvalHandler.ts` correctly patches `sct_request_status` and `cr2a4_pendingtier` exactly as defined in the 5-stage sequential flow.

### Audit Findings: **COMPLIANT (100%)**
- **File:** [approvalHandler.ts](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/approvalHandler.ts#L350-L430)
- **Verified Column Logical Names:**
  - `sct_request_status` (Numeric OptionSet: `854000001` = Submitted, `854000002` = In Review, `854000003` = PO Issued, `854000004` = Rejected, `854000005` = Returned)
  - `cr2a4_pendingtier` (Single-line Text: `'Manager'`, `'Tier 1'`, `'Tier 2'`, `'Tier 3'`, `null`)

### Sequential Flow Matrix Verification:
| Stage / Acting Role | Decision | Next `sct_request_status` | Next `cr2a4_pendingtier` | Code Reference |
| :--- | :--- | :--- | :--- | :--- |
| **New Submission** | Submit | `Submitted (854000001)` | `'Manager'` | `NewRequestForm.tsx` [L711](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/newRequest/NewRequestForm.tsx#L711) |
| **Line Manager** | Approve | `In Review (854000002)` | `'Tier 1'` | `approvalHandler.ts` [L358](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/approvalHandler.ts#L358) |
| **Tier 1 Approver** | Approve | `In Review (854000002)` | `'Tier 2'` | `approvalHandler.ts` [L359](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/approvalHandler.ts#L359) |
| **Tier 2 Approver** | Approve | `In Review (854000002)` | `'Tier 3'` | `approvalHandler.ts` [L360](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/approvalHandler.ts#L360) |
| **Tier 3 Approver** | Approve (Final) | `PO Issued (854000003)` | `null` | `approvalHandler.ts` [L362-L364](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/approvalHandler.ts#L362-L364) |
| **Any Approver** | Reject | `Rejected (854000004)` | `null` | `approvalHandler.ts` [L351-L353](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/approvalHandler.ts#L351-L353) |
| **Any Approver** | Return | `Returned (854000005)` | `null` | `approvalHandler.ts` [L368-L371](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/approvalHandler.ts#L368-L371) |

### Dual Table Transaction Integrity:
1. **Approval Log (POST):** Dispatches to `sct_approvallogs` with `sct_requestid@odata.bind` and approver bind (Lines 374–390).
2. **Parent Request (PATCH):** Dispatches to `sct_requests(${requestId})` updating `sct_request_status` and `cr2a4_pendingtier` (Lines 410–430).
3. **Audit Log (POST):** Dual-auditing writes human-readable transitions to `sct_audits` (Lines 432–480).

---

## 2. Role Filtering & Segregation of Duties (SoD)
**Specification:** Verify that `Dashboard.tsx` or `ReviewQueue.tsx` filters pending queues correctly for Managers and Tiers, and explicitly enforces the SoD rule (`reqId !== currentUserId`).

### Audit Findings: **COMPLIANT (100%)**
- **Files:** 
  - [Dashboard.tsx](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/dashboard/Dashboard.tsx#L857-L920)
  - [ReviewQueue.tsx](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/reviewQueue/ReviewQueue.tsx#L542-L551)

### Rule Enforcement Analysis:
1. **Global SoD Guard (`Dashboard.tsx` Lines 858–862 & `ReviewQueue.tsx` Lines 547–550):**
   ```typescript
   // Evaluated at the top before evaluating role permissions:
   if (!isRequesterRole && safeUserId && (reqId === safeUserId || createdById === safeUserId)) {
     return false; // User cannot view or approve their own request in an approval queue
   }
   ```
2. **Requester Isolation (`Dashboard.tsx` Lines 865–868):**
   - When `activeRole === 'Requester'`, returns `true` strictly if `reqId === safeUserId || createdById === safeUserId`.
3. **Manager Pending Filter (`Dashboard.tsx` Lines 873–886):**
   - Matches `isManagerAssigned && isPendingManagerTier`.
   - Requires `managerId === safeUserId || isDirectReport` AND `(rawPt === 'Manager' || rawStatus === 854000001)`.
   - Excludes requests already approved by the manager.
4. **Tier 1, Tier 2, Tier 3 Approver Queues (`Dashboard.tsx` Lines 889–912):**
   - Maps role to exact Dataverse string (`Tier 1`, `Tier 2`, `Tier 3`).
   - Requires `rawPt === expectedTierValue` AND `sct_request_status === 'In Review'` (`854000002`).

---

## 3. Email Routing & CC Chain
**Specification:** Verify that `GraphMailService.ts` and the approval handler use `/v1.0/me/sendMail`, CC the requester and actor, and pass prior logs for the cumulative HTML table.

### Audit Findings: **COMPLIANT (100%)**
- **Files:**
  - [GraphMailService.ts](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/GraphMailService.ts#L13-L43)
  - [approvalHandler.ts](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/approvalHandler.ts#L540-L603)
  - [NewRequestForm.tsx](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/newRequest/NewRequestForm.tsx#L765-L820)

### Native Microsoft Graph Dispatch:
- Dispatches directly via `graphClient.api('/me/sendMail').version('v1.0').post({ message, saveToSentItems: true })`.
- Runs under the authenticated SPFx session (`/me`), completely eliminating Power Automate license costs, throttling, and latency.

### Full Recipient & CC Routing Matrix:
| Trigger Event | To Recipients | CC Recipients | Cumulative Logs Passed? |
| :--- | :--- | :--- | :--- |
| **New Submission** | Assigned Line Manager | Requester (`currentUserEmail`) | Initial log empty (`[]`) |
| **Manager Approves** | Tier 1 Approvers | `[RequesterEmail, currentUserEmail]` | Yes (Manager decision included) |
| **Tier 1 Approves** | Tier 2 Approvers | `[RequesterEmail, currentUserEmail]` | Yes (Manager + Tier 1 decisions) |
| **Tier 2 Approves** | Tier 3 Approvers | `[RequesterEmail, currentUserEmail]` | Yes (Manager + Tier 1 + Tier 2) |
| **Tier 3 Approves (Final)**| Requester Email | `[currentUserEmail]` | Yes (Complete 3-tier trail) |
| **Reject (Any Tier)** | Requester Email | `[currentUserEmail]` | Yes (All prior attempts + reject) |
| **Return (Any Tier)** | Requester Email | `[currentUserEmail]` | Yes (All prior attempts + return) |

- **Cumulative Log Extraction:** `approvalHandler.ts` [L541-L547](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/approvalHandler.ts#L541-L547) awaits `service.getApprovalLogs(requestId)` immediately after inserting the new decision log, ensuring the current decision is rendered in the HTML table.
- **Runtime Verification:** Live email received in Outlook confirms subject `[SCT Action Required] SCT-2026-1050 — Tier 3 Approval`, To: Aniket Desai & Bhakti Gandhi, CC: Manav Barad, with cumulative history table showing Manager, Tier 1, and Tier 2 approved badges.

---

## 4. GUID Prevention
**Specification:** Verify that the email HTML generator explicitly utilizes `@OData.Community.Display.V1.FormattedValue` for lookup columns.

### Audit Findings: **COMPLIANT (100%)**
- **File:** [GraphMailService.ts](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/services/GraphMailService.ts#L68-L180)

### Lookup Sanitization & Formatted Value Priority:
1. **Requester Name (Lines 83–91):**
   Prioritizes `_sct_requester_value@OData.Community.Display.V1.FormattedValue`, then `_createdby_value@OData.Community.Display.V1.FormattedValue`, guarded by `cleanDisplayVal()`.
2. **Business Unit (Lines 94–101):**
   Prioritizes `sct_businessunit@OData.Community.Display.V1.FormattedValue`.
3. **Site / Location (Lines 104–111):**
   Prioritizes `_sct_site_value@OData.Community.Display.V1.FormattedValue`.
4. **Estimated Spend Total (Lines 114–121):**
   Prioritizes `sct_estimated_total_value@OData.Community.Display.V1.FormattedValue` with localized currency fallback (`$X,XXX.XX`).
5. **Approval History Table Approver (Lines 162–170):**
   Prioritizes `_sct_approver_value@OData.Community.Display.V1.FormattedValue` and `_createdby_value@OData.Community.Display.V1.FormattedValue`.
6. **GUID Rejection Guard (`isGuid()` Regex):**
   ```typescript
   const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
   function cleanDisplayVal(val: unknown, fallback: string): string {
     if (val === undefined || val === null) return fallback;
     const str = String(val).trim();
     if (!str || isGuid(str)) return fallback;
     return str;
   }
   ```
   Ensures raw GUIDs are discarded in favor of human-readable display names.

---

## 5. UI Locking & Sidebar Flyout Experience
**Specification:** Verify that `NewRequestForm` and action popups utilize an `isSubmitting` state to disable buttons during async operations.

### Audit Findings: **COMPLIANT (100%)**
- **Files:**
  - [NewRequestForm.tsx](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/newRequest/NewRequestForm.tsx#L681-L683)
  - [TakeActionModal.tsx](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/common/TakeActionModal.tsx#L143-L330)
  - [ReviewQueue.tsx](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/reviewQueue/ReviewQueue.tsx#L625-L665)

### Verification Details:
1. **New Request Form:**
   - [L682](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/newRequest/NewRequestForm.tsx#L682): `this.setState({ [key]: true, isSubmitting: true, submitError: null })`.
   - [L1773](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/newRequest/NewRequestForm.tsx#L1773): `disabled={isWorking || this.state.isSubmitting || submitting || ...}`.
   - [L1778-L1780](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/newRequest/NewRequestForm.tsx#L1778-L1780): Renders inline `<Spinner size={SpinnerSize.xSmall} /> Submitting…`.
2. **Take Action Popup (Sidebar Flyout):**
   - Renders as a modern right-side panel using Fluent UI `Panel` (`type={PanelType.large}`, Line 235), avoiding center-screen obstruction.
   - [L144](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/common/TakeActionModal.tsx#L144): Sets `this.setState({ isSubmitting: true })`.
   - [L219](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/common/TakeActionModal.tsx#L219): Computes `canAct = !isSubmitting && !decisionLock.isLocked`.
   - [L291, L307, L323](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/common/TakeActionModal.tsx#L291): Dynamically replaces icons with `<Spinner size={SpinnerSize.xSmall} />` and labels with "Saving…".
   - [L262, L329](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/common/TakeActionModal.tsx#L262): Close/Dismiss buttons are disabled with `disabled={isSubmitting}` to prevent closing mid-patch.
3. **Review Queue:**
   - [L625](file:///c:/Users/Aniket.Desai/FortrexDataverse/src/components/reviewQueue/ReviewQueue.tsx#L625): Locks actions with `managerSaving: true` during Dataverse commitments.

---

## Summary of Completed Solution Packages
- **Bundle File:** `dist/ad-fortrex-webpart-web-part_277bd57194077504cd6d.js` (Clean build, 0 errors)
- **SPFx Solution Package:** `sharepoint/solution/ad-dataverse-fortrex.sppkg`
- **Current Version:** `1.0.0.9`
- **Deploy Target:** SharePoint App Catalog (`https://devrishabhsoft.sharepoint.com/sites/FortrexSCT`)
- **Verified Operations:**
  - Sidebar flyout displays sequential approvals and historical logs correctly.
  - Multi-tier sequential approval triggers email notifications automatically via `/me/sendMail`.
  - Recipient lists and CC trails maintain full audit visibility.
  - No raw GUIDs appear in any email notification or UI element.
  - Double submissions are fully prevented across all forms and modals.
