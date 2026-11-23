# Fortrex Spend Control Tracker (SCT) — Approvals Chain & System Overview

This document provides a comprehensive technical and operational breakdown of how the **Approvals Chain** works in the Fortrex SCT application, along with an architectural overview of all project files and components.

---

## 1. Executive Summary: The Two Approval Dimensions

The Fortrex SCT system operates on **two interrelated approval models**:

1. **Policy-Driven Pre-Approval Matrix (Schedule A / Spend Policy v2.0):**
   - Requester must collect and attach written pre-approval evidence (L1, L2, L3) *before* submitting a request.
   - Evaluated by `ApprovalRoutingService.ts` based on spend thresholds and supplier status.
2. **Sequential Active Approval State Machine (SPFx Code-Driven Engine):**
   - Once submitted, requests flow through a live, sequential approval chain directly inside SPFx:
     $$\text{Requester} \longrightarrow \text{Manager} \longrightarrow \text{Tier 1 Approver} \longrightarrow \text{Tier 2 Approver} \longrightarrow \text{Tier 3 Approver (CFO)} \longrightarrow \text{PO Issued} \longrightarrow \text{Receiving / Closed}$$
   - Managed by `approvalHandler.ts`, recorded in Dataverse `sct_approvallogs`, tracked via `cr2a4_pendingtier` on `sct_requests`, and notified via Microsoft Graph `/sendMail`.

---

## 2. Detailed Breakdown: How the Approvals Chain Works

### 2.1 The Sequential Approval State Machine

```
                   [New Request Form]
                           │
                           │  Status: In Review (854000002)
                           │  cr2a4_pendingtier = 'Manager'
                           │  _cr2a4_manager_value = Manager GUID
                           │  Graph API: sendManagerApprovalEmail()
                           ▼
                    ┌──────────────┐
                    │   MANAGER    │
                    │   APPROVAL   │
                    └──────┬───────┘
                           │
       ┌───────────────────┼───────────────────┐
       │ Approve           │ Reject            │ Return for Info
       ▼                   ▼                   ▼
 Status: In Review   Status: Rejected    Status: Returned
 pendingtier:        pendingtier: null   pendingtier: null
 'Tier 1'            Log: Rejected       Log: Return for Info
 Email: Tier 1       Email: Manager FYI  Email: Manager FYI
       │
       ▼
 ┌──────────────┐
 │    TIER 1    │
 │   APPROVER   │
 └─────┬────────┘
       │ Approve
       ▼
 Status: In Review, pendingtier: 'Tier 2', Email: Tier 2
       │
       ▼
 ┌──────────────┐
 │    TIER 2    │
 │   APPROVER   │
 └─────┬────────┘
       │ Approve
       ▼
 Status: In Review, pendingtier: 'Tier 3', Email: Tier 3
       │
       ▼
 ┌──────────────┐
 │    TIER 3    │
 │(CFO / Final) │
 └─────┬────────┘
       │ Approve
       ▼
 Status: PO Issued (854000003)
 pendingtier: null
 (Ready for Procurement PO issuance & Receiving)
```

---

### 2.2 Approval Routing Matrix & State Transitions

The core state machine in `approvalHandler.ts` transitions the Dataverse status column (`sct_request_status`) and queue pointer (`cr2a4_pendingtier`):

| Actor Role | Action | Status Integer (`sct_request_status`) | Pending Tier (`cr2a4_pendingtier`) | UI Status Column Display | Log Tier (`sct_tiername`) | Log Status (`sct_approvalstatus`) | Next Step / Notification |
|---|---|---|---|---|---|---|---|
| **Requester** | Save Draft | `854000000` (`Draft`) | `null` | **Draft** | — | — | Saved in Requester's Drafts |
| **Requester** | Submit | `854000002` (`In Review`) | `'Manager'` | **In Review (Manager)** | — | — | Appears in Manager Queue; Email to Manager |
| **Manager** | Approve | `854000002` (`In Review`) | `'Tier 1'` | **In Review (Tier 1)** | 1 (`Manager`) | 1 (`Approved`) | Advances to Tier 1 Queue; Email to Tier 1 Approvers |
| **Manager** | Reject | `854000004` (`Rejected`) | `null` | **Rejected** | 1 (`Manager`) | 2 (`Rejected`) | Workflow stops; Email with comments to Requester |
| **Manager** | Return | `854000005` (`Returned`) | `null` | **Returned** | 1 (`Manager`) | 3 (`Return for Info`)| Sent back to Requester for edits & resubmission |
| **Tier 1** | Approve | `854000002` (`In Review`) | `'Tier 2'` | **In Review (Tier 2)** | 2 (`Tier 1`) | 1 (`Approved`) | Advances to Tier 2 Queue; Email to Tier 2 Approvers |
| **Tier 1** | Reject | `854000004` (`Rejected`) | `null` | **Rejected** | 2 (`Tier 1`) | 2 (`Rejected`) | Workflow stops; Email with comments to Requester |
| **Tier 1** | Return | `854000005` (`Returned`) | `null` | **Returned** | 2 (`Tier 1`) | 3 (`Return for Info`)| Sent back to Requester for edits & resubmission |
| **Tier 2** | Approve | `854000002` (`In Review`) | `'Tier 3'` | **In Review (Tier 3)** | 3 (`Tier 2`) | 1 (`Approved`) | Advances to Tier 3 Queue; Email to Tier 3 Approvers |
| **Tier 2** | Reject | `854000004` (`Rejected`) | `null` | **Rejected** | 3 (`Tier 2`) | 2 (`Rejected`) | Workflow stops; Email with comments to Requester |
| **Tier 2** | Return | `854000005` (`Returned`) | `null` | **Returned** | 3 (`Tier 2`) | 3 (`Return for Info`)| Sent back to Requester for edits & resubmission |
| **Tier 3** | Approve | `854000003` (`PO Issued`) | `null` | **PO Issued** | 4 (`Tier 3`) | 1 (`Approved`) | Approval chain complete; Email to Requester to issue PO |
| **Tier 3** | Reject | `854000004` (`Rejected`) | `null` | **Rejected** | 4 (`Tier 3`) | 2 (`Rejected`) | Workflow stops; Email with comments to Requester |
| **Tier 3** | Return | `854000005` (`Returned`) | `null` | **Returned** | 4 (`Tier 3`) | 3 (`Return for Info`)| Sent back to Requester for edits & resubmission |
| **Requester** | Resubmit Returned | `854000002` (`In Review`) | `'Manager'` | **In Review (Manager)** | — | — | Re-enters approval chain at Manager step |

---

### 2.3 The Role Resolution Engine

The application dynamically resolves user roles through a **3-step priority hierarchy** (`docs/role-detection-reference.md` and `AppShell.tsx`):

1. **Step 1: Hardcoded Override Map (`USER_ROLE_MAP` / `MANAGER_AADIDS`)**
   - Checks the logged-in user's Azure AD Object ID (`pageContext.aadInfo.userId`).
   - If present, roles are assigned instantly with **zero network latency**.
2. **Step 2a: Manager Role via Dataverse `systemuser` Hierarchy**
   - SPFx calls `DataverseService.getCurrentSystemUser(aadOid)`.
   - Checks `hasDirectReportsInDataverse(systemuserid)`:
     ```odata
     GET systemusers?$filter=_parentsystemuserid_value eq {systemuserid}&$top=1
     ```
   - If 1 or more direct reports exist in Dataverse, the user gets the **Manager** role.
   - Fallback: Graph API `/me/directReports`.
3. **Step 2b: Tier Approvers via Azure AD / Entra ID Security Groups**
   - SPFx calls Microsoft Graph API `POST /me/checkMemberGroups` checking against three Entra ID Security Groups:
     - `TIER1_GROUP_ID` (`fdbf07a0-548b-4ccf-875b-900cb1b24a1c`): SCT-Fortrex_Approvers-Tier1
     - `TIER2_GROUP_ID` (`6a8a4077-c78f-4af2-ae5e-12428d1dcbc3`): SCT-Fortrex_Approvers-Tier2
     - `TIER3_GROUP_ID` (`e63dbc4f-6ad2-4969-a518-b229836756af`): SCT-Fortrex_Approvers-Tier3
4. **Step 3: Default Role**
   - Every user receives the `Requester` role by default.

---

### 2.4 Pre-Approval Evidence Policy (Schedule A / `ApprovalRoutingService.ts`)

Before submission, requesters must attach written pre-approval files based on estimated spend and supplier status:

| Condition | Required Pre-Approval Evidence | Logic in Code |
|---|---|---|
| Existing Supplier, Value $\le \$2,500$ | **L1** — Line / Site Manager | `T_L1 = 2500` |
| Existing Supplier, $\$2,500 < \text{Value} \le \$10,000$ | **L2** — Department Head | `T_L2 = 10000` |
| Existing Supplier, $\text{Value} > \$10,000$ | **L3** — CFO / Executive Leadership | Above `T_L2` |
| New Supplier (any spend value) | **L1 + L2 + L3** (all three required) | `isNewSupplier === true` |

#### Rule R10 — Conflict of Interest / Self-Approval Auto-Escalation
If the resolved approver for any level is the **same person as the requester**, `ApprovalRoutingService.resolveApprovers()` automatically escalates the approval to the next level up (e.g., L1 $\rightarrow$ L2) and flags the record with `R10 ESCALATED`.

#### Active Delegate Substitution
If an approver has configured a delegate in `sct_roles` with `sct_effectivedate` and `sct_expirydate`, and the current date falls within that active window, the delegate's GUID is automatically substituted and marked with a `DELEGATE` badge.

---

### 2.5 Email Notifications via Microsoft Graph (`GraphMailService.ts`)

Instead of relying on fragile or asynchronous Power Automate flows, all emails are sent directly by SPFx using `graphClient.api('/me/sendMail')`:
1. **Manager Notification:** On submission, the requester's direct manager receives an HTML card detailing Request Title, Requester Name, Estimated Amount, and a direct link to review.
2. **Next Tier Notification:** When a manager or tier approver approves, the pre-configured members of the next tier group receive a formatted email notifying them that the request is awaiting their review.

---

## 3. Comprehensive Overview of All Project Files

### 3.1 Documentation & Specification Files

- **`Defination File.txt`**:
  - Detailed Dataverse entity definition and schema map for `sct_request`.
  - Fixes the bug where `sct_name` was sent (which doesn't exist; the primary key/autonumber is `sct_sct_reference`).
  - Lists strict data types: `Edm.Date` (strictly `YYYY-MM-DD`), `Edm.Boolean` (`true`/`false`), and lookup `@odata.bind` properties.
- **`fortrex_sct_project_documentation.md`**:
  - Full system documentation covering business requirements ("No PO, No Pay"), architecture, tech stack (SPFx 1.23.2 + React 17 + Fluent UI + Dataverse v9.2), Dataverse schema, UI design, validation rules, two-stage approval workflow, and deployment.
- **`sct_phased_implementation_plan(phase2).md`**:
  - Strategic architecture plan detailing the transition to code-driven approvals bypassing Power Automate:
    - Phase 1: Dynamic Role & View Resolution via Graph & Dataverse.
    - Phase 2: Multi-step Request Wizard UI.
    - Phase 3: SPFx-Driven Approval Engine & MS Graph mailer.
    - Phase 4: Approver Dashboard & "Take Action" Dialog.
- **`fortrex_sct_phased_execution_plan.md`**:
  - Step-by-step developer implementation guide across Phases 2 through 11.
- **`Pending Modules Implementation Plan.pdf`**:
  - Gap Analysis and Phase 2 Roadmap comparing as-built code against Spend Policy 2.0:
    1. Standing Authorizations & Call-Offs (Blanket POs with 80% cap warnings).
    2. Change Requests & Amendments (`sct_amendments` table; re-approval if cumulative total exceeds \$10k).
    3. Emergency Ratification Workflow (48-hour ratification mode).
    4. Multi-BU Cost Allocation Matrix (split costs across business units).
    5. Supplier Onboarding & Compliance Capture.
    6. Legal Review (for contracts > \$50k) & AP Breach Register (`sct_breaches`).
- **`docs/role-detection-reference.md`**:
  - Reference manual on how `AppShell._loadUserRoles()` determines user roles at startup (hardcoded map $\rightarrow$ Dataverse `systemuser` $\rightarrow$ Entra ID `checkMemberGroups`).
- **`README.md`**:
  - Project setup, building (`heft build`), bundling, and solution packaging guide.

---

### 3.2 Services Layer (`src/services/`)

- **`approvalHandler.ts`**:
  - The sequential approval state machine. Handles actions (`approve`, `reject`, `return`), determines next statuses, updates `cr2a4_pendingtier` on `sct_request`, and creates entries in `sct_approvallogs`.
- **`ApprovalRoutingService.ts`**:
  - The Schedule A pre-approval calculator and dynamic role resolution engine. Resolves L1–L4 approvers from `sct_roles` per site, enforces Rule R10 self-approval escalations, and checks active delegate dates.
- **`DataverseService.ts`**:
  - The primary Dataverse data access layer using `AadHttpClient`.
  - Provides CRUD operations for `sct_requests`, `sct_sites_masters`, `sct_suppliers`, `sct_roles`, `sct_approvallogs`, `sct_auditlogs`, and `annotations` (attachments).
  - Handles 3-source merged data fetching for managers (`getManagerMasterRequests`), tier queue querying (`getTierPendingRequestIds`), and relationship navigation property resolution (`getNavPropName`).
- **`GraphMailService.ts`**:
  - Direct email dispatch service via Microsoft Graph `/me/sendMail` and manager resolution via `/me/manager`.
- **`IDataverseService.ts`**:
  - Interface contracts defining all Dataverse service methods.

---

### 3.3 Models Layer (`src/models/`)

- **`ISCTRequest.ts`**:
  - TypeScript interfaces and enums for `sct_requests`, including `SCTRequestStatus` (Draft=854000000, Submitted=854000001, InReview=854000002, POIssued=854000003, Rejected=854000004, Returned=854000005, Closed=854000006), `SCTRequestType`, `SCTUrgency`, and `ISCTRequestWrite`.
- **`IRole.ts`**:
  - Type definitions for `sct_roles`, tracking `ApprovalLevel` (L1, L2, L3, L4), spend limits (`sct_approvalspendlimit`), site lookups, and primary/delegate person lookups.
- **`ISiteMaster.ts`**:
  - Types for site metadata, cost centres, regions, and default L1 budget owners.
- **`ISupplier.ts`**:
  - Types for pre-approved vendors in `sct_suppliers`.
- **`index.ts`**:
  - Central re-export of all TypeScript models.

---

### 3.4 UI Components (`src/components/`)

- **`shell/AppShell.tsx`**:
  - Web part navigation shell, responsive header, brand styling, tab switching (Dashboard, New Request, My Requests, Review Queue), and the dynamic **Role Switcher Dropdown**.
- **`dashboard/Dashboard.tsx`**:
  - Central dashboard featuring 4 live metric stat cards, searchable and filterable `DetailsList`, tabbed views (Pending vs. Completed / My Requests), and the **Take Action Dialog** for approvers.
- **`dashboard/dashboardUtils.ts`**:
  - Formatting utilities for currencies, dates, status badges, and column mappings.
- **`dashboard/RequestDetailPanel.tsx`**:
  - Flyout panel providing a complete, read-only summary of all fields for a selected request.
- **`newRequest/NewRequestForm.tsx`**:
  - Full request creation form enforcing Schedule C policies, dynamic supplier toggles, compliance checks, pre-approval evidence requirements, attachment uploads, and manager email routing.
- **`reviewQueue/ReviewQueue.tsx`**:
  - Workbench for lifecycle stages:
    - Stage 3: SCT Desk Triage
    - Stage 4: Procurement Sourcing & PO entry
    - Stage 5a: FP&A Review
    - Stage 5b: CFO Gate
    - Stage 6: Goods/Services Receiving
    - Includes a dedicated Manager Review Queue when viewing as Manager.
- **`myRequests/MyRequests.tsx`**:
  - Personal request tracker for requesters, showing visual stage progress pills and request history.
- **`groupRoleConfig/GroupRoleConfig.tsx`**:
  - Administration panel to view and modify `sct_roles` assignments, approval tiers, and spend limits inline.
- **`history/RequestLogsPanel.tsx` & `LogsContent.tsx`**:
  - Panels displaying audit logs and historical approval decision records for complete traceability.
- **`common/` (`AttachmentManager.tsx`, `ErrorBoundary.tsx`, `ErrorBanner.tsx`, `LoadingSpinner.tsx`, `EmptyState.tsx`)**:
  - Reusable presentation components for error boundaries, file uploads, empty states, and loading spinners.

---

### 3.5 Utilities & Web Part Entry Point

- **`utils/auditLogger.ts`**:
  - Utility to log attribute change diffs directly to the Dataverse `sct_auditlogs` table.
- **`webparts/adFortrexWebpart/AdFortrexWebpartWebPart.ts`**:
  - SPFx Web Part manifest entry point, initializing `AadHttpClient`, testing Dataverse connectivity, and mounting the `AppShell`.

---

## 4. Key Takeaways & Operational Recommendations

1. **Dual Routing Synergy:** The app combines pre-submission compliance (uploading signed approval evidence into Dataverse Annotations) with post-submission digital sign-offs (sequential in-app approvals in Dataverse).
2. **Zero Power Automate Dependency:** Approvals and notifications run entirely inside the SPFx client using Dataverse OData v9.2 and Microsoft Graph v3.0, providing instant execution without external flow latency or licensing friction.
3. **Audited & Compliant:** Every single decision (Approve, Reject, Return) is permanently captured in `sct_approvallogs` with the approver's `systemuserid`, tier, timestamp, and mandatory comments.

---

## 5. Vulnerabilities, Loopholes Identified & Resolved Architectural Improvements

During the review of the approval workflow, several critical vulnerabilities, edge cases, and loopholes were identified and resolved to ensure enterprise-grade security, data integrity, and compliance:

### 5.1 Self-Approval Prevention (Policy Rule R10)
- **Vulnerability**: In `approvalHandler.ts` and `ReviewQueue.tsx`, there was no server/client enforcement preventing requesters from approving their own requests if they possessed manager or approver roles.
- **Resolution**:
  - Implemented strict Rule R10 validation in `approvalHandler.ts`. If `userGuid` matches `_sct_requester_value`, the action is rejected immediately with a descriptive error.
  - Filtered out self-submitted requests in `ReviewQueue.tsx` (`_managerQueue`), ensuring users cannot see or approve their own items in the manager queue.

### 5.2 ReviewQueue Pending Stage Calculation Index Bug
- **Vulnerability**: In `ReviewQueue.tsx`, `_calculatePendingStage()` assumed `maxTier` was 0-indexed or mapped 1 to Tier 1, skipping Tier 1 and displaying "All Stages Approved" when only Tier 2 was reached (`maxTier >= 3`).
- **Resolution**:
  - Re-aligned stage mapping to match `TIER_INT` (`Manager = 1`, `Tier 1 = 2`, `Tier 2 = 3`, `Tier 3 = 4`).
  - Corrected stage labels:
    - `maxTier === 0`: Pending with Line Manager
    - `maxTier === 1`: Pending with Tier 1
    - `maxTier === 2`: Pending with Tier 2
    - `maxTier === 3`: Pending with Tier 3 / CFO
    - `maxTier >= 4`: All Stages Approved

### 5.3 Queue Hijacking & Race Condition Protection
- **Vulnerability**: `handleApprovalAction` previously trusted caller state and performed blind patches. Multiple approvers or stale browser tabs could trigger out-of-order execution, bypass tiers, or act on closed/rejected requests.
- **Resolution**:
  - Added pre-flight Dataverse validation via `getRawRequestById()` before any mutation.
  - Validates that the request status is active (`IN_REVIEW`). Requests in `PO_ISSUED`, `REJECTED`, `CLOSED`, or `DRAFT` cannot be acted upon.
  - Validates queue ownership: verifies that `cr2a4_pendingtier` matches the expected tier of the actor (`Manager`, `Tier 1`, `Tier 2`, `Tier 3`).

### 5.4 "Returned for Information" Dead-End & Resubmission Flow
- **Vulnerability**: When a request was returned for clarification (`sct_request_status = Returned`), requesters were blocked from editing because `canEdit` was restricted solely to `DRAFT`. Requesters could not address comments or resubmit.
- **Resolution**:
  - Updated `Dashboard.tsx` to permit editing when `isDraft || isReturned`.
  - Added a dedicated "Resubmit for Review" flow (`_resubmitReturned()`).
  - Resubmission updates status to `IN_REVIEW`, resets `cr2a4_pendingtier = 'Manager'`, refreshes submission timestamps, creates audit logs, and notifies the manager.

### 5.5 Notification Mismatches & Missing Requester Feedback
- **Vulnerability**:
  - On Rejection or Return, `Dashboard.tsx` previously dispatched `sendManagerApprovalEmail()` ("New Request Awaiting Your Approval"), confusing managers.
  - Requesters received no notification when their request was rejected or returned.
  - When Tier 3 final approval was granted, no notification was sent to the requester confirming that the PO could be issued.
- **Resolution**:
  - Added `sendRequesterDecisionEmail()` in `GraphMailService.ts` to notify requesters on Rejection or Return with approver comments and tier information.
  - Added `sendPOIssuedNotificationEmail()` in `GraphMailService.ts` to notify requesters and buyers when Tier 3 approval is granted.
  - Centralized notification dispatch inside `approvalHandler.ts`, eliminating duplicate or contradictory email calls from UI components.

### 5.6 Orphaned Requests (No Line Manager / Self-Manager)
- **Vulnerability**: If an employee lacked a manager in Entra ID / Dataverse or was configured as their own manager, new requests became orphaned in state `In Review` with `cr2a4_pendingtier = 'Manager'` and `_cr2a4_manager_value = null`.
- **Resolution**:
  - In `NewRequestForm.tsx`, detected absent managers or self-managers.
  - Automatically escalated the request directly to `cr2a4_pendingtier = 'Tier 1'` and dispatched approval notifications directly to the Tier 1 approvers group, ensuring zero orphaned submissions.

### 5.7 Delegate Role Date Evaluation Horizon
- **Vulnerability**: In `ApprovalRoutingService.ts`, delegate activation checked `effectiveDate` strictly with `now >= effDate`. If a role had no effective start date configured (`null`), the check returned false, preventing active delegates from acting.
- **Resolution**:
  - Updated condition to `(effDate == null || now >= effDate) && (expDate == null || now <= expDate)`. Delegates without future start dates are immediately active until expiration.

### 5.8 Dual Audit Logging
- **Resolution**: Integrated `auditLogger.logAudit()` in `approvalHandler.ts`, `Dashboard.tsx`, and `ReviewQueue.tsx`, ensuring that in addition to `sct_approvallogs`, every state transition and attribute difference is permanently captured in `sct_auditlogs` for regulatory compliance.


