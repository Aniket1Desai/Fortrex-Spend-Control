# Fortrex Spend Control Tracker (SCT) — Approval Workflow, Status Column & Tier Architecture

This document provides a complete technical and functional specification of how the **Status Column**, the **Approval Chain**, the **Dataverse Hierarchy**, and all **Approval Tiers and Groups** work together in the Fortrex SCT solution.

---

## 1. Executive Summary & Core Architectural Question

> **Question:** *"If we create this logic for the Manager, does it work the same way for all Tiers and Groups?"*
>
> **Answer: YES.** The workflow is architected with a unified, symmetric state machine across all 4 approval tiers:
> 1. **Manager Approval** (Direct Line Manager / Site Manager)
> 2. **Tier 1 Approvers Group** (Procurement & Initial Triage Gate)
> 3. **Tier 2 Approvers Group** (FP&A & Commercial Review Gate)
> 4. **Tier 3 Approvers Group** (CFO & Executive Leadership Gate)

Across **every single tier**, the system enforces three universal rules:
1. **Self-Request Isolation**: No approver (Manager, Tier 1, Tier 2, or Tier 3) sees their own submitted requests in their approval queue. A user's personal requests always live in the dedicated **"My Requests"** tab.
2. **Terminal State Exclusion**: Inactive records (`PO Issued`, `Rejected`, `Closed`, `Draft`, `Returned`) never appear in any "Pending" queue.
3. **Queue Ownership Hand-off**: When an approver clicks **Approve**, the request instantly leaves their queue and appears in the next tier's queue.

---

## 2. The Status Column & Queue Pointer Architecture

The approval engine is powered by two synchronized columns on the Dataverse `sct_requests` table:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                              sct_requests                                   │
├──────────────────────────────────────┬──────────────────────────────────────┤
│ sct_request_status (Choice / Int32)  │ cr2a4_pendingtier (Single Text)     │
├──────────────────────────────────────┼──────────────────────────────────────┤
│ 854000000 = Draft                    │ null                                 │
│ 854000001 = Submitted                │ 'Manager'                            │
│ 854000002 = In Review                │ 'Manager' | 'Tier 1'|'Tier 2'|'Tier 3│
│ 854000003 = PO Issued (Final)        │ null                                 │
│ 854000004 = Rejected                 │ null                                 │
│ 854000005 = Returned for Info        │ null                                 │
│ 854000006 = Closed                   │ null                                 │
└──────────────────────────────────────┴──────────────────────────────────────┘
```

### 2.1 UI Status Column Formatting Logic
In the UI table grids (`Dashboard.tsx`, `MyRequests.tsx`, and `ReviewQueue.tsx`), the **Status** column does not merely show raw database numbers. It dynamically formats the status with the active queue context:

| `sct_request_status` Value | `cr2a4_pendingtier` Value | UI Display Label | Visual Badge Style |
|---|---|---|---|
| `854000000` (`Draft`) | `null` | **Draft** | Grey (`#F3F2F1`, text `#605E5C`) |
| `854000002` (`In Review`) | `'Manager'` | **In Review (Manager)** | Mint Cyan (`#E3F2FB`, text `#0277BD`) |
| `854000002` (`In Review`) | `'Tier 1'` | **In Review (Tier 1)** | Mint Cyan (`#E3F2FB`, text `#0277BD`) |
| `854000002` (`In Review`) | `'Tier 2'` | **In Review (Tier 2)** | Mint Cyan (`#E3F2FB`, text `#0277BD`) |
| `854000002` (`In Review`) | `'Tier 3'` | **In Review (Tier 3)** | Mint Cyan (`#E3F2FB`, text `#0277BD`) |
| `854000003` (`PO Issued`) | `null` | **PO Issued** | Green (`#D4EDDA`, text `#155724`) |
| `854000004` (`Rejected`) | `null` | **Rejected** | Red (`#FDE7E9`, text `#A4262C`) |
| `854000005` (`Returned`) | `null` | **Returned** | Amber (`#FFF4CE`, text `#7A4F01`) |
| `854000006` (`Closed`) | `null` | **Closed** | Grey (`#EDEBE9`, text `#323130`) |

---

## 3. End-to-End State Machine Flow

```
                        [ Requester Submits ]
                                  │
                                  ▼
                   sct_request_status: In Review (854000002)
                   cr2a4_pendingtier: 'Manager'
                   UI Status: "In Review (Manager)"
                                  │
               ┌──────────────────┴──────────────────┐
               │                                     │
           [ Approve ]                           [ Reject ]  ──► Status: Rejected (854000004)
               │                                     │           pendingtier: null
               ▼                                     ▼
     sct_request_status: In Review               [ Return ]  ──► Status: Returned (854000005)
     cr2a4_pendingtier: 'Tier 1'                                 pendingtier: null
     UI Status: "In Review (Tier 1)"                             (Requester edits & resubmits)
               │
               ▼
     [ Tier 1 Approves ]
     sct_request_status: In Review
     cr2a4_pendingtier: 'Tier 2'
     UI Status: "In Review (Tier 2)"
               │
               ▼
     [ Tier 2 Approves ]
     sct_request_status: In Review
     cr2a4_pendingtier: 'Tier 3'
     UI Status: "In Review (Tier 3)"
               │
               ▼
     [ Tier 3 Approves (CFO / Final) ]
     sct_request_status: PO Issued (854000003)
     cr2a4_pendingtier: null
     UI Status: "PO Issued"
     (Ready for Procurement to issue PO in D365)
```

---

## 4. How the Approval Tiers & Groups Work

### 4.1 Tier 1: Line Manager Approval
- **Membership**: Determined by the Dataverse `systemusers` organizational hierarchy.
  - In `systemusers`, each user has `_parentsystemuserid_value` pointing to their manager.
  - When **Bhakti Gandhi** submits a request, her manager is **Aniket Desai**.
- **Pending Queue Rule**:
  - Shows requests where `_cr2a4_manager_value === currentUserId` **OR** the requester is in the manager's `directReportIds`.
  - **Excludes own requests**: Aniket Desai never sees his own spend requests in "Pending My Approval".
  - **Excludes completed requests**: `PO Issued` records do not appear here.
- **Action Result**:
  - `Approve`: Updates status to `In Review`, moves `cr2a4_pendingtier` to `'Tier 1'`, and sends Graph email to Tier 1 Approvers.

### 4.2 Tier 2: Tier 1 Approvers Group
- **Membership**: Azure AD / Entra ID Security Group `SCT-Fortrex_Approvers-Tier1` (`TIER1_GROUP_ID`).
- **Pending Queue Rule**:
  - Shows requests where `sct_request_status === 854000002` (`In Review`) AND `cr2a4_pendingtier === 'Tier 1'`.
  - **Excludes own requests**: If a Tier 1 approver submitted a request, it is excluded from their approval queue.
- **Action Result**:
  - `Approve`: Updates status to `In Review`, advances `cr2a4_pendingtier` to `'Tier 2'`, and sends Graph email to Tier 2 Approvers.

### 4.3 Tier 3: Tier 2 Approvers Group
- **Membership**: Azure AD / Entra ID Security Group `SCT-Fortrex_Approvers-Tier2` (`TIER2_GROUP_ID`).
- **Pending Queue Rule**:
  - Shows requests where `sct_request_status === 854000002` (`In Review`) AND `cr2a4_pendingtier === 'Tier 2'`.
  - Excludes own requests.
- **Action Result**:
  - `Approve`: Updates status to `In Review`, advances `cr2a4_pendingtier` to `'Tier 3'`, and sends Graph email to Tier 3 Approvers.

### 4.4 Tier 4: Tier 3 Approvers Group (CFO / Final Sign-Off)
- **Membership**: Azure AD / Entra ID Security Group `SCT-Fortrex_Approvers-Tier3` (`TIER3_GROUP_ID`).
- **Pending Queue Rule**:
  - Shows requests where `sct_request_status === 854000002` (`In Review`) AND `cr2a4_pendingtier === 'Tier 3'`.
  - Excludes own requests.
- **Action Result**:
  - `Approve`: Transitions status to `854000003` (`PO Issued`), clears `cr2a4_pendingtier = null`, and sends confirmation email to requester confirming PO authorization.

---

## 5. UI Navigation Slices by Active Role

The application adapts dynamically depending on the selected role in the **Role Switcher**:

| Active Role | Dashboard Tab | Card Header Title | Records Displayed |
|---|---|---|---|
| **Requester** | Pending | My Pending Requests | Only requests submitted by the logged-in user that are In Review / Submitted |
| **Requester** | Completed | My Completed Requests | Requests submitted by the user that are PO Issued, Rejected, or Closed |
| **Manager** | Pending | **Pending My Approval** | Requests submitted by direct reports (e.g. Bhakti Gandhi) awaiting manager approval |
| **Manager** | Completed | Completed | Requests previously reviewed/approved by this manager |
| **Approver Tier 1**| Pending | **Pending — Tier 1** | Requests where `cr2a4_pendingtier = 'Tier 1'` (excluding self-submitted items) |
| **Approver Tier 1**| Completed | Completed | Requests this user previously approved at Tier 1 |
| **Approver Tier 2**| Pending | **Pending — Tier 2** | Requests where `cr2a4_pendingtier = 'Tier 2'` (excluding self-submitted items) |
| **Approver Tier 2**| Completed | Completed | Requests this user previously approved at Tier 2 |
| **Approver Tier 3**| Pending | **Pending — Tier 3** | Requests where `cr2a4_pendingtier = 'Tier 3'` (excluding self-submitted items) |
| **Approver Tier 3**| Completed | Completed | Requests this user previously approved at Tier 3 |

### The Global "My Requests" Tab (Top Navigation)
- Regardless of which role is active in the switcher (Manager, Approver, or Requester), the top **"My Requests"** tab **always** loads:
  ```odata
  sct_requests?$filter=_sct_requester_value eq {currentSystemUserId}
  ```
- This guarantees that an approver's personal requests (e.g. `SCT-2026-1017`) are never lost or hidden.

---

## 6. Audit Trail & Logging Specifications

Every decision records dual audit entries in Dataverse:

1. **`sct_approvallogs` (Approval Decision Table)**:
   - `sct_tiername`: Integer reflecting the tier (`1 = Manager`, `2 = Tier 1`, `3 = Tier 2`, `4 = Tier 3`).
   - `sct_approvalstatus`: Integer decision (`1 = Approved`, `2 = Rejected`, `3 = Return for Info`).
   - `sct_comments`: Approver rationale (defaults to `APPROVE by <Tier>` if left empty).
   - `sct_requestid@odata.bind`: Link to the `sct_requests` record.
   - `sct_approverid@odata.bind`: Link to the approver's `systemusers` record.

2. **`sct_auditlogs` (Field Difference Table)**:
   - Captures previous status/tier vs. new status/tier for change tracking and compliance audits.

---

## 7. Package & Deployment Reference

- **Package Artifact**: `sharepoint/solution/ad-dataverse-fortrex.sppkg`
- **Output Path**: `C:\Users\Aniket.Desai\FortrexDataverse\sharepoint\solution\ad-dataverse-fortrex.sppkg`
- **Component ID**: `90faa43b-d5ad-4ef9-bb73-3655876b4cde`
- **Build Status**: Verified with `npx heft build --clean --production` and `npx heft package-solution --production` (Exit Code 0).

