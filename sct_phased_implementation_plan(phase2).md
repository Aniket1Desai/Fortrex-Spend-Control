# Spend Control Tracker (SCT) - Revised Implementation Plan

This document outlines the revised architecture and development phases to transition the SCT application to a fully SPFx code-driven solution with a Wizard-based UI and dynamic role resolution.

---

## Phase 1: Dynamic Role & View Resolution
**Objective:** Replace the static dashboard with a dynamic, context-aware interface.

*   **MS Graph Integration:** On initial load, the web part will call `GET /me/checkMemberGroups` passing the IDs for Tier 1, Tier 2, and Tier 3. It will also query if the user is listed as a `Manager` in Entra ID.
*   **Header UI Logic:**
    *   If the user only has the "Requester" role, display it as plain text (e.g., `Role: Requester`).
    *   If the user has multiple roles (e.g., "Requester" and "Tier 1 Approver"), display a Fluent UI `Dropdown` so they can switch contexts.
*   **Requester Dashboard:**
    *   When viewing as "Requester", render a Fluent UI `Pivot` (Tabs) component with two tabs: 
        1.  **Pending:** Filters Dataverse for `sct_requester = CurrentUser` AND Status in `['Draft', 'Submitted', 'In Review', 'Returned']`.
        2.  **Completed:** Filters for Status in `['Approved', 'Rejected']`.

## Phase 2: The Wizard UI (Multi-Step Form)
**Objective:** Transform the long, overwhelming form into a clean, step-by-step wizard.

*   **Component Structure:** Build a new `RequestWizard.tsx` component using state to track the `currentStep` (integer).
*   **Steps Breakdown (Example):**
    *   *Step 1: General Info* (Business Unit, Site, Purchase Type)
    *   *Step 2: Financials* (Estimated Value, Cost Center, Pricing Basis)
    *   *Step 3: Supplier & Compliance* (Supplier Lookup, Flags like Tech/Data)
    *   *Step 4: Review & Submit* (Read-only summary of all data)
*   **Validation:** The "Next" button on each step will be disabled until all required fields for that specific step are filled correctly.

## Phase 3: SPFx-Driven Approval Engine (Bypassing Power Automate)
**Objective:** Handle all approval routing and email notifications directly via React/TypeScript and MS Graph, eliminating the need for Power Automate flows.

*   **Routing Logic in Code:**
    *   When a user clicks "Submit", the SPFx code sets the status to "In Review - Manager".
    *   The code queries MS Graph to find the requester's Manager.
    *   The code uses MS Graph `/sendMail` endpoint to instantly send an HTML email to the Manager saying "A request requires your approval" with a link to the SPFx web part.
*   **Tier Progression:**
    *   When an approver approves, the code checks the amount/rules, updates the request status to the next tier (e.g., "In Review - Tier 1"), fetches the Tier 1 group members via Graph, and sends them an email.

## Phase 4: Approver Dashboard & "Take Action" UI
**Objective:** Give approvers a dedicated, streamlined view to process requests.

*   **Approver Data Fetch:** When the user switches to an Approver role (e.g., Tier 1), the grid fetches only requests currently stuck at their specific stage.
*   **Grid Action Column Updates:**
    *   **Remove:** The "Edit" (pencil) icon is completely removed for approvers.
    *   **Keep:** The "View" (eye) and "History" (clock) icons remain.
    *   **Add "Take Action":** A prominent primary button or icon in a new column.
*   **The Action Panel:**
    *   Clicking "Take Action" opens a Fluent UI `Dialog` or `Panel`.
    *   It contains a read-only summary of the request, a mandatory "Comments" multi-line text box, and three large action buttons: **[ Approve ]**, **[ Reject ]**, **[ Return for Info ]**.
    *   Clicking one of these buttons instantly posts the decision to the `sct_approvallog` table in Dataverse, updates the main request status, and triggers the next code-based email notification.