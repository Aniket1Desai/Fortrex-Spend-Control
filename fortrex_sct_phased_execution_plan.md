# Fortrex Spend Control Tracker (SCT) — Phased Execution Guide

This document outlines the exact execution sequence for building the Fortrex SCT SPFx web part connected to Microsoft Dataverse (`https://org1bec48f0.crm8.dynamics.com/`). 

---

## Where and How to Run These Prompts

1. **Where to Run:** 
   - Open your project workspace in **Antigravity IDE** or your **Claude** coding session.
   - Ensure you have already run the initial setup prompt provided in the previous step (which configured `package-solution.json` for Dataverse API access and established the foundational TypeScript models).
2. **Execution Rule:** 
   - **Go one phase at a time.** Do not paste all phase prompts at once. 
   - Copy the prompt for the current phase, paste it into your AI assistant, review the code generated, test it (`gulp build`), and only then proceed to the next phase.

---

## Phase Prompts (Execute Sequential Order)

### Phase 2: App Shell & Responsive Layout
```markdown
You are the senior SPFx & React engineer. Implement Phase 2: App Shell.
Build a responsive React UI layout shell for the Fortrex SCT web part.
- Add navigation tabs: "Dashboard", "New Request", "My Requests", and "Review Queue".
- Use accessible labels, keyboard focus, and reusable loading, empty, and error state components.
- Clearly label navigation placeholders as UI shells (no live Dataverse CRUD yet in this phase).
- Run `gulp build` and verify type-checks pass cleanly. Report changes.
```

### Phase 3: Dataverse Data Access Layer (`DataverseService.ts`)
```markdown
You are the senior SPFx & React engineer. Implement Phase 3: Data Access Layer.
Create a typed Dataverse service (`DataverseService.ts`) using SPFx's `AadHttpClient` targeting `https://org1bec48f0.crm8.dynamics.com/api/data/v9.2/`.
- Implement Create, Read, and Update methods for `sct_sites_masters`, `sct_roles`, `sct_suppliers`, and `sct_requests`.
- Handle OData formatting, `@odata.bind` syntax for lookups (e.g., `sct_Site@odata.bind`), and Choice formatted values headers.
- Implement robust error handling for 401/403, throttling, and network failures.
- Run `gulp build` and verify type-checks pass cleanly. Report changes.
```

### Phase 4: Read-Only Dashboard & List View
```markdown
You are the senior SPFx & React engineer. Implement Phase 4: Read-only Dashboard.
Connect the Dashboard tab to the `DataverseService` to load live records from `sct_requests`.
- Display key columns: SCT Reference (Autonumber), Request Status, Business Unit, Purchase Type, Estimated Total Value, and Created On date.
- Implement loading spinners, empty states, error handling, and retry buttons.
- Link table rows to a read-only details preview panel.
- Run `gulp build` and verify type-checks pass cleanly. Report changes.
```

### Phase 5: Create Request Form (Schedule C & Conditional Rules)
```markdown
You are the senior SPFx & React engineer. Implement Phase 5: Create Request Form.
Build the typed Create Request form bound to `sct_requests` schema.
- Load active sites from `sct_sites_masters` and eligible approved suppliers from `sct_suppliers`.
- Enforce the New vs. Approved Supplier toggle (`sct_is_new_supplier`): Yes shows proposed supplier text fields; No shows existing supplier lookup.
- Prevent requesters from editing authoritative category, cost centre, or GL codes.
- Implement conditional UI triggers: Temp labor union check, Repair/Rental asset tag, Emergency declaration reason, and Customer recharge fields.
- Run `gulp build` and verify type-checks pass cleanly. Report changes.
```

### Phase 6: Save, Submit, Autonumber & Attachments
```markdown
You are the senior SPFx & React engineer. Implement Phase 6: Save, Submit & Attachments.
Implement record submission and file attachment handling.
- Allow saving as Draft or submitting to trigger the Dataverse native autonumber `sct_sct_reference` (`SCT-2026-0001`).
- Prevent duplicate submissions on double-clicks or retries.
- Implement file attachments upload to Dataverse Notes (`annotation`) linked to the created request record for mandatory approval evidence and quotes.
- Run `gulp build` and verify type-checks pass cleanly. Report changes.
```

### Phase 7: Stage Workbenches (Triage, Sourcing, FP&A, CFO, PO)
```markdown
You are the senior SPFx & React engineer. Implement Phase 7: Stage Workbenches.
Build role-specific action panels for the 6 Spend Control Tower lifecycle stages:
- Stage 3 (SCT Desk): Triage queue, category/cost centre assignment, and Return for Information action.
- Stage 4 (Procurement): Sourcing quotes collection and D365 PO Number entry.
- Stage 5 (FP&A & CFO): Post-sourcing review and L4 CFO approval triggers for requests >$10k or new suppliers.
- Stage 6 (Receiving): Receipt and completion confirmation.
- Run `gulp build` and verify type-checks pass cleanly. Report changes.
```

### Phase 8 & 9: Approval Routing & Matrix Engine (`ApprovalRoutingService.ts`)
```markdown
You are the senior SPFx & React engineer. Implement Phase 8 & 9: Approval Routing Engine.
Build `ApprovalRoutingService.ts` driven by live records from `sct_roles` and `sct_sites_masters`.
- Compute required L1, L2, L3, and L4 approvers dynamically based on estimated value and new supplier status.
- Implement self-approval escalation rules (escalate to next level up if requester is approver).
- Support active delegate lookup based on effective start/end dates.
- Run `gulp build` and verify type-checks pass cleanly. Report changes.
```

### Phase 10 & 11: Security Hardening, Testing & Packaging (`.sppkg`)
```markdown
You are the senior SPFx & React engineer. Implement Phase 10 & 11: Hardening & Packaging.
Review security boundaries, Dataverse column-level profile permissions, and error handling.
- Verify restricted confidentiality visibility for legal/HR requests.
- Run final type checks, linter, and `gulp bundle --ship && gulp package-solution --ship`.
- Output the verified artifact path for the generated `.sppkg` package file.

```