# Fortrex SCT — Spend Control Tracker

[![SPFx Version](https://img.shields.io/badge/SPFx-1.23.2-0078D4.svg)](https://aka.ms/spfx)
[![Node.js Version](https://img.shields.io/badge/Node.js-22.x-339933.svg)](https://nodejs.org/)
[![React Version](https://img.shields.io/badge/React-17.0.1-61DAFB.svg)](https://reactjs.org/)
[![Fluent UI](https://img.shields.io/badge/Fluent%20UI-8.x-005A9E.svg)](https://developer.microsoft.com/fluentui)
[![Dataverse](https://img.shields.io/badge/Microsoft-Dataverse-742774.svg)](https://powerapps.microsoft.com/)

A modern, enterprise-grade Procurement & Budget Governance platform built as a **SharePoint Framework (SPFx)** client-side web part integrated with **Microsoft Dataverse** and **Microsoft Graph API**.

---

## 📑 Table of Contents

- [Overview & Architecture](#overview--architecture)
- [Key Features](#key-features)
- [Multi-Tier Approval Workflow](#multi-tier-approval-workflow)
- [Dataverse Data Model & Choice Sets](#dataverse-data-model--choice-sets)
- [Prerequisites](#prerequisites)
- [Installation & Local Setup](#installation--local-setup)
- [Building & Packaging](#building--packaging)
- [Deployment Guide](#deployment-guide)
- [Configuration & Web Part Properties](#configuration--web-part-properties)
- [Contributing & Version History](#contributing--version-history)

---

## 🏛️ Overview & Architecture

Fortrex Spend Control Tracker (SCT) replaces fragmented manual procurement forms and spreadsheets with a centralized, auditable procurement lifecycle portal embedded in SharePoint Online and Microsoft Teams.

```
┌────────────────────────────────────────────────────────────────────────┐
│                        SharePoint Framework (SPFx)                     │
│                             AppShell Portal                            │
│  [ Dashboard ]     [ + New Request ]     [ My Requests ]   [ Review Queue ]│
└──────────────────┬─────────────────────────────────┬───────────────────┘
                   │                                 │
                   ▼                                 ▼
   ┌───────────────────────────────┐  ┌──────────────────────────────────┐
   │      Microsoft Dataverse      │  │        Microsoft Graph API       │
   │  - sct_requests               │  │  - User profile & /me/manager    │
   │  - sct_approvallogs           │  │  - Security Group membership     │
   │  - sct_auditlogs              │  │  - Automated HTML email alerts   │
   │  - systemusers, sites, vendors│  └──────────────────────────────────┘
   └───────────────────────────────┘
```

---

## ✨ Key Features

1. **Role-Based Workspaces & Dynamic Role Switcher**:
   - Seamlessly switch between roles: `Requester`, `Line Manager`, `Approver — Tier 1`, `Approver — Tier 2`, and `Approver — Tier 3`.
   - Real-time pending badge counters fetched dynamically from Dataverse.

2. **Originator Hierarchy Skip (Dynamic Routing)**:
   - Managers and Approvers can initiate requests directly from any view.
   - When a Manager or Approver raises a request, redundant prior approval steps are automatically bypassed, routing directly to the appropriate downstream tier.

3. **Strict Segregation of Duties (SoD) & Solo Test Mode**:
   - **Strict SoD Mode**: Enforces enterprise compliance—approvers are blocked from self-approving their own purchase requests.
   - **Solo Test Mode**: A dedicated toggle in the header enabling single-developer testing through all lifecycle stages without needing multi-user tenant accounts.

4. **In-Progress Form Protection**:
   - Navigation away from in-progress request forms (via tabs or role switcher) triggers a confirmation modal dialog (`Leave & Discard` vs. `Stay & Keep Editing`).
   - Browser unload guard prevents accidental page refresh data loss.

5. **Live Microsoft Graph Email Dispatch**:
   - Generates beautifully styled HTML approval notifications.
   - Sends alerts directly to Entra ID Security Groups (`Tier 1`, `Tier 2`, `Tier 3`) with automated CCs to the requester.

6. **Excel / CSV Export Engine**:
   - Exports all table records with clean, human-readable choice labels (e.g. `PSSI`, `Repair`, `Standard`) instead of raw Dataverse internal integers.

---

## 🔄 Multi-Tier Approval Workflow

```
   ┌─────────────────┐
   │ Standard User   │ ──► [ Line Manager ] ──► [ Tier 1: Triage/Sourcing ]
   │ (Requester)     │
   └─────────────────┘
   ┌─────────────────┐
   │ Department Mgr  │ ───────────────────────► [ Tier 1: Triage/Sourcing ]
   └─────────────────┘
   ┌─────────────────┐
   │ Tier 1 Approver │ ───────────────────────────────────────────────► [ Tier 2: FP&A Review ]
   └─────────────────┘
   ┌─────────────────┐
   │ Tier 2 Approver │ ────────────────────────────────────────────────────────────────────────► [ Tier 3: CFO/Executive ] ──► [ PO Issued ]
   └─────────────────┘
```

---

## 🗄️ Dataverse Data Model & Choice Sets

### Request Status (`sct_request_status`)
| Integer Value | Display Label | Purpose |
| :--- | :--- | :--- |
| `854000001` | **Draft** | Incomplete draft saved by user |
| `854000002` | **Submitted** | Pending review by Line Manager |
| `854000003` | **In Review** | In-flight across Tier 1, Tier 2, or Tier 3 |
| `854000004` | **PO Issued** | Fully approved; purchase order generated |
| `854000005` | **Rejected** | Declined by an approver |
| `854000006` | **Returned** | Returned to author for clarification |
| `854000007` | **Closed** | Order fulfilled or cancelled |

### Azure AD Security Groups
- **Tier 1 Approvers:** `fdbf07a0-548b-4ccf-875b-900cb1b24a1c` (`SCT-Fortrex_Approvers-Tier1`)
- **Tier 2 Approvers:** `6a8a4077-c78f-4af2-ae5e-12428d1dcbc3` (`SCT-Fortrex_Approvers-Tier2`)
- **Tier 3 Approvers:** `e63dbc4f-6ad2-4969-a518-b229836756af` (`SCT-Fortrex_Approvers-Tier3`)

---

## ⚙️ Prerequisites

- **Node.js**: `v22.x` (supported range: `>= 22.14.0 < 23.0.0`)
- **Package Manager**: `npm` (v10+)
- **Build Engine**: `@rushstack/heft`
- **SharePoint**: Microsoft 365 SharePoint Online App Catalog admin permissions
- **Dataverse**: Power Apps / Dynamics environment with imported `Fortrex SCT` solution

---

## 🚀 Installation & Local Setup

1. **Clone the repository**:
   ```bash
   git clone <YOUR_REPOSITORY_URL>
   cd FortrexDataverse
   ```

2. **Install project dependencies**:
   ```bash
   npm install
   ```

3. **Start local development server**:
   ```bash
   npm run start
   ```
   This compiles assets and opens the local SPFx development server.

4. **Test in SharePoint Hosted Workbench**:
   Navigate to:
   ```text
   https://<your-tenant>.sharepoint.com/sites/<your-site>/_layouts/15/workbench.aspx
   ```

---

## 📦 Building & Packaging

To build the production bundle and generate the `.sppkg` package file:

```bash
npm run build
```

This command automatically executes:
1. `heft build --clean --production`: Cleans previous outputs, transpiles TypeScript (`v5.8+`), compiles SCSS, and produces optimized Webpack bundles.
2. `heft package-solution --production`: Generates the production SharePoint package file at:
   ```text
   sharepoint/solution/ad-dataverse-fortrex.sppkg
   ```

---

## 🚢 Deployment Guide

1. Navigate to your Microsoft 365 **SharePoint Admin Center**.
2. Go to **Apps** $\rightarrow$ **Manage apps** (App Catalog).
3. Upload `sharepoint/solution/ad-dataverse-fortrex.sppkg`.
4. Check **"Enable this app and add it to all sites"** (or target specific site collections).
5. Click **Deploy**.
6. **API Approvals**:
   - Go to **SharePoint Admin Center** $\rightarrow$ **API access**.
   - Approve the pending API requests for:
     - `Dataverse (Dynamics CRM)` $\rightarrow$ `user_impersonation`
     - `Microsoft Graph` $\rightarrow$ `User.Read`, `GroupMember.Read.All`, `Mail.Send`

---

## 🔧 Web Part Properties

When editing the web part in SharePoint, configure the Property Pane:
- **Dataverse Environment URL**: e.g., `https://<org>.crm8.dynamics.com`
- **Title**: Header title displayed on the banner (default: `Fortrex Spend Control`)
- **Default Test Mode**: Enable or disable default Solo Test Mode for staging environments.

---

## 📄 License & Attribution

- **Developer**: Fortrex / Rishabh Software
- **Solution ID**: `69bca647-493e-4ee7-86b1-5bde01ad9a26`
- **Current Version**: `1.0.0.25`
