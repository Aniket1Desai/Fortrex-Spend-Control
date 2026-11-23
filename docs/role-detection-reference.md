# Role Detection Reference — Fortrex SCT

## How roles are assigned at login

When the web part loads, `AppShell._loadUserRoles()` runs and builds the dropdown
options the user sees. It follows a three-step priority chain:

```
1. Hardcoded map (instant, no network)
        ↓ (only if user NOT in map)
2. Dataverse systemuser table  → Manager role
   Azure AD group membership   → Tier Approver roles   (run in parallel)
        ↓
3. Graceful fallback: show only "Requester" if everything fails
```

---

## Step 1 — Hardcoded map (`USER_ROLE_MAP` / `MANAGER_AADIDS`)

Defined at the top of `AppShell.tsx`.

| Constant | Purpose |
|---|---|
| `USER_ROLE_MAP` | Maps an AAD Object ID → array of tier role keys |
| `MANAGER_AADIDS` | Set of AAD Object IDs that are always managers |

The AAD OID comes from `pageContext.aadInfo.userId` — no network call needed.
If the OID is found in either map, the function returns immediately and
**Graph / Dataverse calls are skipped entirely**.

---

## Step 2a — Manager via Dataverse `systemuser` table

**Source of truth for the Manager role in this app.**

### Why Dataverse (not Graph `/me/manager`)?

Requests are routed to managers via `_cr2a4_manager_value` on `sct_request`,
which is a lookup to the Dataverse `systemuser` table.
The Dataverse `parentsystemuserid` hierarchy is therefore more reliable than
the AAD org chart for this app.

### How it works

```
SPFx pageContext.aadInfo.userId (AAD OID)
        ↓
DataverseService.getCurrentSystemUser(aadOid)
  → GET systemusers?$filter=azureactivedirectoryobjectid eq '{aadOid}'
  → returns: { systemuserid, fullname, internalemailaddress, _parentsystemuserid_value }
        ↓
DataverseService.hasDirectReportsInDataverse(systemuserid)
  → GET systemusers?$filter=_parentsystemuserid_value eq {systemuserid}&$top=1
  → if 1+ rows returned → isManager = true → "Manager" added to dropdown
```

### Relevant Dataverse columns on `systemuser`

| Column | Type | Meaning |
|---|---|---|
| `systemuserid` | GUID | Primary key — used in all lookup bindings |
| `azureactivedirectoryobjectid` | GUID | Maps Dataverse user ↔ AAD user |
| `internalemailaddress` | string | Used as fallback lookup key |
| `_parentsystemuserid_value` | GUID | The GUID of this user's manager in Dataverse |

A user is treated as a **Manager** if at least one other `systemuser` record
has `_parentsystemuserid_value` equal to that user's `systemuserid`.

---

## Step 2b — Tier Approver roles via Azure AD groups

Three AD security group GUIDs are configured in `AppShell.tsx`:

| Constant | Maps to role |
|---|---|
| `TIER1_GROUP_ID` | `approver_tier1` — "Approver — Tier 1" |
| `TIER2_GROUP_ID` | `approver_tier2` — "Approver — Tier 2" |
| `TIER3_GROUP_ID` | `approver_tier3` — "Approver — Tier 3" |

The check uses the Graph `checkMemberGroups` API:

```
POST /me/checkMemberGroups
Body: { groupIds: [TIER1_GROUP_ID, TIER2_GROUP_ID, TIER3_GROUP_ID] }
→ returns the subset the current user actually belongs to
```

This is a single Graph call regardless of how many groups are checked.
If the user is in `TIER1_GROUP_ID`, the "Approver — Tier 1" option is added, etc.

---

## Role keys used throughout the app

| Dropdown text | `activeRole` key | Who sees it |
|---|---|---|
| Requester | `requester` | Every user (always present) |
| Manager | `manager` | Users with direct reports in Dataverse |
| Approver — Tier 1 | `approver_tier1` | Members of `TIER1_GROUP_ID` in AAD |
| Approver — Tier 2 | `approver_tier2` | Members of `TIER2_GROUP_ID` in AAD |
| Approver — Tier 3 | `approver_tier3` | Members of `TIER3_GROUP_ID` in AAD |

---

## What to change when onboarding a new manager or approver

| Scenario | Action |
|---|---|
| New manager (org chart set up in Dataverse) | Nothing — `hasDirectReportsInDataverse` detects it automatically |
| New manager (org chart NOT yet in Dataverse) | Add their AAD OID to `MANAGER_AADIDS` in `AppShell.tsx` as a temporary override |
| New Tier approver | Add the user to the appropriate AAD security group (`TIER1_GROUP_ID` / `TIER2_GROUP_ID` / `TIER3_GROUP_ID`) |
| Override any role manually | Add entry to `USER_ROLE_MAP` in `AppShell.tsx` |
