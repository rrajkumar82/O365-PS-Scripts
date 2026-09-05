# Assign-EntraServicePrincipalToGroup — README

This helper script (Assign-EntraServicePrincipalToGroup.ps1) is a small interactive Microsoft Graph / Entra ID utility that adds a service principal (app) to an Entra ID group. This is useful when you want to grant a service principal the same access that is managed by a group (for example, adding the SP to a group that your compliance admins map to a Purview eDiscovery role).

This README provides a quick summary, prerequisites, usage, sample outputs, and common errors/troubleshooting tips.

---

## What it does
- Prompts an administrator to sign in to Microsoft Graph (Entra) using delegated (interactive) auth.
- Looks up a service principal by AppId (application/client id).
- Looks up a group by display name.
- Adds the service principal as a member of the group using the Microsoft Graph API.
- Exits with distinct non-zero codes for common failure cases.

---

## Prerequisites
- PowerShell 7.x or Windows PowerShell 5.1.
- The Microsoft.Graph PowerShell module will be installed automatically by the script if missing (CurrentUser scope).
- The signed-in account must have privileges to read service principals and modify group membership.
- The group should already exist in Entra ID (create it in the portal if needed).

---

## Usage

Interactive run from an elevated PowerShell session (example):

```powershell
PowerShell -NoProfile -ExecutionPolicy Bypass -File .\Assign-EntraServicePrincipalToGroup.ps1 -AppId '<your-app-id>' -GroupName 'eDiscovery Managers'
```

Optional parameter to request additional delegated scopes (default: Application.Read.All Group.ReadWrite.All):

```powershell
PowerShell -File .\Assign-EntraServicePrincipalToGroup.ps1 -AppId '<app-id>' -GroupName 'eDiscovery Managers' -Scopes 'Application.Read.All Group.ReadWrite.All User.Read.All'
```

---

## Sample outputs

Successful add (console):

[INFO] Connecting to Microsoft Graph with delegated scopes: Application.Read.All Group.ReadWrite.All
[INFO] Signed in as admin@contoso.com
[INFO] Locating service principal for AppId: 11111111-2222-3333-4444-555555555555
[INFO] Found service principal: Id=44444444-aaaa-bbbb-cccc-777777777777 DisplayName=MyApp
[INFO] Locating group with display name: eDiscovery Managers
[INFO] Found group: Id=99999999-dddd-eeee-ffff-888888888888 DisplayName=eDiscovery Managers
[INFO] Checking whether the service principal is already a member of the group...
[INFO] Adding service principal to group...
[INFO] Successfully added service principal (Id: 44444444-aaaa-bbbb-cccc-777777777777) to group 'eDiscovery Managers' (Id: 99999999-dddd-eeee-ffff-888888888888).

If the service principal is already a member:

[INFO] Service principal is already a member of group 'eDiscovery Managers' (Id: 99999999-dddd-eeee-ffff-888888888888). No action taken.

---

## Exit codes
- 0 — Success (member added or already present).
- 1 — Unexpected error.
- 2 — Service principal with supplied AppId not found.
- 3 — Target group not found.
- 4 — Failed to add service principal to group (API error). Check details in console output.

---

## Common errors & troubleshooting

1) "Service principal with AppId ... not found"
- Cause: The AppId is incorrect, the app registration doesn’t exist in the tenant, or the signed-in user lacks permission to read service principals.
- Fix: Verify the AppId in Entra portal → App registrations. Ensure the signed-in admin has Directory.Read.All or appropriate view permissions.

2) "Group '...' not found"
- Cause: Typo in the group display name or the group lives in a different tenant.
- Fix: Confirm the group display name exactly (including spacing/case) in the Entra portal → Groups.

3) "Insufficient privileges" or Graph permission errors during Connect-MgGraph
- Cause: The signed-in user did not consent to requested delegated scopes and cannot perform operations.
- Fix: Sign in with a privileged account (Global Admin or Privileged Role Administrator) or request admin consent for the scopes. Alternatively, run an admin-consent flow to grant the tenant the required delegated permissions.

4) "Failed to add service principal to group" (API 400/409 responses)
- Cause: Duplicate membership, group type restrictions, or policy preventing adding certain principal types.
- Fix: Inspect the full error message printed by the script. Some groups (e.g., some directory role groups) cannot accept service principals as members. Use a security group or mail-enabled security group suitable for role mapping.

5) Microsoft.Graph module installation fails
- Cause: Network or policy restrictions on the runner/machine.
- Fix: Install the Microsoft.Graph module manually as an elevated user or pre-install it on the runner. Example:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
```

---

## Security notes
- This script uses delegated interactive auth (Connect-MgGraph). Do not run it unattended with stored credentials.
- For automation, use an app-only (certificate) flow and follow least-privilege principles.
- After adding the service principal to a group, assign that group to the desired Purview/eDiscovery role in the Purview portal (the script does not directly modify Purview role mappings).

---

If you want, I can also:
- Add an app-only (certificate) variant of this helper as a separate script, or
- Add an automated test harness that validates the lookup and membership flows against a test tenant.

Tell me which you prefer and I’ll add it to the branch.