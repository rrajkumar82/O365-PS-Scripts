# Purview Role Mapping — short guide

This short README explains how to map an Entra ID group (that contains your automation service principal or user accounts) into Microsoft Purview (Security & Compliance) role groups so the group members can create and run content searches and exports.

Use this file alongside `Export O365 mailbox from Security and Compliance.ps1`.

---

## Goal
Give a service principal (app) or a set of users the Purview eDiscovery permissions they need by adding them to an Entra ID security group and mapping that group into the appropriate Purview role group (e.g., eDiscovery Manager).

---

## High-level steps
1. Create or identify an Entra ID security group.
2. Add the service principal (or user accounts) into the group.
3. In the Microsoft Purview compliance portal, add the Entra ID group to the appropriate role group (e.g., eDiscovery Manager).
4. Verify the role mapping and test by running a small content search/export.

---

## Step-by-step (detailed)

1) Create an Entra ID security group
- Portal: https://entra.microsoft.com
- Entra ID -> Groups -> New group
  - Group type: Security
  - Group name: e.g., `eDiscovery Managers`
  - Membership type: Assigned (recommended)
  - Create the group

2) Add the service principal or users to the group
- Option A: Use the helper script in this branch to add the service principal to the group interactively (`Assign-EntraServicePrincipalToGroup.ps1`).
- Option B (portal): Entra ID -> Groups -> select the group -> Members -> Add members -> search for the user(s) or service principal -> Add.

Notes for service principals
- In the portal, service principals are listed under "Enterprise applications" or may be searchable in Groups -> Add members. If not visible, use the Microsoft Graph helper script or Graph PowerShell (interactive) to add the SP as a member.

3) Map the Entra group into a Purview role group
- Portal: https://compliance.microsoft.com (Microsoft Purview compliance portal)
- Left nav -> Permissions (or Roles)
- Choose the role group you want to modify (e.g., `eDiscovery Manager`) and click it
- Click **Edit role group** (or Manage role group) -> Add -> select **Group** and search for the Entra ID group you created above -> Add -> Save

Important
- Some tenants may require adding individual users rather than groups; check your tenant policies and Purview configuration.
- The UI may change — look for the Permissions / Roles area in the Purview portal if labels differ.

4) Verify access
- After assigning the group to the Purview role, check the role group members include the Entra ID group.
- Test by running a small compliance search and creating an export using the account or app in the group. If using app-only auth, ensure the app has the appropriate application permissions and is included in the group.

---

## Common troubleshooting
- Group not found in Purview role group UI: ensure the group is a Security group (not Microsoft 365 group) and is in the same tenant.
- Service principal not visible when adding members: use the helper script or Graph PowerShell to add the SP to the group.
- Exports failing with permission errors: confirm the account or app has the Purview role and that Exchange/Microsoft Graph APIs are accessible with the chosen auth method.

---

## Security notes
- Use least privilege: assign only the role group needed (e.g., eDiscovery Manager) rather than higher-privilege roles like Compliance Administrator unless required.
- Audit role assignments and monitor export activity in Purview audit logs.

---

## Next steps
If you want, I can also add a short PowerShell or Microsoft Graph snippet in this file showing how to verify the group membership and the Purview role group members via API — say the word and I will add it.
