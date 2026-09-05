# Export O365 mailbox from Security & Compliance — Usage

This branch contains a modernized PowerShell script to export a single Office 365 mailbox from Microsoft Purview (Security & Compliance) to local storage using the Unified Export Tool (UET), and optionally copy the results to a NAS folder.

This README documents prerequisites, parameters, examples, and operational notes for the script: `Export O365 mailbox from Security and Compliance.ps1`.

---

## Key features
- Uses the ExchangeOnlineManagement module (recommended v3.9.0+).
- Supports interactive MFA sign-in or app-only (certificate) authentication for automation.
- Connects to Purview with `Connect-IPPSSession -EnableSearchOnlySession` per modern guidance.
- Robust parsing of export details, safer logging (avoids printing full SAS tokens), improved error handling.

---

## Prerequisites
- PowerShell 7.0+ or Windows PowerShell 5.1 (script uses common core/Windows cmdlets).
- ExchangeOnlineManagement module (install / update):

```powershell
Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser
```

- The Microsoft Unified Export Tool (UET) must be installed on the machine running the script. The script looks for the UET under `%LOCALAPPDATA%\Apps\2.0`.

- Appropriate permissions:
  - For interactive runs: the signed-in admin account must have the required eDiscovery / Compliance roles (e.g., eDiscovery Manager, Compliance Admin) to create and export content searches.
  - For app-only automation: the Entra ID app (service principal) must be granted the necessary Purview/eDiscovery roles (assign using Purview portal / PowerShell as required).

- Network access to download from the export container (SAS URL) and to copy to the NAS destination (if used).

---

## File(s) added/changed
- `Export O365 mailbox from Security and Compliance.ps1` — the modernized script (committed to branch `modernize/export-compliance-script`).
- `Export O365 mailbox from Security and Compliance - README.md` — this usage document (this file).

---

## Script parameters
- `-Mailbox` (string) — email address of the mailbox to export (required).
- `-UserFullname` (string) — used to build a safe search/export name (required).
- `-LocalExportLocation` (string) — local path to save the download (required). The script will create a subfolder using `MailboxUserCountry`.
- `-MailboxUserCountry` (string) — subfolder name (required).
- `-NASDrive` (string) — optional, default `D:`. Target NAS drive letter/path for copying results.
- `-UseAppOnly` (switch) — enable app-only authentication mode (automation).
- `-AppId` (string) — app registration client (application) id (required with `-UseAppOnly`).
- `-CertificateThumbprint` (string) — certificate thumbprint used by the app (required with `-UseAppOnly`).
- `-TenantId` (string) — optional tenant id (Organization parameter to Connect-ExchangeOnline).
- `-EnableTranscript` (switch) — enable Start-Transcript logging to `%PUBLIC%\Documents\O365_PSscriptExecution`.
- `-PollIntervalSeconds` (int) — polling interval for status checks (default 10).

---

## Example usage

Interactive (manual MFA sign-in):

```powershell
.\"Export O365 mailbox from Security and Compliance.ps1" -Mailbox 'user@contoso.com' -UserFullname 'User Name' -LocalExportLocation 'C:\PSTExports' -MailboxUserCountry 'Australia' -EnableTranscript
```

App-only (automation):

```powershell
.\"Export O365 mailbox from Security and Compliance.ps1" -Mailbox 'user@contoso.com' -UserFullname 'User Name' -LocalExportLocation 'C:\PSTExports' -MailboxUserCountry 'Australia' -UseAppOnly -AppId '<app-id>' -CertificateThumbprint '<thumbprint>' -TenantId '<tenant-id>' -EnableTranscript
```

---

## Troubleshooting & Permissions

Common issues and how to resolve them:

- Permission denied / insufficient privileges when creating a compliance search or export:
  - Ensure the account (or service principal) has the eDiscovery permissions. For interactive sign-in give the user one of the following roles in Purview: `eDiscovery Manager`, `Compliance Administrator`, or a custom role with `Content Search` and `Export` privileges.
  - For app-only auth, register an Entra ID app and grant it the minimum application-level permissions required for Exchange and Purview. Microsoft also requires assigning the app to an RBAC role in the Purview portal or using PowerShell to assign the necessary permissions.

- `Connect-IPPSSession` errors or cmdlets not found:
  - Confirm you are running ExchangeOnlineManagement v3.9.0+ and call `Connect-IPPSSession -EnableSearchOnlySession` as shown in the script. Older module versions or missing switches will cause failures.

- UET not found / cannot start Unified Export Tool:
  - Make sure the Microsoft 365 compliance export tool (UET) is installed on the host and the executable is present under `%LOCALAPPDATA%\Apps\2.0` (or update the script to point to the installed location).

- Export container URL or SAS token not present yet:
  - The script waits and polls for the export action to be ready. If it never appears, check the compliance search status in the Purview portal and verify your service principal/user has export rights.

- Network / firewall issues while downloading export with UET:
  - Ensure the host can reach the export Container URL (storage endpoint) and outbound HTTPS is allowed.

Example: assign eDiscovery (manual steps)

1. Create or identify an Entra ID app registration for automation and upload a certificate (or use client secret — certificate is recommended).
2. Give the app the minimum Entra ID permissions if required for Exchange (app-only Exchange access uses application permissions set by Microsoft guidance).
3. In the Purview portal (or via PowerShell), add the service principal or an Entra ID group to an appropriate eDiscovery role group (e.g., `eDiscovery Manager`) so it can create searches and export content.

Note: Role assignment in Purview is required in addition to Entra ID application permissions. Follow Microsoft Purview documentation for granting service principals access to eDiscovery features.

---

## CI / Scheduled-run guidance (automation best practices)

When you run this script from a CI pipeline (Azure DevOps, GitHub Actions) or as a scheduled automation (Azure Automation, runbook, or VM scheduled task), prefer app-only authentication or managed identity flows and avoid interactive sign-in.

Options:

1. App-only with certificate (recommended for on-prem/VM or self-hosted runners)
   - Store the certificate securely (local Cert store or Key Vault) and provide the thumbprint to the script.
   - The Entra ID app must be granted the required Purview roles as noted above.

2. Managed Identity (recommended for Azure-hosted automation)
   - If executing in Azure (VM, Function, WebJob, Automation account with managed identity), prefer a system-assigned or user-assigned managed identity.
   - Grant the managed identity the same Purview/eDiscovery role in Purview (add the principal to the role group).
   - Retrieve tokens / connect using supported SDKs or use the certificate-backed approach if ExchangeOnlineManagement supports it in your scenario.

3. Secrets & Certificate storage (KeyVault)
   - Store certificate private keys and any client secrets in Azure Key Vault.
   - Grant the pipeline or managed identity access to Key Vault via RBAC or access policies.
   - In a pipeline or runbook, retrieve the certificate at runtime and use it to call `Connect-ExchangeOnline -AppId <id> -CertificateThumbprint <thumbprint>` (import to the local cert store first if required).

Example: GitHub Actions (high level)
- Use a self-hosted runner (Windows) that has the UET installed and can reach Microsoft endpoints + NAS.
- Use GitHub Secrets to store AppId, TenantId and KeyVault references. Use Azure/login and Azure/keyvault actions to fetch certificates into the runner at job time.
- Run PowerShell step to execute the script using the retrieved certificate thumbprint.

Example: Azure Automation / Runbook (high level)
- Use a hybrid worker or Azure VM with the UET installed.
- Use a managed identity or store certificate in Key Vault and retrieve at runtime.
- Schedule the runbook with appropriate monitoring/alerting on failures.

Security notes for automation
- Prefer certificate-based app-only auth over client secrets.
- Limit eDiscovery roles and scope to the minimum required.
- Rotate certificates regularly and monitor audit logs for export usage.

---

## Operational tips
- Avoid running multiple exports with identical search names — the script builds a search name from `UserFullname`; ensure uniqueness if running multiple exports in parallel.
- Keep `PollIntervalSeconds` conservative (10s default) to avoid high API call rates.
- The script truncates SAS tokens in logs — treat those tokens as sensitive and avoid shipping logs to public places.

---

If you want, I can:
- Add a sample PowerShell script/snippet that shows how to assign the Purview eDiscovery role to a service principal using REST or PowerShell (subject to tenant permissions), or
- Draft a GitHub Actions workflow that demonstrates fetching a certificate from Key Vault and running this script on a self-hosted Windows runner.

Tell me which of those you'd like next and I will add it to the branch.