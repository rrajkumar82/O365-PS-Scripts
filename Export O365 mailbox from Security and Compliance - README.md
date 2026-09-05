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
  - For app-only automation: the Azure AD app (service principal) must be granted the necessary Purview/eDiscovery roles (assign using Purview portal / PowerShell as required).

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
.\