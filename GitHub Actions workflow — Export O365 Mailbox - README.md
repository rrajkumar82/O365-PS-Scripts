GitHub Actions workflow — Export O365 Mailbox
This README describes the GitHub Actions workflow located at .github/workflows/export-mailbox.yml. Use this as a companion guide for configuring and running the workflow on a self-hosted Windows runner.

Purpose

Retrieve a certificate (PFX) from a secure source and import it into the runner's certificate store.
Run the main export script (Export O365 mailbox from Security and Compliance.ps1) in app-only (certificate) mode to create and download a Purview export.
Prerequisites

Self-hosted Windows runner with:
PowerShell (pwsh) available
Unified Export Tool (UET) installed and accessible
Network access to Microsoft storage endpoints and your NAS
GitHub repository secrets configured (or Key Vault action configured):
AZURE_CREDENTIALS (if using azure/login)
PFX_BASE64 (base64-encoded PFX file) OR configure the Key Vault action to pull secrets
PFX_SECRET_PASSWORD
APP_ID
TENANT_ID
The Entra ID app (APP_ID) has been created and the certificate uploaded to it, and the app/group has been mapped to Purview roles as described in the repo docs.
Usage

Trigger manually from the Actions tab (workflow_dispatch) or allow scheduled runs (cron configured in the workflow).
Edit the workflow to parameterize mailbox, output path, and other inputs (currently the job uses inline sample values).
Security notes

Prefer retrieving PFX from Azure Key Vault instead of storing PFX_BASE64 directly in repository secrets. Use the azure/keyvault-secrets action to retrieve secrets at runtime.
Do not log certificate private material. The workflow only stores PFX temporarily on the runner; ensure the runner is secured and transient if possible.
Cleanup recommendations

Remove the imported certificate from the certificate store at the end of the job if the runner is shared. Example cleanup commands (pwsh):

$thumb = 'THUMBPRINT' Get-ChildItem Cert:\LocalMachine\My | Where-Object Thumbprint -EQ $thumb | Remove-Item -Force Remove-Item .\agentcert.pfx -Force

Troubleshooting

If the runner cannot import the PFX: verify the PFX password and the runner has permission to modify Cert:\LocalMachine\My.
If UET fails to download: check network/firewall logs and ensure the runner can reach the storage endpoint in the export container.
If export commands fail with permission errors: confirm the Entra ID app and group are assigned the Purview eDiscovery role.
Notes for customization

Parameterize mailbox and output location by using repository/organization-level secrets or workflow inputs.
Consider using a dedicated service account or transient self-hosted runner for better security. '@
$az = @
