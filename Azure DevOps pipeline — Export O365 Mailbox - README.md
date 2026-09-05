Azure DevOps pipeline — Export O365 Mailbox
This README describes the Azure DevOps pipeline example (azure-pipelines.yml) included in the repository. It explains required service connections, variables, and runner considerations.

Purpose

Retrieve a PFX certificate from Azure Key Vault and import it into the agent's cert store.
Run the main export script (Export O365 mailbox from Security and Compliance.ps1) in app-only mode using the imported certificate.
Prerequisites

An Azure service connection (service principal) with access to the Key Vault used by the pipeline.
A Key Vault secret that stores the PFX (base64 or raw). The pipeline expects a secret (example name: my-pfx-secret) and a pipeline variable for the PFX password (PfxPassword).
Pipeline variables/secrets: AppId, TenantId, PfxPassword (and Key Vault secret name mapping).
Self-hosted Windows agent (recommended) with Unified Export Tool (UET) installed OR a custom image that has UET installed. The example uses windows-latest but UET must be available.
Usage

Configure an Azure DevOps service connection with permission to read Key Vault secrets.
Add pipeline variables/secret variables for AppId, TenantId, and PfxPassword.
Run the pipeline manually or schedule it using pipeline triggers.
Security notes

Use Key Vault to store the PFX and retrieve it at runtime rather than storing raw PFX content in pipeline variables.
Limit access to the service connection and Key Vault to only the users and service principals that need it.
Cleanup recommendations

Remove the certificate from the LocalMachine store at the end of the job if the agent is shared.
Troubleshooting

Key Vault secret not found: ensure service connection has Get/List permissions for secrets in the Key Vault.
Certificate import errors: ensure the agent account has permissions to import into Cert:\LocalMachine:\My, or run the agent as a user with required permissions.
Export failures: check that the enrolled Entra ID app has been mapped to the Purview role and that the agent can reach the export storage endpoint.
Customization

Parameterize mailbox address and output folder via pipeline variables so you can run the same pipeline for different mailboxes.
Consider using a dedicated hybrid worker or ephemeral agent to reduce long-lived secrets exposure. '@
