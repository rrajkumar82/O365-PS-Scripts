<#
.SYNOPSIS
Assign an Entra ID (Azure AD) service principal (app) to an Entra ID group using Microsoft Graph PowerShell (interactive).

.DESCRIPTION
This script locates a service principal by AppId and an AAD/Entra group by display name, and adds the service principal as a group member. Use this when you want to grant a service principal the same access your team manages through a group (for example, adding the SP to a group that is mapped to a Purview eDiscovery role).

It uses interactive Connect-MgGraph and requires a user with sufficient privileges to read service principals and modify group membership.

.PARAMETER AppId
The application (client) ID of the Entra ID app registration (service principal) to add to the group.

.PARAMETER GroupName
The display name of the Entra ID group to which the service principal will be added.

.PARAMETER Scopes
Optional: space-separated Microsoft Graph delegated scopes to request during interactive sign-in. Default: "Application.Read.All Group.ReadWrite.All".

.EXAMPLE
.
PowerShell -File .\Assign-EntraServicePrincipalToGroup.ps1 -AppId 'xxxx-xxxx-xxxx' -GroupName 'eDiscovery Managers'

#>

param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$AppId,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$GroupName,

    [Parameter(Mandatory=$false)]
    [string]$Scopes = 'Application.Read.All Group.ReadWrite.All'
)

Set-StrictMode -Version Latest

function Log { param($m) Write-Host "[INFO] $m" }
function Err { param($m) Write-Host "[ERROR] $m" -ForegroundColor Red }

try {
    # Ensure Microsoft Graph PowerShell module
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Log 'Microsoft.Graph module not found. Installing (CurrentUser scope)...'
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module Microsoft.Graph -Force

    # Connect interactively to Microsoft Graph (Entra ID)
    Log "Connecting to Microsoft Graph with delegated scopes: $Scopes"
    $scopeArray = $Scopes -split ' '\n    Connect-MgGraph -Scopes $scopeArray -ErrorAction Stop

    # Confirm sign-in
    $who = Get-MgUserMe -ErrorAction SilentlyContinue
    if ($who) { Log "Signed in as $($who.UserPrincipalName)" } else { Log "Signed in (no user info available)." }

    # Find the service principal
    Log "Locating service principal for AppId: $AppId"
    $sp = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue
    if (-not $sp) {
        Err "Service principal with AppId $AppId not found. Ensure the app exists and you have permission to read service principals."
        exit 2
    }
    Log "Found service principal: Id=$($sp.Id) DisplayName=$($sp.DisplayName)"

    # Find the group
    Log "Locating group with display name: $GroupName"
    $group = Get-MgGroup -Filter "displayName eq '$GroupName'" -ErrorAction SilentlyContinue
    if (-not $group) {
        Err "Group '$GroupName' not found. Ensure the group exists and you have permission to read groups."
        exit 3
    }
    Log "Found group: Id=$($group.Id) DisplayName=$($group.DisplayName)"

    # Check membership (best-effort): attempt to get members filtered by id
    Log "Checking whether the service principal is already a member of the group..."
    $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue
    $isMember = $false
    if ($members) {
        foreach ($m in $members) {
            if ($m.Id -eq $sp.Id) { $isMember = $true; break }
        }
    }

    if ($isMember) {
        Log "Service principal is already a member of group '$GroupName' (Id: $($group.Id)). No action taken."
        exit 0
    }

    # Add service principal to group using REST call to ensure correct reference creation
    Log "Adding service principal to group..."
    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)" } | ConvertTo-Json
    try {
        Invoke-MgGraphRequest -Method POST -Uri "/groups/$($group.Id)/members/\$ref" -Body $body -ContentType 'application/json' -ErrorAction Stop
        Log "Successfully added service principal (Id: $($sp.Id)) to group '$($group.DisplayName)' (Id: $($group.Id))."
    } catch {
        # If already member, Graph returns 400 with specific message; capture and surface
        $e = $_.Exception
        Err "Failed to add service principal to group: $($e.Message)"
        exit 4
    }

} catch {
    $err = $_.Exception
    Err "Unexpected error: $($err.Message)"
    exit 1
} finally {
    # Disconnect session
    try { Disconnect-MgGraph -Confirm:$false -ErrorAction SilentlyContinue } catch {}
}
