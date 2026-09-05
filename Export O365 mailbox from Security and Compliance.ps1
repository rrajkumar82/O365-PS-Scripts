# Modernized mailbox export script
# - Requires ExchangeOnlineManagement module (v3.9.0+ recommended)
# - Supports interactive or app-only authentication
# - Uses Connect-IPPSSession -EnableSearchOnlySession for compliance cmdlets
# - Safer path handling, logging, error handling, and export parsing

Param(
    [Parameter(Mandatory=$true)][string]$Mailbox,
    [Parameter(Mandatory=$true)][string]$UserFullname,
    [Parameter(Mandatory=$true)][string]$LocalExportLocation,
    [Parameter(Mandatory=$true)][string]$MailboxUserCountry,
    [Parameter(Mandatory=$false)][string]$NASDrive = "D:",
    [switch]$UseAppOnly,
    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$TenantId,
    [switch]$EnableTranscript,
    [int]$PollIntervalSeconds = 10
)

Set-StrictMode -Version Latest

# ---------- Logging ----------
$LogFolder = Join-Path -Path $env:Public -ChildPath "Documents\O365_PSscriptExecution"
if (!(Test-Path -Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null }
$Date = (Get-Date).ToString('yyyyMMdd_HHmmss')
$LogFile = Join-Path $LogFolder "$Date`_Mbx_exportjob.log"
if ($EnableTranscript) {
    Start-Transcript -Path $LogFile -Append
}

# Minimal console logging helper
function Log { param($Message) Write-Host $Message; if ($EnableTranscript) { $Message | Out-File -FilePath $LogFile -Append } }

try {
    Log "Starting mailbox export for $Mailbox (User: $UserFullname)"
    # ---------- Ensure module ----------
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Log "ExchangeOnlineManagement module not found. Installing..."
        Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser
    }
    Import-Module ExchangeOnlineManagement -Force

    # Recommend current module
    $EOModuleVersion = (Get-Module ExchangeOnlineManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version
    Log "ExchangeOnlineManagement module version: $EOModuleVersion"

    # ---------- Connect to Exchange Online (interactive or app-only) ----------
    if ($UseAppOnly) {
        if (-not $AppId -or -not $CertificateThumbprint) {
            throw "App-only auth requested but AppId or CertificateThumbprint missing."
        }
        Log "Connecting to Exchange Online (app-only) using AppId $AppId..."
        # For app-only Connect-ExchangeOnline parameters:
        # -AppId and -CertificateThumbprint are supported for Connect-ExchangeOnline
        if ($TenantId) {
            Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $TenantId -ErrorAction Stop
        } else {
            Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop
        }
    } else {
        Log "Connecting to Exchange Online (interactive). Please complete MFA prompt if shown..."
        Connect-ExchangeOnline -ErrorAction Stop
    }

    # ---------- Connect to Compliance / Purview (IPPSSession) ----------
    Log "Connecting to Security & Compliance endpoint (search-only session)..."
    # Use -EnableSearchOnlySession per modern guidance
    Connect-IPPSSession -EnableSearchOnlySession -ErrorAction Stop

    # ---------- Validate mailbox existence ----------
    function Get-MailboxAvailability {
        param([string]$MailboxToCheck)
        try {
            $mbx = Get-EXOMailbox -Identity $MailboxToCheck -ErrorAction Stop
            return $true
        } catch {
            return $false
        }
    }

    if (-not (Get-MailboxAvailability -MailboxToCheck $Mailbox)) {
        throw "Mailbox $Mailbox not found in tenant."
    }

    # ---------- Prepare local paths ----------
    $ExportRoot = Join-Path -Path $LocalExportLocation -ChildPath $MailboxUserCountry
    if (-not (Test-Path -Path $ExportRoot)) { New-Item -Path $ExportRoot -ItemType Directory -Force | Out-Null }
    Log "Export root: $ExportRoot"

    # ---------- Build search and export names ----------
    $SafeName = ($UserFullname -replace '[^a-zA-Z0-9\-_\.? ]','_').Trim()
    $SearchName = "${SafeName}_mailbox"
    $ExportName = "${SearchName}_Export"
    Log "Search name: $SearchName; Export name: $ExportName"

    # ---------- Check existing jobs ----------
    function Get-ExistingSearchJobResult {
        param([string]$SearchName, [string]$ExportName)
        $result = @{ SearchExists = $false; ExportExists = $false }
        try {
            if (Get-ComplianceSearch -Identity $SearchName -ErrorAction Stop) { $result.SearchExists = $true }
        } catch { $result.SearchExists = $false }
        try {
            if (Get-ComplianceSearchAction -Identity $ExportName -ErrorAction Stop) { $result.ExportExists = $true }
        } catch { $result.ExportExists = $false }
        return $result
    }

    $existing = Get-ExistingSearchJobResult -SearchName $SearchName -ExportName $ExportName
    if ($existing.SearchExists) {
        Log "A compliance search named $SearchName already exists. Please review or specify a different UserFullname."
        throw "Conflicting compliance search name"
    }
    if ($existing.ExportExists) {
        Log "An export action named $ExportName already exists. Please review & delete it before proceeding."
        throw "Conflicting compliance export action name"
    }

    # ---------- Create and run compliance search ----------
    Log "Creating compliance search..."
    $newSearch = New-ComplianceSearch -Name $SearchName -ExchangeLocation $Mailbox -Description "Email: $Mailbox" -AllowNotFoundExchangeLocationsEnabled $true -ErrorAction Stop
    Log "Starting compliance search..."
    Start-ComplianceSearch -Identity $SearchName -ErrorAction Stop

    # Wait for search completion
    Log "Waiting for compliance search to complete..."
    while ($true) {
        Start-Sleep -Seconds $PollIntervalSeconds
        $status = (Get-ComplianceSearch -Identity $SearchName -ErrorAction Stop).Status
        Log "Search status: $status"
        if ($status -eq 'Completed') { break }
        if ($status -match 'Failed|Stopped') { throw "Compliance search entered status: $status" }
    }
    Log "Compliance search completed."

    # ---------- Create export action ----------
    Log "Creating export action..."
    $exportAction = New-ComplianceSearchAction -SearchName $SearchName -Export -Format FxStream -ExchangeArchiveFormat PerUserPst -Scope BothIndexedAndUnindexedItems -EnableDedupe $true -ErrorAction Stop

    # Wait for export details and extract Container URL and SAS token using regex (robust against field order)
    Log "Waiting for export action to be prepared..."
    $containerUrl = $null; $sasToken = $null
    while ($true) {
        Start-Sleep -Seconds $PollIntervalSeconds
        $action = Get-ComplianceSearchAction -Identity $ExportName -IncludeCredential -Details -ErrorAction SilentlyContinue
        if (-not $action) { Log "Export action not ready yet..."; continue }
        $raw = $action.Results -as [string]
        if ($raw) {
            # Try to extract Container url and SAS token with regex
            $cMatch = [regex]::Match($raw, 'Container url:\s*(?<url>https?://[^\s;]+)', 'IgnoreCase')
            $sMatch = [regex]::Match($raw, 'SAS token:\s*(?<sas>[^;]+)', 'IgnoreCase')
            if ($cMatch.Success -and $sMatch.Success) {
                $containerUrl = $cMatch.Groups['url'].Value.Trim()
                $sasToken = $sMatch.Groups['sas'].Value.Trim()
                break
            } else {
                Log "Export action created but container url/SAS not available yet."
            }
        }
    }

    if (-not $containerUrl -or -not $sasToken) { throw "Unable to obtain export Container URL or SAS token." }

    # Do not log the full SAS token; log truncated version only
    Log ("Found export container URL: {0}" -f $containerUrl)
    Log ("Found SAS token: {0}****" -f ($sasToken.Substring(0, [Math]::Min(8,$sasToken.Length))))

    # ---------- Find Unified Export Tool (UET) ----------
    Log "Locating the Microsoft Unified Export Tool..."
    $exportExe = Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA 'Apps\2.0') -Filter 'microsoft.office.client.discovery.unifiedexporttool.exe' -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notmatch '_none_' } | Select-Object -First 1
    if (-not $exportExe) {
        throw "Unified Export Tool not found under %LOCALAPPDATA%\Apps\2.0. Please ensure the Microsoft 365 export tool is installed on the machine running this script."
    }
    $exportExePath = $exportExe.FullName
    Log "Unified Export Tool: $exportExePath"

    # ---------- Start UET to download the export ----------
    $dest = $ExportRoot
    if (-not (Test-Path -Path $dest)) { New-Item -Path $dest -ItemType Directory -Force | Out-Null }

    $args = @(
        "-name", "`"$SearchName`"",
        "-source", "`"$containerUrl`"",
        "-key", "`"$sasToken`"",
        "-dest", "`"$dest`"",
        "-trace", "true"
    )

    Log "Starting Unified Export Tool to download export. Destination: $dest"
    Start-Process -FilePath $exportExePath -ArgumentList $args -NoNewWindow -PassThru | Out-Null

    # Monitor the UET process and export progress (best-effort)
    Log "Monitoring download process..."
    while (Get-Process -Name 'microsoft.office.client.discovery.unifiedexporttool' -ErrorAction SilentlyContinue) {
        Start-Sleep -Seconds $PollIntervalSeconds
        # Attempt to compute bytes downloaded (best-effort)
        try {
            $downloaded = (Get-ChildItem -Path (Join-Path $dest "$SearchName`_Export") -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            if ($downloaded) { Log ("Downloaded bytes so far: {0}" -f $downloaded) }
        } catch {}
    }
    Log "Download process finished (UET exited)."

    # ---------- Copy to NAS ----------
    $nasFolder = Join-Path -Path $NASDrive.TrimEnd(':') + ':' -ChildPath $MailboxUserCountry
    # If user passed D: as 'D:' join will yield correct path. Normalize:
    if ($NASDrive -match '^[A-Za-z]:$') { $nasFolder = Join-Path -Path $NASDrive -ChildPath $MailboxUserCountry }
    if (-not (Test-Path -Path $nasFolder)) { New-Item -Path $nasFolder -ItemType Directory -Force | Out-Null }
    Log "Copying exported files to NAS location: $nasFolder"
    Copy-Item -Path $ExportRoot -Destination $nasFolder -Recurse -Force -ErrorAction Stop
    Log "Copy completed."

    Log "Mailbox export workflow finished successfully for $Mailbox."

} catch {
    $err = $_.Exception
    Log "ERROR: $($err.Message)"
    if ($EnableTranscript) { Write-Error $err }
    exit 1
} finally {
    # Cleanup: disconnect sessions if connected
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try { Remove-PSSession -ErrorAction SilentlyContinue } catch {}
    if ($EnableTranscript) { Stop-Transcript }
}
