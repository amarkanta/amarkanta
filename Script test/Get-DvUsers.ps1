<#
.SYNOPSIS
 Reports Dataverse users in an environment and refreshes (syncs from Entra ID) selected users.

.DESCRIPTION
 Signs in to a specific tenant, then either:
   - Report mode   : lists all Dataverse users and exports a CSV (optionally refreshes users from the report)
   - Targeted mode : refreshes only the user(s) supplied via -EntraId or -InputFile (no full report)

.PARAMETER TenantId        Tenant GUID or domain (e.g. contoso.onmicrosoft.com) to sign in to.
.PARAMETER EnvironmentName Power Platform environment GUID.
.PARAMETER EntraId         One or more Entra object IDs to refresh.
.PARAMETER InputFile       .csv (column EntraId / ObjectId / PrincipalObjectId / azureactivedirectoryobjectid) or .txt (one GUID per line).
.PARAMETER Refresh         Report mode: refresh users with missing email/UPN.
.PARAMETER RefreshAll      Report mode: refresh every user with an Entra id.

.EXAMPLE
 .\Get-DvUsers.ps1 -TenantId <tenant> -EnvironmentName <env-guid>                          # report only
 .\Get-DvUsers.ps1 -TenantId <tenant> -EnvironmentName <env-guid> -Refresh                 # refresh users with missing details
 .\Get-DvUsers.ps1 -TenantId <tenant> -EnvironmentName <env-guid> -RefreshAll              # refresh all
 .\Get-DvUsers.ps1 -TenantId <tenant> -EnvironmentName <env-guid> -EntraId <user-guid>     # refresh a single user
 .\Get-DvUsers.ps1 -TenantId <tenant> -EnvironmentName <env-guid> -InputFile .\users.csv   # refresh users from file

 .\Get-DvUsers.ps1 -TenantId <tenant> -EnvironmentName <env-guid> -EntraId <user-guid> -Enable            # refresh AND enable
 .\Get-DvUsers.ps1 -TenantId <tenant> -EnvironmentName <env-guid> -EntraId <user-guid> -Enable -SkipRefresh   # enable only
 .\Get-DvUsers.ps1 -TenantId <tenant> -EnvironmentName <env-guid> -Enable -WhatIf                  # preview enabling all disabled users

 Prereqs: Install-Module Az.Accounts, Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Report')]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$EnvironmentName,

    [Parameter(ParameterSetName = 'Targeted')][string[]]$EntraId,
    [Parameter(ParameterSetName = 'Targeted')][string]$InputFile,

    [Parameter(ParameterSetName = 'Report')][switch]$Refresh,
    [Parameter(ParameterSetName = 'Report')][switch]$RefreshAll,
    [Parameter(ParameterSetName = 'Report')][string]$OutDir = ".",

    # Enable users that are disabled in Dataverse (isdisabled = true -> false). Needs System Administrator in the environment.
    # Report mode: applies to the refresh targets; with no -Refresh/-RefreshAll it applies to ALL disabled (non-app) users.
    [switch]$Enable,
    # Skip the Entra ID refresh (e.g. only enable; avoids the Power Apps sign-in that Conditional Access may block).
    [switch]$SkipRefresh
)

$ErrorActionPreference = 'Stop'

# ---------- Sign in to the tenant ----------
$ctx = Get-AzContext
if (-not $ctx -or ($ctx.Tenant.Id -ne $TenantId -and $ctx.Tenant.Id -ne (Get-AzTenant -TenantId $TenantId -ErrorAction SilentlyContinue).Id)) {
    try { Connect-AzAccount -Tenant $TenantId -ErrorAction Stop | Out-Null }
    catch {
        Write-Warning "Interactive sign-in failed ($($_.Exception.Message)). Falling back to device code."
        Connect-AzAccount -Tenant $TenantId -UseDeviceAuthentication | Out-Null
    }
}
Write-Host "Signed in to tenant $TenantId" -ForegroundColor Cyan

# The PowerApps admin module signs in with its own app ("Microsoft PowerPlatform Powershell"), which
# Conditional Access may block (AADSTS53003). Only sign in with it when a refresh is actually needed.
$script:ppSignedIn = $false
function Connect-PowerAppsOnDemand {
    if ($script:ppSignedIn) { return }
    try { Add-PowerAppsAccount -TenantID $TenantId | Out-Null; $script:ppSignedIn = $true }
    catch { throw "Power Apps sign-in failed (often a Conditional Access block, error 53003 - use a compliant device/VPN or contact your admin). Details: $($_.Exception.Message)" }
}

function Get-PlainToken([string]$Resource) {
    $t = Get-AzAccessToken -ResourceUrl $Resource -TenantId $TenantId
    if ($t.Token -is [securestring]) { [Net.NetworkCredential]::new('', $t.Token).Password } else { $t.Token }
}

# Resolve the environment's Dataverse URL (BAP REST API, Az token) and build request headers once.
function Initialize-Dataverse {
    if ($script:org) { return }
    $bap = 'https://api.bap.microsoft.com'
    $bapUrl = "$bap/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${EnvironmentName}?api-version=2021-04-01"
    $e = Invoke-RestMethod -Uri $bapUrl -Headers @{ Authorization = "Bearer $(Get-PlainToken $bap)"; Accept = 'application/json' }
    if (-not $e) { throw "Environment $EnvironmentName not found." }
    $o = $e.properties.linkedEnvironmentMetadata.instanceApiUrl
    if (-not $o) { throw "Environment has no Dataverse database." }
    $script:org = $o.TrimEnd('/')
    $script:h = @{ Authorization = "Bearer $(Get-PlainToken $script:org)"; Accept = 'application/json'; 'OData-Version' = '4.0' }
}

# Enable a Dataverse user (by Entra object id) if currently disabled.
function Enable-DvUser {
    [CmdletBinding(SupportsShouldProcess)]
    param([string]$EntraObjectId, [string]$Label)
    Initialize-Dataverse
    $q = "$script:org/api/data/v9.2/systemusers?`$select=systemuserid,isdisabled,applicationid&`$filter=azureactivedirectoryobjectid eq $EntraObjectId"
    $u = @((Invoke-RestMethod -Uri $q -Headers $script:h).value)[0]
    if (-not $u) { Write-Warning "$Label : not found in Dataverse, cannot enable."; return }
    if ($u.applicationid) { Write-Warning "$Label : application user, skipped."; return }
    if (-not $u.isdisabled) { Write-Host "$Label already enabled" -ForegroundColor DarkGray; return }
    if ($PSCmdlet.ShouldProcess($Label, "Enable user in Dataverse")) {
        try {
            Invoke-RestMethod -Method Patch -Uri "$script:org/api/data/v9.2/systemusers($($u.systemuserid))" -Headers $script:h `
                -ContentType 'application/json' -Body '{"isdisabled":false}' | Out-Null
            Write-Host "Enabled $Label" -ForegroundColor Green
        } catch { Write-Warning "Failed to enable ${Label}: $($_.Exception.Message)" }
    }
}

# ---------- Targeted mode: refresh specific users ----------
if ($PSCmdlet.ParameterSetName -eq 'Targeted') {
    $ids = @()
    if ($EntraId) { $ids += $EntraId }
    if ($InputFile) {
        if (-not (Test-Path $InputFile)) { throw "Input file not found: $InputFile" }
        if ([IO.Path]::GetExtension($InputFile) -eq '.csv') {
            $rows = Import-Csv $InputFile
            $col  = 'EntraId','ObjectId','PrincipalObjectId','azureactivedirectoryobjectid' |
                    Where-Object { $rows[0].PSObject.Properties.Name -contains $_ } | Select-Object -First 1
            if (-not $col) { throw "CSV needs a column named EntraId, ObjectId, PrincipalObjectId or azureactivedirectoryobjectid." }
            $ids += $rows.$col
        }
        else { $ids += Get-Content $InputFile }
    }
    $ids = $ids | ForEach-Object { "$_".Trim().Trim('"', "'") } | Where-Object { $_ } | Select-Object -Unique
    if (-not $ids) { throw "Provide -EntraId or -InputFile with at least one Entra object ID." }

    if (-not $SkipRefresh) { Connect-PowerAppsOnDemand }
    $ok = 0; $fail = 0
    foreach ($id in $ids) {
        # Convert the input string to a GUID (accepts dashed, no-dash and {braced} forms) and normalise to lowercase dashed format
        $g = [guid]::Empty
        if (-not [guid]::TryParse($id, [ref]$g)) { Write-Warning "Skipping invalid GUID: $id"; $fail++; continue }
        $id = $g.ToString('D')
        if (-not $SkipRefresh -and $PSCmdlet.ShouldProcess($id, "Refresh Dataverse user in $EnvironmentName")) {
            try {
                Add-AdminPowerAppsSyncUser -EnvironmentName $EnvironmentName -PrincipalObjectId $id | Out-Null
                Write-Host "Refreshed $id" -ForegroundColor Green; $ok++
            } catch { Write-Warning "Failed ${id}: $($_.Exception.Message)"; $fail++; continue }
        }
        if ($Enable) {
            if (-not $SkipRefresh) { Start-Sleep -Seconds 2 }   # let a newly synced user appear
            Enable-DvUser -EntraObjectId $id -Label $id -WhatIf:$WhatIfPreference
        }
    }
    Write-Host "Done. Refreshed: $ok, Failed/Skipped: $fail"
    return
}

# ---------- Report mode ----------
Initialize-Dataverse   # environment lookup via BAP REST + Az token (no PowerApps module sign-in needed)
$org = $script:org; $h = $script:h

$url = "$org/api/data/v9.2/systemusers?`$select=systemuserid,fullname,internalemailaddress,domainname,azureactivedirectoryobjectid,isdisabled,accessmode,applicationid,modifiedon&`$orderby=fullname"
$users = @()
while ($url) {
    $r = Invoke-RestMethod -Uri $url -Headers $h
    $users += $r.value
    $url = $r.'@odata.nextLink'
}

$report = $users | ForEach-Object {
    [pscustomobject]@{
        Name         = $_.fullname
        Email        = $_.internalemailaddress
        UPN          = $_.domainname
        EntraId      = $_.azureactivedirectoryobjectid
        Enabled      = -not $_.isdisabled
        AccessMode   = switch ($_.accessmode) { 0 {'Read-Write'} 1 {'Administrative'} 2 {'Read'} 3 {'Support'} 4 {'Non-interactive'} 5 {'Delegated admin'} default {"$($_.accessmode)"} }
        IsAppUser    = [bool]$_.applicationid
        LastModified = $_.modifiedon
    }
}
$csv = Join-Path $OutDir "users_$EnvironmentName.csv"
$report | Export-Csv $csv -NoTypeInformation
Write-Host "$($report.Count) users. Report: $csv"
$report | Group-Object Enabled | Select-Object @{n='Enabled';e={$_.Name}}, Count | Format-Table

$targets = if ($RefreshAll) { $report }
           elseif ($Refresh) { $report | Where-Object { $_.EntraId -and (-not $_.Email -or -not $_.UPN) } }
           elseif ($Enable) { $report | Where-Object { -not $_.Enabled } }
           else { @() }

$targets = @($targets | Where-Object { $_.EntraId -and -not $_.IsAppUser })
$doRefresh = -not $SkipRefresh -and ($Refresh -or $RefreshAll)
if ($targets.Count -gt 0 -and $doRefresh -and -not $WhatIfPreference) { Connect-PowerAppsOnDemand }

foreach ($u in $targets) {
    if ($doRefresh -and $PSCmdlet.ShouldProcess($u.Name, "Refresh from Entra ID")) {
        try {
            Add-AdminPowerAppsSyncUser -EnvironmentName $EnvironmentName -PrincipalObjectId $u.EntraId | Out-Null
            Write-Host "Refreshed $($u.Name)" -ForegroundColor Green
        } catch { Write-Warning "Failed $($u.Name): $($_.Exception.Message)"; continue }
    }
    if ($Enable) {
        if ($doRefresh) { Start-Sleep -Seconds 1 }
        Enable-DvUser -EntraObjectId $u.EntraId -Label $u.Name -WhatIf:$WhatIfPreference
    }
}
