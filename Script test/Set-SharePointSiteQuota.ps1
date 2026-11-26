#Requires -Modules Microsoft.Online.SharePoint.PowerShell

<#
.SYNOPSIS
    Reports all SharePoint site collections and adds additional storage quota to each.

.PARAMETER TenantAdminUrl
    SharePoint admin center URL.
    Example: https://contoso-admin.sharepoint.com

.PARAMETER AdditionalQuotaGB
    GB to add to every site's current quota. Default: 50

.PARAMETER ReportOnly
    Generate the report without modifying any quotas.

.EXAMPLE
    .\Set-SharePointSiteQuota.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com"

.EXAMPLE
    .\Set-SharePointSiteQuota.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
                                  -AdditionalQuotaGB 100 -ReportOnly
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [string]$TenantAdminUrl,

    [Parameter()]
    [ValidateRange(1, 25600)]
    [int]$AdditionalQuotaGB = 50,

    [Parameter()]
    [switch]$ReportOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Constants ─────────────────────────────────────────────────────────────────
$MB_PER_GB    = 1024
$AddedQuotaMB = $AdditionalQuotaGB * $MB_PER_GB

# ── Connect ───────────────────────────────────────────────────────────────────
Write-Host "`nConnecting to SharePoint Online..." -ForegroundColor Cyan
Connect-SPOService -Url $TenantAdminUrl

# ── Get all site collections ──────────────────────────────────────────────────
Write-Host "Retrieving all site collections..." -ForegroundColor Cyan

$allSites = Get-SPOSite -IncludePersonalSite $false -Limit All -Detailed

if (-not $allSites) {
    Write-Warning "No sites found. Exiting."
    Disconnect-SPOService
    return
}

Write-Host "Found $($allSites.Count) site(s).`n" -ForegroundColor Green

# ── Process sites ─────────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter  = 0

foreach ($site in $allSites) {
    $counter++
    $pct = [math]::Round(($counter / $allSites.Count) * 100)
    Write-Progress -Activity "Processing sites" `
                   -Status "$counter of $($allSites.Count) — $($site.Url)" `
                   -PercentComplete $pct

    $currentQuotaMB = $site.StorageQuota
    $usedMB         = $site.StorageUsageCurrent
    $newQuotaMB     = $currentQuotaMB + $AddedQuotaMB
    $usedPct        = if ($currentQuotaMB -gt 0) {
                          [math]::Round(($usedMB / $currentQuotaMB) * 100, 1)
                      } else { 0 }

    $record = [PSCustomObject]@{
        Title            = $site.Title
        Url              = $site.Url
        Template         = $site.Template
        Owner            = $site.Owner
        CurrentQuota_GB  = [math]::Round($currentQuotaMB / $MB_PER_GB, 2)
        Used_GB          = [math]::Round($usedMB          / $MB_PER_GB, 2)
        Used_Pct         = "$usedPct%"
        Added_GB         = $AdditionalQuotaGB
        NewQuota_GB      = [math]::Round($newQuotaMB      / $MB_PER_GB, 2)
        Status           = 'Pending'
        Error            = ''
    }

    if ($ReportOnly) {
        $record.Status = 'ReportOnly'
    }
    elseif ($PSCmdlet.ShouldProcess($site.Url, "Increase quota by $AdditionalQuotaGB GB")) {
        try {
            Set-SPOSite -Identity $site.Url -StorageQuota $newQuotaMB
            $record.Status = 'Updated'
            Write-Host "  [OK] $($site.Url)" -ForegroundColor Green
        }
        catch {
            $record.Status = 'Failed'
            $record.Error  = $_.Exception.Message
            Write-Warning "  [FAIL] $($site.Url) — $($_.Exception.Message)"
        }
    }

    $results.Add($record)
}

Write-Progress -Activity "Processing sites" -Completed

# ── Console summary ───────────────────────────────────────────────────────────
Write-Host "`n===== REPORT =====" -ForegroundColor Cyan
$results | Format-Table Title, Url, CurrentQuota_GB, Used_GB, Used_Pct, NewQuota_GB, Status -AutoSize

$updated = ($results | Where-Object Status -eq 'Updated').Count
$failed  = ($results | Where-Object Status -eq 'Failed').Count
$skipped = ($results | Where-Object Status -eq 'ReportOnly').Count

Write-Host "Total sites : $($results.Count)"   -ForegroundColor Cyan
Write-Host "Updated     : $updated"             -ForegroundColor Green
Write-Host "Failed      : $failed"              -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Green' })
if ($ReportOnly) {
Write-Host "Report-only : $skipped (no changes made)" -ForegroundColor Yellow }

# ── Export CSV ────────────────────────────────────────────────────────────────
$reportPath = ".\SPO_QuotaReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nReport saved → $reportPath" -ForegroundColor Cyan

# ── Disconnect ────────────────────────────────────────────────────────────────
Disconnect-SPOService
Write-Host "Done.`n" -ForegroundColor Green
