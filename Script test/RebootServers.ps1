<#
.SYNOPSIS
    Remotely reboots servers and verifies they come back online within a time limit.

.EXAMPLE
    .\RebootServers.ps1
    .\RebootServers.ps1 -Servers 'SRV1','SRV2' -TimeoutSeconds 120 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string[]]$Servers = @(
        'XSPW70K128A.pharma.aventis.com',
        'XSPW70C889W.pharma.aventis.com',
        'XSPW70W276S.pharma.aventis.com',
        'XSPW70C488F.pharma.aventis.com'
    ),
    [int]$TimeoutSeconds = 120,   # how long to wait for each server to come back
    [int]$PollSeconds    = 5,
    [System.Management.Automation.PSCredential]$Credential
)

$cred = @{}
if ($Credential) { $cred.Credential = $Credential }

function Get-BootTime($server) {
    try {
        (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $server -ErrorAction Stop @cred).LastBootUpTime
    } catch { $null }
}

# 1. Capture pre-reboot boot time and issue the reboot
$state = @{}
foreach ($s in $Servers) {
    $before = Get-BootTime $s
    if (-not $before) {
        Write-Warning "$s : unreachable, skipping."
        $state[$s] = [pscustomobject]@{ Server = $s; Before = $null; Status = 'SKIPPED (unreachable)'; After = $null }
        continue
    }
    if ($PSCmdlet.ShouldProcess($s, 'Restart-Computer -Force')) {
        try {
            Restart-Computer -ComputerName $s -Force -ErrorAction Stop @cred
            Write-Host "$s : reboot issued (last boot $before)" -ForegroundColor Yellow
            $state[$s] = [pscustomobject]@{ Server = $s; Before = $before; Status = 'PENDING'; After = $null }
        } catch {
            Write-Warning "$s : reboot failed - $($_.Exception.Message)"
            $state[$s] = [pscustomobject]@{ Server = $s; Before = $before; Status = 'REBOOT FAILED'; After = $null }
        }
    }
}

# 2. Poll until each server reports a newer boot time, or timeout
$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
Start-Sleep -Seconds 10   # give servers time to actually go down
while ((Get-Date) -lt $deadline) {
    $pending = $state.Values | Where-Object Status -eq 'PENDING'
    if (-not $pending) { break }
    foreach ($p in $pending) {
        $now = Get-BootTime $p.Server
        if ($now -and $now -gt $p.Before) {
            $p.After  = $now
            $p.Status = 'ONLINE (rebooted)'
            Write-Host "$($p.Server) : back online (new boot $now)" -ForegroundColor Green
        }
    }
    Start-Sleep -Seconds $PollSeconds
}

foreach ($p in ($state.Values | Where-Object Status -eq 'PENDING')) {
    $p.Status = "NOT BACK within $TimeoutSeconds s"
}

# 3. Summary
Write-Host "`n=== Reboot Summary ===" -ForegroundColor Cyan
$state.Values | Select-Object Server, Status, Before, After | Format-Table -AutoSize

if ($state.Values | Where-Object { $_.Status -notlike 'ONLINE*' }) { exit 1 }
