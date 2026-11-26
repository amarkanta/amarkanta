<#
.SYNOPSIS
    Rolling-upgrade helper for ONE On-premises data gateway cluster member.

.DESCRIPTION
    Run on each gateway server, secondary members first, primary last.
    Disabling/enabling the member in the portal is manual (do it before/after).

    Actions:
      Upgrade : record version -> stop service -> download/verify installer -> silent install -> (optional) restart
      Verify  : post-restart checks (service, version, event log, outbound connectivity)

.EXAMPLE
    .\Update-DataGateway.ps1 -Action Upgrade -Restart
    .\Update-DataGateway.ps1 -Action Verify -TargetVersion 3000.xxx.xx
    .\Update-DataGateway.ps1 -Action Upgrade -InstallerPath D:\Temp\GatewayInstall.exe -WhatIf
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][ValidateSet('Upgrade', 'Verify')]
    [string]$Action,

    # Use a pre-downloaded installer instead of downloading
    [string]$InstallerPath,

    [string]$DownloadUrl = 'https://aka.ms/on-premises-data-gateway-installer',

    # Reboot automatically after a successful install
    [switch]$Restart,

    # Verify: expected version (prefix match), optional
    [string]$TargetVersion,

    # Seconds to let in-flight work drain after stopping (portal disable should be done first)
    [int]$DrainSeconds = 60,

    [string]$LogDir = 'C:\Temp\GatewayUpgrade'
)

$ErrorActionPreference = 'Stop'
$ServiceName = 'PBIEgwService'
$GatewayExe  = 'C:\Program Files\On-premises data gateway\Microsoft.PowerBI.EnterpriseGateway.exe'

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$LogFile = Join-Path $LogDir ("{0}-{1}-{2:yyyyMMdd-HHmmss}.log" -f $env:COMPUTERNAME, $Action, (Get-Date))
Start-Transcript -Path $LogFile | Out-Null

function Get-GatewayVersion {
    if (Test-Path $GatewayExe) { (Get-Item $GatewayExe).VersionInfo.FileVersion } else { $null }
}

function Write-Step($msg) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $msg) -ForegroundColor Cyan }

try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
               ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) { throw 'Run this script in an elevated (Administrator) PowerShell session.' }

    switch ($Action) {

        'Upgrade' {
            Write-Step "Host: $env:COMPUTERNAME"
            $oldVersion = Get-GatewayVersion
            if (-not $oldVersion) { throw "Gateway not found at $GatewayExe. Aborting." }
            Write-Step "Current version: $oldVersion"

            $svc = Get-Service $ServiceName
            Write-Step "Service status: $($svc.Status)"

            Write-Warning 'Confirm this member is DISABLED in the portal and other cluster members are Online.'

            # Installer
            if ($InstallerPath) {
                if (-not (Test-Path $InstallerPath)) { throw "Installer not found: $InstallerPath" }
                $installer = (Resolve-Path $InstallerPath).Path
            } else {
                $dlDir = Join-Path $LogDir ("download-{0:yyyyMMdd-HHmmss}" -f (Get-Date))
                New-Item -ItemType Directory -Force -Path $dlDir | Out-Null
                $installer = Join-Path $dlDir 'GatewayInstall.exe'
                Write-Step "Downloading installer from $DownloadUrl"
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Invoke-WebRequest -Uri $DownloadUrl -OutFile $installer -UseBasicParsing
            }

            # Verify signature before running
            $sig = Get-AuthenticodeSignature $installer
            if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
                throw "Installer signature check failed (Status=$($sig.Status); Signer=$($sig.SignerCertificate.Subject))."
            }
            $newVersion = (Get-Item $installer).VersionInfo.FileVersion
            Write-Step "Installer signature OK. Installer version: $newVersion"

            if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Stop $ServiceName and install gateway $newVersion")) {

                if ($svc.Status -eq 'Running') {
                    Write-Step "Draining for $DrainSeconds s, then stopping $ServiceName"
                    Start-Sleep -Seconds $DrainSeconds
                    Stop-Service $ServiceName -Force
                    $svc.WaitForStatus('Stopped', [TimeSpan]::FromMinutes(2))
                }

                Write-Step 'Installing (silent)...'
                $p = Start-Process -FilePath $installer -ArgumentList '/quiet', 'ACCEPTEULA=yes' -Wait -PassThru
                Write-Step "Installer exit code: $($p.ExitCode)"
                if ($p.ExitCode -notin 0, 3010) { throw "Installer failed with exit code $($p.ExitCode)." }

                Write-Step "Installed version on disk: $(Get-GatewayVersion)"

                if ($Restart) {
                    Write-Step 'Restarting server in 15 seconds. After reboot run: -Action Verify'
                    Stop-Transcript | Out-Null
                    Start-Sleep -Seconds 15
                    Restart-Computer -Force
                    return
                } else {
                    Write-Warning 'Restart the server, then run this script with -Action Verify.'
                }
            }
        }

        'Verify' {
            $results = [ordered]@{}

            $svc = Get-Service $ServiceName -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -ne 'Running') {
                Write-Step "Service is $($svc.Status); attempting start"
                Start-Service $ServiceName
                $svc.WaitForStatus('Running', [TimeSpan]::FromMinutes(2))
                $svc.Refresh()
            }
            $results['Service running']   = ($svc -and $svc.Status -eq 'Running')
            $results['Startup automatic'] = ((Get-CimInstance Win32_Service -Filter "Name='$ServiceName'").StartMode -eq 'Auto')

            $ver = Get-GatewayVersion
            Write-Step "Installed version: $ver"
            $results['Version matches target'] = if ($TargetVersion) { [bool]($ver -and $ver.StartsWith($TargetVersion)) } else { 'n/a (no -TargetVersion)' }

            # Recent gateway errors (last 30 min)
            $errs = Get-WinEvent -FilterHashtable @{
                LogName = 'Application'; Level = 1, 2; StartTime = (Get-Date).AddMinutes(-30)
            } -ErrorAction SilentlyContinue | Where-Object { $_.ProviderName -match 'gateway|PBIEgw' }
            $results['No recent gateway errors'] = (-not $errs)
            if ($errs) { $errs | Select-Object -First 5 TimeCreated, ProviderName, Message | Format-List }

            # Basic outbound connectivity (HTTPS)
            foreach ($h in 'login.microsoftonline.com', 'api.powerbi.com') {
                $results["HTTPS to $h"] = (Test-NetConnection $h -Port 443 -WarningAction SilentlyContinue).TcpTestSucceeded
            }

            Write-Host "`n===== Verification summary =====" -ForegroundColor Yellow
            $results.GetEnumerator() | ForEach-Object {
                $ok = $_.Value -eq $true -or $_.Value -like 'n/a*'
                Write-Host ("{0,-30} {1}" -f $_.Key, $_.Value) -ForegroundColor ($(if ($ok) { 'Green' } else { 'Red' }))
            }
            Write-Host "`nNext (manual): enable the member in the portal, Test connection on a data source, run a test refresh." -ForegroundColor Yellow

            if ($results.Values -contains $false) { throw 'One or more checks failed. See log.' }
        }
    }
}
catch {
    Write-Error $_
    exit 1
}
finally {
    try { Stop-Transcript | Out-Null } catch {}
    Write-Host "Log: $LogFile"
}
