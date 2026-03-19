<#
.SYNOPSIS
    Checks that the repo's required modules are installed and that M365AuthConfig.psd1 is valid.

.DESCRIPTION
    This script:
      - Validates that M365AuthConfig.psd1 is present and contains the required keys
      - Scans the repo for Import-Module statements and checks that those modules are installed
      - Optionally installs missing modules (use -InstallMissing or answer Yes when prompted)

.NOTES
    Run from the repo root:
        .\Test-M365Environment.ps1
    To auto-install missing modules:
        .\Test-M365Environment.ps1 -InstallMissing
#>

[CmdletBinding()]
param(
    [switch]$InstallMissing
)

$repoRoot = $PSScriptRoot
$configPath = Join-Path $repoRoot 'M365AuthConfig.psd1'

function Test-AuthConfig {
    param(
        [string]$Path
    )

    Write-Host "`nChecking auth config: $Path" -ForegroundColor Cyan

    if (-not (Test-Path $Path)) {
        Write-Host "  Config file not found." -ForegroundColor Red
        return $false
    }

    try {
        $config = Import-PowerShellDataFile -Path $Path
    } catch {
        Write-Host "  Failed to parse config file: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    $required = @('TenantId','ClientId','CertificateThumbprint')
    $missing = @()

    foreach ($key in $required) {
        if (-not $config.ContainsKey($key) -or [string]::IsNullOrWhiteSpace($config[$key])) {
            $missing += $key
        }
    }

    if ($missing.Count -gt 0) {
        Write-Host "  Missing required config values: $($missing -join ', ')" -ForegroundColor Yellow
        Write-Host "    (This will prevent certificate-based authentication from working.)" -ForegroundColor Yellow
        return $false
    }

    Write-Host "  Config file is present and contains required values." -ForegroundColor Green
    return $true
}

function Get-ProjectModules {
    param(
        [string]$Root
    )

    $scriptFiles = Get-ChildItem -Path $Root -Recurse -Include *.ps1 -File -ErrorAction SilentlyContinue
    $moduleNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($file in $scriptFiles) {
        $lines = Get-Content -LiteralPath $file.FullName -ErrorAction SilentlyContinue
        $inBlockComment = $false
        foreach ($line in $lines) {
            if ($line -match '<#') {
                $inBlockComment = $true
            }
            if ($inBlockComment) {
                if ($line -match '#>') {
                    $inBlockComment = $false
                }
                continue
            }

            if ($line -match '\bImport-Module\b') {
                # Skip Import-Module occurrences inside comments.
                $importPos = $line.IndexOf('Import-Module')
                $hashPos = $line.IndexOf('#')
                if ($hashPos -ge 0 -and $hashPos -lt $importPos) {
                    continue
                }

                $tokens = -split $line
                for ($i = 0; $i -lt $tokens.Count; $i++) {
                    if ($tokens[$i] -ieq 'Import-Module') {
                        $nameIndex = $i + 1
                        if ($nameIndex -lt $tokens.Count) {
                            $candidate = $tokens[$nameIndex]

                            if ($candidate -ieq '-Name' -or $candidate -ieq '-Module') {
                                $nameIndex++
                                if ($nameIndex -lt $tokens.Count) {
                                    $candidate = $tokens[$nameIndex]
                                } else {
                                    continue
                                }
                            }

                            $name = $candidate.Trim('"','''')

                            if ($name -and $name -notlike '*\*' -and $name -notlike '*/*' -and $name -notlike '*.psm1' -and $name -notlike '*.psd1') {
                                $moduleNames.Add($name) | Out-Null
                            }
                        }
                    }
                }
            }
        }
    }

    return $moduleNames
}

function Test-ModulesInstalled {
    param(
        [string[]]$ModuleNames
    )

    $results = [ordered]@{}
    foreach ($name in $ModuleNames) {
        $installed = Get-InstalledModule -Name $name -ErrorAction SilentlyContinue
        $results[$name] = $null -ne $installed
    }

    return $results
}

function Install-MissingModules {
    param(
        [string[]]$ModuleNames
    )

    foreach ($name in $ModuleNames) {
        Write-Host "Installing module: $name" -ForegroundColor Cyan
        try {
            Install-Module -Name $name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "  Installed $name" -ForegroundColor Green
        } catch {
            Write-Host ("  Failed to install {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
        }
    }
}

Write-Host "=== M365 Repo Environment Check ===" -ForegroundColor Cyan

$configOk = Test-AuthConfig -Path $configPath

Write-Host "`nDetecting referenced modules in repo scripts..." -ForegroundColor Cyan
$modules = Get-ProjectModules -Root $repoRoot
if ($modules.Count -eq 0) {
    Write-Host "  No module references found." -ForegroundColor Yellow
} else {
    Write-Host "  Found $($modules.Count) module(s): $($modules -join ', ')" -ForegroundColor Cyan

    $moduleResult = Test-ModulesInstalled -ModuleNames $modules
    $missing = $moduleResult.Keys | Where-Object { -not $moduleResult[$_] }

    if ($missing.Count -gt 0) {
        Write-Host "`nMissing required module(s):" -ForegroundColor Red
        foreach ($m in $missing) { Write-Host "  - $m" -ForegroundColor Red }

        $shouldInstall = $InstallMissing
        if (-not $shouldInstall) {
            $answer = Read-Host "Install missing modules now? [Y/N]"
            $shouldInstall = $answer -match '^[Yy]'
        }

        if ($shouldInstall) {
            Install-MissingModules -ModuleNames $missing

            # Re-check after install attempt.
            $moduleResult = Test-ModulesInstalled -ModuleNames $modules
            $missing = $moduleResult.Keys | Where-Object { -not $moduleResult[$_] }
        } else {
            Write-Host "`nInstall missing modules with: Install-Module <ModuleName> -Scope CurrentUser" -ForegroundColor Yellow
        }
    } else {
        Write-Host "`nAll referenced modules are installed." -ForegroundColor Green
    }
}

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$authConfigColor = if ($configOk) { 'Green' } else { 'Red' }
Write-Host "  Auth config: $([bool]$configOk)" -ForegroundColor $authConfigColor

if ($modules.Count -gt 0) {
    $missing = $moduleResult.Keys | Where-Object { -not $moduleResult[$_] }
    $missingColor = if ($missing.Count -eq 0) { 'Green' } else { 'Red' }
    Write-Host "  Modules missing: $($missing.Count)" -ForegroundColor $missingColor
}

$missingCount = if ($modules.Count -gt 0) { ($moduleResult.Keys | Where-Object { -not $moduleResult[$_] }).Count } else { 0 }
if (-not $configOk -or ($missingCount -gt 0)) {
    exit 1
}

exit 0
