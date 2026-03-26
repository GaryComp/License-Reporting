# ============================================================
# OneDrive Modified Files Indexer - Azure CLI Auth
# ============================================================
# PREREQUISITES:
#   1. Azure CLI installed and logged in: az login
#   2. Account must have access to the target user's OneDrive
# ============================================================

# --- CONFIG ---
$UserPrincipalName = "gary.compagnon@clutchsolutions.com"        # Target user's UPN
$FolderPath        = ""          # Relative path in OneDrive (blank = root)
$ModifiedAfter     = [datetime]"2025-03-23"        # Cutoff date
$OutputCsv         = "\OneDrive_Modified_Files.csv"
# --------------

# --- GET TOKEN VIA AZURE CLI ---
function Get-GraphToken {
    Write-Host "Getting token from Azure CLI..." -ForegroundColor Cyan
    $tokenJson = az account get-access-token --resource https://graph.microsoft.com 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Azure CLI token fetch failed. Make sure you're logged in with: az login"
    }
    return ($tokenJson | ConvertFrom-Json).accessToken
}

# --- RECURSIVE FILE INDEXER ---
$script:FoldersScanned = 0
$script:FilesMatched   = 0

function Get-DriveItems {
    param (
        [string]$Token,
        [string]$ItemId,
        [string]$CurrentPath
    )

    $script:FoldersScanned++
    Write-Progress -Activity "Scanning OneDrive" `
                   -Status "Folder: $CurrentPath" `
                   -CurrentOperation "$script:FilesMatched file(s) matched so far  |  $script:FoldersScanned folder(s) scanned"

    $headers = @{ Authorization = "Bearer $Token" }
    $results = @()
    $url = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/drive/items/$ItemId/children?`$top=999"

    do {
        try {
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        } catch {
            Write-Warning "Failed to list items at '$CurrentPath': $_"
            break
        }

        foreach ($item in $response.value) {
            $itemPath = "$CurrentPath/$($item.name)"
            if ($item.folder) {
                $results += Get-DriveItems -Token $Token -ItemId $item.id -CurrentPath $itemPath
            } elseif ($item.file) {
                $lastModified = [datetime]$item.lastModifiedDateTime
                if ($lastModified -gt $ModifiedAfter) {
                    $script:FilesMatched++
                    $results += [PSCustomObject]@{
                        Name             = $item.name
                        Path             = $itemPath
                        SizeKB           = [math]::Round($item.size / 1KB, 2)
                        LastModified     = $lastModified.ToString("yyyy-MM-dd HH:mm:ss")
                        LastModifiedBy   = $item.lastModifiedBy.user.displayName
                        WebUrl           = $item.webUrl
                    }
                }
            }
        }
        $url = $response.'@odata.nextLink'
    } while ($url)

    return $results
}

# --- MAIN ---
$token = Get-GraphToken
$headers = @{ Authorization = "Bearer $token" }

# Resolve starting folder
try {
    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        $startItem = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/drive/root" `
            -Headers $headers
    } else {
        $encodedPath = [Uri]::EscapeDataString($FolderPath)
        $startItem = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/drive/root:/$encodedPath" `
            -Headers $headers
    }
} catch {
    throw "Could not resolve folder '$FolderPath' for user '$UserPrincipalName'. Check the UPN and folder path. Error: $_"
}

Write-Host "Indexing: '$($startItem.name)' for $UserPrincipalName..." -ForegroundColor Cyan
$files = Get-DriveItems -Token $token -ItemId $startItem.id -CurrentPath $startItem.name
Write-Progress -Activity "Scanning OneDrive" -Completed

Write-Host "Found $($files.Count) file(s) modified after $($ModifiedAfter.ToString('yyyy-MM-dd'))" -ForegroundColor Green

if ($files.Count -gt 0) {
    $files | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $OutputCsv" -ForegroundColor Green
} else {
    Write-Host "No files matched. No CSV created." -ForegroundColor Yellow
}
```