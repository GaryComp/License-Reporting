<#
=============================================================================================
Name: Get Exchange Online Mailbox Folder Statistics Using PowerShell
~~~~~~~~~~~~~~~~~
Script Highlights: 
~~~~~~~~~~~~~~~~~
1. The script verifies and installs Exchange PowerShell module (if not installed already) upon your confirmation.
2. Retrieve folder statistics for all mailbox folders.
3. Retrieve statistics for specific mailbox folders.
4. Provides folder statistics for a single user and bulk users.
5. Allows to use filter to get folder statistics for all user mailboxes.
6. Allows to use filter to get folder statistics for all shared mailboxes.
7. The script can be executed with an MFA-enabled account too.
8. Exports report results to CSV.
9. The script is scheduler friendly.
10. It can be executed with certificate-based authentication (CBA) too.

============================================================================================
#>

    Param (
    [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [string]$TenantId,
        [string]$CertificateThumbprint,
        [string]$UserName,
        [SecureString]$Password,
        [string]$MailboxUPN  ,
        [string]$MailBoxCSV ,
        [switch]$UserMailboxOnly,
        [switch]$SharedMailboxOnly,
        [string]$FolderPaths #Must folderpaths like(/Inbox,/Sent Items,/Inbox/SubFolder)
         )
    Import-Module "$PSScriptRoot\..\M365AuthModule.psm1" -Force
    Connect-M365Services -Services "ExchangeOnline" -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -UserName $UserName -Password $Password

#Outputfile path
   $ExportsDir = Join-Path $PSScriptRoot '..' 'Exports'
   if (-not (Test-Path $ExportsDir)) { New-Item -Path $ExportsDir -ItemType Directory | Out-Null }
   $OutputCSV = Join-Path $ExportsDir "MailboxFolderStatisticsReports_$((Get-Date -format 'dd-MMM-yyyy-ddd hh-mm-ss').ToString()).csv"
  
#Function to FolderStatistics
 Function GetFolderStatistics($Folder){ 
        $FolderDetail = [PSCustomObject]@{
                        "Display Name" = $DisplayName
                        "UPN" =$MailboxUPN
                        "Folder Name" =$Folder.Name
                        "Folder Path" =$Folder.FolderPath
                        "Items In Folder" =$Folder.ItemsInFolder
                        "Folder Size" =$Folder.FolderSize.ToString().split("(") | Select-Object -Index 0
                        "Items In Folder And Subfolders" =$Folder.ItemsInFolderAndSubfolders 
                        "Folder And Subfolder Size" =$Folder.FolderAndSubfolderSize.ToString().split("(") | Select-Object -Index 0
                        "Deleted Items In Folder" =$Folder.DeletedItemsInFolder
                        "DeletedItems In Folder And Subfolders" =$Folder.DeletedItemsInFolderAndSubfolders
                        "Visible Items In Folder" =$Folder.VisibleItemsInFolder
                        "Hidden Items In Folder" =$Folder.HiddenItemsInFolder
                        "Mailbox Type" =$Mailboxtype
                        "Folder Type" =$Folder.FolderType
                        "Creation Time" =$Folder.CreationTime
                        "Last Modified Time" =$Folder.LastModifiedTime
                          }  
         $FolderDetail | Export-csv -Path $OutputCSV -Append -NoTypeInformation -Force 
        
}  
#Function to statistics for all folders 
Function ProcessAllMailboxFolders {
  Param([string]$MailboxUPN)
          $FolderDetails = Get-EXOMailboxFolderStatistics -Identity $MailboxUPN 
          if($FolderDetails){
          Foreach($Folder in $FolderDetails){
            GetFolderStatistics $Folder
          }
          Add-Content -Path "$OutputCSV" -Value "" 
  }
}
#Function to statistics for specific Folders
Function ProcessSpecificMailboxFolder {
    Param(
        [string]$MailboxUPN,
        [string]$FolderPaths
         )
        $Folders = $FolderPaths -split ',' |ForEach-Object { $_.Trim() }
        $FolderPathDetails = Get-EXOMailboxFolderStatistics -Identity $MailboxUPN | Where-Object { $_.FolderPath -in $Folders}
        if($FolderPathDetails){
           Foreach($Folder in $FolderPathDetails){
      
            GetFolderStatistics $Folder
            } 
            Add-Content -Path "$OutputCSV" -Value ""  
     }
 }
 #Function to get mailbox details
 Function Getmailbox{
   Param($MailboxUPN)
       $MailboxInfo = Get-EXOMailbox -UserPrincipalName $MailboxUPN
       $DisplayName = $MailboxInfo.DisplayName
       $MailboxType = $MailboxInfo.RecipientTypeDetails
       Invoke-MailboxReport -MailboxUPN $MailboxUPN
 }
 Function Invoke-MailboxReport {
   Param ($MailboxUPN)

        Write-Progress -Activity "Processing mailboxes" -Status "Processed: $ProgressIndex | Currently Processing: $MailboxUPN"
        if ($FolderPaths -eq "") {
           ProcessAllMailboxFolders -MailboxUPN $MailboxUPN
        }
        else {
           ProcessSpecificMailboxFolder -MailboxUPN $MailboxUPN -FolderPaths $FolderPaths
        }
}
 #Single user
 if ($MailboxUPN){ 
     $ProgressIndex =1 
     Getmailbox -MailboxUPN $MailboxUPN
     
 }

#Multiple CSV users
elseif($MailBoxCSV){ 
        $Mailboxes = Import-Csv -Path $MailBoxCSV
        $ProgressIndex = 0
        foreach ($Mailbox in $Mailboxes) {
            $ProgressIndex++
            Getmailbox -MailboxUPN $Mailbox.Mailboxes
        }
} 

else{ 
    $ProgressIndex =0
     if($SharedMailboxOnly.IsPresent -or $UserMailboxOnly.IsPresent){
  
         if ($SharedMailboxOnly.IsPresent) {
           $RecipientType ="SharedMailbox"
         }
         else {
           $RecipientType = "UserMailbox"
         }
           Get-EXOMailbox -RecipientTypeDetails $RecipientType -ResultSize Unlimited | ForEach-Object {
           $ProgressIndex++
           $DisplayName =$_.DisplayName
           $MailboxUPN = $_.UserPrincipalName
           $MailboxType = $_.RecipientTypeDetails
      
           Invoke-MailboxReport -MailboxUPN $MailboxUPN
           }
      }     
      else{ 
             Get-EXOMailbox -ResultSize Unlimited | ForEach-Object {
             $ProgressIndex++
             $DisplayName =$_.DisplayName
             $MailboxUPN = $_.UserPrincipalName
             $MailboxType = $_.RecipientTypeDetails
             Invoke-MailboxReport -MailboxUPN $MailboxUPN
             }
         }     
}     
Write-Progress -Activity "Processing mailboxes" -Completed

Disconnect-ExchangeOnline -Confirm:$false

if (Test-Path -Path $OutputCSV) {
    $ItemCounts = (Import-Csv -Path $OutputCSV).Count
    Write-Host "`nThe output file contains $($ProgressIndex) mailboxes and $($ItemCounts) records." -ForegroundColor Cyan
    Write-Host "Report available at: " -NoNewline -ForegroundColor Yellow
    Write-Host $OutputCSV
} else {
    Write-Host "No records found." -ForegroundColor Yellow
}
