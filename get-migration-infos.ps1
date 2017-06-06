Import-Module ActiveDirectory

("userPrincipalName,OPCO,Country,Enabled,Mailbox,ServerName,LitigationHoldEnabled,SingleItemRecoveryEnabled")  | Out-File -FilePath ("C:\scripts\MigrationReport\Output\MigrationReport.csv")
$AllADAccounts = Get-ADUser -Filter * -SearchBase "OU=Production,DC=contoso,DC=com" -Properties UserPrincipalName,SamAccountName,physicalDeliveryOfficeName,co,Enabled,sIDHistory,lastLogonTimestamp -Server domaincontroller.contoso.com
ForEach ($Account in $AllADAccounts){
If ($Account.Enabled -eq $true) { $Enabled = "Yes"} Else { $Enabled = "No" }

$OnPremMailbox = Get-Mailbox -Identity $Account.SamAccountName -ErrorAction SilentlyContinue | Select LitigationHoldEnabled,SingleItemRecoveryEnabled,ServerName
If($OnPremMailbox -ne $null){
    $CheckOnDuplicate = $OnPremMailbox | measure
    If($CheckOnDuplicate.Count -eq 1){
        $MailboxExists = "Yes"
        $LitigationHoldEnabled = $OnPremMailbox.LitigationHoldEnabled
        $SingleItemRecoveryEnabled = $OnPremMailbox.SingleItemRecoveryEnabled
        $AuditEnabled = $OnPremMailbox.AuditEnabled
        $ServerName = $OnPremMailbox.ServerName
    }
    Else { }
}
Else {
$CloudMailbox = Get-RemoteMailbox -Identity $Account.SamAccountName -ErrorAction SilentlyContinue | Select LitigationHoldEnabled,SingleItemRecoveryEnabled
    If($CloudMailbox -ne $null){
    $MailboxExists = "Yes"
    $LitigationHoldEnabled = $CloudMailbox.LitigationHoldEnabled
    $SingleItemRecoveryEnabled = $CloudMailbox.SingleItemRecoveryEnabled
    $AuditEnabled = $CloudMailbox.AuditEnabled
    $ServerName = "Office 365"
    }
    Else { 
    $MailboxExists = "No"
    $LitigationHoldEnabled = ""
    $SingleItemRecoveryEnabled = ""
    $AuditEnabled = ""
    $ServerName = ""
    }
}

$OPCO = $Account.physicalDeliveryOfficeName

# Corrections in the country field
$Country = $Account.co

($Account.UserPrincipalName + "," + $OPCO + "," + $Country + "," + $MailboxExists + "," + $ServerName + "," + $LitigationHoldEnabled + "," + $SingleItemRecoveryEnabled)  | Out-File -FilePath ("C:\scripts\MigrationReport\Output\MigrationReport.csv") -Append


# Getting history trend data
$Timestamp = Get-Date -Format d
$Report = Import-Csv "C:\scripts\MigrationReport\Output\MigrationReport.csv"

$MigratedAccounts = $Report | ?{$_.Migrated -eq "Yes"} | measure
$MigratedMailboxes = $Report | ?{$_.Migrated -eq "Yes" -and $_.Mailbox -eq "Yes"} | measure
$MigratedEmployees = $Report | ?{$_.Migrated -eq "Yes" -and $_.Mailbox -eq "Yes" -and $_.LoggedIn -eq "Yes"} | measure
$MailboxesOnPrem = $Report | ?{$_.ServerName -ne "Office 365" -and $_.ServerName -ne ""} | measure
$MailboxesCloud = $Report | ?{$_.ServerName -eq "Office 365"} | measure
$AccountsReady = $Report | ?{$_.Migrated -eq "Yes" -and $_.Mailbox -eq "No"} | measure
$NewMailboxes = $Report | ?{$_.Migrated -eq "No" -and $_.Mailbox -eq "Yes"} | measure
$NewUsers = $Report | ?{$_.Migrated -eq "No" -and $_.Mailbox -eq "Yes" -and $_.LoggedIn -eq "Yes"} | measure
$TotalAccounts = $Report | measure
$TotalMailboxes = $Report | ?{$_.ServerName -ne ""} | measure

# Get total amount of mailbox data
$TotalOnPremMbxSize = Get-MailboxDatabase -Status | sort name | select name,@{Name='DB Size (Gb)';Expression={$_.DatabaseSize.ToGb()}} | measure "DB Size (Gb)" -sum

# Write values to CSV
($Timestamp + "," + $MigratedAccounts.Count + "," + $MigratedMailboxes.Count + "," + $MigratedEmployees.Count + "," + $MailboxesOnPrem.Count + "," +  + $MailboxesCloud.Count + "," + $AccountsReady.Count + "," + $NewMailboxes.Count + "," + $NewUsers.Count + "," + $TotalAccounts.Count + "," + $TotalMailboxes.Count + "," + $TotalOnPremMbxSize.Sum/1000) | Out-File -FilePath ("C:\scripts\MigrationReport\Output\MigrationReportTrends.csv") -Append

("TotalAccounts,TotalMailboxes,TotalOnPremMbxSize") | Out-File -FilePath ("C:\scripts\MigrationReport\Output\MigrationTotals.csv")
([string]$TotalAccounts.Count + "," + [string]$TotalMailboxes.Count + "," + [string]$TotalOnPremMbxSize.Sum/1000) | Out-File -FilePath ("C:\scripts\MigrationReport\Output\MigrationTotals.csv") -Append

