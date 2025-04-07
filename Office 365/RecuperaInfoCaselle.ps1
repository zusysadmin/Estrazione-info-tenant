<#
    Date:         07/04/2025
    Author:       Alberti Fabrizio (fabrizio.alberti@zucchetti.it)
    Description:  Estrazione dati Exchange OnLine
	Version:      1.0
#>


$Global:ExOnlineSession=Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange"}

if ($Global:ExOnlineSession){
    return $Global:ExOnlineSession
}else{
    try{
	    Connect-ExchangeOnline
    }catch{
	    # Già connesso
    }
}

# Variabili
$MbxReport = @()
$i = 0

$Mbxs = Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited | Select userPrincipalName | Sort userPrincipalName
$totMbx = ($Mbxs).Count

# Per ogni casella
foreach ($userMbx in $Mbxs)
{
$nrcaselle = $nrcaselle + 1

$userMbxMailInfo = Get-Mailbox -Identity $userMbx.UserPrincipalName | Select DisplayName, PrimarySmtpAddress, RecipientType, ArchiveDatabase

$userMbxMailSize= Get-Mailbox $userMbx.UserPrincipalName | Get-MailboxStatistics | Select-Object @{
    Name = 'MailboxSizeMB'
    Expression = {
        $raw = $_.TotalItemSize.ToString()
        if ($raw -match '\(([\d,]+) bytes\)') {
            $bytes = $matches[1] -replace ',', ''
            [math]::Round([double]$bytes / 1MB, 2)
        } else {
            0
        }
    }
} 


if ($userMbxMailInfo.ArchiveDatabase -ne $null) {
    $userMbxArchiveSize = Get-Mailbox $userMbx.UserPrincipalName | Get-MailboxStatistics -Archive | Select-Object @{
    Name = 'ArchiveMailboxSizeMB'
    Expression = {
        $raw = $_.TotalItemSize.ToString()
        if ($raw -match '\(([\d,]+) bytes\)') {
            $bytes = $matches[1] -replace ',', ''
            [math]::Round([double]$bytes / 1MB, 2)
        } else {
            0
        }
     }
  } 
} 


$reportUserMbx = New-Object PSObject
$reportUserMbx | Add-Member NoteProperty -Name "Nome" -Value $userMbxMailInfo.DisplayName
$reportUserMbx | Add-Member NoteProperty -Name "Indirizzo Primario" -Value $userMbxMailInfo.PrimarySmtpAddress
$reportUserMbx | Add-Member NoteProperty -Name "Dimensione Posta (MB)" -Value $userMbxMailSize.MailboxSizeMB
$reportUserMbx | Add-Member NoteProperty -Name "Dimensione Archivio (MB)" -Value $userMbxArchiveSize.ArchiveMailboxSizeMB
$reportUserMbx | Add-Member NoteProperty -Name "Tipo Casella" -Value $userMbxMailInfo.RecipientType

[array]$reportMBX += $reportUserMbx

  
[int]$i = $nrcaselle / ($Mbxs).Count * 100
Write-Progress -Activity "Elaborazione della casella di posta $nr di $totMbx in corso..." -Status "$i% completato" -PercentComplete $i
Start-Sleep -Milliseconds 100

}


# Estrazione a Video Caselle
# Write-Output $reportMBX | Out-GridView 

# Estrazione su Excel

$FilePath = "$PSScriptRoot\DatiExchangeOnline.xlsx"

# Primo foglio UserMailbox
$UserMBXExcel = $reportMBX | Where-Object { $_."Tipo Casella" -eq "UserMAilbox" } | Export-Excel -Path $FilePath -WorksheetName "Caselle Utente" -AutoSize -ClearSheet -FreezeTopRow


# Secondo foglio ShareMailbox
$SharedMBXExcel = $reportMBX | Where-Object { $_."Tipo Casella" -eq "SharedMailbox" } | Export-Excel -Path $FilePath -WorksheetName "Caselle Condivise" -AutoSize -Append -MoveToEnd -FreezeTopRow
#>

# Distribution List
$DLs = Get-DistributionGroup -ResultSize Unlimited | Where-Object { $_."GroupType" -eq "Universal" } | Select DisplayName, Alias, PrimarySMTPAddress | Sort DisplayName
$totDL = $DLs.Count

# Variabili
$DLReport = @()
$pdl = 0

foreach ($DL in $DLs)

{
$nrdl = $nrdl + 1
$DLInfo = Get-DistributionGroup -Identity $DL.PrimarySMTPAddress | Select DisplayName, Alias, PrimarySMTPAddress | Sort DisplayName


$reportDL = New-Object PSObject
$reportDL | Add-Member NoteProperty -Name "Nome" -Value $DLInfo.DisplayName
$reportDL | Add-Member NoteProperty -Name "Indirizzo Primario" -Value $DLInfo.PrimarySmtpAddress
$reportDL | Add-Member NoteProperty -Name "Alias" -Value $DLInfo.Alias

[array]$reportDLs += $reportDL


[int]$pdl = $nrdl / ($DLs).Count * 100
Write-Progress -Activity "Elaborazione della DL $nr di $totDL in corso..." -Status "$pdl% completato" -PercentComplete $pdl
Start-Sleep -Milliseconds 100

}


# Estrazione Excel
$DLExcel = $reportDL | Export-Excel -Path $FilePath -WorksheetName "Distribution List" -AutoSize -Append -MoveToEnd -Show

Disconnect-ExchangeOnline -Confirm:$false