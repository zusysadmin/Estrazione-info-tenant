<#
    Date:         07/04/2025
    Author:       Alberti Fabrizio (fabrizio.alberti@zucchetti.it)
    Description:  Estrazione dati Sharepoint e OneDrive
	Version:      1.0
#>

Param(
        [Parameter(Mandatory=$true,HelpMessage="Tenant Sharepoint Admin URL (es. https://zucchetti-admin.sharepoint.com")]
        [String]
        $SharepointAdminUrl
    )

Connect-SPOService -url $SharepointAdminUrl
$siti = Get-SPOSite -IncludePersonalSite $true -Limit All | Select Title, Url, Owner, StorageUsageCurrent

$FilePath = "$PSScriptRoot\DatiSharepointOnedrive.xlsx"

$SharePointExcel = $siti |  Where-Object { $_.Url -notlike  "*-my.sharepoint.com/personal/*" } | Sort Title |  Export-Excel -Path $FilePath -WorksheetName "Sharepoint" -AutoSize -ClearSheet -FreezeTopRow

$OneDriveExcel = $siti |  Where-Object { $_.Url -like  "*-my.sharepoint.com/personal/*" } | Sort Title | Export-Excel -Path $FilePath -WorksheetName "OneDrive" -AutoSize -Append -MoveToEnd -FreezeTopRow -Show