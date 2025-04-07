<#
    Date:         07/04/2025
    Author:       Alberti Fabrizio (fabrizio.alberti@zucchetti.it)
    Description:  Estrazione dati Teams
	Version:      1.0
#>

Connect-MicrosoftTeams
$teams = Get-Team | Select DisplayName, Description

$FilePath = "$PSScriptRoot\DatiTeams.xlsx"

$teams | Sort Title | Export-Excel -Path $FilePath -WorksheetName "Teams" -AutoSize -ClearSheet -MoveToEnd -FreezeTopRow -Show