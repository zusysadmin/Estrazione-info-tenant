<#
    Date:         07/04/2025
    Author:       Alberti Fabrizio (fabrizio.alberti@zucchetti.it)
    Description:  Setup installazione moduli Powershell Exchange Online, Sharepoint, Teams
	Version:      1.0
#>

#Requires -RunAsAdministrator

Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

$Moduli = @("ExchangeOnlineManagement", "Microsoft.Online.SharePoint.PowerShell", "MicrosoftTeams")

foreach ($Modulo in $Moduli)

{

    if (-not (Get-InstalledModule -Name $Modulo -ErrorAction SilentlyContinue)) {
        Write-Output "Modulo '$Modulo' non trovato. Procedo con l'installazione..."
        try {
            Install-Module -Name $Modulo -Force -Scope CurrentUser -AllowClobber
            Write-Output "Modulo $Modulo installato con successo."
        } catch {
            Write-Error "Errore durante l'installazione del modulo $Modulo"
        }
    } else {
        Write-Output "Modulo '$Modulo' già installato. Controllo aggiornamenti..."
        try {
            Update-Module -Name $Modulo -Force
            Write-Output "Modulo $Modulo aggiornato all'ultima versione."
        } catch {
            Write-Error "Errore durante l'aggiornamento del modulo $Modulo"
        }
    }

}