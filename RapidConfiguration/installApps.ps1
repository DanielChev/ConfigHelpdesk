<#
.SYNOPSIS

Installer et configurer en moins de gestes un ordinateur de prêt

.DESCRIPTION 

Traîter le contenu du fichier Json et effectuer des modifications sur la machine

#>

function getSelection{
    Write-Host "Daniel Chevalley - Stagiaire Helpdesk Cheseaux - Version 0.1" -ForegroundColor Green
    Write-Host "Ce script a besoin d'une connexion internet pour installer les applications" -ForeGroundColor Cyan
    Write-Host "!!! DECONNECTER L'UTILISATEUR HELPDESK AVANT LE SCRIPT SINON LA SUPPRESSION DU FOLDER N'EST PAS POSSIBLE !!!" -ForegroundColor Red
    Write-Host "!!! EFFECTUER LES MISES à JOUR WINDOWS POUR UNE INSTALLATION CORRECTE D'OFFICE !!!" -ForegroundColor Red
    Write-Host "Bienvenue dans l'utilitaire de configuration d'ordinateur de prêt, sélectionnez parmis les actions suivantes
        1) Installer et mettre à jour les applications pour un Lenovo
        2) Installer et mettre à jour les applications pour un Dell
        3) Installer uniquement les applications sans Vantage ou DellCMD
        4) Installer uniquement le pack Office LTSC (Désinstaller les vieilles versions si déjà présentes)
        5) Effectuer les mises à jour des applications
        6) Reset le compte Helpdesk et supprimer les comptes obsolètes uniquement
        7) Lancer une réinitialisation de l'ordinateur, si la configuration de l'ordinateur est trop spécialisée"
    Write-Host "Sélectionnez l'option en tapant le numéro associé à l'action" -ForegroundColor Green
    $Selection = Read-Host
    while(($Selection -ne 1) -and ($Selection -ne 2) -and ($Selection -ne 3) -and ($Selection -ne 4) -and ($Selection -ne 5) -and ($Selection -ne 5) -and ($Selection -ne 6) -and ($Selection -ne 7) -and ($Selection -ne 0)){
        Write-Host "Input non autorisée, choisir uniquement par 1,2,3,4,5,6,7"
        $Selection = Read-Host
    }
    return $Selection
}

#Requires -RunAsAdministrator
#Start Transcript logging in Temp folder
Start-Transcript $ENV:TEMP\install.log

#Set-Executionpolicy and no prompting
Set-ExecutionPolicy Bypass -Force:$True -Confirm:$false -ErrorAction SilentlyContinue
Set-Variable -Name 'ConfirmPreference' -Value 'None' -Scope Global

#For faster downloads
$ProgressPreference = 'SilentlyContinue'

#Import list of apps, features and modules that can be installed using json file
$json = Get-Content "$($PSScriptRoot)\installApps.json" | ConvertFrom-Json

$Selection = getSelection

while($Selection -ne 0){

    #Reset de l'ordinateur
    if ($Selection -eq 7){
    Write-Host "!!! Si le programme est sur clé USB, il est préférable de copier le dossier sur le bureau et d'enlever la clé !!!" -ForeGroundColor Red
    Write-Host "Vous êtes sur le point de lancer un factory reset, êtes vous sûr ? (Y/N)"
    $Confirm = Read-Host
        if(($Confirm -ne "Y") -and ($Confirm -ne "N") -and ($Confirm -ne "y") -and ($Confirm -ne "n")){
            while(($Confirm -ne "Y") -and ($Confirm -ne "N") -and ($Confirm -ne "y") -and ($Confirm -ne "n")){
                Write-Host "Input non autorisée, choisir uniquement par y ou n"
                $Confirm = Read-Host
            }
        }

        if(($Confirm -eq "Y") -or ($Confirm -eq "y")){
            systemreset -factoryreset
            $Selection = 0
        }else{
            $Selection = getSelection
        }
    }

        #Installation de WinGet
    if (!(Get-AppxPackage -Name Microsoft.Winget.Source)) {
        Write-Host "Winget n'est pas installé, installation en cours" -ForeGroundColor Yellow
        Invoke-Webrequest -uri https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx -Outfile $ENV:TEMP\Microsoft.VCLibs.x64.14.00.Desktop.appx
        Invoke-Webrequest -uri https://aka.ms/getwinget -Outfile $ENV:TEMP\winget.msixbundle    
        Add-AppxPackage $ENV:TEMP\Microsoft.VCLibs.x64.14.00.Desktop.appx -ErrorAction SilentlyContinue
        Add-AppxPackage -Path $ENV:TEMP\winget.msixbundle -ErrorAction SilentlyContinue
        winget.exe source reset --force
        winget.exe source update
    }else{
        winget.exe source reset --force
        winget.exe source update
    }

    #Management des utilisateurs
    if($Selection -eq 6){
        Write-Host "Tentative de suppression des dossiers d'utilisateurs indésirables" -ForegroundColor Cyan
        $ErrorActionPreference= 'silentlycontinue'
        $Users = Get-WmiObject -Class Win32_UserProfile
        get-childitem -Path "C:\users" -Exclude "Supadm", "Default", "Public", "supadm" -Force -Directory| Remove-Item -Force -Recurse
        Remove-LocalUser -Name "Helpdesk"
        New-LocalUser -Name "Helpdesk" -Description "Compte déstiné à l'utilisateur" -NoPassword -UserMayNotChangePassword -AccountNeverExpires | Set-LocalUser -PasswordNeverExpires $true
        Write-Host "Helpdesk compte local ajouté"
        Add-LocalGroupMember -SID "S-1-5-32-544" -Member "Helpdesk"
        Write-Host "Helpdesk ajouté au groupe Administrateurs"
        $Selection = 0
    }

    #Installer le pack office
    if($Selection -eq 4){
        Write-Host Tentative de désinstallation du pack Office avant réinstallation LTSC -ForeGroundColor Cyan
        # Exécuter la commande winget pour obtenir la liste des packages Microsoft Office
        $officePackages = winget list --id Microsoft.Office --accept-source-agreements 
        # Parcourir chaque ligne de sortie pour extraire l'ID du package et le désinstaller
        foreach ($line in $officePackages) {
            $parts = $line -split '\s+'
            foreach ($part in $parts) {
                if ($part -match '^Microsoft\.Office\.(.*?)$') {
                    $packageId = $matches[0].Trim()
                    Write-Host "Désinstallation de $packageId ..."
                    winget.exe uninstall --id $packageId --silent --force --accept-source-agreements
                }
            }
        }
        Foreach ($App in $json.Office) {
            Write-Host Checking if $App is already installed...
            winget.exe list --id $App --accept-source-agreements | Out-Null
            if ($LASTEXITCODE -eq '-1978335212') {
                Write-Host $App.Split('.')[1] was not found and installing now -ForegroundColor Yellow
                winget.exe install $App --silent --force --accept-package-agreements --accept-source-agreements
            }
        }
            $Selection = 0                        
    }

    #Effectuer uniquement les mises à jour
    if($Selection -eq 5){
        winget.exe upgrade --all --force --silent --include-unknown --accept-package-agreements --accept-source-agreements
        $Selection = 0
    }

    #Installer l'ensemble des applications et manager les comptes
    if (($Selection -eq 1) -or ($Selection -eq 2) -or ($Selection -eq 3)) {
        Write-Host Installing Applications but skipping install if already present -ForegroundColor Green
        Write-Host Tentative de désinstallation du pack Office avant réinstallation LTSC -ForeGroundColor Cyan
        $officePackages = winget.exe list --id Microsoft.Office --accept-source-agreements

        foreach ($line in $officePackages) {
            $parts = $line -split '\s+'
            foreach ($part in $parts) {
                if ($part -match '^Microsoft\.Office\.(.*?)$') {
                    $packageId = $matches[0].Trim()
                    Write-Host "Désinstallation de $packageId ..."
                    winget.exe uninstall --id $packageId --silent --force --accept-source-agreements
                }
            }
        }

        Foreach ($App in $json.Apps) {
            Write-Host Checking if $App is already installed...
            winget.exe list --id $App --accept-source-agreements | Out-Null
            if ($LASTEXITCODE -eq '-1978335212') {
                Write-Host $App.Split('.')[1] was not found and installing now -ForegroundColor Yellow
                winget.exe install $App --silent --force --accept-package-agreements --accept-source-agreements
            }
        }
        if($Selection -eq 1){
            Foreach ($App in $json.Lenovo) {
            Write-Host Checking if $App is already installed...
            winget.exe list --id $App --accept-source-agreements | Out-Null
                if ($LASTEXITCODE -eq '-1978335212') {
                    Write-Host $App.Split('.')[1] was not found and installing now -ForegroundColor Yellow
                    winget.exe install $App --silent --force --accept-package-agreements --accept-source-agreements
                }
            }
        }
        if($Selection -eq 2){
            Foreach ($App in $json.Dell) {
            Write-Host Checking if $App is already installed...
            winget.exe list --id $App --accept-source-agreements | Out-Null
                if ($LASTEXITCODE -eq '-1978335212') {
                    Write-Host $App.Split('.')[1] was not found and installing now -ForegroundColor Yellow
                    winget.exe install $App --silent --force --accept-package-agreements --accept-source-agreements
                }
            }
        }
        winget.exe upgrade --all --force --silent --accept-package-agreements --accept-source-agreements

        #Clean-up downloaded Winget Packages
        Remove-Item $ENV:TEMP\Winget -Recurse -Force:$True -ErrorAction:SilentlyContinue

        Write-Host "Tentative de suppression des dossiers d'utilisateurs indésirables" -ForegroundColor Cyan
        $ErrorActionPreference= 'silentlycontinue'
        $Users = Get-WmiObject -Class Win32_UserProfile
        get-childitem -Path "C:\users" -Exclude "Supadm", "Default", "Public", "supadm" -Force -Directory| Remove-Item -Force -Recurse
        Remove-LocalUser -Name "Helpdesk"
        New-LocalUser -Name "Helpdesk" -Description "Compte déstiné à l'utilisateur" -NoPassword -UserMayNotChangePassword -AccountNeverExpires | Set-LocalUser -PasswordNeverExpires $true
        Write-Host "Helpdesk compte local ajouté"
        Add-LocalGroupMember -SID "S-1-5-32-544" -Member "Helpdesk"
        Write-Host "Helpdesk ajouté au groupe Administrateurs"
        $Selection = 0    
    }
}

#Stop Transcript logging
Stop-Transcript