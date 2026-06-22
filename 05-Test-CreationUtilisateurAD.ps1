#requires -Version 5.1
<#
.SYNOPSIS
    Crée un unique utilisateur AD de test dans une OU existante.
.PREREQUIS
    Module ActiveDirectory (RSAT), OU de test existante et délégation permettant
    de créer un utilisateur. Le mot de passe n'est jamais stocké en clair.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$Prenom = 'Jean',
    [string]$Nom = 'Migration',
    [string]$Login = 'migration.test',
    [string]$DomaineUPN = 'intra.ght53.fr',
    [string]$OUTest = 'OU=Utilisateurs-Test,DC=intra,DC=ght53,DC=fr',
    [SecureString]$MotDePasseTemporaire,
    [switch]$DryRun
)

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution { if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }; if ((Read-Host "Tapez OUI pour créer l'utilisateur '$Login'") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 } }

Import-Module ActiveDirectory -ErrorAction Stop
if (-not (Get-ADOrganizationalUnit -Identity $OUTest -ErrorAction SilentlyContinue)) { throw "L'OU de test '$OUTest' n'existe pas. Créez ou corrigez cette OU avant de continuer." }
$loginFiltre = $Login.Replace("'", "''")
$existants = @(Get-ADUser -Filter "SamAccountName -eq '$loginFiltre'" -Properties UserPrincipalName)
if ($existants.Count -gt 1) { throw "Plusieurs objets portent le SamAccountName '$Login'. Intervention manuelle requise." }
if ($existants.Count -eq 1) {
    Write-Log "L'utilisateur existe déjà : $($existants[0].DistinguishedName). Aucune création." 'OK'
    exit 0
}

$upn = "$Login@$DomaineUPN"
Confirm-Execution
if (Test-Simulation) {
    Write-Log "SIMULATION - Créer '$Prenom $Nom' ($Login, $upn) dans '$OUTest', compte activé, changement de mot de passe obligatoire."
    exit 0
}
if (-not $MotDePasseTemporaire) { $MotDePasseTemporaire = Read-Host 'Mot de passe temporaire' -AsSecureString }
if ($PSCmdlet.ShouldProcess($upn, "Créer l'utilisateur AD dans $OUTest")) {
    New-ADUser -Name "$Prenom $Nom" -GivenName $Prenom -Surname $Nom -DisplayName "$Prenom $Nom" `
        -SamAccountName $Login -UserPrincipalName $upn -Path $OUTest -AccountPassword $MotDePasseTemporaire `
        -Enabled $true -ChangePasswordAtLogon $true -Description 'Compte de test migration' -ErrorAction Stop
    Write-Log "Utilisateur créé : $upn" 'OK'
}
