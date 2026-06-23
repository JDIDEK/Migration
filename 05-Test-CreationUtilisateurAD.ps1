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
    [string]$OUTest = 'OU=Utilisateurs,OU=HLER,DC=intra,DC=ght53,DC=fr',
    [string]$ServeurAD = 'intra.ght53.fr',
    [PSCredential]$IdentifiantsAD,
    [SecureString]$MotDePasseTemporaire,
    [switch]$DryRun
)

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution { if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }; if ((Read-Host "Tapez OUI pour créer l'utilisateur '$Login'") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 } }

Import-Module ActiveDirectory -ErrorAction Stop
Write-Log "Serveur AD cible : $ServeurAD"
$identiteCourante = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$sessionIntra = ($identiteCourante -like 'INTRA\*' -or $identiteCourante -like "*@$DomaineUPN")
Write-Log "Session Windows : $identiteCourante"
if (-not $IdentifiantsAD -and -not (Test-Simulation) -and -not $sessionIntra) {
    $IdentifiantsAD = Get-Credential -Message "Compte admin INTRA autorisé à créer l'utilisateur, ex: INTRA\admin ou admin@$DomaineUPN"
}
$parametresAD = @{ Server = $ServeurAD; ErrorAction = 'Stop' }
if ($IdentifiantsAD) {
    $parametresAD.Credential = $IdentifiantsAD
    Write-Log "Compte AD utilisé : $($IdentifiantsAD.UserName)"
}
elseif ($sessionIntra) {
    Write-Log "Compte AD utilisé : session courante ($identiteCourante)" 'INFO'
}
elseif (Test-Simulation) {
    Write-Log 'Simulation : aucun compte admin AD ne sera demande.' 'INFO'
}
try { $null = Get-ADDomain @parametresAD }
catch { throw "Impossible de contacter Active Directory via '$ServeurAD'. Vérifiez DNS, réseau, RSAT/ADWS et droits. Détail : $($_.Exception.Message)" }
if (-not (Get-ADOrganizationalUnit -Identity $OUTest @parametresAD)) { throw "L'OU de test '$OUTest' n'existe pas. Créez ou corrigez cette OU avant de continuer." }
$loginFiltre = $Login.Replace("'", "''")
$existants = @(Get-ADUser -Filter "SamAccountName -eq '$loginFiltre'" -Properties UserPrincipalName @parametresAD)
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
    $parametresCreation = @{
        Name                  = "$Prenom $Nom"
        GivenName             = $Prenom
        Surname               = $Nom
        DisplayName           = "$Prenom $Nom"
        SamAccountName        = $Login
        UserPrincipalName     = $upn
        Path                  = $OUTest
        AccountPassword       = $MotDePasseTemporaire
        Enabled               = $true
        ChangePasswordAtLogon = $true
        Description           = 'Compte de test migration'
    }
    foreach ($cle in $parametresAD.Keys) { $parametresCreation[$cle] = $parametresAD[$cle] }
    $parametresCreation.PassThru = $true
    $nouveau = New-ADUser @parametresCreation
    $cree = Get-ADUser -Identity $nouveau.ObjectGUID -Properties UserPrincipalName,DistinguishedName,ObjectGUID @parametresAD
    $dansOuCible = @(Get-ADUser -SearchBase $OUTest -Filter "SamAccountName -eq '$loginFiltre'" -Properties UserPrincipalName,DistinguishedName @parametresAD)
    if ($dansOuCible.Count -ne 1) {
        throw "L'utilisateur a été créé ou trouvé par GUID, mais il n'est pas retrouvé de façon unique dans l'OU cible '$OUTest'. DN actuel : $($cree.DistinguishedName)"
    }
    Write-Log "Utilisateur créé : $($cree.UserPrincipalName)" 'OK'
    Write-Log "Emplacement AD : $($cree.DistinguishedName)" 'OK'
    Write-Log "ObjectGUID : $($cree.ObjectGUID)" 'OK'
    Write-Log "Contrôlez dans ADUC sur le domaine/serveur '$ServeurAD' et actualisez l'OU cible." 'INFO'
}
