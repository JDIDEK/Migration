#requires -Version 5.1
<#
.SYNOPSIS
    Prépare le dossier d'un profil itinérant et l'affecte à un utilisateur AD de test.
.PREREQUIS
    À lancer sur le serveur de fichiers. Modules ActiveDirectory et SmbShare,
    droits administrateur local et droit de modifier l'utilisateur dans l'AD.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$UtilisateurTest = 'migration.test',
    [string]$ServeurFichiers = 'SRV-FICHIERS-TEST',
    [string]$RacineLocaleProfils = 'D:\Profils-Test',
    [string]$NomPartageProfils = 'Profils-Test$',
    [switch]$DryRun
)

$CheminProfil = "\\$ServeurFichiers\$NomPartageProfils\$UtilisateurTest"

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution {
    if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }
    if ((Read-Host "Tapez OUI pour préparer et affecter le profil de '$UtilisateurTest'") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 }
}
function Invoke-Change {
    param([string]$Cible,[string]$Action,[scriptblock]$Commande)
    if ($DryRun) { Write-Log "DRYRUN - $Action sur $Cible"; return }
    if ($PSCmdlet.ShouldProcess($Cible,$Action)) { & $Commande }
}

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module SmbShare -ErrorAction Stop
$utilisateur = Get-ADUser -Identity $UtilisateurTest -Properties ProfilePath -ErrorAction Stop
$administrateurs = ([Security.Principal.SecurityIdentifier]'S-1-5-32-544').Translate([Security.Principal.NTAccount]).Value
$systeme = ([Security.Principal.SecurityIdentifier]'S-1-5-18').Translate([Security.Principal.NTAccount]).Value
$compteUtilisateur = $utilisateur.SID.Translate([Security.Principal.NTAccount]).Value
$dossierUtilisateur = Join-Path $RacineLocaleProfils $UtilisateurTest
$dossierUtilisateurNouveau = -not (Test-Path -LiteralPath $dossierUtilisateur -PathType Container)
Confirm-Execution

if (-not (Test-Path -LiteralPath $RacineLocaleProfils -PathType Container)) {
    Invoke-Change $RacineLocaleProfils 'Créer la racine des profils' { New-Item -ItemType Directory -Path $RacineLocaleProfils -Force | Out-Null }
}
$partage = Get-SmbShare -Name $NomPartageProfils -ErrorAction SilentlyContinue
if ($partage -and ([IO.Path]::GetFullPath($partage.Path) -ne [IO.Path]::GetFullPath($RacineLocaleProfils))) {
    throw "Le partage '$NomPartageProfils' pointe déjà vers '$($partage.Path)'."
}
if (-not $partage) {
    Invoke-Change $NomPartageProfils 'Créer le partage SMB des profils' {
        New-SmbShare -Name $NomPartageProfils -Path $RacineLocaleProfils -FullAccess $administrateurs -ChangeAccess $compteUtilisateur | Out-Null
    }
}
if (-not (Test-Simulation)) {
    $droitPartage = Get-SmbShareAccess -Name $NomPartageProfils -ErrorAction SilentlyContinue | Where-Object {
        $_.AccountName -eq $compteUtilisateur -and $_.AccessControlType -eq 'Allow' -and $_.AccessRight -in @('Change','Full')
    }
    if (-not $droitPartage) {
        Invoke-Change $NomPartageProfils "Accorder le droit de partage Modifier à $compteUtilisateur" {
            Grant-SmbShareAccess -Name $NomPartageProfils -AccountName $compteUtilisateur -AccessRight Change -Force | Out-Null
        }
    }
}
else { Write-Log "DRYRUN - Vérifier/ajouter le droit de partage Modifier pour $compteUtilisateur." }
if (-not (Test-Path -LiteralPath $dossierUtilisateur -PathType Container)) {
    Invoke-Change $dossierUtilisateur 'Créer le dossier du profil utilisateur' { New-Item -ItemType Directory -Path $dossierUtilisateur -Force | Out-Null }
}

if (-not (Test-Simulation)) {
    $acl = Get-Acl -LiteralPath $dossierUtilisateur
    if ($dossierUtilisateurNouveau -and -not $acl.AreAccessRulesProtected) {
        # Sur un nouveau dossier uniquement : retirer les droits hérités de la racine
        # afin que le profil ne soit accessible qu'aux trois identités explicites.
        $acl.SetAccessRuleProtection($true,$false)
    }
    $identites = @($compteUtilisateur, $administrateurs, $systeme)
    foreach ($identite in $identites) {
        $droit = 'FullControl'
        $droitEnum = [Security.AccessControl.FileSystemRights]::$droit
        $existe = $acl.Access | Where-Object {
            $_.IdentityReference.Value -eq $identite -and $_.AccessControlType -eq 'Allow' -and
            (($_.FileSystemRights -band $droitEnum) -eq $droitEnum)
        }
        if (-not $existe) {
            Invoke-Change $dossierUtilisateur "Accorder $droit à $identite" {
                $regle = New-Object Security.AccessControl.FileSystemAccessRule($identite,$droit,'ContainerInherit,ObjectInherit','None','Allow')
                $acl.AddAccessRule($regle) | Out-Null
                Set-Acl -LiteralPath $dossierUtilisateur -AclObject $acl
            }
        }
    }
    if (-not (Test-Path -LiteralPath $CheminProfil -PathType Container)) {
        throw "Le chemin UNC '$CheminProfil' n'est pas accessible. Le ProfilePath n'a pas été modifié."
    }
    Write-Log "Chemin UNC accessible : $CheminProfil" 'OK'
}
else {
    Write-Log "DRYRUN - Vérifier/ajouter les ACL sur '$dossierUtilisateur', puis tester '$CheminProfil'."
}

if ($utilisateur.ProfilePath -eq $CheminProfil) {
    Write-Log 'Le ProfilePath est déjà configuré avec la valeur attendue.' 'OK'
}
else {
    Invoke-Change $utilisateur.DistinguishedName "Définir ProfilePath sur $CheminProfil" {
        Set-ADUser -Identity $utilisateur -ProfilePath $CheminProfil
    }
}
Write-Log 'Configuration du profil itinérant terminée.' 'OK'
