#requires -Version 5.1
<#
.SYNOPSIS
    Crée un dossier et un partage SMB de test avec des droits NTFS de modification.
.PREREQUIS
    Rôle Serveur de fichiers, module SmbShare, volume D:, droits administrateur
    local et groupe AD de test déjà existant.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$CheminLocal = 'D:\Partages\Test',
    [string]$NomPartage = 'Migration-Test',
    [string]$DescriptionPartage = 'Partage de validation migration - TEST UNIQUEMENT',
    [string]$GroupeAD = 'INTRA\GG_Migration_Test_RW',
    [switch]$DryRun
)

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'OK', 'ATTENTION', 'ERREUR')][string]$Niveau = 'INFO')
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Niveau, $Message)
}
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution {
    if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }
    if ((Read-Host "Tapez OUI pour créer/configurer le dossier et le partage '$NomPartage'") -cne 'OUI') {
        Write-Log 'Opération annulée.' 'ATTENTION'; exit 0
    }
}
function Invoke-Change {
    param([string]$Cible, [string]$Action, [scriptblock]$Commande)
    if ($DryRun) { Write-Log "DRYRUN - $Action sur $Cible"; return }
    if ($PSCmdlet.ShouldProcess($Cible, $Action)) { & $Commande }
}

Import-Module SmbShare -ErrorAction Stop
try { $null = ([Security.Principal.NTAccount]$GroupeAD).Translate([Security.Principal.SecurityIdentifier]) }
catch { throw "Le groupe '$GroupeAD' est introuvable ou non résolu depuis ce serveur." }

$administrateurs = ([Security.Principal.SecurityIdentifier]'S-1-5-32-544').Translate([Security.Principal.NTAccount]).Value
Confirm-Execution

if (-not (Test-Path -LiteralPath $CheminLocal -PathType Container)) {
    Invoke-Change $CheminLocal 'Créer le dossier' { New-Item -ItemType Directory -Path $CheminLocal -Force | Out-Null }
}
else { Write-Log "Dossier déjà présent : $CheminLocal" 'OK' }

$partage = Get-SmbShare -Name $NomPartage -ErrorAction SilentlyContinue
if ($partage -and ([IO.Path]::GetFullPath($partage.Path) -ne [IO.Path]::GetFullPath($CheminLocal))) {
    throw "Le partage '$NomPartage' existe déjà sur '$($partage.Path)'. Aucun changement n'a été fait sur ce partage."
}
if (-not $partage) {
    Invoke-Change $NomPartage 'Créer le partage SMB' {
        New-SmbShare -Name $NomPartage -Path $CheminLocal -Description $DescriptionPartage `
            -FullAccess $administrateurs -ChangeAccess $GroupeAD | Out-Null
    }
}
else { Write-Log "Partage SMB déjà présent sur le bon chemin." 'OK' }

if (-not (Test-Simulation) -and (Test-Path -LiteralPath $CheminLocal)) {
    $acl = Get-Acl -LiteralPath $CheminLocal
    $droitsSuffisants = $acl.Access | Where-Object {
        $_.IdentityReference.Value -eq $GroupeAD -and
        $_.AccessControlType -eq [Security.AccessControl.AccessControlType]::Allow -and
        (($_.FileSystemRights -band [Security.AccessControl.FileSystemRights]::Modify) -eq [Security.AccessControl.FileSystemRights]::Modify) -and
        (($_.InheritanceFlags -band [Security.AccessControl.InheritanceFlags]::ContainerInherit) -ne 0) -and
        (($_.InheritanceFlags -band [Security.AccessControl.InheritanceFlags]::ObjectInherit) -ne 0)
    }
    if (-not $droitsSuffisants) {
        Invoke-Change $CheminLocal "Accorder les droits NTFS Modifier à $GroupeAD" {
            $regle = New-Object Security.AccessControl.FileSystemAccessRule(
                $GroupeAD, 'Modify', 'ContainerInherit,ObjectInherit', 'None', 'Allow'
            )
            $acl.AddAccessRule($regle) | Out-Null
            Set-Acl -LiteralPath $CheminLocal -AclObject $acl
        }
    }
    else { Write-Log 'Droits NTFS déjà présents.' 'OK' }

    $droitPartage = Get-SmbShareAccess -Name $NomPartage -ErrorAction SilentlyContinue | Where-Object {
        $_.AccountName -eq $GroupeAD -and $_.AccessControlType -eq 'Allow' -and $_.AccessRight -in @('Change', 'Full')
    }
    if (-not $droitPartage) {
        Invoke-Change $NomPartage "Accorder le droit de partage Modifier à $GroupeAD" {
            Grant-SmbShareAccess -Name $NomPartage -AccountName $GroupeAD -AccessRight Change -Force | Out-Null
        }
    }
}
elseif (Test-Simulation) {
    Write-Log "DRYRUN - Vérifier/ajouter les droits NTFS Modifier et les droits de partage pour $GroupeAD."
}

Write-Log 'Contrôle du partage terminé.' 'OK'
