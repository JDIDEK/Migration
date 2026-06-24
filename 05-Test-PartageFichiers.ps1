#requires -Version 5.1
<#
.SYNOPSIS
    Crée un dossier et un partage SMB de test avec des droits NTFS de modification.
.PREREQUIS
    Rôle Serveur de fichiers, WinRM actif sur le serveur distant, module SmbShare,
    droits administrateur sur le serveur et groupe AD de test déjà existant.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$ServeurFichiers = 'SRV-FICHIERS-TEST',
    [string]$CheminLocal = 'D:\Partages\Test',
    [string]$NomPartage = 'Migration-Test',
    [string]$DescriptionPartage = 'Partage de validation migration - TEST UNIQUEMENT',
    [string]$GroupeAD = 'INTRA\GG_Migration_Test_RW',
    [PSCredential]$IdentifiantsServeur,
    [string]$IdentifiantsServeurPath = '',
    [switch]$DryRun
)

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'OK', 'ATTENTION', 'ERREUR')][string]$Niveau = 'INFO')
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Niveau, $Message)
}
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution {
    if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }
    if ((Read-Host "Tapez OUI pour créer/configurer le partage '$NomPartage' sur '$ServeurFichiers'") -cne 'OUI') {
        Write-Log 'Opération annulée.' 'ATTENTION'; exit 0
    }
}

if (-not $IdentifiantsServeur -and $IdentifiantsServeurPath -and (Test-Path -LiteralPath $IdentifiantsServeurPath -PathType Leaf)) {
    $IdentifiantsServeur = Import-Clixml -LiteralPath $IdentifiantsServeurPath
}

$nomServeurCourt = (($ServeurFichiers -split '\.')[0]).ToUpperInvariant()
$executionLocale = $nomServeurCourt -in @('.', 'LOCALHOST', $env:COMPUTERNAME.ToUpperInvariant())
if (-not $executionLocale -and -not $IdentifiantsServeur) {
    $IdentifiantsServeur = Get-Credential -Message "Compte administrateur sur le serveur $ServeurFichiers"
}

Confirm-Execution

$configurationPartage = {
    param(
        [string]$CheminLocal,
        [string]$NomPartage,
        [string]$DescriptionPartage,
        [string]$GroupeAD,
        [bool]$Simulation
    )

    $ErrorActionPreference = 'Stop'
    function New-Resultat {
        param([string]$Niveau, [string]$Message)
        [PSCustomObject]@{
            Serveur = $env:COMPUTERNAME
            Niveau  = $Niveau
            Message = $Message
        }
    }

    Import-Module SmbShare -ErrorAction Stop
    try {
        $null = ([Security.Principal.NTAccount]$GroupeAD).Translate([Security.Principal.SecurityIdentifier])
    }
    catch {
        throw "Le groupe '$GroupeAD' est introuvable ou non résolu depuis $env:COMPUTERNAME."
    }

    $administrateurs = ([Security.Principal.SecurityIdentifier]'S-1-5-32-544').Translate(
        [Security.Principal.NTAccount]
    ).Value

    if ($Simulation) {
        New-Resultat 'INFO' "DRYRUN - Créer/vérifier '$CheminLocal'."
        New-Resultat 'INFO' "DRYRUN - Créer/vérifier le partage '\\$env:COMPUTERNAME\$NomPartage'."
        New-Resultat 'INFO' "DRYRUN - Accorder Modifier à '$GroupeAD' sur le partage et le dossier."
        return
    }

    if (-not (Test-Path -LiteralPath $CheminLocal -PathType Container)) {
        New-Item -ItemType Directory -Path $CheminLocal -Force | Out-Null
        New-Resultat 'OK' "Dossier créé : $CheminLocal"
    }
    else {
        New-Resultat 'OK' "Dossier déjà présent : $CheminLocal"
    }

    $partage = Get-SmbShare -Name $NomPartage -ErrorAction SilentlyContinue
    if ($partage -and ([IO.Path]::GetFullPath($partage.Path) -ne [IO.Path]::GetFullPath($CheminLocal))) {
        throw "Le partage '$NomPartage' existe déjà sur '$($partage.Path)'."
    }
    if (-not $partage) {
        New-SmbShare -Name $NomPartage -Path $CheminLocal -Description $DescriptionPartage `
            -FullAccess $administrateurs -ChangeAccess $GroupeAD -ErrorAction Stop | Out-Null
        New-Resultat 'OK' "Partage créé : \\$env:COMPUTERNAME\$NomPartage"
    }
    else {
        New-Resultat 'OK' "Partage déjà présent : \\$env:COMPUTERNAME\$NomPartage"
    }

    $acl = Get-Acl -LiteralPath $CheminLocal
    $droitsSuffisants = $acl.Access | Where-Object {
        $_.IdentityReference.Value -eq $GroupeAD -and
        $_.AccessControlType -eq [Security.AccessControl.AccessControlType]::Allow -and
        (($_.FileSystemRights -band [Security.AccessControl.FileSystemRights]::Modify) -eq [Security.AccessControl.FileSystemRights]::Modify) -and
        (($_.InheritanceFlags -band [Security.AccessControl.InheritanceFlags]::ContainerInherit) -ne 0) -and
        (($_.InheritanceFlags -band [Security.AccessControl.InheritanceFlags]::ObjectInherit) -ne 0)
    }
    if (-not $droitsSuffisants) {
        $regle = New-Object Security.AccessControl.FileSystemAccessRule(
            $GroupeAD, 'Modify', 'ContainerInherit,ObjectInherit', 'None', 'Allow'
        )
        $acl.AddAccessRule($regle) | Out-Null
        Set-Acl -LiteralPath $CheminLocal -AclObject $acl
        New-Resultat 'OK' "Droits NTFS Modifier accordés à $GroupeAD."
    }
    else {
        New-Resultat 'OK' 'Droits NTFS déjà présents.'
    }

    $droitPartage = Get-SmbShareAccess -Name $NomPartage -ErrorAction SilentlyContinue | Where-Object {
        $_.AccountName -eq $GroupeAD -and $_.AccessControlType -eq 'Allow' -and $_.AccessRight -in @('Change', 'Full')
    }
    if (-not $droitPartage) {
        Grant-SmbShareAccess -Name $NomPartage -AccountName $GroupeAD -AccessRight Change -Force | Out-Null
        New-Resultat 'OK' "Droit de partage Modifier accordé à $GroupeAD."
    }
    else {
        New-Resultat 'OK' 'Droit de partage déjà présent.'
    }
}

Write-Log "Serveur de fichiers : $ServeurFichiers"
if ($executionLocale) {
    $resultats = & $configurationPartage $CheminLocal $NomPartage $DescriptionPartage $GroupeAD ([bool](Test-Simulation))
}
else {
    try {
        Test-WSMan -ComputerName $ServeurFichiers -Credential $IdentifiantsServeur `
            -Authentication Negotiate -ErrorAction Stop | Out-Null
    }
    catch {
        throw "WinRM est inaccessible sur '$ServeurFichiers' : $($_.Exception.Message)"
    }

    try {
        $resultats = Invoke-Command `
            -ComputerName $ServeurFichiers `
            -Credential $IdentifiantsServeur `
            -Authentication Negotiate `
            -ScriptBlock $configurationPartage `
            -ArgumentList $CheminLocal,$NomPartage,$DescriptionPartage,$GroupeAD,([bool](Test-Simulation)) `
            -ErrorAction Stop
    }
    catch {
        throw "Échec de la configuration distante sur '$ServeurFichiers' : $($_.Exception.Message)"
    }
}

foreach ($resultat in @($resultats)) {
    Write-Log "[$($resultat.Serveur)] $($resultat.Message)" $resultat.Niveau
}
Write-Log "Contrôle terminé. Chemin réseau : \\$ServeurFichiers\$NomPartage" 'OK'
