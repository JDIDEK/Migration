#requires -Version 5.1
<#
.SYNOPSIS
    Copie un petit jeu de fichiers avec Robocopy en conservant les métadonnées NTFS.
.PREREQUIS
    Robocopy (Windows), accès source/destination et droits de sauvegarde/restauration
    pour /COPYALL. Le script ne supprime jamais les fichiers de destination.
.PARAMETER Simulation
    Exécute réellement Robocopy avec /L : aucun fichier n'est copié, mais un fichier
    de log est créé. -DryRun/-WhatIf n'exécutent même pas Robocopy et ne créent rien.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$DossierSource = 'C:\Migration-Test\Source',
    [string]$DossierDestination = '\\SRV-FICHIERS-TEST\Migration-Test\Robocopy',
    [string]$DossierLogs = 'C:\Migration-Test\Logs',
    [switch]$DryRun,
    [switch]$Simulation
)

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-SimulationTotale { [bool]($DryRun -or $WhatIfPreference) }

if (-not (Get-Command robocopy.exe -ErrorAction SilentlyContinue)) { throw 'robocopy.exe est introuvable.' }
if (-not (Test-Path -LiteralPath $DossierSource -PathType Container)) { throw "Le dossier source '$DossierSource' n'existe pas." }
$horodatage = Get-Date -Format 'yyyyMMdd-HHmmss'
$fichierLog = Join-Path $DossierLogs "Robocopy-Test-$horodatage.log"
$arguments = @(
    $DossierSource, $DossierDestination,
    '/E', '/COPYALL', '/DCOPY:DAT', '/R:2', '/W:3', '/XJ', '/NP', '/TEE',
    "/LOG:$fichierLog"
)
if ($Simulation) { $arguments += '/L' }
$commandeLisible = 'robocopy.exe ' + (($arguments | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }) -join ' ')

if (Test-SimulationTotale) {
    Write-Log "DRYRUN - Commande prévue : $commandeLisible"
    Write-Log 'Aucun dossier, log ou fichier ne sera créé.'
    exit 0
}

$mode = if ($Simulation) { 'la simulation Robocopy /L (seul le log sera créé)' } else { 'la copie Robocopy réelle' }
if ((Read-Host "Tapez OUI pour lancer $mode") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 }
if ($PSCmdlet.ShouldProcess($DossierDestination,$mode)) {
    if (-not (Test-Path -LiteralPath $DossierLogs -PathType Container)) { New-Item -ItemType Directory -Path $DossierLogs -Force | Out-Null }
    Write-Log "Commande : $commandeLisible"
    & robocopy.exe @arguments
    $code = $LASTEXITCODE
    if ($code -ge 8) { throw "Robocopy a retourné le code d'échec $code. Consultez '$fichierLog'." }
    Write-Log "Robocopy terminé avec le code $code. Log : $fichierLog" 'OK'
}
