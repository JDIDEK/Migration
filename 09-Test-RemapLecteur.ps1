#requires -Version 5.1
<#
.SYNOPSIS
    Mappe un lecteur réseau persistant sur la VM de test locale.
.PREREQUIS
    Exécuter dans la session interactive de l'utilisateur de test. Le partage doit
    être accessible. Un lecteur créé en session élevée peut ne pas apparaître dans
    la session non élevée à cause de l'isolation UAC.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$LettreLecteur = 'Z',
    [string]$CheminUNC = '\\SRV-FICHIERS-TEST\Migration-Test',
    [switch]$DryRun
)

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution { if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }; if ((Read-Host "Tapez OUI pour mapper $LettreLecteur`: vers '$CheminUNC'") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 } }

$LettreLecteur = $LettreLecteur.TrimEnd(':').ToUpperInvariant()
if ($LettreLecteur -notmatch '^[A-Z]$') { throw 'La lettre de lecteur doit contenir une seule lettre de A à Z.' }
$existant = Get-PSDrive -Name $LettreLecteur -ErrorAction SilentlyContinue
if ($existant) {
    $racineExistante = if ($existant.DisplayRoot) { $existant.DisplayRoot } else { $existant.Root }
    if ($racineExistante.TrimEnd('\') -ieq $CheminUNC.TrimEnd('\')) { Write-Log "$LettreLecteur`: est déjà mappé vers le bon partage." 'OK'; exit 0 }
    throw "$LettreLecteur`: est déjà utilisé par '$racineExistante'. Aucun remappage automatique n'est effectué."
}
if (-not (Test-Simulation) -and -not (Test-Path -LiteralPath $CheminUNC -PathType Container)) { throw "Le partage '$CheminUNC' n'est pas accessible depuis cette session." }
Confirm-Execution
if ($DryRun) { Write-Log "DRYRUN - Mapper $LettreLecteur`: vers '$CheminUNC' de façon persistante."; exit 0 }
if ($PSCmdlet.ShouldProcess("$LettreLecteur`:","Mapper vers $CheminUNC")) {
    New-PSDrive -Name $LettreLecteur -PSProvider FileSystem -Root $CheminUNC -Persist -Scope Global -ErrorAction Stop | Out-Null
    Write-Log "$LettreLecteur`: mappé vers '$CheminUNC'." 'OK'
}
