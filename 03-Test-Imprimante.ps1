#requires -Version 5.1
<#
.SYNOPSIS
    Ajoute un port TCP/IP documentaire et une imprimante partagée de test.
.PREREQUIS
    Rôle Serveur d'impression, module PrintManagement et droits administrateur.
    Le pilote « Generic / Text Only » doit être présent dans le magasin de pilotes.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$NomPort = 'IP_TEST_192.0.2.10',
    [string]$AdresseIPTest = '192.0.2.10', # Plage TEST-NET-1
    [string]$NomPilote = 'Generic / Text Only',
    [string]$NomImprimante = 'Imprimante-Migration-Test',
    [string]$NomPartage = 'IMP-Migration-Test',
    [switch]$DryRun
)

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution { if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }; if ((Read-Host "Tapez OUI pour créer l'imprimante de test") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 } }
function Invoke-Change { param([string]$Cible,[string]$Action,[scriptblock]$Commande); if ($DryRun) { Write-Log "DRYRUN - $Action sur $Cible"; return }; if ($PSCmdlet.ShouldProcess($Cible,$Action)) { & $Commande } }

Import-Module PrintManagement -ErrorAction Stop
Confirm-Execution

$port = Get-PrinterPort -Name $NomPort -ErrorAction SilentlyContinue
if ($port -and $port.PrinterHostAddress -and $port.PrinterHostAddress -ne $AdresseIPTest) { throw "Le port '$NomPort' existe avec l'adresse $($port.PrinterHostAddress)." }
if (-not $port) { Invoke-Change $NomPort "Créer le port TCP/IP vers $AdresseIPTest" { Add-PrinterPort -Name $NomPort -PrinterHostAddress $AdresseIPTest } }
else { Write-Log 'Port déjà présent.' 'OK' }

if (-not (Get-PrinterDriver -Name $NomPilote -ErrorAction SilentlyContinue)) {
    Invoke-Change $NomPilote 'Installer le pilote depuis le magasin de pilotes' { Add-PrinterDriver -Name $NomPilote -ErrorAction Stop }
}
else { Write-Log 'Pilote déjà présent.' 'OK' }

$imprimante = Get-Printer -Name $NomImprimante -ErrorAction SilentlyContinue
if ($imprimante -and ($imprimante.PortName -ne $NomPort -or $imprimante.DriverName -ne $NomPilote)) {
    throw "L'imprimante existe avec un port ou un pilote différent. Aucun remplacement automatique n'est effectué."
}
if (-not $imprimante) {
    Invoke-Change $NomImprimante "Créer et partager l'imprimante" {
        Add-Printer -Name $NomImprimante -DriverName $NomPilote -PortName $NomPort -Shared -ShareName $NomPartage
    }
}
elseif (-not $imprimante.Shared -or $imprimante.ShareName -ne $NomPartage) {
    Invoke-Change $NomImprimante "Activer le partage $NomPartage" { Set-Printer -Name $NomImprimante -Shared $true -ShareName $NomPartage }
}
else { Write-Log 'Imprimante déjà présente et partagée comme attendu.' 'OK' }
Write-Log "Configuration de l'imprimante de test terminée." 'OK'
