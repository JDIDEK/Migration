#requires -Version 5.1
<#
.SYNOPSIS
    Exporte une boîte Exchange on-premise vers un PST, suit puis nettoie la requête.
.PREREQUIS
    Exchange Management Shell on-premise ; rôle RBAC « Mailbox Import Export ».
    Le chemin PST doit être UNC et le groupe Exchange Trusted Subsystem doit avoir
    les droits d'écriture sur le partage et le dossier NTFS.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$BoiteTest = 'migration.test',
    [string]$DossierPST = '\\SRV-FICHIERS-TEST\PST-Test$',
    [string]$NomFichierPST = 'migration.test.pst',
    [string]$NomRequete = 'Export-Migration-Test',
    [ValidateRange(1,3600)][int]$IntervalleSecondes = 15,
    [ValidateRange(1,1440)][int]$DelaiMaximumMinutes = 120,
    [switch]$DryRun,
    [switch]$NePasAttendre
)

$CheminPST = Join-Path $DossierPST $NomFichierPST

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution { if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }; if ((Read-Host "Tapez OUI pour exporter '$BoiteTest' vers '$CheminPST'") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 } }
function Invoke-Change { param([string]$Cible,[string]$Action,[scriptblock]$Commande); if ($DryRun) { Write-Log "DRYRUN - $Action sur $Cible"; return }; if ($PSCmdlet.ShouldProcess($Cible,$Action)) { & $Commande } }
function Get-RequeteTest {
    @(Get-MailboxExportRequest -Mailbox $BoiteTest -ErrorAction SilentlyContinue | Where-Object Name -eq $NomRequete)
}

foreach ($commande in @('Get-Mailbox','New-MailboxExportRequest','Get-MailboxExportRequest','Get-MailboxExportRequestStatistics','Remove-MailboxExportRequest')) {
    if (-not (Get-Command $commande -ErrorAction SilentlyContinue)) { throw "Cmdlet '$commande' absente. Ouvrez l'Exchange Management Shell et vérifiez le rôle Mailbox Import Export." }
}
$null = Get-Mailbox -Identity $BoiteTest -ErrorAction Stop
$requetes = @(Get-RequeteTest)
if ($requetes.Count -gt 1) { throw "Plusieurs requêtes '$NomRequete' existent pour cette boîte. Nettoyage manuel requis." }
$requete = $requetes | Select-Object -First 1

if (-not $requete -and -not (Test-Simulation)) {
    if (-not (Test-Path -LiteralPath $DossierPST -PathType Container)) { throw "Le dossier UNC '$DossierPST' n'est pas accessible depuis cette session." }
    if (Test-Path -LiteralPath $CheminPST) { throw "Le PST '$CheminPST' existe déjà. Il est conservé ; renommez-le ou modifiez la variable de destination." }
}

Confirm-Execution
if (-not $requete) {
    Invoke-Change $BoiteTest "Créer la requête d'export vers $CheminPST" {
        New-MailboxExportRequest -Mailbox $BoiteTest -FilePath $CheminPST -Name $NomRequete -ErrorAction Stop | Out-Null
    }
    if (Test-Simulation) { Write-Log 'DRYRUN - Suivre la requête, puis la supprimer uniquement après succès.'; exit 0 }
    $requete = Get-RequeteTest | Select-Object -First 1
}
else { Write-Log "Requête existante détectée : $($requete.Status). Elle n'est pas dupliquée." 'INFO' }

if (Test-Simulation) {
    $statistiques = $requete | Get-MailboxExportRequestStatistics
    Write-Log ("SIMULATION - Requête actuelle : statut={0}, progression={1}" -f $statistiques.Status,$statistiques.PercentComplete)
    if ($statistiques.Status -in @('Completed','CompletedWithWarning')) {
        Write-Log 'SIMULATION - La requête terminée serait supprimée ; le fichier PST serait conservé.'
    }
    else {
        Write-Log "SIMULATION - La requête serait suivie jusqu'à son terme, sans changement pendant cette exécution."
    }
    exit 0
}

$debut = Get-Date
do {
    $requete = Get-RequeteTest | Select-Object -First 1
    if (-not $requete) { throw "La requête d'export est introuvable pendant le suivi." }
    $statistiques = $requete | Get-MailboxExportRequestStatistics -IncludeReport
    Write-Log ("Export : statut={0}, progression={1}" -f $statistiques.Status, $statistiques.PercentComplete)

    if ($statistiques.Status -in @('Completed','CompletedWithWarning')) {
        Invoke-Change "$BoiteTest/$NomRequete" "Supprimer la requête d'export terminée" {
            $requete | Remove-MailboxExportRequest -Confirm:$false -ErrorAction Stop
        }
        Write-Log "Export terminé. PST : $CheminPST" 'OK'
        break
    }
    if ($statistiques.Status -eq 'Failed') {
        Write-Log 'Export en échec. La requête est conservée pour permettre le diagnostic du rapport.' 'ERREUR'
        throw $statistiques.Message
    }
    if ($NePasAttendre) { Write-Log 'Suivi interrompu à la demande. Relancez ce script pour contrôler et nettoyer la requête.' 'INFO'; break }
    if (((Get-Date) - $debut).TotalMinutes -ge $DelaiMaximumMinutes) { throw "Délai de $DelaiMaximumMinutes minutes dépassé. La requête reste en place." }
    Start-Sleep -Seconds $IntervalleSecondes
} while ($true)
