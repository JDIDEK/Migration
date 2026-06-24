#requires -Version 5.1
<#
.SYNOPSIS
    Lance ou suit un déplacement distant (cross-forest) de la boîte de test.
.PREREQUIS
    À lancer dans l'Exchange Management Shell de la forêt CIBLE. L'objet cible doit
    être préparé comme MailUser (pas comme UserMailbox), MRS Proxy doit être actif
    côté source, et le compte source doit avoir les droits de migration.
.IMPORTANT
    « intra.ght53.fr » est le TargetDeliveryDomain, pas une destination suffisante
    à lui seul. Renseignez également le FQDN Exchange/MRS Proxy source. Le script 10
    et ce scénario cross-forest sont des scénarios alternatifs pour le même compte.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$IdentiteCible = 'migration.test@intra.ght53.fr',
    [string]$ServeurExchangeSource = 'exchange-source.ancien-domaine.local',
    [string]$DomaineLivraisonCible = 'intra.ght53.fr',
    [string]$BaseCible = '', # Vide = sélection automatique par Exchange
    [switch]$DryRun,
    [switch]$Relancer,
    [PSCredential]$IdentifiantsSource
)

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution { param([string]$Texte); if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }; if ((Read-Host "Tapez OUI pour $Texte") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 } }
function Show-MoveStatus {
    param($Requete)
    $stats = $Requete | Get-MoveRequestStatistics -IncludeReport
    Write-Log ("Migration : statut={0}, progression={1}, éléments transférés={2}" -f $stats.Status,$stats.PercentComplete,$stats.ItemsTransferred)
    return $stats
}

foreach ($commande in @('Get-Recipient','Get-MoveRequest','Get-MoveRequestStatistics','New-MoveRequest','Remove-MoveRequest')) {
    if (-not (Get-Command $commande -ErrorAction SilentlyContinue)) { throw "Cmdlet '$commande' absente. Ouvrez l'Exchange Management Shell on-premise de la cible." }
}
if ($ServeurExchangeSource -like '*ancien-domaine*') { Write-Log "Le FQDN source '$ServeurExchangeSource' ressemble à une valeur d'exemple : modifiez-le avant une exécution réelle." 'ATTENTION' }
$destinataire = Get-Recipient -Identity $IdentiteCible -ErrorAction Stop
if ($destinataire.RecipientTypeDetails -ne 'MailUser') { throw "L'objet cible doit être un MailUser pour un remote move ; type actuel : $($destinataire.RecipientTypeDetails). N'exécutez pas le script 10 avant ce scénario." }

$move = Get-MoveRequest -Identity $IdentiteCible -ErrorAction SilentlyContinue
if ($move) {
    $stats = Show-MoveStatus $move
    if (-not $Relancer) {
        Write-Log "Une requête existe déjà. Elle est conservée. Utilisez -Relancer uniquement pour supprimer puis recréer une requête en échec ou bloquée." 'INFO'
        exit 0
    }
    Confirm-Execution "supprimer puis relancer la requête existante de '$IdentiteCible'"
    if ($DryRun) { Write-Log "DRYRUN - Supprimer la requête existante ($($stats.Status)), puis en créer une nouvelle."; exit 0 }
    if ($PSCmdlet.ShouldProcess($IdentiteCible,"Supprimer la MoveRequest existante ($($stats.Status))")) {
        $move | Remove-MoveRequest -Confirm:$false -ErrorAction Stop
        Write-Log 'Ancienne requête supprimée.' 'OK'
    }
}
else { Confirm-Execution "lancer le déplacement distant de '$IdentiteCible'" }

if (Test-Simulation) {
    Write-Log "DRYRUN - New-MoveRequest -Remote via '$ServeurExchangeSource', TargetDeliveryDomain '$DomaineLivraisonCible', base '$BaseCible'."
    exit 0
}
if ($ServeurExchangeSource -like '*ancien-domaine*') {
    throw "Remplacez la valeur d'exemple de `$ServeurExchangeSource par le FQDN MRS Proxy réel avant de lancer la migration."
}
if (-not $IdentifiantsSource) { $IdentifiantsSource = Get-Credential -Message 'Compte autorisé dans la forêt Exchange source' }
$parametres = @{
    Identity             = $IdentiteCible
    Remote               = $true
    RemoteHostName       = $ServeurExchangeSource
    RemoteCredential     = $IdentifiantsSource
    TargetDeliveryDomain = $DomaineLivraisonCible
    ErrorAction          = 'Stop'
}
if ($BaseCible) { $parametres.TargetDatabase = $BaseCible }
if ($PSCmdlet.ShouldProcess($IdentiteCible,"Créer la MoveRequest distante vers $DomaineLivraisonCible")) {
    New-MoveRequest @parametres | Out-Null
    Write-Log 'Requête de déplacement créée.' 'OK'
    $nouvelleRequete = Get-MoveRequest -Identity $IdentiteCible -ErrorAction Stop
    $null = Show-MoveStatus $nouvelleRequete
}
