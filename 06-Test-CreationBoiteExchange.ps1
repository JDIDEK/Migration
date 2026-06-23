#requires -Version 5.1
<#
.SYNOPSIS
    Active une boîte Exchange on-premise pour l'utilisateur AD de test existant.
.PREREQUIS
    À lancer dans l'Exchange Management Shell on-premise avec les rôles RBAC
    nécessaires. L'utilisateur AD doit exister et ne doit pas déjà être un autre
    type de destinataire Exchange.
.NOTE
    Ce script utilise Enable-Mailbox (et non New-Mailbox) afin de conserver
    l'utilisateur AD créé par le script 05.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$UtilisateurTest = 'migration.test',
    [string]$AliasMessagerie = 'migration.test',
    [string]$BaseDeDonnees = '', # Vide = sélection automatique par Exchange
    [string]$ServeurExchange = 'ght-exchangew-1',
    [string]$UriPowerShellExchange = 'http://ght-exchangew-1/PowerShell/',
    [switch]$DryRun
)

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution { if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }; if ((Read-Host "Tapez OUI pour activer la boîte Exchange de '$UtilisateurTest'") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 } }
function Import-ExchangeSessionIfNeeded {
    param(
        [string[]]$CommandesAttendues,
        [string]$ConnectionUri
    )

    $manquantes = @($CommandesAttendues | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($manquantes.Count -eq 0) { return }

    Write-Log ("Cmdlets Exchange absentes : {0}. Import de session distante..." -f ($manquantes -join ', ')) 'ATTENTION'
    try {
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos -ErrorAction Stop
    }
    catch {
        Write-Log "Connexion Kerberos implicite impossible : $($_.Exception.Message)" 'ATTENTION'
        $identifiants = Get-Credential -Message "Compte Exchange/INTRA autorisé, ex: INTRA\admin ou admin@intra.ght53.fr"
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos -Credential $identifiants -ErrorAction Stop
    }
    Import-PSSession $session -DisableNameChecking -AllowClobber -ErrorAction Stop | Out-Null
    Write-Log "Session Exchange importée depuis $ConnectionUri." 'OK'
}

$commandesExchange = @('Get-Recipient','Get-User','Get-Mailbox','Enable-Mailbox')
Import-ExchangeSessionIfNeeded -CommandesAttendues $commandesExchange -ConnectionUri $UriPowerShellExchange
foreach ($commande in $commandesExchange) {
    if (-not (Get-Command $commande -ErrorAction SilentlyContinue)) { throw "Cmdlet '$commande' absente. Ouvrez l'Exchange Management Shell on-premise." }
}

$boite = Get-Mailbox -Identity $UtilisateurTest -ErrorAction SilentlyContinue
if ($boite) { Write-Log "La boîte existe déjà : $($boite.PrimarySmtpAddress). Aucune action." 'OK'; exit 0 }
$destinataire = Get-Recipient -Identity $UtilisateurTest -ErrorAction SilentlyContinue
if ($destinataire) { throw "L'objet est déjà un destinataire Exchange de type '$($destinataire.RecipientTypeDetails)'. Enable-Mailbox n'est pas appliqué automatiquement." }
$utilisateur = Get-User -Identity $UtilisateurTest -ErrorAction Stop
if ($utilisateur.RecipientTypeDetails -notin @('User','DisabledUser')) { throw "Type d'objet inattendu : $($utilisateur.RecipientTypeDetails)." }

Confirm-Execution
if ($DryRun) { Write-Log "DRYRUN - Activer la boîte de '$UtilisateurTest' avec l'alias '$AliasMessagerie'$(if ($BaseDeDonnees) { " dans '$BaseDeDonnees'" })."; exit 0 }
if ($PSCmdlet.ShouldProcess($UtilisateurTest, 'Activer la boîte Exchange on-premise')) {
    $parametres = @{ Identity = $UtilisateurTest; Alias = $AliasMessagerie; ErrorAction = 'Stop' }
    if ($BaseDeDonnees) { $parametres.Database = $BaseDeDonnees }
    Enable-Mailbox @parametres | Out-Null
    $boite = Get-Mailbox -Identity $UtilisateurTest -ErrorAction Stop
    Write-Log "Boîte activée : $($boite.PrimarySmtpAddress), base $($boite.Database)." 'OK'
}
