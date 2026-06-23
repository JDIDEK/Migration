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
    [string]$DomaineUPN = 'intra.ght53.fr',
    [string]$DomaineSMTP = 'intra.ght53.fr',
    [string]$AdresseSMTPPrincipale = '',
    [string]$BaseDeDonnees = '', # Vide = sélection automatique par Exchange
    [string]$ServeurExchange = 'ght-exchangew-1',
    [string]$UriPowerShellExchange = 'http://ght-exchangew-1/PowerShell/',
    [string]$ControleurDomaineExchange = '',
    [PSCredential]$IdentifiantsExchange,
    [string]$IdentifiantsExchangePath = '',
    [switch]$DryRun
)

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution { if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }; if ((Read-Host "Tapez OUI pour activer la boîte Exchange de '$UtilisateurTest'") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 } }
function Get-DomainControllersToTry {
    if ($ControleurDomaineExchange) { return @($ControleurDomaineExchange) }

    $controleurs = @()
    try {
        $sortieNltest = & nltest "/dclist:$DomaineUPN" 2>&1
        if ($LASTEXITCODE -eq 0) {
            $controleurs = @($sortieNltest | ForEach-Object {
                if ($_ -match '\\\\([^\s\[]+)') {
                    $nom = $matches[1].Trim()
                    if ($nom -and $nom -notlike '*.*') { "$nom.$DomaineUPN" } else { $nom }
                }
            } | Where-Object { $_ } | Sort-Object -Unique)
        }
        else {
            Write-Log "nltest /dclist:$DomaineUPN a échoué : $($sortieNltest -join ' | ')" 'ATTENTION'
        }
    }
    catch {
        Write-Log "Impossible de lister les DC avec nltest : $($_.Exception.Message)" 'ATTENTION'
    }

    @('') + $controleurs
}
function Get-ExistingExchangeObject {
    param(
        [string[]]$Identites,
        [ValidateSet('Mailbox','Recipient','User')][string]$Type,
        [string[]]$DomainControllers
    )

    foreach ($controleur in $DomainControllers) {
        foreach ($identite in $Identites) {
            try {
                $parametres = @{ Identity = $identite; ErrorAction = 'Stop' }
                if ($controleur) { $parametres.DomainController = $controleur }
                switch ($Type) {
                    'Mailbox' { $objet = Get-Mailbox @parametres }
                    'Recipient' { $objet = Get-Recipient @parametres }
                    'User' { $objet = Get-User @parametres }
                }
                $script:ControleurDomaineTrouve = $controleur
                return $objet
            }
            catch {
                $nomException = $_.Exception.GetType().FullName
                $idErreur = $_.FullyQualifiedErrorId
                if ($nomException -like '*ManagementObjectNotFoundException*' -or $idErreur -like '*ManagementObjectNotFoundException*') { continue }
                throw
            }
        }
    }
    return $null
}
function Import-ExchangeSessionIfNeeded {
    param(
        [string[]]$CommandesAttendues,
        [string]$ConnectionUri,
        [PSCredential]$Credential
    )

    $manquantes = @($CommandesAttendues | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($manquantes.Count -eq 0) { return }

    Write-Log ("Cmdlets Exchange absentes : {0}. Import de session distante..." -f ($manquantes -join ', ')) 'ATTENTION'
    try {
        if ($Credential) {
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos -Credential $Credential -ErrorAction Stop
        }
        else {
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos -ErrorAction Stop
        }
    }
    catch {
        Write-Log "Connexion Kerberos implicite impossible : $($_.Exception.Message)" 'ATTENTION'
        if (-not $Credential) { $Credential = Get-Credential -Message "Compte Exchange/INTRA autorisé, ex: INTRA\admin ou admin@intra.ght53.fr" }
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos -Credential $Credential -ErrorAction Stop
    }
    Import-PSSession $session -DisableNameChecking -AllowClobber -ErrorAction Stop | Out-Null
    Write-Log "Session Exchange importée depuis $ConnectionUri." 'OK'
}

$commandesExchange = @('Get-Recipient','Get-User','Get-Mailbox','Enable-Mailbox','Set-Mailbox')
if (-not $IdentifiantsExchange -and $IdentifiantsExchangePath -and (Test-Path -LiteralPath $IdentifiantsExchangePath -PathType Leaf)) {
    $IdentifiantsExchange = Import-Clixml -LiteralPath $IdentifiantsExchangePath
}
Import-ExchangeSessionIfNeeded -CommandesAttendues $commandesExchange -ConnectionUri $UriPowerShellExchange -Credential $IdentifiantsExchange
foreach ($commande in $commandesExchange) {
    if (-not (Get-Command $commande -ErrorAction SilentlyContinue)) { throw "Cmdlet '$commande' absente. Ouvrez l'Exchange Management Shell on-premise." }
}
if ($ControleurDomaineExchange) { Write-Log "Contrôleur de domaine Exchange forcé : $ControleurDomaineExchange" }
$controleursDomaine = @(Get-DomainControllersToTry)
Write-Log ("Contrôleurs de domaine testés : {0}" -f (($controleursDomaine | ForEach-Object { if ($_) { $_ } else { '<Exchange défaut>' } }) -join ', '))

$identites = @($UtilisateurTest)
if ($UtilisateurTest -notlike '*@*') { $identites += "$UtilisateurTest@$DomaineUPN" }
if (-not $AdresseSMTPPrincipale) { $AdresseSMTPPrincipale = "$AliasMessagerie@$DomaineSMTP" }
Write-Log "Adresse SMTP principale cible : $AdresseSMTPPrincipale"

$script:ControleurDomaineTrouve = $null
$boite = Get-ExistingExchangeObject -Identites $identites -Type Mailbox -DomainControllers $controleursDomaine
if ($boite) { Write-Log "La boîte existe déjà : $($boite.PrimarySmtpAddress). Aucune action." 'OK'; exit 0 }
$destinataire = Get-ExistingExchangeObject -Identites $identites -Type Recipient -DomainControllers $controleursDomaine
if ($destinataire) { throw "L'objet est déjà un destinataire Exchange de type '$($destinataire.RecipientTypeDetails)'. Enable-Mailbox n'est pas appliqué automatiquement." }
$utilisateur = Get-ExistingExchangeObject -Identites $identites -Type User -DomainControllers $controleursDomaine
if (-not $utilisateur) {
    throw "Utilisateur '$UtilisateurTest' introuvable par Exchange sur les DC testés. Vérifiez la réplication AD ou renseignez -ControleurDomaineExchange."
}
if ($script:ControleurDomaineTrouve) { Write-Log "Utilisateur trouvé via le DC : $script:ControleurDomaineTrouve" 'OK' }
else { Write-Log "Utilisateur trouvé via le contrôleur choisi par Exchange." 'OK' }
if ($utilisateur.RecipientTypeDetails -notin @('User','DisabledUser')) { throw "Type d'objet inattendu : $($utilisateur.RecipientTypeDetails)." }

Confirm-Execution
if ($DryRun) { Write-Log "DRYRUN - Activer la boîte de '$UtilisateurTest' avec l'alias '$AliasMessagerie'$(if ($BaseDeDonnees) { " dans '$BaseDeDonnees'" })."; exit 0 }
if ($PSCmdlet.ShouldProcess($UtilisateurTest, 'Activer la boîte Exchange on-premise')) {
    $parametres = @{ Identity = $UtilisateurTest; Alias = $AliasMessagerie; ErrorAction = 'Stop' }
    $controleurOperation = if ($ControleurDomaineExchange) { $ControleurDomaineExchange } else { $script:ControleurDomaineTrouve }
    if ($controleurOperation) { $parametres.DomainController = $controleurOperation }
    if ($BaseDeDonnees) { $parametres.Database = $BaseDeDonnees }
    Enable-Mailbox @parametres | Out-Null
    $parametresAdresse = @{
        Identity                  = $UtilisateurTest
        EmailAddressPolicyEnabled = $false
        PrimarySmtpAddress        = $AdresseSMTPPrincipale
        ErrorAction               = 'Stop'
    }
    if ($controleurOperation) { $parametresAdresse.DomainController = $controleurOperation }
    Set-Mailbox @parametresAdresse
    $boite = Get-ExistingExchangeObject -Identites $identites -Type Mailbox -DomainControllers $controleursDomaine
    Write-Log "Boîte activée : $($boite.PrimarySmtpAddress), base $($boite.Database)." 'OK'
}
