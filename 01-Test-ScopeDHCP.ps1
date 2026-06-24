#requires -Version 5.1
<#
.SYNOPSIS
    Crée un scope DHCP inactif et réserve une adresse pour la VM locale.
.PREREQUIS
    Rôle/Outils DHCP (module DhcpServer), module NetAdapter, droits administrateur
    DHCP. La VM qui exécute le script doit être la VM à réserver.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$ServeurDHCP = 'SRV-DHCP-TEST',
    [string]$NomScope = 'TEST-MIGRATION-192.168.99.0',
    [string]$ScopeId = '192.168.99.0',
    [string]$DebutPlage = '192.168.99.100',
    [string]$FinPlage = '192.168.99.150',
    [string]$Masque = '255.255.255.0',
    [string]$IPReserveeVM = '192.168.99.110',
    [string]$NomCarteReseau = '', # Vide = détection de l'unique carte active
    [switch]$DryRun
)

function Write-Log { param([string]$Message,[string]$Niveau='INFO'); Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Niveau,$Message) }
function Test-Simulation { [bool]($DryRun -or $WhatIfPreference) }
function Confirm-Execution { if (Test-Simulation) { Write-Log 'Simulation : aucune modification ne sera effectuée.'; return }; if ((Read-Host "Tapez OUI pour créer le scope DHCP INACTIF et sa réservation") -cne 'OUI') { Write-Log 'Opération annulée.' 'ATTENTION'; exit 0 } }
function Invoke-Change { param([string]$Cible,[string]$Action,[scriptblock]$Commande); if ($DryRun) { Write-Log "DRYRUN - $Action sur $Cible"; return }; if ($PSCmdlet.ShouldProcess($Cible,$Action)) { & $Commande } }

Import-Module DhcpServer -ErrorAction Stop
if ($NomCarteReseau) { $cartes = @(Get-NetAdapter -Name $NomCarteReseau -ErrorAction Stop) }
else { $cartes = @(Get-NetAdapter | Where-Object { $_.Status -eq 'Up' -and $_.MacAddress }) }
if ($cartes.Count -ne 1) { throw "Détection de carte ambiguë ($($cartes.Count) carte(s)). Renseignez `$NomCarteReseau en haut du script." }
$mac = ($cartes[0].MacAddress -replace '[-:.]', '').ToUpperInvariant()
Write-Log "Carte retenue : $($cartes[0].Name), MAC $mac"
Confirm-Execution

$scope = Get-DhcpServerv4Scope -ComputerName $ServeurDHCP -ScopeId $ScopeId -ErrorAction SilentlyContinue
if (-not $scope) {
    Invoke-Change "$ServeurDHCP/$ScopeId" 'Créer le scope DHCP désactivé' {
        Add-DhcpServerv4Scope -ComputerName $ServeurDHCP -Name $NomScope -StartRange $DebutPlage -EndRange $FinPlage -SubnetMask $Masque -State InActive -Description 'TEST UNIQUEMENT - ne pas activer sans validation'
    }
}
else {
    if ($scope.StartRange.IPAddressToString -ne $DebutPlage -or $scope.EndRange.IPAddressToString -ne $FinPlage -or $scope.SubnetMask.IPAddressToString -ne $Masque) {
        throw 'Un scope existe déjà avec le même ScopeId mais une plage ou un masque différent.'
    }
    if ($scope.State -ne 'Inactive') { throw "Le scope existant est ACTIF. Le script refuse de le modifier ; désactivez-le et vérifiez son usage manuellement." }
    Write-Log 'Scope déjà présent et inactif.' 'OK'
}

if ($scope -or -not (Test-Simulation)) {
    $reservations = @(Get-DhcpServerv4Reservation -ComputerName $ServeurDHCP -ScopeId $ScopeId -ErrorAction SilentlyContinue)
    $memeIP = $reservations | Where-Object { $_.IPAddress.IPAddressToString -eq $IPReserveeVM }
    $memeMac = $reservations | Where-Object { ($_.ClientId -replace '[-:.]', '').ToUpperInvariant() -eq $mac }
    if (($memeIP -and (($memeIP.ClientId -replace '[-:.]', '').ToUpperInvariant() -ne $mac)) -or ($memeMac -and $memeMac.IPAddress.IPAddressToString -ne $IPReserveeVM)) {
        throw 'Une réservation existante utilise déjà cette IP ou cette MAC avec une autre association.'
    }
    if (-not $memeIP -and -not $memeMac) {
        Invoke-Change "$IPReserveeVM / $mac" 'Créer la réservation DHCP de la VM' {
            Add-DhcpServerv4Reservation -ComputerName $ServeurDHCP -ScopeId $ScopeId -IPAddress $IPReserveeVM -ClientId $mac -Name $env:COMPUTERNAME -Description 'VM de test migration'
        }
    }
    else { Write-Log 'Réservation déjà présente.' 'OK' }
}
else { Write-Log "DRYRUN - Créer la réservation $IPReserveeVM pour la MAC $mac." }

Write-Log 'Le scope reste INACTIF. Son activation doit être effectuée manuellement après vérification.' 'ATTENTION'
