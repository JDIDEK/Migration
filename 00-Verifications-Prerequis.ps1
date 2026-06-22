#requires -Version 5.1
<#
.SYNOPSIS
    Vérifie les prérequis des scripts de test de migration.
.DESCRIPTION
    Contrôle sans rien modifier : élévation, modules Windows et cmdlets Exchange.
.PREREQUIS
    PowerShell 5.1. À lancer de préférence sur chaque serveur concerné, car les
    rôles et outils RSAT peuvent différer d'un serveur à l'autre.
#>
[CmdletBinding()]
param(
    [string]$ModulesRequis = 'ActiveDirectory;DhcpServer;PrintManagement;SmbShare',
    [string]$CommandesExchange = 'Get-Mailbox;Enable-Mailbox;New-MailboxExportRequest;Get-MailboxExportRequest;New-MoveRequest;Get-MoveRequest',
    [switch]$DryRun
)

$listeModulesRequis = @($ModulesRequis.Split(';') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$listeCommandesExchange = @($CommandesExchange.Split(';') | ForEach-Object { $_.Trim() } | Where-Object { $_ })

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'OK', 'ATTENTION', 'ERREUR')][string]$Niveau = 'INFO')
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Niveau, $Message)
}

if ($DryRun) {
    Write-Log "-DryRun n'est pas nécessaire ici : ce script est strictement en lecture seule." 'INFO'
}

$identite = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identite)
$estAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Write-Log ("Session élevée : {0}" -f $(if ($estAdmin) { 'oui' } else { 'non' })) $(if ($estAdmin) { 'OK' } else { 'ATTENTION' })

$echecs = 0
foreach ($module in $listeModulesRequis) {
    if (Get-Module -ListAvailable -Name $module) {
        Write-Log "Module disponible : $module" 'OK'
    }
    else {
        Write-Log "Module absent sur cette machine : $module" 'ATTENTION'
        $echecs++
    }
}

$commandesExchangeAbsentes = @($listeCommandesExchange | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
if ($commandesExchangeAbsentes.Count -eq 0) {
    Write-Log 'Exchange Management Shell chargée : toutes les cmdlets attendues sont disponibles.' 'OK'
}
else {
    Write-Log ("Cmdlets Exchange absentes de la session : {0}" -f ($commandesExchangeAbsentes -join ', ')) 'ATTENTION'
    Write-Log "Ouvrez l'Exchange Management Shell on-premise ou importez une session distante Exchange." 'INFO'
    $echecs++
}

if (Get-Command Get-PSSnapin -ErrorAction SilentlyContinue) {
    $snapin = Get-PSSnapin -Registered -ErrorAction SilentlyContinue |
        Where-Object Name -Like 'Microsoft.Exchange.Management.PowerShell*'
    Write-Log ("Snap-in Exchange enregistré localement : {0}" -f $(if ($snapin) { 'oui' } else { 'non ou non applicable' })) 'INFO'
}

if ($echecs -eq 0) {
    Write-Log 'Tous les contrôles génériques sont satisfaits sur cette machine.' 'OK'
    exit 0
}

Write-Log "$echecs catégorie(s) de prérequis sont absentes. Consultez les messages ci-dessus." 'ATTENTION'
exit 1
