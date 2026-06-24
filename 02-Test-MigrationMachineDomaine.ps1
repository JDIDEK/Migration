#requires -Version 5.1
<#
.SYNOPSIS
    Bascule une liste de machines d'un ancien domaine vers le domaine cible.
.PREREQUIS
    - Lancer PowerShell en administrateur
    - Les machines doivent etre allumees et joignables
    - WinRM doit etre disponible pour les machines distantes
    - DNS correctement configure vers le nouveau domaine
    - Compte autorise a se connecter aux PC et a les sortir de l'ancien domaine
    - Compte autorise a joindre les PC au nouveau domaine
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ComputerListPath = '.\machines-test.txt',
    [string]$NewDomainFqdn = 'intra.ght53.fr',
    [string]$OUPath = 'OU=Postes,OU=HLER,DC=intra,DC=ght53,DC=fr',
    [string]$LogPath = '.\Logs\migration-domaine-resultats.csv',
    [switch]$Restart,
    [switch]$DryRun,
    [int]$RestartDelaySeconds = 30
)

$ErrorActionPreference = 'Stop'
if ($DryRun) { $WhatIfPreference = $true }
$WhatIfMode = [bool]$WhatIfPreference
$LocalComputerName = $env:COMPUTERNAME.ToUpperInvariant()

function Normalize-ComputerName {
    param([string]$Name)
    (($Name.Trim() -split '\.')[0]).ToUpperInvariant()
}

function Export-CsvPourExcel {
    param(
        [Parameter(Mandatory)][object[]]$Donnees,
        [Parameter(Mandatory)][string]$Chemin
    )

    $lignes = @('sep=;') + @($Donnees | ConvertTo-Csv -NoTypeInformation -Delimiter ';')
    # Excel Windows détecte l'UTF-16 LE de façon fiable, même avec les paramètres régionaux français.
    $encodage = [Text.Encoding]::Unicode
    [IO.File]::WriteAllLines([IO.Path]::GetFullPath($Chemin), [string[]]$lignes, $encodage)
}

if (-not (Test-Path -LiteralPath $ComputerListPath -PathType Leaf)) {
    throw "Fichier liste introuvable : $ComputerListPath"
}

$ComputersRaw = @(Get-Content -LiteralPath $ComputerListPath |
    Where-Object { $_ -and $_.Trim() -ne '' } |
    ForEach-Object { $_.Trim() } |
    Sort-Object -Unique)

if ($ComputersRaw.Count -eq 0) {
    throw "Aucune machine dans le fichier : $ComputerListPath"
}

# On met la machine locale en dernier si elle est dans la liste.
$Computers = @($ComputersRaw | Sort-Object {
    if ((Normalize-ComputerName $_) -eq $LocalComputerName) { 1 } else { 0 }
})

Write-Host ''
Write-Host "Machines a traiter : $($Computers.Count)" -ForegroundColor Cyan
Write-Host "Machine locale     : $env:COMPUTERNAME" -ForegroundColor Cyan
Write-Host "Domaine cible      : $NewDomainFqdn" -ForegroundColor Cyan
Write-Host "OU cible           : $OUPath" -ForegroundColor Cyan
Write-Host "Mode simulation    : $WhatIfMode" -ForegroundColor Cyan

Write-Host ''
Write-Host 'Credential 1/2 - Ancien domaine' -ForegroundColor Cyan
Write-Host "Ce compte sert a se connecter aux PC et a les sortir de l'ancien domaine." -ForegroundColor DarkCyan
Write-Host 'Exemple : admin@ernee.local ou ERNEE\admin' -ForegroundColor DarkCyan
$OldDomainCredential = Get-Credential -Message 'Compte ancien domaine / admin local des PC'

Write-Host ''
Write-Host 'Credential 2/2 - Nouveau domaine' -ForegroundColor Cyan
Write-Host 'Ce compte sert a joindre les PC au nouveau domaine.' -ForegroundColor DarkCyan
Write-Host "IMPORTANT : utilise le format nom@intra.ght53.fr ou INTRA\admin" -ForegroundColor DarkCyan
$NewDomainCredential = Get-Credential -Message 'Compte nouveau domaine - format nom@intra.ght53.fr ou INTRA\admin'

$MigrationScriptBlock = {
    param(
        [string]$NewDomainFqdn,
        [string]$OUPath,
        [pscredential]$OldDomainCredential,
        [pscredential]$NewDomainCredential,
        [bool]$Restart,
        [int]$RestartDelaySeconds,
        [bool]$WhatIfMode
    )

    $ErrorActionPreference = 'Stop'
    $ComputerSystem = Get-CimInstance Win32_ComputerSystem
    $CurrentComputerName = $env:COMPUTERNAME
    $CurrentDomain = $ComputerSystem.Domain
    $PartOfDomain = [bool]$ComputerSystem.PartOfDomain

    if ($PartOfDomain -and ($CurrentDomain -ieq $NewDomainFqdn)) {
        return [PSCustomObject]@{
            ComputerName  = $CurrentComputerName
            CurrentDomain = $CurrentDomain
            PartOfDomain  = $PartOfDomain
            TargetDomain  = $NewDomainFqdn
            Status        = 'SKIPPED - already in target domain'
            Error         = ''
        }
    }

    $NltestOutput = & nltest "/dsgetdc:$NewDomainFqdn" 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Impossible de trouver un controleur de domaine pour $NewDomainFqdn. nltest : $($NltestOutput -join ' | ')"
    }

    if ($WhatIfMode) {
        return [PSCustomObject]@{
            ComputerName  = $CurrentComputerName
            CurrentDomain = $CurrentDomain
            PartOfDomain  = $PartOfDomain
            TargetDomain  = $NewDomainFqdn
            Status        = 'SIMULATION - no change made'
            Error         = ''
        }
    }

    $AddComputerParams = @{
        DomainName  = $NewDomainFqdn
        Credential  = $NewDomainCredential
        OUPath      = $OUPath
        Force       = $true
        ErrorAction = 'Stop'
    }

    if ($PartOfDomain) {
        $AddComputerParams.UnjoinDomainCredential = $OldDomainCredential
    }

    Add-Computer @AddComputerParams

    if ($Restart) {
        shutdown.exe /r /t $RestartDelaySeconds /c "Migration domaine vers $NewDomainFqdn - redemarrage automatique"
        $Status = 'OK - joined, reboot scheduled'
    }
    else {
        $Status = 'OK - joined, manual reboot required'
    }

    [PSCustomObject]@{
        ComputerName  = $CurrentComputerName
        CurrentDomain = $CurrentDomain
        PartOfDomain  = $PartOfDomain
        TargetDomain  = $NewDomainFqdn
        Status        = $Status
        Error         = ''
    }
}

$Results = foreach ($Computer in $Computers) {
    $StartedAt = Get-Date
    $NormalizedComputer = Normalize-ComputerName $Computer
    $IsLocalComputer = ($NormalizedComputer -eq $LocalComputerName)

    Write-Host ''
    Write-Host "[$Computer] Debut du traitement..." -ForegroundColor Yellow

    try {
        if ($IsLocalComputer) {
            Write-Host "[$Computer] Machine locale detectee, execution locale." -ForegroundColor Cyan
            $MigrationResult = & $MigrationScriptBlock `
                $NewDomainFqdn `
                $OUPath `
                $OldDomainCredential `
                $NewDomainCredential `
                ([bool]$Restart) `
                $RestartDelaySeconds `
                $WhatIfMode
        }
        else {
            Write-Host "[$Computer] Machine distante, test WinRM..." -ForegroundColor Cyan
            Test-WSMan -ComputerName $Computer -ErrorAction Stop | Out-Null

            Write-Host "[$Computer] Execution distante via Kerberos..." -ForegroundColor Cyan
            $MigrationResult = Invoke-Command `
                -ComputerName $Computer `
                -Credential $OldDomainCredential `
                -Authentication Kerberos `
                -ScriptBlock $MigrationScriptBlock `
                -ArgumentList $NewDomainFqdn, $OUPath, $OldDomainCredential, $NewDomainCredential, ([bool]$Restart), $RestartDelaySeconds, $WhatIfMode `
                -ErrorAction Stop
        }

        foreach ($Item in $MigrationResult) {
            [PSCustomObject]@{
                RequestedComputer = $Computer
                ReportedComputer  = $Item.ComputerName
                CurrentDomain     = $Item.CurrentDomain
                PartOfDomain      = $Item.PartOfDomain
                TargetDomain      = $Item.TargetDomain
                Status            = $Item.Status
                Error             = $Item.Error
                StartedAt         = $StartedAt
                FinishedAt        = Get-Date
            }

            if ($Item.Status -like 'OK*') {
                Write-Host "[$Computer] $($Item.Status)" -ForegroundColor Green
            }
            elseif ($Item.Status -like 'SIMULATION*') {
                Write-Host "[$Computer] $($Item.Status)" -ForegroundColor DarkYellow
            }
            elseif ($Item.Status -like 'SKIPPED*') {
                Write-Host "[$Computer] $($Item.Status)" -ForegroundColor Yellow
            }
            else {
                Write-Host "[$Computer] $($Item.Status)" -ForegroundColor White
            }
        }
    }
    catch {
        $Message = $_.Exception.Message

        [PSCustomObject]@{
            RequestedComputer = $Computer
            ReportedComputer  = ''
            CurrentDomain     = ''
            PartOfDomain      = ''
            TargetDomain      = $NewDomainFqdn
            Status            = 'KO'
            Error             = $Message
            StartedAt         = $StartedAt
            FinishedAt        = Get-Date
        }

        Write-Host "[$Computer] ERREUR : $Message" -ForegroundColor Red
    }
}

$logDirectory = Split-Path -Parent $LogPath
if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

$resultatsExcel = $Results | Select-Object `
    @{ Name = 'Machine demandée'; Expression = { $_.RequestedComputer } },
    @{ Name = 'Machine détectée'; Expression = { $_.ReportedComputer } },
    @{ Name = 'Domaine actuel'; Expression = { $_.CurrentDomain } },
    @{ Name = 'Membre domaine'; Expression = { if ($_.PartOfDomain -eq $true) { 'Oui' } elseif ($_.PartOfDomain -eq $false) { 'Non' } else { '' } } },
    @{ Name = 'Domaine cible'; Expression = { $_.TargetDomain } },
    @{ Name = 'Statut'; Expression = { $_.Status } },
    @{ Name = 'Erreur'; Expression = { $_.Error } },
    @{ Name = 'Début'; Expression = { if ($_.StartedAt) { $_.StartedAt.ToString('dd/MM/yyyy HH:mm:ss') } } },
    @{ Name = 'Fin'; Expression = { if ($_.FinishedAt) { $_.FinishedAt.ToString('dd/MM/yyyy HH:mm:ss') } } },
    @{ Name = 'Durée (secondes)'; Expression = { if ($_.StartedAt -and $_.FinishedAt) { [Math]::Round(($_.FinishedAt - $_.StartedAt).TotalSeconds, 1) } } }

Export-CsvPourExcel -Donnees @($resultatsExcel) -Chemin $LogPath

Write-Host ''
Write-Host "Termine. Log genere : $LogPath" -ForegroundColor Cyan
$resultatsExcel | Format-Table -AutoSize
