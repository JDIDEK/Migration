#requires -Version 5.1
<#
.SYNOPSIS
    Recrée dans le domaine cible une liste d'utilisateurs provenant d'un autre domaine.
.DESCRIPTION
    Le SamAccountName cible est construit au format prenom.nom, sans accent, en
    minuscules et limité à 20 caractères. En cas de collision, un suffixe numérique
    est ajouté. Les mots de passe et SIDHistory ne peuvent pas être migrés avec les
    cmdlets ActiveDirectory : un mot de passe temporaire cible est donc demandé.
.PREREQUIS
    Module ActiveDirectory (RSAT), accès en lecture au domaine source et délégation
    de création d'utilisateurs dans l'OU cible.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$ListeUtilisateursPath = '.\utilisateurs-migration.txt',
    [string]$ServeurADSource = 'ernee.local',
    [string]$ServeurADCible = 'intra.ght53.fr',
    [string]$OUCible = 'OU=Utilisateurs,OU=HLER,DC=intra,DC=ght53,DC=fr',
    [string]$DomaineUPNCible = 'intra.ght53.fr',
    [string]$LogPath = '.\migration-utilisateurs-resultats.csv',
    [PSCredential]$IdentifiantsSource,
    [PSCredential]$IdentifiantsCible,
    [string]$IdentifiantsCiblePath = '',
    [SecureString]$MotDePasseTemporaire,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message, [string]$Niveau = 'INFO')
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Niveau, $Message)
}

function Test-Simulation {
    [bool]($DryRun -or $WhatIfPreference)
}

function ConvertTo-NomCompte {
    param([Parameter(Mandatory)][string]$Valeur)

    $decompose = $Valeur.Trim().ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD)
    $sansAccents = New-Object Text.StringBuilder
    foreach ($caractere in $decompose.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($caractere) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sansAccents.Append($caractere)
        }
    }

    $normalise = $sansAccents.ToString().Normalize([Text.NormalizationForm]::FormC)
    $normalise = $normalise -replace '[^a-z0-9]+', '.'
    $normalise.Trim('.')
}

function Get-SamAccountNameDisponible {
    param(
        [Parameter(Mandatory)][string]$Prenom,
        [Parameter(Mandatory)][string]$Nom,
        [Parameter(Mandatory)][hashtable]$ParametresCible
    )

    $base = "$(ConvertTo-NomCompte $Prenom).$(ConvertTo-NomCompte $Nom)".Trim('.')
    if ([string]::IsNullOrWhiteSpace($base)) {
        throw "Impossible de construire un SamAccountName à partir de '$Prenom $Nom'."
    }

    for ($numero = 1; $numero -le 999; $numero++) {
        $suffixe = if ($numero -eq 1) { '' } else { [string]$numero }
        $longueurBase = 20 - $suffixe.Length
        $candidat = $base.Substring(0, [Math]::Min($base.Length, $longueurBase)).TrimEnd('.') + $suffixe
        $filtre = $candidat.Replace("'", "''")
        $existant = Get-ADUser -Filter "SamAccountName -eq '$filtre'" @ParametresCible
        if (-not $existant) { return $candidat }
    }

    throw "Aucun SamAccountName disponible pour '$Prenom $Nom' après 999 tentatives."
}

function Confirm-Execution {
    param([int]$Nombre)
    if (Test-Simulation) {
        Write-Log 'Simulation : aucune modification ne sera effectuée.'
        return
    }
    if ((Read-Host "Tapez OUI pour migrer $Nombre utilisateur(s) vers '$ServeurADCible'") -cne 'OUI') {
        Write-Log 'Opération annulée.' 'ATTENTION'
        exit 0
    }
}

Import-Module ActiveDirectory -ErrorAction Stop

if (-not (Test-Path -LiteralPath $ListeUtilisateursPath -PathType Leaf)) {
    throw "Fichier de liste introuvable : $ListeUtilisateursPath"
}

$identites = @(Get-Content -LiteralPath $ListeUtilisateursPath |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -and -not $_.StartsWith('#') } |
    Sort-Object -Unique)

if ($identites.Count -eq 0) {
    throw "Aucun utilisateur à migrer dans '$ListeUtilisateursPath'. Ajoutez un SamAccountName ou UPN par ligne."
}

if (-not $IdentifiantsCible -and $IdentifiantsCiblePath -and (Test-Path -LiteralPath $IdentifiantsCiblePath -PathType Leaf)) {
    $IdentifiantsCible = Import-Clixml -LiteralPath $IdentifiantsCiblePath
}
if (-not $IdentifiantsSource) {
    $IdentifiantsSource = Get-Credential -Message "Compte autorisé à lire les utilisateurs sur $ServeurADSource"
}
if (-not $IdentifiantsCible) {
    $IdentifiantsCible = Get-Credential -Message "Compte autorisé à créer les utilisateurs sur $ServeurADCible"
}

$parametresSource = @{
    Server      = $ServeurADSource
    Credential  = $IdentifiantsSource
    ErrorAction = 'Stop'
}
$parametresCible = @{
    Server      = $ServeurADCible
    Credential  = $IdentifiantsCible
    ErrorAction = 'Stop'
}

try {
    $null = Get-ADDomain @parametresSource
    Write-Log "Connexion au domaine source '$ServeurADSource' : OK" 'OK'
}
catch {
    throw "Impossible de lire le domaine source '$ServeurADSource' : $($_.Exception.Message)"
}

try {
    $null = Get-ADDomain @parametresCible
    $null = Get-ADOrganizationalUnit -Identity $OUCible @parametresCible
    Write-Log "Connexion au domaine cible '$ServeurADCible' et OU cible : OK" 'OK'
}
catch {
    throw "Impossible de valider le domaine ou l'OU cible '$OUCible' : $($_.Exception.Message)"
}

Write-Host ''
Write-Host "Utilisateurs demandés : $($identites.Count)" -ForegroundColor Cyan
Write-Host "Domaine source        : $ServeurADSource" -ForegroundColor Cyan
Write-Host "Domaine cible         : $ServeurADCible" -ForegroundColor Cyan
Write-Host "OU cible              : $OUCible" -ForegroundColor Cyan
Write-Host "Norme de compte       : prenom.nom" -ForegroundColor Cyan
Write-Host "Mode simulation       : $(Test-Simulation)" -ForegroundColor Cyan

Confirm-Execution -Nombre $identites.Count
if (-not (Test-Simulation) -and -not $MotDePasseTemporaire) {
    $MotDePasseTemporaire = Read-Host 'Mot de passe temporaire commun aux comptes migrés' -AsSecureString
}

$resultats = foreach ($identiteSource in $identites) {
    $debut = Get-Date
    try {
        $utilisateurSource = Get-ADUser -Identity $identiteSource `
            -Properties GivenName,Surname,DisplayName,Enabled,Description,Department,Title,Company,Office,OfficePhone,MobilePhone `
            @parametresSource

        if ([string]::IsNullOrWhiteSpace($utilisateurSource.GivenName) -or [string]::IsNullOrWhiteSpace($utilisateurSource.Surname)) {
            throw "Les attributs GivenName (prénom) et Surname (nom) sont obligatoires."
        }

        $marqueurSource = "MigrationSourceGuid=$($utilisateurSource.ObjectGUID)"
        $dejaMigres = @(Get-ADUser -LDAPFilter "(info=$marqueurSource)" -Properties UserPrincipalName @parametresCible)
        if ($dejaMigres.Count -gt 1) {
            throw "Plusieurs comptes cibles portent le marqueur '$marqueurSource'."
        }
        if ($dejaMigres.Count -eq 1) {
            Write-Log "'$identiteSource' est déjà migré vers '$($dejaMigres[0].SamAccountName)'." 'ATTENTION'
            [PSCustomObject]@{
                IdentiteSource       = $identiteSource
                SamAccountNameSource = $utilisateurSource.SamAccountName
                Prenom               = $utilisateurSource.GivenName
                Nom                  = $utilisateurSource.Surname
                SamAccountNameCible  = $dejaMigres[0].SamAccountName
                UserPrincipalName    = $dejaMigres[0].UserPrincipalName
                Statut               = 'IGNORÉ - déjà migré'
                Erreur               = ''
                Debut                = $debut
                Fin                  = Get-Date
            }
            continue
        }

        $samCible = Get-SamAccountNameDisponible `
            -Prenom $utilisateurSource.GivenName `
            -Nom $utilisateurSource.Surname `
            -ParametresCible $parametresCible
        $upnCible = "$samCible@$DomaineUPNCible"
        $nomAffiche = if ($utilisateurSource.DisplayName) {
            $utilisateurSource.DisplayName
        }
        else {
            "$($utilisateurSource.GivenName) $($utilisateurSource.Surname)"
        }

        if (Test-Simulation) {
            $statut = 'SIMULATION'
            Write-Log "SIMULATION - '$identiteSource' deviendrait '$samCible' ($upnCible)."
        }
        elseif ($PSCmdlet.ShouldProcess($upnCible, "Créer le compte cible depuis '$identiteSource'")) {
            $creation = @{
                Name                  = $nomAffiche
                GivenName             = $utilisateurSource.GivenName
                Surname               = $utilisateurSource.Surname
                DisplayName           = $nomAffiche
                SamAccountName        = $samCible
                UserPrincipalName     = $upnCible
                Path                  = $OUCible
                AccountPassword       = $MotDePasseTemporaire
                Enabled               = [bool]$utilisateurSource.Enabled
                ChangePasswordAtLogon = $true
                PassThru              = $true
                OtherAttributes       = @{ info = $marqueurSource }
            }

            foreach ($propriete in @('Description','Department','Title','Company','Office','OfficePhone','MobilePhone')) {
                if (-not [string]::IsNullOrWhiteSpace([string]$utilisateurSource.$propriete)) {
                    $creation[$propriete] = $utilisateurSource.$propriete
                }
            }
            foreach ($cle in $parametresCible.Keys) { $creation[$cle] = $parametresCible[$cle] }

            $nouveau = New-ADUser @creation
            $verification = Get-ADUser -Identity $nouveau.ObjectGUID -Properties UserPrincipalName @parametresCible
            $statut = 'OK'
            Write-Log "Utilisateur créé : $($verification.SamAccountName) - $($verification.UserPrincipalName)" 'OK'
        }
        else {
            $statut = 'ANNULÉ'
        }

        [PSCustomObject]@{
            IdentiteSource       = $identiteSource
            SamAccountNameSource = $utilisateurSource.SamAccountName
            Prenom               = $utilisateurSource.GivenName
            Nom                  = $utilisateurSource.Surname
            SamAccountNameCible  = $samCible
            UserPrincipalName    = $upnCible
            Statut               = $statut
            Erreur               = ''
            Debut                = $debut
            Fin                  = Get-Date
        }
    }
    catch {
        Write-Log "Échec pour '$identiteSource' : $($_.Exception.Message)" 'ERREUR'
        [PSCustomObject]@{
            IdentiteSource       = $identiteSource
            SamAccountNameSource = ''
            Prenom               = ''
            Nom                  = ''
            SamAccountNameCible  = ''
            UserPrincipalName    = ''
            Statut               = 'KO'
            Erreur               = $_.Exception.Message
            Debut                = $debut
            Fin                  = Get-Date
        }
    }
}

$dossierLog = Split-Path -Parent $LogPath
if ($dossierLog -and -not (Test-Path -LiteralPath $dossierLog -PathType Container)) {
    New-Item -ItemType Directory -Path $dossierLog -Force | Out-Null
}
$resultats | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Log "Traitement terminé. Log : $LogPath" 'OK'
$resultats | Format-Table IdentiteSource,SamAccountNameCible,UserPrincipalName,Statut -AutoSize
