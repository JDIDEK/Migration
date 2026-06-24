#requires -Version 5.1
<#
.SYNOPSIS
    Interface graphique de configuration et de lancement des tests de migration.
.DESCRIPTION
    Toutes les valeurs propres à l'environnement sont modifiables dans la grille.
    Le mode DryRun est actif par défaut. Les scripts s'ouvrent dans une console
    séparée afin de permettre les confirmations et saisies sécurisées.
#>
[CmdletBinding()]
param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()

$racineScripts = $PSScriptRoot
if (-not $racineScripts) { $racineScripts = Split-Path -Parent $MyInvocation.MyCommand.Path }
$powershellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

function Test-CompteIntra {
    param([string]$NomUtilisateur)
    $NomUtilisateur -match '^(INTRA\\.+|[^\\@]+@intra\.ght53\.fr)$'
}
function ConvertFrom-SecureStringToPlainText {
    param([Security.SecureString]$SecureString)
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally {
        if ($bstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }
}
function Test-IdentifiantsIntra {
    param([PSCredential]$Credential)

    if (-not $Credential) { return $false }
    try {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement -ErrorAction Stop
        $nomUtilisateur = $Credential.UserName
        if ($nomUtilisateur -like '*\*') { $nomUtilisateur = ($nomUtilisateur -split '\\',2)[1] }
        elseif ($nomUtilisateur -like '*@*') { $nomUtilisateur = ($nomUtilisateur -split '@',2)[0] }

        $motDePasse = ConvertFrom-SecureStringToPlainText $Credential.Password
        $contexte = New-Object DirectoryServices.AccountManagement.PrincipalContext(
            [DirectoryServices.AccountManagement.ContextType]::Domain,
            'intra.ght53.fr'
        )
        $contexte.ValidateCredentials($nomUtilisateur, $motDePasse)
    }
    catch {
        $false
    }
    finally {
        $motDePasse = $null
        if ($contexte) { $contexte.Dispose() }
    }
}

$script:identifiantsExecution = $null
do {
    $script:identifiantsExecution = Get-Credential -Message "Compte INTRA utilise pour lancer les scripts, ex: INTRA\admin ou admin@intra.ght53.fr"
    if (-not $script:identifiantsExecution) { throw 'Compte INTRA obligatoire pour lancer l interface.' }
    if (-not (Test-CompteIntra $script:identifiantsExecution.UserName)) {
        [Windows.Forms.MessageBox]::Show(
            "Le compte doit etre au format INTRA\compte ou compte@intra.ght53.fr.",
            'Compte INTRA requis',
            'OK',
            'Warning'
        ) | Out-Null
        $script:identifiantsExecution = $null
    }
    elseif (-not (Test-IdentifiantsIntra $script:identifiantsExecution)) {
        [Windows.Forms.MessageBox]::Show(
            "Authentification refusee par intra.ght53.fr. L interface ne peut pas demarrer avec ce compte.",
            'Compte INTRA invalide',
            'OK',
            'Error'
        ) | Out-Null
        $script:identifiantsExecution = $null
    }
} until ($script:identifiantsExecution)
$script:cheminIdentifiantsExecution = Join-Path $env:TEMP 'MigrationErnee-IdentifiantsIntra.clixml'
$script:identifiantsExecution | Export-Clixml -LiteralPath $script:cheminIdentifiantsExecution

function New-Definition {
    param(
        [string]$Nom,
        [string]$Libelle,
        [string]$Valeur,
        [string]$Aide,
        [bool]$Obligatoire = $true,
        [ValidateSet('Texte','Entier','Login')][string]$Type = 'Texte'
    )
    [pscustomobject]@{ Nom=$Nom; Libelle=$Libelle; Valeur=$Valeur; Aide=$Aide; Obligatoire=$Obligatoire; Type=$Type }
}

$descriptions = @{
    '00-Verifications-Prerequis.ps1'    = 'Contrôle les modules Windows et les cmdlets Exchange. Lecture seule.'
    '01-Test-ScopeDHCP.ps1'             = 'Crée un scope DHCP inactif et réserve une IP pour la VM locale.'
    '02-Test-MigrationMachineDomaine.ps1' = 'Bascule une liste de machines du domaine ERNEE vers le domaine INTRA, avec log CSV.'
    '03-MigrationUtilisateursAD.ps1'    = 'Recrée une liste d utilisateurs dans INTRA avec un SamAccountName normalisé au format prenom.nom.'
    '04-Test-CreationUtilisateurAD.ps1' = "Crée l'utilisateur dans l'OU choisie. Le mot de passe est demandé dans la console."
    '05-Test-PartageFichiers.ps1'       = 'Crée le dossier, le partage SMB et les permissions du groupe AD.'
    '06-Test-ProfilItinerant.ps1'       = "Prépare le stockage du profil et configure le ProfilePath de l'utilisateur."
    '07-Test-Imprimante.ps1'            = "Crée le port, le pilote et l'imprimante partagée de test."
    '08-Test-Robocopy.ps1'              = 'Copie les fichiers avec /COPYALL ou simule Robocopy avec /L.'
    '09-Test-RemapLecteur.ps1'           = 'Mappe un lecteur réseau persistant dans la session de la VM.'
    '10-Test-CreationBoiteExchange.ps1' = "Active une boîte Exchange on-premise pour l'utilisateur AD."
    '11-Test-ExportPST.ps1'             = "Exporte la boîte vers PST, suit l'avancement et nettoie la requête après succès."
    '12-Test-MigrationBoite.ps1'        = 'Lance ou suit un déplacement Exchange cross-forest. Scénario alternatif au script 10.'
}

$definitions = @{
    '00-Verifications-Prerequis.ps1' = @(
        (New-Definition 'ModulesRequis' 'Modules requis' 'ActiveDirectory;DhcpServer;PrintManagement;SmbShare' 'Séparer les noms par un point-virgule.'),
        (New-Definition 'CommandesExchange' 'Cmdlets Exchange' 'Get-Mailbox;Enable-Mailbox;New-MailboxExportRequest;Get-MailboxExportRequest;New-MoveRequest;Get-MoveRequest' 'Séparer les noms par un point-virgule.')
    )
    '05-Test-PartageFichiers.ps1' = @(
        (New-Definition 'CheminLocal' 'Chemin local' 'D:\Partages\Test' 'Dossier à créer sur le serveur de fichiers.'),
        (New-Definition 'NomPartage' 'Nom du partage' 'Migration-Test' 'Nom SMB sans barre oblique.'),
        (New-Definition 'DescriptionPartage' 'Description' 'Partage de validation migration - TEST UNIQUEMENT' 'Description visible dans la gestion des partages.'),
        (New-Definition 'GroupeAD' 'Groupe AD RW' 'INTRA\GG_Migration_Test_RW' 'Format DOMAINE\Groupe.')
    )
    '06-Test-ProfilItinerant.ps1' = @(
        (New-Definition 'UtilisateurTest' 'Utilisateur AD' 'migration.test' "SamAccountName de l'utilisateur."),
        (New-Definition 'ServeurFichiers' 'Serveur de fichiers' 'SRV-FICHIERS-TEST' 'Nom DNS utilisé pour construire le chemin UNC.'),
        (New-Definition 'RacineLocaleProfils' 'Racine locale' 'D:\Profils-Test' 'Dossier local sur le serveur de fichiers.'),
        (New-Definition 'NomPartageProfils' 'Partage profils' 'Profils-Test$' 'Nom du partage SMB, éventuellement caché avec $.')
    )
    '07-Test-Imprimante.ps1' = @(
        (New-Definition 'NomPort' 'Nom du port' 'IP_TEST_192.0.2.10' 'Nom du port TCP/IP.'),
        (New-Definition 'AdresseIPTest' 'Adresse IP' '192.0.2.10' 'Adresse documentaire ou adresse de test.'),
        (New-Definition 'NomPilote' 'Pilote' 'Generic / Text Only' 'Nom exact du pilote installé ou présent dans le magasin.'),
        (New-Definition 'NomImprimante' 'Nom imprimante' 'Imprimante-Migration-Test' "Nom local de la file d'impression."),
        (New-Definition 'NomPartage' 'Nom du partage' 'IMP-Migration-Test' "Nom réseau de l'imprimante partagée.")
    )
    '01-Test-ScopeDHCP.ps1' = @(
        (New-Definition 'ServeurDHCP' 'Serveur DHCP' 'SRV-DHCP-TEST' 'Nom DNS du serveur DHCP.'),
        (New-Definition 'NomScope' 'Nom du scope' 'TEST-MIGRATION-192.168.99.0' 'Libellé clairement identifié comme test.'),
        (New-Definition 'ScopeId' 'ID du scope' '192.168.99.0' 'Adresse réseau du scope.'),
        (New-Definition 'DebutPlage' 'Début de plage' '192.168.99.100' 'Première adresse distribuable.'),
        (New-Definition 'FinPlage' 'Fin de plage' '192.168.99.150' 'Dernière adresse distribuable.'),
        (New-Definition 'Masque' 'Masque' '255.255.255.0' 'Masque IPv4 du scope.'),
        (New-Definition 'IPReserveeVM' 'IP réservée VM' '192.168.99.110' 'Adresse de réservation comprise dans le scope.'),
        (New-Definition 'NomCarteReseau' 'Carte réseau VM' '' "Vide : détecter l'unique carte active." $false)
    )
    '04-Test-CreationUtilisateurAD.ps1' = @(
        (New-Definition 'Prenom' 'Prénom' 'Jean' 'Prénom du compte de test.'),
        (New-Definition 'Nom' 'Nom' 'Migration' 'Nom du compte de test.'),
        (New-Definition 'Login' 'Login' 'migration.test' 'SamAccountName.' $true 'Login'),
        (New-Definition 'DomaineUPN' 'Domaine UPN' 'intra.ght53.fr' 'Suffixe UPN, sans @.'),
        (New-Definition 'OUTest' 'OU de destination' 'OU=Utilisateurs,OU=HLER,DC=intra,DC=ght53,DC=fr' "Distinguished Name complet de l'OU."),
        (New-Definition 'ServeurAD' 'Serveur AD' 'intra.ght53.fr' 'Nom DNS du domaine ou du contrôleur de domaine cible.')
    )
    '10-Test-CreationBoiteExchange.ps1' = @(
        (New-Definition 'UtilisateurTest' 'Utilisateur AD' 'migration.test' "Identité Exchange de l'utilisateur."),
        (New-Definition 'AliasMessagerie' 'Alias messagerie' 'migration.test' 'Alias Exchange souhaité.'),
        (New-Definition 'DomaineUPN' 'Domaine UPN' 'intra.ght53.fr' "Suffixe UPN utilisé si l'identité courte n'est pas trouvée."),
        (New-Definition 'DomaineSMTP' 'Domaine SMTP' 'intra.ght53.fr' 'Domaine de l''adresse SMTP principale.'),
        (New-Definition 'AdresseSMTPPrincipale' 'SMTP principale' '' 'Vide : alias@DomaineSMTP. Exemple : migration.test@intra.ght53.fr.' $false),
        (New-Definition 'BaseDeDonnees' 'Base Exchange' '' 'Vide : laisser Exchange sélectionner la base.' $false),
        (New-Definition 'ServeurExchange' 'Serveur Exchange' 'ght-exchangew-1' 'Nom du serveur Exchange pour importer une session distante.'),
        (New-Definition 'UriPowerShellExchange' 'URI PowerShell' 'http://ght-exchangew-1/PowerShell/' 'URI PowerShell Exchange distante.'),
        (New-Definition 'ControleurDomaineExchange' 'DC Exchange' '' 'Optionnel. Exemple : GHT-ADW-2.intra.ght53.fr.' $false)
    )
    '11-Test-ExportPST.ps1' = @(
        (New-Definition 'BoiteTest' 'Boîte à exporter' 'migration.test' 'Identité Exchange de la boîte.'),
        (New-Definition 'DossierPST' 'Dossier PST UNC' '\\SRV-FICHIERS-TEST\PST-Test$' 'Chemin UNC accessible par Exchange Trusted Subsystem.'),
        (New-Definition 'NomFichierPST' 'Nom du PST' 'migration.test.pst' 'Nom de fichier avec extension .pst.'),
        (New-Definition 'NomRequete' 'Nom de requête' 'Export-Migration-Test' 'Nom Exchange unique pour cette requête.'),
        (New-Definition 'IntervalleSecondes' 'Intervalle (secondes)' '15' 'Fréquence du suivi, de 1 à 3600.' $true 'Entier'),
        (New-Definition 'DelaiMaximumMinutes' 'Délai maximal (minutes)' '120' 'Délai de suivi, de 1 à 1440.' $true 'Entier')
    )
    '12-Test-MigrationBoite.ps1' = @(
        (New-Definition 'IdentiteCible' 'Identité cible' 'migration.test@intra.ght53.fr' 'MailUser préparé dans la forêt cible.'),
        (New-Definition 'ServeurExchangeSource' 'MRS Proxy source' 'exchange-source.ancien-domaine.local' 'FQDN Exchange/MRS Proxy réel de la source.'),
        (New-Definition 'DomaineLivraisonCible' 'Domaine cible' 'intra.ght53.fr' 'TargetDeliveryDomain du déplacement.'),
        (New-Definition 'BaseCible' 'Base cible' '' 'Vide : laisser Exchange sélectionner la base.' $false)
    )
    '09-Test-RemapLecteur.ps1' = @(
        (New-Definition 'LettreLecteur' 'Lettre du lecteur' 'Z' 'Une lettre de A à Z, avec ou sans deux-points.'),
        (New-Definition 'CheminUNC' 'Chemin UNC' '\\SRV-FICHIERS-TEST\Migration-Test' 'Partage réseau à mapper.')
    )
    '08-Test-Robocopy.ps1' = @(
        (New-Definition 'DossierSource' 'Dossier source' 'C:\Migration-Test\Source' 'Petit jeu de fichiers de test.'),
        (New-Definition 'DossierDestination' 'Destination UNC' '\\SRV-FICHIERS-TEST\Migration-Test\Robocopy' 'Dossier cible de la copie.'),
        (New-Definition 'DossierLogs' 'Dossier des logs' 'C:\Migration-Test\Logs' 'Emplacement local des journaux Robocopy.')
    )
    '02-Test-MigrationMachineDomaine.ps1' = @(
        (New-Definition 'ComputerListPath' 'Liste machines' '.\machines-test.txt' 'Fichier texte : une machine par ligne.'),
        (New-Definition 'NewDomainFqdn' 'Domaine cible FQDN' 'intra.ght53.fr' 'Nom DNS complet du nouveau domaine.'),
        (New-Definition 'OUPath' 'OU poste cible' 'OU=Postes,OU=HLER,DC=intra,DC=ght53,DC=fr' 'Distinguished Name complet de l''OU des postes.'),
        (New-Definition 'LogPath' 'Log CSV' '.\migration-domaine-resultats.csv' 'Chemin du fichier CSV de resultat.'),
        (New-Definition 'RestartDelaySeconds' 'Delai reboot' '30' 'Delai avant redemarrage automatique, en secondes.' $true 'Entier')
    )
    '03-MigrationUtilisateursAD.ps1' = @(
        (New-Definition 'ListeUtilisateursPath' 'Liste utilisateurs' '.\utilisateurs-migration.txt' 'Un SamAccountName ou UPN source par ligne.'),
        (New-Definition 'ServeurADSource' 'AD source' 'ernee.local' 'Domaine ou contrôleur de domaine source.'),
        (New-Definition 'ServeurADCible' 'AD cible' 'intra.ght53.fr' 'Domaine ou contrôleur de domaine cible.'),
        (New-Definition 'OUCible' 'OU cible' 'OU=Utilisateurs,OU=HLER,DC=intra,DC=ght53,DC=fr' 'OU dans laquelle créer les comptes.'),
        (New-Definition 'DomaineUPNCible' 'UPN cible' 'intra.ght53.fr' 'Suffixe UPN des nouveaux comptes.'),
        (New-Definition 'LogPath' 'Log CSV' '.\migration-utilisateurs-resultats.csv' 'Résultat détaillé de la migration.')
    )
}

function Get-ScriptsDisponibles {
    @(Get-ChildItem -LiteralPath $racineScripts -Filter '*.ps1' -File |
        Where-Object Name -Match '^\d{2}-' |
        Sort-Object Name)
}

function ConvertTo-ArgumentNatif {
    param([AllowEmptyString()][string]$Valeur)
    if ($Valeur -match '["\r\n]') { throw 'Les guillemets et retours à la ligne ne sont pas autorisés.' }
    $echappe = $Valeur
    $nbBarresFinales = 0
    for ($i=$echappe.Length-1; $i -ge 0 -and $echappe[$i] -eq '\'; $i--) { $nbBarresFinales++ }
    if ($nbBarresFinales -gt 0) { $echappe += ('\' * $nbBarresFinales) }
    '"' + $echappe + '"'
}

$script:valeursSession = @{}
$script:optionsSession = @{}
$script:scriptCharge = $null

$form = New-Object Windows.Forms.Form
$form.Text = 'Configuration des tests de migration AD / Exchange'
$form.StartPosition = 'CenterScreen'
$form.Size = [Drawing.Size]::new(1140,760)
$form.MinimumSize = [Drawing.Size]::new(1140,760)
$form.MaximizeBox = $false
$form.Font = [Drawing.Font]::new('Segoe UI',9)
$form.BackColor = [Drawing.Color]::FromArgb(245,247,250)

$titre = New-Object Windows.Forms.Label
$titre.Text = 'Tests de migration — configuration et lancement'
$titre.Location = [Drawing.Point]::new(20,14)
$titre.Size = [Drawing.Size]::new(1000,38)
$titre.Font = [Drawing.Font]::new('Segoe UI Semibold',18)
$titre.ForeColor = [Drawing.Color]::FromArgb(35,55,75)
$form.Controls.Add($titre)

$intro = New-Object Windows.Forms.Label
$intro.Text = "Toutes les valeurs sont modifiables ci-dessous. Scripts lances avec : $($script:identifiantsExecution.UserName)"
$intro.Location = [Drawing.Point]::new(23,51)
$intro.Size = [Drawing.Size]::new(900,24)
$intro.ForeColor = [Drawing.Color]::DimGray
$form.Controls.Add($intro)

$groupeScripts = New-Object Windows.Forms.GroupBox
$groupeScripts.Text = 'Scripts disponibles'
$groupeScripts.Location = [Drawing.Point]::new(20,80)
$groupeScripts.Size = [Drawing.Size]::new(390,555)
$form.Controls.Add($groupeScripts)

$listeScripts = New-Object Windows.Forms.ListBox
$listeScripts.Location = [Drawing.Point]::new(12,24)
$listeScripts.Size = [Drawing.Size]::new(365,478)
$listeScripts.IntegralHeight = $false
$listeScripts.Font = [Drawing.Font]::new('Consolas',9)
$groupeScripts.Controls.Add($listeScripts)

$boutonActualiser = New-Object Windows.Forms.Button
$boutonActualiser.Text = 'Actualiser la liste'
$boutonActualiser.Location = [Drawing.Point]::new(12,512)
$boutonActualiser.Size = [Drawing.Size]::new(177,30)
$groupeScripts.Controls.Add($boutonActualiser)

$boutonDefauts = New-Object Windows.Forms.Button
$boutonDefauts.Text = 'Valeurs par défaut'
$boutonDefauts.Location = [Drawing.Point]::new(200,512)
$boutonDefauts.Size = [Drawing.Size]::new(177,30)
$groupeScripts.Controls.Add($boutonDefauts)

$groupeDescription = New-Object Windows.Forms.GroupBox
$groupeDescription.Text = 'Description'
$groupeDescription.Location = [Drawing.Point]::new(425,80)
$groupeDescription.Size = [Drawing.Size]::new(685,105)
$form.Controls.Add($groupeDescription)

$texteDescription = New-Object Windows.Forms.TextBox
$texteDescription.Location = [Drawing.Point]::new(13,23)
$texteDescription.Size = [Drawing.Size]::new(658,66)
$texteDescription.Multiline = $true
$texteDescription.ReadOnly = $true
$texteDescription.BackColor = [Drawing.Color]::White
$groupeDescription.Controls.Add($texteDescription)

$groupeMode = New-Object Windows.Forms.GroupBox
$groupeMode.Text = "Mode d'exécution"
$groupeMode.Location = [Drawing.Point]::new(425,195)
$groupeMode.Size = [Drawing.Size]::new(685,95)
$form.Controls.Add($groupeMode)

$radioDryRun = New-Object Windows.Forms.RadioButton
$radioDryRun.Text = 'DryRun — aucune modification (recommandé)'
$radioDryRun.Location = [Drawing.Point]::new(15,24)
$radioDryRun.Size = [Drawing.Size]::new(310,24)
$radioDryRun.Checked = $true
$groupeMode.Controls.Add($radioDryRun)

$radioReel = New-Object Windows.Forms.RadioButton
$radioReel.Text = 'Exécution réelle'
$radioReel.Location = [Drawing.Point]::new(350,24)
$radioReel.Size = [Drawing.Size]::new(155,24)
$radioReel.ForeColor = [Drawing.Color]::DarkRed
$groupeMode.Controls.Add($radioReel)

$caseValidation = New-Object Windows.Forms.CheckBox
$caseValidation.Text = "J'ai vérifié tous les paramètres affichés ci-dessous"
$caseValidation.Location = [Drawing.Point]::new(15,56)
$caseValidation.Size = [Drawing.Size]::new(450,24)
$caseValidation.Enabled = $false
$groupeMode.Controls.Add($caseValidation)

$groupeConfiguration = New-Object Windows.Forms.GroupBox
$groupeConfiguration.Text = 'Paramètres du script sélectionné'
$groupeConfiguration.Location = [Drawing.Point]::new(425,300)
$groupeConfiguration.Size = [Drawing.Size]::new(685,335)
$form.Controls.Add($groupeConfiguration)

$grille = New-Object Windows.Forms.DataGridView
$grille.Location = [Drawing.Point]::new(12,23)
$grille.Size = [Drawing.Size]::new(660,248)
$grille.AllowUserToAddRows = $false
$grille.AllowUserToDeleteRows = $false
$grille.AllowUserToResizeRows = $false
$grille.RowHeadersVisible = $false
$grille.SelectionMode = 'CellSelect'
$grille.EditMode = 'EditOnEnter'
$grille.BackgroundColor = [Drawing.Color]::White
$grille.AutoSizeRowsMode = 'AllCells'
[void]$grille.Columns.Add('Libelle','Paramètre')
[void]$grille.Columns.Add('Valeur','Valeur modifiable')
[void]$grille.Columns.Add('Aide','Aide')
$grille.Columns[0].ReadOnly = $true
$grille.Columns[0].Width = 155
$grille.Columns[1].Width = 225
$grille.Columns[2].ReadOnly = $true
$grille.Columns[2].Width = 255
$grille.Columns[2].DefaultCellStyle.WrapMode = 'True'
$groupeConfiguration.Controls.Add($grille)

$caseNePasAttendre = New-Object Windows.Forms.CheckBox
$caseNePasAttendre.Text = "Ne pas attendre la fin de l'export PST"
$caseNePasAttendre.Location = [Drawing.Point]::new(15,287)
$caseNePasAttendre.Size = [Drawing.Size]::new(280,24)
$caseNePasAttendre.Visible = $false
$groupeConfiguration.Controls.Add($caseNePasAttendre)

$caseRelancer = New-Object Windows.Forms.CheckBox
$caseRelancer.Text = 'Supprimer puis recréer la MoveRequest existante'
$caseRelancer.Location = [Drawing.Point]::new(15,287)
$caseRelancer.Size = [Drawing.Size]::new(370,24)
$caseRelancer.ForeColor = [Drawing.Color]::DarkRed
$caseRelancer.Visible = $false
$groupeConfiguration.Controls.Add($caseRelancer)

$caseSimulationRobocopy = New-Object Windows.Forms.CheckBox
$caseSimulationRobocopy.Text = 'Simulation Robocopy /L — seul le log est créé'
$caseSimulationRobocopy.Location = [Drawing.Point]::new(15,287)
$caseSimulationRobocopy.Size = [Drawing.Size]::new(370,24)
$caseSimulationRobocopy.Visible = $false
$groupeConfiguration.Controls.Add($caseSimulationRobocopy)

$caseRedemarrer = New-Object Windows.Forms.CheckBox
$caseRedemarrer.Text = 'Redémarrer les machines après jonction au domaine'
$caseRedemarrer.Location = [Drawing.Point]::new(15,287)
$caseRedemarrer.Size = [Drawing.Size]::new(370,24)
$caseRedemarrer.ForeColor = [Drawing.Color]::DarkRed
$caseRedemarrer.Visible = $false
$groupeConfiguration.Controls.Add($caseRedemarrer)

$labelSaisieSecurisee = New-Object Windows.Forms.Label
$labelSaisieSecurisee.Text = 'Les mots de passe et identifiants sont demandés dans la console ; ils ne sont jamais conservés ici.'
$labelSaisieSecurisee.Location = [Drawing.Point]::new(15,287)
$labelSaisieSecurisee.Size = [Drawing.Size]::new(645,32)
$labelSaisieSecurisee.ForeColor = [Drawing.Color]::DimGray
$groupeConfiguration.Controls.Add($labelSaisieSecurisee)

$boutonDossier = New-Object Windows.Forms.Button
$boutonDossier.Text = 'Ouvrir le dossier'
$boutonDossier.Location = [Drawing.Point]::new(20,655)
$boutonDossier.Size = [Drawing.Size]::new(145,36)
$form.Controls.Add($boutonDossier)

$statut = New-Object Windows.Forms.Label
$statut.Text = 'Prêt — DryRun est activé.'
$statut.Location = [Drawing.Point]::new(185,663)
$statut.Size = [Drawing.Size]::new(645,28)
$statut.ForeColor = [Drawing.Color]::DarkGreen
$form.Controls.Add($statut)

$boutonExecuter = New-Object Windows.Forms.Button
$boutonExecuter.Text = 'Exécuter'
$boutonExecuter.Location = [Drawing.Point]::new(845,653)
$boutonExecuter.Size = [Drawing.Size]::new(125,39)
$boutonExecuter.Font = [Drawing.Font]::new('Segoe UI Semibold',10)
$boutonExecuter.BackColor = [Drawing.Color]::FromArgb(40,115,180)
$boutonExecuter.ForeColor = [Drawing.Color]::White
$boutonExecuter.FlatStyle = 'Flat'
$form.Controls.Add($boutonExecuter)

$boutonFermer = New-Object Windows.Forms.Button
$boutonFermer.Text = 'Fermer'
$boutonFermer.Location = [Drawing.Point]::new(985,653)
$boutonFermer.Size = [Drawing.Size]::new(125,39)
$form.Controls.Add($boutonFermer)

function Save-ConfigurationCourante {
    if (-not $scriptCharge) { return }
    $valeurs = @{}
    foreach ($ligne in $grille.Rows) {
        $definition = $ligne.Tag
        if ($definition) { $valeurs[$definition.Nom] = [string]$ligne.Cells[1].Value }
    }
    $valeursSession[$scriptCharge] = $valeurs
    $optionsSession[$scriptCharge] = @{
        NePasAttendre = $caseNePasAttendre.Checked
        Relancer = $caseRelancer.Checked
        Simulation = $caseSimulationRobocopy.Checked
        Redemarrer = $caseRedemarrer.Checked
    }
}

function Show-Configuration {
    Save-ConfigurationCourante
    $script:scriptCharge = [string]$listeScripts.SelectedItem
    $grille.Rows.Clear()
    $caseNePasAttendre.Visible = $false
    $caseRelancer.Visible = $false
    $caseSimulationRobocopy.Visible = $false
    $caseRedemarrer.Visible = $false
    $labelSaisieSecurisee.Visible = $true
    $caseValidation.Checked = $false
    if (-not $scriptCharge) { $texteDescription.Text='Aucun script disponible.'; $boutonExecuter.Enabled=$false; return }
    $boutonExecuter.Enabled = $true
    $texteDescription.Text = $descriptions[$scriptCharge]
    $valeursSauvees = $valeursSession[$scriptCharge]
    foreach ($definition in @($definitions[$scriptCharge])) {
        $valeur = if ($valeursSauvees -and $valeursSauvees.ContainsKey($definition.Nom)) { $valeursSauvees[$definition.Nom] } else { $definition.Valeur }
        $index = $grille.Rows.Add($definition.Libelle,$valeur,$definition.Aide)
        $grille.Rows[$index].Tag = $definition
        if ($definition.Obligatoire) { $grille.Rows[$index].Cells[0].Style.Font = [Drawing.Font]::new($form.Font,[Drawing.FontStyle]::Bold) }
    }
    $options = $optionsSession[$scriptCharge]
    $caseNePasAttendre.Checked = [bool]($options -and $options.NePasAttendre)
    $caseRelancer.Checked = [bool]($options -and $options.Relancer)
    $caseSimulationRobocopy.Checked = [bool]($options -and $options.Simulation)
    $caseRedemarrer.Checked = [bool]($options -and $options.Redemarrer)
    switch ($scriptCharge) {
        '11-Test-ExportPST.ps1' { $caseNePasAttendre.Visible=$true; $labelSaisieSecurisee.Visible=$false }
        '12-Test-MigrationBoite.ps1' { $caseRelancer.Visible=$true }
        '08-Test-Robocopy.ps1' { $caseSimulationRobocopy.Visible=$true; $labelSaisieSecurisee.Visible=$false }
        '02-Test-MigrationMachineDomaine.ps1' { $caseRedemarrer.Visible=$true }
        '00-Verifications-Prerequis.ps1' { $labelSaisieSecurisee.Visible=$false }
        '01-Test-ScopeDHCP.ps1' { $labelSaisieSecurisee.Visible=$false }
        '05-Test-PartageFichiers.ps1' { $labelSaisieSecurisee.Visible=$false }
        '06-Test-ProfilItinerant.ps1' { $labelSaisieSecurisee.Visible=$false }
        '07-Test-Imprimante.ps1' { $labelSaisieSecurisee.Visible=$false }
        '09-Test-RemapLecteur.ps1' { $labelSaisieSecurisee.Visible=$false }
    }
}

function Update-ListeScripts {
    $selection = [string]$listeScripts.SelectedItem
    Save-ConfigurationCourante
    $listeScripts.Items.Clear()
    foreach ($script in Get-ScriptsDisponibles) { [void]$listeScripts.Items.Add($script.Name) }
    if ($selection -and $listeScripts.Items.Contains($selection)) { $listeScripts.SelectedItem=$selection }
    elseif ($listeScripts.Items.Count -gt 0) { $listeScripts.SelectedIndex=0 }
}

function Get-ArgumentsConfiguration {
    $arguments = @()
    foreach ($ligne in $grille.Rows) {
        $definition = $ligne.Tag
        $valeur = ([string]$ligne.Cells[1].Value).Trim()
        if ($definition.Obligatoire -and [string]::IsNullOrWhiteSpace($valeur)) { throw "Le paramètre '$($definition.Libelle)' est obligatoire." }
        if ($valeur -match '["\r\n]') { throw "Le paramètre '$($definition.Libelle)' contient un caractère interdit." }
        if ($definition.Type -eq 'Entier') {
            $nombre = 0
            if (-not [int]::TryParse($valeur,[ref]$nombre) -or $nombre -lt 1) { throw "Le paramètre '$($definition.Libelle)' doit être un entier positif." }
        }
        if ($definition.Type -eq 'Login' -and $valeur -notmatch '^[a-zA-Z0-9._-]+$') { throw "Le login contient un caractère non autorisé." }
        $arguments += @("-$($definition.Nom)",(ConvertTo-ArgumentNatif $valeur))
    }
    $arguments
}

$listeScripts.Add_SelectedIndexChanged({ Show-Configuration })
$boutonActualiser.Add_Click({ Update-ListeScripts })
$boutonDefauts.Add_Click({
    if ($scriptCharge) {
        $valeursSession.Remove($scriptCharge)
        $optionsSession.Remove($scriptCharge)
        $ancien = $scriptCharge
        $script:scriptCharge = $null
        $listeScripts.SelectedItem = $ancien
        Show-Configuration
    }
})
$radioReel.Add_CheckedChanged({
    $caseValidation.Enabled = $radioReel.Checked
    if ($radioReel.Checked) { $statut.Text='Mode réel sélectionné — vérifiez chaque paramètre.'; $statut.ForeColor=[Drawing.Color]::DarkRed }
    else { $caseValidation.Checked=$false; $statut.Text='Prêt — DryRun est activé.'; $statut.ForeColor=[Drawing.Color]::DarkGreen }
})
$boutonDossier.Add_Click({ Start-Process -FilePath 'explorer.exe' -ArgumentList (ConvertTo-ArgumentNatif $racineScripts) })
$boutonFermer.Add_Click({ $form.Close() })

$boutonExecuter.Add_Click({
    if (-not $scriptCharge) { return }
    $script = Join-Path $racineScripts $scriptCharge
    if (-not (Test-Path -LiteralPath $script -PathType Leaf)) { [Windows.Forms.MessageBox]::Show("Script introuvable : $script",'Erreur','OK','Error')|Out-Null; return }
    try { $configurationArguments = @(Get-ArgumentsConfiguration) }
    catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Paramètre invalide','OK','Warning')|Out-Null; return }

    if ($radioReel.Checked -and $scriptCharge -ne '00-Verifications-Prerequis.ps1') {
        if (-not $caseValidation.Checked) { [Windows.Forms.MessageBox]::Show('Validez la vérification des paramètres avant une exécution réelle.','Vérification requise','OK','Warning')|Out-Null; return }
        $reponse = [Windows.Forms.MessageBox]::Show("Lancer $scriptCharge en mode RÉEL ?`r`n`r`nLe script demandera ensuite de taper OUI dans la console.",'Confirmation','YesNo','Warning')
        if ($reponse -ne [Windows.Forms.DialogResult]::Yes) { return }
    }

    $arguments = @('-NoProfile','-NoExit','-File',(ConvertTo-ArgumentNatif $script)) + $configurationArguments
    if ($radioDryRun.Checked) { $arguments += '-DryRun' }
    if ($scriptCharge -eq '11-Test-ExportPST.ps1' -and $caseNePasAttendre.Checked) { $arguments += '-NePasAttendre' }
    if ($scriptCharge -eq '12-Test-MigrationBoite.ps1' -and $caseRelancer.Checked) { $arguments += '-Relancer' }
    if ($scriptCharge -eq '08-Test-Robocopy.ps1' -and $caseSimulationRobocopy.Checked) { $arguments += '-Simulation' }
    if ($scriptCharge -eq '02-Test-MigrationMachineDomaine.ps1' -and $caseRedemarrer.Checked) { $arguments += '-Restart' }
    if ($scriptCharge -eq '04-Test-CreationUtilisateurAD.ps1') { $arguments += @('-IdentifiantsADPath',(ConvertTo-ArgumentNatif $script:cheminIdentifiantsExecution)) }
    if ($scriptCharge -eq '10-Test-CreationBoiteExchange.ps1') { $arguments += @('-IdentifiantsExchangePath',(ConvertTo-ArgumentNatif $script:cheminIdentifiantsExecution)) }
    if ($scriptCharge -eq '03-MigrationUtilisateursAD.ps1') { $arguments += @('-IdentifiantsCiblePath',(ConvertTo-ArgumentNatif $script:cheminIdentifiantsExecution)) }

    try {
        $parametresProcessus = @{
            FilePath         = $powershellExe
            ArgumentList     = $arguments
            WorkingDirectory = $racineScripts
            WindowStyle      = 'Normal'
            ErrorAction      = 'Stop'
        }
        Start-Process @parametresProcessus | Out-Null
        $mode = if ($radioDryRun.Checked) { 'DryRun' } else { 'RÉEL' }
        $statut.Text = "$scriptCharge lancé en mode $mode avec $($script:identifiantsExecution.UserName)."
        $statut.ForeColor = if ($radioDryRun.Checked) { [Drawing.Color]::DarkGreen } else { [Drawing.Color]::DarkRed }
    }
    catch { [Windows.Forms.MessageBox]::Show("Impossible de lancer le script :`r`n$($_.Exception.Message)",'Erreur','OK','Error')|Out-Null }
})

$form.AcceptButton = $boutonExecuter
$form.CancelButton = $boutonFermer
$form.Add_FormClosed({
    Remove-Item -LiteralPath $script:cheminIdentifiantsExecution -Force -ErrorAction SilentlyContinue
})
Update-ListeScripts
[void]$form.ShowDialog()
