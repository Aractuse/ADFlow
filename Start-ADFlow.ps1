#Requires -Version 5.1
<#
.SYNOPSIS
    ADFlow CLI - Interface CLI pour le deploiement d'objets Active Directory
.DESCRIPTION
    Script interactif avec navigation par menus pour deployer des utilisateurs,
    groupes et ordinateurs dans Active Directory a partir de fichiers CSV.
    Lancez le script sans argument pour acceder a l'interface interactive.
.PARAMETER ExportTemplate
    Exporte les templates CSV dans le dossier .\CSV\
.PARAMETER Help
    Affiche l'aide
.PARAMETER Version
    Affiche la version
.EXAMPLE
    .\Start-ADFlow.ps1
    Lance l'interface interactive (CLI)
.EXAMPLE
    .\Start-ADFlow.ps1 -ExportTemplate
    Exporte les templates CSV
.NOTES
    Auteur: Taeckens.M
    Version: 2.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(ParameterSetName = 'ExportTemplate', Mandatory = $true)]
    [switch]$ExportTemplate,

    [Parameter(Mandatory = $false)]
    [switch]$Help,

    [Parameter(Mandatory = $false)]
    [switch]$Version
)

#region Variables Globales
$script:AppVersion = "2.0.0"
$script:AppName = "ADFlow CLI"
$script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:CSVDefaultPath = Join-Path -Path $script:ScriptPath -ChildPath "CSV"
$script:TemplatesPath = Join-Path -Path $script:ScriptPath -ChildPath "Templates"
$script:ToolsPath = Join-Path -Path $script:ScriptPath -ChildPath "Tools"
$script:ExportsPath = Join-Path -Path $script:ScriptPath -ChildPath "Exports"

# Chemins par defaut des fichiers CSV
$script:DefaultCSVPaths = @{
    Users     = Join-Path -Path $script:CSVDefaultPath -ChildPath "Users.csv"
    Groups    = Join-Path -Path $script:CSVDefaultPath -ChildPath "Groups.csv"
    Computers = Join-Path -Path $script:CSVDefaultPath -ChildPath "Computers.csv"
}

# Configuration du deploiement (modifiable via menu)
$script:DeployConfig = @{
    WhatIf                  = $false
    Update                  = $false
    CreateOUs               = $false
    EncryptPasswords        = $false
    PasswordLength          = 16
    ContinueOnMissingMember = $false
}

# Variables de logging
$script:LogFilePath = $null
$script:ErrorSummary = [System.Collections.ArrayList]::new()
$script:WarningSummary = [System.Collections.ArrayList]::new()
$script:SuccessSummary = [System.Collections.ArrayList]::new()

# Variables de mots de passe
$script:GeneratedPasswords = [System.Collections.ArrayList]::new()

# Colonnes requises pour chaque type d'objet
$script:RequiredColumns = @{
    Users     = @('SamAccountName', 'GivenName', 'Surname', 'OU')
    Groups    = @('Name', 'SamAccountName', 'GroupScope', 'GroupCategory', 'OU')
    Computers = @('Name', 'SamAccountName', 'OU')
}

# Colonnes optionnelles pour chaque type d'objet
$script:OptionalColumns = @{
    Users     = @('DisplayName', 'Email', 'Password', 'Groups', 'Enabled', 'Description')
    Groups    = @('Description', 'Members')
    Computers = @('Description', 'Enabled')
}

# Information domaine (rempli apres verification)
$script:DomainInfo = @{
    BaseDN           = $null
    DomainName       = $null
    DomainController = $null
}
#endregion

#region Fonctions - Interface CLI
#==========================================================================
# Fonction    : Clear-Screen
# Arguments   : aucun
# Return      : void
# Description : Efface l'ecran de la console
#==========================================================================
function Clear-Screen {
    [Console]::Clear()
    [Console]::SetCursorPosition(0, 0)
}

#==========================================================================
# Fonction    : Show-Banner
# Arguments   : aucun
# Return      : void
# Description : Affiche la banniere ASCII avec version et domaine
#==========================================================================
function Show-Banner {
    $domainText = if ($script:DomainInfo.DomainName) { $script:DomainInfo.DomainName } else { "Non connecte" }

    Write-Host ""
    Write-Host "  ╭─────────────────────────────────────────────────────────────╮" -ForegroundColor Cyan
    Write-Host "  │                                                             │" -ForegroundColor Cyan
    Write-Host "  │     █████╗ ██████╗ ███████╗██╗      ██████╗ ██╗    ██╗      │" -ForegroundColor Cyan
    Write-Host "  │    ██╔══██╗██╔══██╗██╔════╝██║     ██╔═══██╗██║    ██║      │" -ForegroundColor Cyan
    Write-Host "  │    ███████║██║  ██║█████╗  ██║     ██║   ██║██║ █╗ ██║      │" -ForegroundColor Cyan
    Write-Host "  │    ██╔══██║██║  ██║██╔══╝  ██║     ██║   ██║██║███╗██║      │" -ForegroundColor Cyan
    Write-Host "  │    ██║  ██║██████╔╝██║     ███████╗╚██████╔╝╚███╔███╔╝      │" -ForegroundColor Cyan
    Write-Host "  │    ╚═╝  ╚═╝╚═════╝ ╚═╝     ╚══════╝ ╚═════╝  ╚══╝╚══╝       │" -ForegroundColor Cyan
    Write-Host "  │                                                             │" -ForegroundColor Cyan
    Write-Host "  │              Deploiement Active Directory                   │" -ForegroundColor Cyan
    Write-Host "  │                                                             │" -ForegroundColor Cyan
    Write-Host "  │  Version: " -ForegroundColor Cyan -NoNewline
    Write-Host $script:AppVersion.PadRight(50) -ForegroundColor Yellow -NoNewline
    Write-Host "│" -ForegroundColor Cyan
    Write-Host "  │  Domaine: " -ForegroundColor Cyan -NoNewline
    Write-Host $domainText.PadRight(50) -ForegroundColor Green -NoNewline
    Write-Host "│" -ForegroundColor Cyan
    Write-Host "  │                                      " -ForegroundColor Cyan -NoNewline
    Write-Host "Auteur: Taeckens.M" -ForegroundColor DarkGray -NoNewline
    Write-Host "     │" -ForegroundColor Cyan
    Write-Host "  ╰─────────────────────────────────────────────────────────────╯" -ForegroundColor Cyan
    Write-Host ""
}

#==========================================================================
# Fonction    : Show-MenuHeader
# Arguments   : string title
# Return      : void
# Description : Affiche un en-tete de menu avec bordure arrondie
#==========================================================================
function Show-MenuHeader {
    param([string]$Title)

    $width = 50
    $titlePadded = " $Title ".PadRight($width - 4).PadLeft($width - 2)

    Write-Host ""
    Write-Host "  ╭──$("─" * ($width - 4))──╮" -ForegroundColor Yellow
    Write-Host "  │$titlePadded  │" -ForegroundColor Yellow
    Write-Host "  ╰──$("─" * ($width - 4))──╯" -ForegroundColor Yellow
    Write-Host ""
}

#==========================================================================
# Fonction    : Show-MenuItem
# Arguments   : string number, string text, string color
# Return      : void
# Description : Affiche un element de menu avec numero
#==========================================================================
function Show-MenuItem {
    param(
        [string]$Number,
        [string]$Text,
        [string]$Color = "White"
    )

    Write-Host "    " -NoNewline
    Write-Host "$Number." -ForegroundColor Green -NoNewline
    Write-Host " $Text" -ForegroundColor $Color
}

#==========================================================================
# Fonction    : Get-UserChoice
# Arguments   : string prompt, array validChoices
# Return      : string choix utilisateur
# Description : Demande et valide un choix utilisateur
#==========================================================================
function Get-UserChoice {
    param(
        [string]$Prompt = "Votre choix",
        [array]$ValidChoices = @()
    )

    Write-Host ""
    Write-Host "  $Prompt : " -ForegroundColor Cyan -NoNewline
    $choice = Read-Host

    if ($ValidChoices.Count -gt 0 -and $choice -notin $ValidChoices) {
        Write-Host "  Choix invalide. Veuillez reessayer." -ForegroundColor Red
        Start-Sleep -Milliseconds 1000
        return $null
    }

    return $choice
}

#==========================================================================
# Fonction    : Show-Confirmation
# Arguments   : string message
# Return      : bool confirmation
# Description : Demande une confirmation O/N
#==========================================================================
function Show-Confirmation {
    param([string]$Message)

    Write-Host ""
    Write-Host "  $Message (O/N) : " -ForegroundColor Yellow -NoNewline
    $response = Read-Host
    return $response -match '^[OoYy]'
}

#==========================================================================
# Fonction    : Wait-KeyPress
# Arguments   : string message
# Return      : void
# Description : Attend Entree ou Echap pour continuer
#==========================================================================
function Wait-KeyPress {
    param([string]$Message = "Appuyez sur Entree ou Echap pour continuer...")

    Write-Host ""
    Write-Host "  $Message" -ForegroundColor Gray
    do {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } while ($key.VirtualKeyCode -ne 13 -and $key.VirtualKeyCode -ne 27)
}

#==========================================================================
# Fonction    : Show-StatusMessage
# Arguments   : string message, string level
# Return      : void
# Description : Affiche un message avec statut colore
#==========================================================================
function Show-StatusMessage {
    param(
        [string]$Message,
        [ValidateSet("OK", "ERROR", "WARN", "INFO", "SKIP")]
        [string]$Level = "INFO"
    )

    $prefix = switch ($Level) {
        "OK"    { "[OK]"; $color = "Green" }
        "ERROR" { "[ERREUR]"; $color = "Red" }
        "WARN"  { "[WARN]"; $color = "Yellow" }
        "SKIP"  { "[SKIP]"; $color = "Yellow" }
        default { "[INFO]"; $color = "Cyan" }
    }

    Write-Host "  $prefix " -ForegroundColor $color -NoNewline
    Write-Host $Message
}

#==========================================================================
# Fonction    : Show-ProgressItem
# Arguments   : string identity, string action, string status
# Return      : void
# Description : Affiche un element de progression de deploiement
#==========================================================================
function Show-ProgressItem {
    param(
        [string]$Identity,
        [string]$Action,
        [ValidateSet("OK", "ERROR", "SKIP", "UPDATE")]
        [string]$Status
    )

    $statusText = switch ($Status) {
        "OK"     { "[OK]"; $color = "Green" }
        "ERROR"  { "[ERREUR]"; $color = "Red" }
        "SKIP"   { "[SKIP]"; $color = "Yellow" }
        "UPDATE" { "[MAJ]"; $color = "Cyan" }
    }

    Write-Host "    $statusText " -ForegroundColor $color -NoNewline
    Write-Host "$Action : " -NoNewline
    Write-Host $Identity -ForegroundColor White
}
#endregion

#region Fonctions - Prerequis
#==========================================================================
# Fonction    : Test-Prerequisites
# Arguments   : bool silent
# Return      : PSCustomObject resultat des verifications
# Description : Verifie tous les prerequis necessaires au fonctionnement
#==========================================================================
function Test-Prerequisites {
    param([switch]$Silent)

    $results = [PSCustomObject]@{
        PowerShellVersion = $false
        ADModule          = $false
        ADModuleLoaded    = $false
        DomainJoined      = $false
        DCConnectivity    = $false
        ADPermissions     = $false
        AllPassed         = $false
        Errors            = [System.Collections.ArrayList]::new()
    }

    if (-not $Silent) {
        Show-MenuHeader -Title "Verification des prerequis"
    }

    # 1. Version PowerShell
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -ge 5 -and $psVersion.Minor -ge 1) {
        $results.PowerShellVersion = $true
        if (-not $Silent) { Show-StatusMessage -Message "PowerShell $($psVersion.Major).$($psVersion.Minor)" -Level "OK" }
    }
    else {
        [void]$results.Errors.Add("PowerShell 5.1 ou superieur requis (actuel: $($psVersion.Major).$($psVersion.Minor))")
        if (-not $Silent) {
            Show-StatusMessage -Message "PowerShell 5.1 requis (actuel: $($psVersion.Major).$($psVersion.Minor))" -Level "ERROR"
        }
    }

    # 2. Module ActiveDirectory installe
    $adModule = Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue
    if ($adModule) {
        $results.ADModule = $true
        if (-not $Silent) { Show-StatusMessage -Message "Module ActiveDirectory installe (v$($adModule.Version))" -Level "OK" }

        # 3. Charger le module
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            $results.ADModuleLoaded = $true
            if (-not $Silent) { Show-StatusMessage -Message "Module ActiveDirectory charge" -Level "OK" }
        }
        catch {
            [void]$results.Errors.Add("Impossible de charger le module ActiveDirectory")
            if (-not $Silent) { Show-StatusMessage -Message "Impossible de charger le module ActiveDirectory" -Level "ERROR" }
        }
    }
    else {
        [void]$results.Errors.Add("Module ActiveDirectory non installe")
        if (-not $Silent) {
            Show-StatusMessage -Message "Module ActiveDirectory non installe" -Level "ERROR"
            Write-Host ""
            Write-Host "  Pour installer le module:" -ForegroundColor Yellow
            Write-Host "  - Windows 10/11: " -ForegroundColor Gray -NoNewline
            Write-Host "Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ForegroundColor White
            Write-Host "  - Windows Server: " -ForegroundColor Gray -NoNewline
            Write-Host "Install-WindowsFeature -Name RSAT-AD-PowerShell" -ForegroundColor White
        }
    }

    # 4. Machine jointe au domaine
    try {
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
        if ($computerSystem.PartOfDomain) {
            $results.DomainJoined = $true
            $script:DomainInfo.DomainName = $computerSystem.Domain
            if (-not $Silent) { Show-StatusMessage -Message "Machine jointe au domaine: $($computerSystem.Domain)" -Level "OK" }
        }
        else {
            [void]$results.Errors.Add("La machine n'est pas jointe a un domaine")
            if (-not $Silent) { Show-StatusMessage -Message "Machine non jointe a un domaine" -Level "ERROR" }
        }
    }
    catch {
        [void]$results.Errors.Add("Impossible de verifier l'appartenance au domaine")
        if (-not $Silent) { Show-StatusMessage -Message "Impossible de verifier l'appartenance au domaine" -Level "ERROR" }
    }

    # 5. Connectivite DC
    if ($results.ADModuleLoaded) {
        try {
            $dc = Get-ADDomainController -Discover -ErrorAction Stop
            $results.DCConnectivity = $true
            $script:DomainInfo.DomainController = $dc.HostName[0]
            $script:DomainInfo.BaseDN = (Get-ADDomain).DistinguishedName
            if (-not $Silent) { Show-StatusMessage -Message "Connexion DC: $($dc.HostName[0])" -Level "OK" }
        }
        catch {
            [void]$results.Errors.Add("Impossible de contacter un controleur de domaine")
            if (-not $Silent) { Show-StatusMessage -Message "Impossible de contacter un controleur de domaine" -Level "ERROR" }
        }
    }

    # 6. Verification des privileges AD
    if ($results.DCConnectivity) {
        try {
            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
            $currentUserName = $currentUser.Name
            $missingPermissions = [System.Collections.ArrayList]::new()

            # Groupes privilegies a verifier
            $privilegedGroups = @(
                "Domain Admins",
                "Administrateurs du domaine",
                "Enterprise Admins",
                "Administrateurs de l'entreprise",
                "Account Operators",
                "Operateurs de compte"
            )

            # Verifier l'appartenance aux groupes privilegies
            $isPrivilegedMember = $false
            $principal = New-Object System.Security.Principal.WindowsPrincipal($currentUser)

            foreach ($group in $privilegedGroups) {
                try {
                    $sid = (Get-ADGroup -Identity $group -ErrorAction SilentlyContinue).SID
                    if ($sid -and $principal.IsInRole($sid)) {
                        $isPrivilegedMember = $true
                        break
                    }
                }
                catch { }
            }

            # Verifier les permissions specifiques sur le domaine
            $domainDN = $script:DomainInfo.BaseDN
            $acl = Get-Acl "AD:$domainDN" -ErrorAction SilentlyContinue

            if ($acl) {
                $userSid = $currentUser.User.Value
                $userGroups = $currentUser.Groups | ForEach-Object { $_.Value }

                # GUIDs des types d'objets AD
                $userGuid = [guid]"bf967aba-0de6-11d0-a285-00aa003049e2"    # User
                $groupGuid = [guid]"bf967a9c-0de6-11d0-a285-00aa003049e2"   # Group
                $computerGuid = [guid]"bf967a86-0de6-11d0-a285-00aa003049e2" # Computer
                $ouGuid = [guid]"bf967aa5-0de6-11d0-a285-00aa003049e2"      # OrganizationalUnit

                $canCreateUsers = $false
                $canCreateGroups = $false
                $canCreateComputers = $false
                $canCreateOUs = $false

                foreach ($ace in $acl.Access) {
                    $identityRef = $ace.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value

                    if ($identityRef -eq $userSid -or $userGroups -contains $identityRef) {
                        if ($ace.ActiveDirectoryRights -match "CreateChild" -or $ace.ActiveDirectoryRights -match "GenericAll") {
                            if ($ace.ObjectType -eq [guid]::Empty -or $ace.ObjectType -eq $userGuid) { $canCreateUsers = $true }
                            if ($ace.ObjectType -eq [guid]::Empty -or $ace.ObjectType -eq $groupGuid) { $canCreateGroups = $true }
                            if ($ace.ObjectType -eq [guid]::Empty -or $ace.ObjectType -eq $computerGuid) { $canCreateComputers = $true }
                            if ($ace.ObjectType -eq [guid]::Empty -or $ace.ObjectType -eq $ouGuid) { $canCreateOUs = $true }
                        }
                    }
                }

                # Si membre d'un groupe privilegie, on considere les droits OK
                if ($isPrivilegedMember) {
                    $canCreateUsers = $true
                    $canCreateGroups = $true
                    $canCreateComputers = $true
                    $canCreateOUs = $true
                }

                if (-not $canCreateUsers) { [void]$missingPermissions.Add("Creer des utilisateurs") }
                if (-not $canCreateGroups) { [void]$missingPermissions.Add("Creer des groupes") }
                if (-not $canCreateComputers) { [void]$missingPermissions.Add("Creer des ordinateurs") }
                if (-not $canCreateOUs) { [void]$missingPermissions.Add("Creer des unites d'organisation (OU)") }
            }

            if ($missingPermissions.Count -eq 0) {
                $results.ADPermissions = $true
                if (-not $Silent) { Show-StatusMessage -Message "Privileges AD: $currentUserName" -Level "OK" }
            }
            else {
                [void]$results.Errors.Add("Droits AD insuffisants pour: $($missingPermissions -join ', ')")
                if (-not $Silent) {
                    Show-StatusMessage -Message "Privileges AD insuffisants" -Level "ERROR"
                    Write-Host ""
                    Write-Host "  Droits manquants pour le compte '$currentUserName':" -ForegroundColor Yellow
                    foreach ($perm in $missingPermissions) {
                        Write-Host "    - $perm" -ForegroundColor Red
                    }
                    Write-Host ""
                    Write-Host "  Solutions possibles:" -ForegroundColor Yellow
                    Write-Host "    - Executez le script avec un compte administrateur AD" -ForegroundColor Gray
                    Write-Host "    - Demandez la delegation des droits necessaires" -ForegroundColor Gray
                }
            }
        }
        catch {
            # En cas d'erreur de verification, on continue avec un warning
            $results.ADPermissions = $true
            if (-not $Silent) { Show-StatusMessage -Message "Privileges AD: verification non conclusive (continuer avec prudence)" -Level "WARN" }
        }
    }

    $results.AllPassed = $results.PowerShellVersion -and $results.ADModule -and $results.ADModuleLoaded -and $results.DomainJoined -and $results.DCConnectivity -and $results.ADPermissions

    if (-not $Silent) {
        Write-Host ""
        if ($results.AllPassed) {
            Write-Host "  Tous les prerequis sont satisfaits." -ForegroundColor Green
        }
        else {
            Write-Host "  Certains prerequis ne sont pas satisfaits." -ForegroundColor Red
        }
    }

    return $results
}

#==========================================================================
# Fonction    : Test-ADObjectExists
# Arguments   : string identity, string objectType
# Return      : bool existe
# Description : Verifie si un objet AD existe
#==========================================================================
function Test-ADObjectExists {
    param(
        [string]$Identity,
        [ValidateSet("User", "Group", "Computer")]
        [string]$ObjectType
    )

    try {
        switch ($ObjectType) {
            "User" { Get-ADUser -Identity $Identity -ErrorAction Stop | Out-Null }
            "Group" { Get-ADGroup -Identity $Identity -ErrorAction Stop | Out-Null }
            "Computer" { Get-ADComputer -Identity $Identity -ErrorAction Stop | Out-Null }
        }
        return $true
    }
    catch { return $false }
}

#==========================================================================
# Fonction    : Test-ADObjectExistsAny
# Arguments   : string identity
# Return      : bool existe
# Description : Verifie si un objet AD existe (user, group ou computer)
#==========================================================================
function Test-ADObjectExistsAny {
    param([string]$Identity)

    if (Test-ADObjectExists -Identity $Identity -ObjectType "User") { return $true }
    if (Test-ADObjectExists -Identity $Identity -ObjectType "Group") { return $true }
    if (Test-ADObjectExists -Identity $Identity -ObjectType "Computer") { return $true }
    return $false
}
#endregion

#region Fonctions - Validation
#==========================================================================
# Fonction    : Test-CSVFile
# Arguments   : string path, string objectType
# Return      : PSCustomObject resultat validation
# Description : Valide un fichier CSV
#==========================================================================
function Test-CSVFile {
    param(
        [string]$Path,
        [ValidateSet("Users", "Groups", "Computers")]
        [string]$ObjectType
    )

    $result = [PSCustomObject]@{
        IsValid    = $false
        FilePath   = $Path
        ObjectType = $ObjectType
        RowCount   = 0
        Errors     = [System.Collections.ArrayList]::new()
        Warnings   = [System.Collections.ArrayList]::new()
        Data       = $null
    }

    # Verifier existence
    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        [void]$result.Errors.Add("Le fichier n'existe pas: $Path")
        return $result
    }

    # Verifier extension
    if ([System.IO.Path]::GetExtension($Path).ToLower() -ne ".csv") {
        [void]$result.Errors.Add("Le fichier n'est pas un CSV")
        return $result
    }

    # Lire le CSV
    try {
        $csvData = Import-Csv -Path $Path -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        try {
            $csvData = Import-Csv -Path $Path -ErrorAction Stop
        }
        catch {
            [void]$result.Errors.Add("Impossible de lire le fichier CSV: $($_.Exception.Message)")
            return $result
        }
    }

    if ($null -eq $csvData -or $csvData.Count -eq 0) {
        [void]$result.Errors.Add("Le fichier CSV est vide")
        return $result
    }

    $result.RowCount = $csvData.Count
    $fileColumns = $csvData[0].PSObject.Properties.Name
    $requiredCols = $script:RequiredColumns[$ObjectType]
    $missingCols = $requiredCols | Where-Object { $_ -notin $fileColumns }

    if ($missingCols.Count -gt 0) {
        [void]$result.Errors.Add("Colonnes requises manquantes: $($missingCols -join ', ')")
    }

    # Colonnes inconnues
    $allKnownCols = $requiredCols + $script:OptionalColumns[$ObjectType]
    $unknownCols = $fileColumns | Where-Object { $_ -notin $allKnownCols }
    if ($unknownCols.Count -gt 0) {
        [void]$result.Warnings.Add("Colonnes non reconnues: $($unknownCols -join ', ')")
    }

    if ($result.Errors.Count -gt 0) { return $result }

    # Valider chaque ligne
    $rowIndex = 0
    foreach ($row in $csvData) {
        $rowIndex++
        $rowErrors = Test-CSVRow -Row $row -ObjectType $ObjectType -RowIndex $rowIndex
        foreach ($err in $rowErrors.Errors) { [void]$result.Errors.Add($err) }
        foreach ($warn in $rowErrors.Warnings) { [void]$result.Warnings.Add($warn) }
    }

    $result.IsValid = ($result.Errors.Count -eq 0)
    $result.Data = $csvData
    return $result
}

#==========================================================================
# Fonction    : Test-CSVRow
# Arguments   : PSObject row, string objectType, int rowIndex
# Return      : hashtable erreurs et warnings
# Description : Valide une ligne de donnees CSV
#==========================================================================
function Test-CSVRow {
    param(
        [PSObject]$Row,
        [ValidateSet("Users", "Groups", "Computers")]
        [string]$ObjectType,
        [int]$RowIndex
    )

    $result = @{
        Errors   = [System.Collections.ArrayList]::new()
        Warnings = [System.Collections.ArrayList]::new()
    }

    switch ($ObjectType) {
        "Users" {
            if ([string]::IsNullOrWhiteSpace($Row.SamAccountName)) {
                [void]$result.Errors.Add("Ligne $RowIndex : SamAccountName est vide")
            }
            elseif ($Row.SamAccountName.Length -gt 20) {
                [void]$result.Errors.Add("Ligne $RowIndex : SamAccountName depasse 20 caracteres")
            }
            elseif ($Row.SamAccountName -notmatch '^[a-zA-Z0-9._-]+$') {
                [void]$result.Errors.Add("Ligne $RowIndex : SamAccountName contient des caracteres invalides")
            }
            if ([string]::IsNullOrWhiteSpace($Row.GivenName)) { [void]$result.Errors.Add("Ligne $RowIndex : GivenName est vide") }
            if ([string]::IsNullOrWhiteSpace($Row.Surname)) { [void]$result.Errors.Add("Ligne $RowIndex : Surname est vide") }
            if ([string]::IsNullOrWhiteSpace($Row.OU)) { [void]$result.Errors.Add("Ligne $RowIndex : OU est vide") }
        }
        "Groups" {
            if ([string]::IsNullOrWhiteSpace($Row.Name)) { [void]$result.Errors.Add("Ligne $RowIndex : Name est vide") }
            if ([string]::IsNullOrWhiteSpace($Row.SamAccountName)) { [void]$result.Errors.Add("Ligne $RowIndex : SamAccountName est vide") }
            if ([string]::IsNullOrWhiteSpace($Row.GroupScope) -or $Row.GroupScope -notmatch '^(DomainLocal|Global|Universal)$') {
                [void]$result.Errors.Add("Ligne $RowIndex : GroupScope invalide (DomainLocal, Global, Universal)")
            }
            if ([string]::IsNullOrWhiteSpace($Row.GroupCategory) -or $Row.GroupCategory -notmatch '^(Security|Distribution)$') {
                [void]$result.Errors.Add("Ligne $RowIndex : GroupCategory invalide (Security, Distribution)")
            }
            if ([string]::IsNullOrWhiteSpace($Row.OU)) { [void]$result.Errors.Add("Ligne $RowIndex : OU est vide") }
        }
        "Computers" {
            if ([string]::IsNullOrWhiteSpace($Row.Name)) { [void]$result.Errors.Add("Ligne $RowIndex : Name est vide") }
            elseif ($Row.Name.Length -gt 15) { [void]$result.Errors.Add("Ligne $RowIndex : Name depasse 15 caracteres") }
            elseif ($Row.Name -notmatch '^[a-zA-Z0-9-]+$') { [void]$result.Errors.Add("Ligne $RowIndex : Name contient des caracteres invalides") }
            if ([string]::IsNullOrWhiteSpace($Row.SamAccountName)) { [void]$result.Errors.Add("Ligne $RowIndex : SamAccountName est vide") }
            if ([string]::IsNullOrWhiteSpace($Row.OU)) { [void]$result.Errors.Add("Ligne $RowIndex : OU est vide") }
        }
    }
    return $result
}

#==========================================================================
# Fonction    : Test-OUsExist
# Arguments   : array data, string baseDN
# Return      : PSCustomObject resultat verification OUs
# Description : Verifie que toutes les OUs referencees existent
#==========================================================================
function Test-OUsExist {
    param(
        [array]$Data,
        [string]$BaseDN
    )

    $result = [PSCustomObject]@{
        AllExist   = $true
        MissingOUs = [System.Collections.ArrayList]::new()
        CheckedOUs = @{}
    }

    $uniqueOUs = $Data | Select-Object -ExpandProperty OU -Unique

    foreach ($ou in $uniqueOUs) {
        if ([string]::IsNullOrWhiteSpace($ou)) { continue }

        $fullDN = if ($ou -notmatch "DC=") { "$ou,$BaseDN" } else { $ou }

        if ($result.CheckedOUs.ContainsKey($fullDN)) { continue }

        $exists = $false
        try {
            Get-ADOrganizationalUnit -Identity $fullDN -ErrorAction Stop | Out-Null
            $exists = $true
        }
        catch { $exists = $false }

        $result.CheckedOUs[$fullDN] = $exists
        if (-not $exists) {
            $result.AllExist = $false
            [void]$result.MissingOUs.Add($fullDN)
        }
    }
    return $result
}

#==========================================================================
# Fonction    : New-MissingOUs
# Arguments   : ArrayList missingOUs, string baseDN, bool whatIfMode
# Return      : PSCustomObject resultat creation
# Description : Cree les OUs manquantes de maniere recursive
#==========================================================================
function New-MissingOUs {
    param(
        [System.Collections.ArrayList]$MissingOUs,
        [string]$BaseDN,
        [switch]$WhatIfMode
    )

    $result = [PSCustomObject]@{
        Success    = $true
        CreatedOUs = [System.Collections.ArrayList]::new()
        Errors     = [System.Collections.ArrayList]::new()
    }

    $allOUsToCreate = [System.Collections.ArrayList]::new()

    foreach ($ouDN in $MissingOUs) {
        $ouSegments = Get-OUSegmentsFromDN -OUDN $ouDN -BaseDN $BaseDN
        $currentPath = $BaseDN

        foreach ($segment in $ouSegments) {
            $currentPath = "$segment,$currentPath"
            $ouExists = $false
            try {
                Get-ADOrganizationalUnit -Identity $currentPath -ErrorAction Stop | Out-Null
                $ouExists = $true
            }
            catch { $ouExists = $false }

            if (-not $ouExists -and $allOUsToCreate -notcontains $currentPath) {
                [void]$allOUsToCreate.Add($currentPath)
            }
        }
    }

    $sortedOUs = $allOUsToCreate | Sort-Object { ($_ -split ',').Count }

    foreach ($ouToCreate in $sortedOUs) {
        $ouName = ($ouToCreate -split ',')[0] -replace '^OU=', ''
        $parentPath = ($ouToCreate -split ',', 2)[1]

        if ($WhatIfMode) {
            [void]$result.CreatedOUs.Add($ouToCreate)
        }
        else {
            try {
                New-ADOrganizationalUnit -Name $ouName -Path $parentPath -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
                [void]$result.CreatedOUs.Add($ouToCreate)
            }
            catch {
                $result.Success = $false
                [void]$result.Errors.Add("Impossible de creer l'OU '$ouToCreate': $($_.Exception.Message)")
            }
        }
    }
    return $result
}

#==========================================================================
# Fonction    : Get-OUSegmentsFromDN
# Arguments   : string ouDN, string baseDN
# Return      : array segments OU
# Description : Extrait les segments OU d'un DN complet
#==========================================================================
function Get-OUSegmentsFromDN {
    param(
        [string]$OUDN,
        [string]$BaseDN
    )

    $ouPart = $OUDN
    if ($OUDN.EndsWith(",$BaseDN")) {
        $ouPart = $OUDN.Substring(0, $OUDN.Length - $BaseDN.Length - 1)
    }

    $segments = [System.Collections.ArrayList]::new()
    $currentSegment = ""
    $chars = $ouPart.ToCharArray()

    for ($i = 0; $i -lt $chars.Length; $i++) {
        $char = $chars[$i]
        if ($char -eq ',' -and ($i -eq 0 -or $chars[$i - 1] -ne '\')) {
            if (-not [string]::IsNullOrWhiteSpace($currentSegment)) {
                [void]$segments.Add($currentSegment.Trim())
            }
            $currentSegment = ""
        }
        else { $currentSegment += $char }
    }

    if (-not [string]::IsNullOrWhiteSpace($currentSegment)) {
        [void]$segments.Add($currentSegment.Trim())
    }

    $ouSegments = $segments | Where-Object { $_ -match '^OU=' }
    [array]::Reverse($ouSegments)
    return $ouSegments
}

#==========================================================================
# Fonction    : Test-DuplicateSamAccountNames
# Arguments   : array data
# Return      : array doublons
# Description : Verifie les doublons de SamAccountName dans un CSV
#==========================================================================
function Test-DuplicateSamAccountNames {
    param([array]$Data)

    return $Data | Group-Object -Property SamAccountName | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name
}

#==========================================================================
# Fonction    : Convert-EnabledValue
# Arguments   : string value, bool default
# Return      : bool valeur convertie
# Description : Convertit une valeur Enabled en booleen
#==========================================================================
function Convert-EnabledValue {
    param(
        [string]$Value,
        [bool]$Default = $true
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
    switch ($Value.ToLower().Trim()) {
        { $_ -in @('true', '1', 'yes', 'oui') } { return $true }
        { $_ -in @('false', '0', 'no', 'non') } { return $false }
        default { return $Default }
    }
}
#endregion

#region Fonctions - Logging
#==========================================================================
# Fonction    : Initialize-Logging
# Arguments   : string logPath, string sessionName
# Return      : bool succes
# Description : Initialise le systeme de logging
#==========================================================================
function Initialize-Logging {
    param(
        [string]$LogPath,
        [string]$SessionName = "Deploy"
    )

    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path -Path $script:ExportsPath -ChildPath "Logs"
    }

    if (-not (Test-Path -Path $LogPath)) {
        try {
            New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
        }
        catch { return $false }
    }

    $timestamp = Get-Date -Format "ddMMyyyy_HHmm"
    $script:LogFilePath = Join-Path -Path $LogPath -ChildPath "${SessionName}_${timestamp}.log"
    $script:ErrorSummary.Clear()
    $script:WarningSummary.Clear()
    $script:SuccessSummary.Clear()

    $header = @"
================================================================================
  ADFlow CLI - Session de deploiement
  Demarre le: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
  Utilisateur: $env:USERNAME
  Machine: $env:COMPUTERNAME
================================================================================

"@

    try {
        $header | Out-File -FilePath $script:LogFilePath -Encoding UTF8
        return $true
    }
    catch { return $false }
}

#==========================================================================
# Fonction    : Write-Log
# Arguments   : string message, string level, bool noConsole
# Return      : void
# Description : Ecrit une entree dans le fichier de log
#==========================================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO",
        [switch]$NoConsole
    )

    $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    if ($script:LogFilePath -and (Test-Path -Path (Split-Path $script:LogFilePath -Parent))) {
        try { $logEntry | Out-File -FilePath $script:LogFilePath -Append -Encoding UTF8 } catch { }
    }

    switch ($Level) {
        "ERROR" { [void]$script:ErrorSummary.Add($Message) }
        "WARNING" { [void]$script:WarningSummary.Add($Message) }
        "SUCCESS" { [void]$script:SuccessSummary.Add($Message) }
    }

    if (-not $NoConsole) {
        $color = switch ($Level) {
            "INFO" { "Cyan" }
            "WARNING" { "Yellow" }
            "ERROR" { "Red" }
            "SUCCESS" { "Green" }
            "DEBUG" { "Gray" }
        }
        Write-Host "  $logEntry" -ForegroundColor $color
    }
}

#==========================================================================
# Fonction    : Write-FinalSummary
# Arguments   : aucun
# Return      : PSCustomObject resume
# Description : Affiche et log le resume final de la session
#==========================================================================
function Write-FinalSummary {
    Write-Host ""
    Write-Host "  ╭─────────────────────────────────────────╮" -ForegroundColor Cyan
    Write-Host "  │          Resume de la session           │" -ForegroundColor Cyan
    Write-Host "  ╰─────────────────────────────────────────╯" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    Operations reussies : " -NoNewline
    Write-Host "$($script:SuccessSummary.Count)" -ForegroundColor Green
    Write-Host "    Avertissements      : " -NoNewline
    Write-Host "$($script:WarningSummary.Count)" -ForegroundColor Yellow
    Write-Host "    Erreurs             : " -NoNewline
    Write-Host "$($script:ErrorSummary.Count)" -ForegroundColor Red

    if ($script:LogFilePath) {
        Write-Host ""
        Write-Host "    Fichier de log: " -NoNewline
        Write-Host $script:LogFilePath -ForegroundColor Gray
    }

    return [PSCustomObject]@{
        SuccessCount = $script:SuccessSummary.Count
        WarningCount = $script:WarningSummary.Count
        ErrorCount   = $script:ErrorSummary.Count
    }
}
#endregion

#region Fonctions - Mots de passe
#==========================================================================
# Fonction    : New-SecurePassword
# Arguments   : int length
# Return      : string mot de passe
# Description : Genere un mot de passe aleatoire securise
#==========================================================================
function New-SecurePassword {
    param([int]$Length = 16)

    $lowercase = 'abcdefghijkmnopqrstuvwxyz'
    $uppercase = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $numbers = '23456789'
    $special = '!@#$%^&*()_+-=[]{}|;:,.<>?'
    $allChars = $lowercase + $uppercase + $numbers + $special

    do {
        $password = ""
        $charArray = $allChars.ToCharArray()
        $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
        $bytes = New-Object byte[]($Length)
        $rng.GetBytes($bytes)

        for ($i = 0; $i -lt $Length; $i++) {
            $password += $charArray[$bytes[$i] % $charArray.Length]
        }

        $hasLower = $password -cmatch '[a-z]'
        $hasUpper = $password -cmatch '[A-Z]'
        $hasNumber = $password -match '[0-9]'
        $hasSpecial = $password -match '[!@#$%^&*()_+\-=\[\]{}|;:,.<>?]'
        $meetsComplexity = $hasLower -and $hasUpper -and $hasNumber -and $hasSpecial
    } while (-not $meetsComplexity)

    return $password
}

#==========================================================================
# Fonction    : Add-GeneratedPassword
# Arguments   : string samAccountName, string displayName, string password
# Return      : void
# Description : Ajoute un mot de passe genere a la liste pour export
#==========================================================================
function Add-GeneratedPassword {
    param(
        [string]$SamAccountName,
        [string]$DisplayName,
        [string]$Password
    )

    $entry = [PSCustomObject]@{
        SamAccountName = $SamAccountName
        DisplayName    = $DisplayName
        Password       = $Password
        GeneratedAt    = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    }
    [void]$script:GeneratedPasswords.Add($entry)
}

#==========================================================================
# Fonction    : Export-GeneratedPasswords
# Arguments   : bool encrypt, string toolsPath
# Return      : PSCustomObject resultat export
# Description : Exporte les mots de passe generes dans un fichier
#==========================================================================
function Export-GeneratedPasswords {
    param(
        [switch]$Encrypt,
        [string]$ToolsPath
    )

    if ($script:GeneratedPasswords.Count -eq 0) {
        return [PSCustomObject]@{
            Success            = $true
            Message            = "Aucun mot de passe a exporter"
            FilePath           = $null
            Encrypted          = $false
            PasswordCount      = 0
            EncryptionPassword = $null
        }
    }

    $ExportPath = Join-Path -Path $script:ExportsPath -ChildPath "Passwords"

    if (-not (Test-Path -Path $ExportPath)) {
        try { New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null }
        catch {
            return [PSCustomObject]@{
                Success = $false
                Message = "Impossible de creer le dossier d'export"
                FilePath = $null
                Encrypted = $false
                PasswordCount = 0
                EncryptionPassword = $null
            }
        }
    }

    $timestamp = Get-Date -Format "ddMMyyyy_HHmm"
    $csvFileName = "Passwords_${timestamp}.csv"
    $csvFilePath = Join-Path -Path $ExportPath -ChildPath $csvFileName

    try {
        $script:GeneratedPasswords | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            Message = "Erreur lors de l'export CSV"
            FilePath = $null
            Encrypted = $false
            PasswordCount = 0
            EncryptionPassword = $null
        }
    }

    $finalPath = $csvFilePath
    $encrypted = $false
    $generatedEncryptionPassword = $null

    if ($Encrypt) {
        $sevenZipPath = if ($ToolsPath) { Join-Path -Path $ToolsPath -ChildPath "7za.exe" }
                        else { Join-Path -Path $script:ToolsPath -ChildPath "7za.exe" }

        if (Test-Path -Path $sevenZipPath -PathType Leaf) {
            $EncryptionPassword = New-SecurePassword -Length 20
            $generatedEncryptionPassword = $EncryptionPassword
            $zipFileName = "Passwords_${timestamp}.7z"
            $zipFilePath = Join-Path -Path $ExportPath -ChildPath $zipFileName

            try {
                # Construire les arguments en une seule chaine avec guillemets pour gerer espaces et caracteres speciaux
                $argumentString = "a -t7z -mhe=on `"-p$EncryptionPassword`" `"$zipFilePath`" `"$csvFilePath`""

                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = $sevenZipPath
                $psi.Arguments = $argumentString
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true

                $process = [System.Diagnostics.Process]::Start($psi)
                $process.WaitForExit()

                if ($process.ExitCode -eq 0) {
                    Remove-Item -Path $csvFilePath -Force
                    $finalPath = $zipFilePath
                    $encrypted = $true
                }
            }
            catch { $generatedEncryptionPassword = $null }
        }
    }

    return [PSCustomObject]@{
        Success            = $true
        Message            = "Export reussi"
        FilePath           = $finalPath
        Encrypted          = $encrypted
        PasswordCount      = $script:GeneratedPasswords.Count
        EncryptionPassword = $generatedEncryptionPassword
    }
}
#endregion

#region Fonctions - Utilisateurs
#==========================================================================
# Fonction    : New-ADUserFromCSV
# Arguments   : PSObject userData, string baseDN, int passwordLength, bool whatIfMode
# Return      : PSCustomObject resultat operation
# Description : Cree un utilisateur AD a partir d'une ligne CSV
#==========================================================================
function New-ADUserFromCSV {
    param(
        [PSObject]$UserData,
        [string]$BaseDN,
        [int]$PasswordLength = 16,
        [switch]$WhatIfMode
    )

    $result = [PSCustomObject]@{
        Success        = $false
        Action         = "Create"
        SamAccountName = $UserData.SamAccountName
        Message        = ""
        Password       = $null
        GroupsAdded    = [System.Collections.ArrayList]::new()
        GroupsFailed   = [System.Collections.ArrayList]::new()
        MissingGroups  = [System.Collections.ArrayList]::new()
    }

    try {
        $ouDN = if ($UserData.OU -notmatch "DC=") { "$($UserData.OU),$BaseDN" } else { $UserData.OU }

        # Verifier que tous les groupes existent AVANT de creer l'utilisateur
        if (-not [string]::IsNullOrWhiteSpace($UserData.Groups)) {
            $groups = $UserData.Groups -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
            foreach ($group in $groups) {
                if (-not (Test-ADObjectExists -Identity $group -ObjectType "Group")) {
                    [void]$result.MissingGroups.Add($group)
                }
            }
            # Si des groupes manquent, arreter le deploiement de cet utilisateur
            if ($result.MissingGroups.Count -gt 0) {
                $result.Success = $false
                $result.Message = "Groupes inexistants: $($result.MissingGroups -join ', ')"
                return $result
            }
        }

        $password = if (-not [string]::IsNullOrWhiteSpace($UserData.Password)) { $UserData.Password }
                    else { New-SecurePassword -Length $PasswordLength }
        $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
        $result.Password = $password

        $displayName = if (-not [string]::IsNullOrWhiteSpace($UserData.DisplayName)) { $UserData.DisplayName }
                       else { "$($UserData.GivenName) $($UserData.Surname)" }

        $enabled = $true
        if ($UserData.PSObject.Properties.Name -contains 'Enabled') {
            $enabled = Convert-EnabledValue -Value $UserData.Enabled -Default $true
        }

        $userParams = @{
            SamAccountName        = $UserData.SamAccountName
            UserPrincipalName     = "$($UserData.SamAccountName)@$((Get-ADDomain).DNSRoot)"
            Name                  = $displayName
            GivenName             = $UserData.GivenName
            Surname               = $UserData.Surname
            DisplayName           = $displayName
            AccountPassword       = $securePassword
            Enabled               = $enabled
            Path                  = $ouDN
            ChangePasswordAtLogon = $true
        }

        if (-not [string]::IsNullOrWhiteSpace($UserData.Email)) { $userParams['EmailAddress'] = $UserData.Email }
        if (-not [string]::IsNullOrWhiteSpace($UserData.Description)) { $userParams['Description'] = $UserData.Description }

        if ($WhatIfMode) {
            $result.Success = $true
            $result.Message = "[WhatIf] L'utilisateur '$($UserData.SamAccountName)' serait cree"
            return $result
        }

        New-ADUser @userParams -ErrorAction Stop
        $result.Success = $true
        $result.Message = "Utilisateur '$($UserData.SamAccountName)' cree avec succes"

        # Ajouter aux groupes (tous verifies comme existants)
        if (-not [string]::IsNullOrWhiteSpace($UserData.Groups)) {
            $groups = $UserData.Groups -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
            foreach ($group in $groups) {
                try {
                    Add-ADGroupMember -Identity $group -Members $UserData.SamAccountName -ErrorAction Stop
                    [void]$result.GroupsAdded.Add($group)
                }
                catch { [void]$result.GroupsFailed.Add("$group ($($_.Exception.Message))") }
            }
        }

        Add-GeneratedPassword -SamAccountName $UserData.SamAccountName -DisplayName $displayName -Password $password
    }
    catch {
        $result.Success = $false
        $result.Message = "Erreur: $($_.Exception.Message)"
    }

    return $result
}

#==========================================================================
# Fonction    : Update-ADUserFromCSV
# Arguments   : PSObject userData, string baseDN, bool whatIfMode
# Return      : PSCustomObject resultat operation
# Description : Met a jour un utilisateur AD existant
#==========================================================================
function Update-ADUserFromCSV {
    param(
        [PSObject]$UserData,
        [string]$BaseDN,
        [switch]$WhatIfMode
    )

    $result = [PSCustomObject]@{
        Success        = $false
        Action         = "Update"
        SamAccountName = $UserData.SamAccountName
        Message        = ""
        GroupsAdded    = [System.Collections.ArrayList]::new()
        GroupsFailed   = [System.Collections.ArrayList]::new()
        UpdatedFields  = [System.Collections.ArrayList]::new()
    }

    try {
        $existingUser = Get-ADUser -Identity $UserData.SamAccountName -Properties GivenName, Surname, DisplayName, EmailAddress, Description, Enabled -ErrorAction Stop
        $updateParams = @{}

        if ($UserData.GivenName -ne $existingUser.GivenName) {
            $updateParams['GivenName'] = $UserData.GivenName
            [void]$result.UpdatedFields.Add("GivenName")
        }
        if ($UserData.Surname -ne $existingUser.Surname) {
            $updateParams['Surname'] = $UserData.Surname
            [void]$result.UpdatedFields.Add("Surname")
        }

        $displayName = if (-not [string]::IsNullOrWhiteSpace($UserData.DisplayName)) { $UserData.DisplayName }
                       else { "$($UserData.GivenName) $($UserData.Surname)" }
        if ($displayName -ne $existingUser.DisplayName) {
            $updateParams['DisplayName'] = $displayName
            [void]$result.UpdatedFields.Add("DisplayName")
        }

        if (-not [string]::IsNullOrWhiteSpace($UserData.Email) -and $UserData.Email -ne $existingUser.EmailAddress) {
            $updateParams['EmailAddress'] = $UserData.Email
            [void]$result.UpdatedFields.Add("EmailAddress")
        }
        if (-not [string]::IsNullOrWhiteSpace($UserData.Description) -and $UserData.Description -ne $existingUser.Description) {
            $updateParams['Description'] = $UserData.Description
            [void]$result.UpdatedFields.Add("Description")
        }

        if ($WhatIfMode) {
            $result.Success = $true
            $result.Message = "[WhatIf] L'utilisateur '$($UserData.SamAccountName)' serait mis a jour"
            return $result
        }

        if ($updateParams.Count -gt 0) {
            Set-ADUser -Identity $UserData.SamAccountName @updateParams -ErrorAction Stop
        }

        # Gerer les groupes
        if (-not [string]::IsNullOrWhiteSpace($UserData.Groups)) {
            $groups = $UserData.Groups -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
            foreach ($group in $groups) {
                try {
                    if (Test-ADObjectExists -Identity $group -ObjectType "Group") {
                        $members = Get-ADGroupMember -Identity $group -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SamAccountName
                        if ($members -notcontains $UserData.SamAccountName) {
                            Add-ADGroupMember -Identity $group -Members $UserData.SamAccountName -ErrorAction Stop
                            [void]$result.GroupsAdded.Add($group)
                        }
                    }
                    else { [void]$result.GroupsFailed.Add("$group (inexistant)") }
                }
                catch { [void]$result.GroupsFailed.Add("$group ($($_.Exception.Message))") }
            }
        }

        $result.Success = $true
        $result.Message = if ($result.UpdatedFields.Count -gt 0 -or $result.GroupsAdded.Count -gt 0) {
            "Utilisateur '$($UserData.SamAccountName)' mis a jour"
        } else { "Utilisateur '$($UserData.SamAccountName)' deja a jour" }
    }
    catch {
        $result.Success = $false
        $result.Message = "Erreur: $($_.Exception.Message)"
    }

    return $result
}
#endregion

#region Fonctions - Groupes
#==========================================================================
# Fonction    : New-ADGroupFromCSV
# Arguments   : PSObject groupData, string baseDN, bool whatIfMode
# Return      : PSCustomObject resultat operation
# Description : Cree un groupe AD a partir d'une ligne CSV
#==========================================================================
function New-ADGroupFromCSV {
    param(
        [PSObject]$GroupData,
        [string]$BaseDN,
        [switch]$WhatIfMode,
        [switch]$ContinueOnMissingMember
    )

    $result = [PSCustomObject]@{
        Success        = $false
        Action         = "Create"
        SamAccountName = $GroupData.SamAccountName
        Message        = ""
        MembersAdded   = [System.Collections.ArrayList]::new()
        MembersFailed  = [System.Collections.ArrayList]::new()
        MissingMembers = [System.Collections.ArrayList]::new()
    }

    try {
        $ouDN = if ($GroupData.OU -notmatch "DC=") { "$($GroupData.OU),$BaseDN" } else { $GroupData.OU }

        # Verifier les membres AVANT de creer le groupe si l'option ContinueOnMissingMember n'est pas activee
        if (-not [string]::IsNullOrWhiteSpace($GroupData.Members) -and -not $ContinueOnMissingMember) {
            $members = $GroupData.Members -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
            foreach ($member in $members) {
                if (-not (Test-ADObjectExistsAny -Identity $member)) {
                    [void]$result.MissingMembers.Add($member)
                }
            }
            # Si des membres manquent et l'option n'est pas activee, arreter le deploiement
            if ($result.MissingMembers.Count -gt 0) {
                $result.Success = $false
                $result.Message = "Membres inexistants: $($result.MissingMembers -join ', ')"
                return $result
            }
        }

        $groupParams = @{
            Name           = $GroupData.Name
            SamAccountName = $GroupData.SamAccountName
            GroupScope     = $GroupData.GroupScope
            GroupCategory  = $GroupData.GroupCategory
            Path           = $ouDN
        }

        if (-not [string]::IsNullOrWhiteSpace($GroupData.Description)) {
            $groupParams['Description'] = $GroupData.Description
        }

        if ($WhatIfMode) {
            $result.Success = $true
            $result.Message = "[WhatIf] Le groupe '$($GroupData.Name)' serait cree"
            return $result
        }

        New-ADGroup @groupParams -ErrorAction Stop
        $result.Success = $true
        $result.Message = "Groupe '$($GroupData.Name)' cree avec succes"

        # Ajouter les membres
        if (-not [string]::IsNullOrWhiteSpace($GroupData.Members)) {
            $members = $GroupData.Members -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
            foreach ($member in $members) {
                try {
                    if (Test-ADObjectExistsAny -Identity $member) {
                        Add-ADGroupMember -Identity $GroupData.SamAccountName -Members $member -ErrorAction Stop
                        [void]$result.MembersAdded.Add($member)
                    }
                    else {
                        # Membre inexistant mais on continue (ContinueOnMissingMember est active)
                        [void]$result.MembersFailed.Add("$member (inexistant - ignore)")
                    }
                }
                catch { [void]$result.MembersFailed.Add("$member ($($_.Exception.Message))") }
            }
        }
    }
    catch {
        $result.Success = $false
        $result.Message = "Erreur: $($_.Exception.Message)"
    }

    return $result
}

#==========================================================================
# Fonction    : Update-ADGroupFromCSV
# Arguments   : PSObject groupData, string baseDN, bool whatIfMode
# Return      : PSCustomObject resultat operation
# Description : Met a jour un groupe AD existant
#==========================================================================
function Update-ADGroupFromCSV {
    param(
        [PSObject]$GroupData,
        [string]$BaseDN,
        [switch]$WhatIfMode,
        [switch]$ContinueOnMissingMember
    )

    $result = [PSCustomObject]@{
        Success        = $false
        Action         = "Update"
        SamAccountName = $GroupData.SamAccountName
        Message        = ""
        MembersAdded   = [System.Collections.ArrayList]::new()
        MembersFailed  = [System.Collections.ArrayList]::new()
        MissingMembers = [System.Collections.ArrayList]::new()
        UpdatedFields  = [System.Collections.ArrayList]::new()
    }

    try {
        $existingGroup = Get-ADGroup -Identity $GroupData.SamAccountName -Properties Description -ErrorAction Stop
        $updateParams = @{}

        if (-not [string]::IsNullOrWhiteSpace($GroupData.Description) -and $GroupData.Description -ne $existingGroup.Description) {
            $updateParams['Description'] = $GroupData.Description
            [void]$result.UpdatedFields.Add("Description")
        }

        # Verifier les membres AVANT si l'option ContinueOnMissingMember n'est pas activee
        if (-not [string]::IsNullOrWhiteSpace($GroupData.Members) -and -not $ContinueOnMissingMember) {
            $members = $GroupData.Members -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
            foreach ($member in $members) {
                if (-not (Test-ADObjectExistsAny -Identity $member)) {
                    [void]$result.MissingMembers.Add($member)
                }
            }
            if ($result.MissingMembers.Count -gt 0) {
                $result.Success = $false
                $result.Message = "Membres inexistants: $($result.MissingMembers -join ', ')"
                return $result
            }
        }

        if ($WhatIfMode) {
            $result.Success = $true
            $result.Message = "[WhatIf] Le groupe '$($GroupData.SamAccountName)' serait mis a jour"
            return $result
        }

        if ($updateParams.Count -gt 0) {
            Set-ADGroup -Identity $GroupData.SamAccountName @updateParams -ErrorAction Stop
        }

        # Gerer les membres
        if (-not [string]::IsNullOrWhiteSpace($GroupData.Members)) {
            $members = $GroupData.Members -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
            $currentMembers = Get-ADGroupMember -Identity $GroupData.SamAccountName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SamAccountName

            foreach ($member in $members) {
                try {
                    if ($currentMembers -notcontains $member) {
                        if (Test-ADObjectExistsAny -Identity $member) {
                            Add-ADGroupMember -Identity $GroupData.SamAccountName -Members $member -ErrorAction Stop
                            [void]$result.MembersAdded.Add($member)
                        }
                        else {
                            # Membre inexistant mais on continue (ContinueOnMissingMember est active)
                            [void]$result.MembersFailed.Add("$member (inexistant - ignore)")
                        }
                    }
                }
                catch { [void]$result.MembersFailed.Add("$member ($($_.Exception.Message))") }
            }
        }

        $result.Success = $true
        $result.Message = if ($result.UpdatedFields.Count -gt 0 -or $result.MembersAdded.Count -gt 0) {
            "Groupe '$($GroupData.SamAccountName)' mis a jour"
        } else { "Groupe '$($GroupData.SamAccountName)' deja a jour" }
    }
    catch {
        $result.Success = $false
        $result.Message = "Erreur: $($_.Exception.Message)"
    }

    return $result
}
#endregion

#region Fonctions - Ordinateurs
#==========================================================================
# Fonction    : New-ADComputerFromCSV
# Arguments   : PSObject computerData, string baseDN, bool whatIfMode
# Return      : PSCustomObject resultat operation
# Description : Cree un ordinateur AD a partir d'une ligne CSV
#==========================================================================
function New-ADComputerFromCSV {
    param(
        [PSObject]$ComputerData,
        [string]$BaseDN,
        [switch]$WhatIfMode
    )

    $result = [PSCustomObject]@{
        Success        = $false
        Action         = "Create"
        SamAccountName = $ComputerData.SamAccountName
        Message        = ""
    }

    try {
        $ouDN = if ($ComputerData.OU -notmatch "DC=") { "$($ComputerData.OU),$BaseDN" } else { $ComputerData.OU }

        $enabled = $true
        if ($ComputerData.PSObject.Properties.Name -contains 'Enabled') {
            $enabled = Convert-EnabledValue -Value $ComputerData.Enabled -Default $true
        }

        $samAccountName = $ComputerData.SamAccountName
        if (-not $samAccountName.EndsWith('$')) { $samAccountName = "$samAccountName$" }

        $computerParams = @{
            Name           = $ComputerData.Name
            SamAccountName = $samAccountName
            Path           = $ouDN
            Enabled        = $enabled
        }

        if (-not [string]::IsNullOrWhiteSpace($ComputerData.Description)) {
            $computerParams['Description'] = $ComputerData.Description
        }

        if ($WhatIfMode) {
            $result.Success = $true
            $result.Message = "[WhatIf] L'ordinateur '$($ComputerData.Name)' serait cree"
            return $result
        }

        New-ADComputer @computerParams -ErrorAction Stop
        $result.Success = $true
        $result.Message = "Ordinateur '$($ComputerData.Name)' cree avec succes"
    }
    catch {
        $result.Success = $false
        $result.Message = "Erreur: $($_.Exception.Message)"
    }

    return $result
}

#==========================================================================
# Fonction    : Update-ADComputerFromCSV
# Arguments   : PSObject computerData, string baseDN, bool whatIfMode
# Return      : PSCustomObject resultat operation
# Description : Met a jour un ordinateur AD existant
#==========================================================================
function Update-ADComputerFromCSV {
    param(
        [PSObject]$ComputerData,
        [string]$BaseDN,
        [switch]$WhatIfMode
    )

    $result = [PSCustomObject]@{
        Success        = $false
        Action         = "Update"
        SamAccountName = $ComputerData.SamAccountName
        Message        = ""
        UpdatedFields  = [System.Collections.ArrayList]::new()
    }

    try {
        $existingComputer = Get-ADComputer -Identity $ComputerData.Name -Properties Description, Enabled -ErrorAction Stop
        $updateParams = @{}

        if (-not [string]::IsNullOrWhiteSpace($ComputerData.Description) -and $ComputerData.Description -ne $existingComputer.Description) {
            $updateParams['Description'] = $ComputerData.Description
            [void]$result.UpdatedFields.Add("Description")
        }

        if ($ComputerData.PSObject.Properties.Name -contains 'Enabled' -and -not [string]::IsNullOrWhiteSpace($ComputerData.Enabled)) {
            $enabled = Convert-EnabledValue -Value $ComputerData.Enabled -Default $true
            if ($enabled -ne $existingComputer.Enabled) {
                $updateParams['Enabled'] = $enabled
                [void]$result.UpdatedFields.Add("Enabled")
            }
        }

        if ($WhatIfMode) {
            $result.Success = $true
            $result.Message = "[WhatIf] L'ordinateur '$($ComputerData.Name)' serait mis a jour"
            return $result
        }

        if ($updateParams.Count -gt 0) {
            Set-ADComputer -Identity $ComputerData.Name @updateParams -ErrorAction Stop
        }

        $result.Success = $true
        $result.Message = if ($result.UpdatedFields.Count -gt 0) {
            "Ordinateur '$($ComputerData.Name)' mis a jour"
        } else { "Ordinateur '$($ComputerData.Name)' deja a jour" }
    }
    catch {
        $result.Success = $false
        $result.Message = "Erreur: $($_.Exception.Message)"
    }

    return $result
}
#endregion

#region Fonctions - Rollback
#==========================================================================
# Fonction    : Invoke-Rollback
# Arguments   : string rollbackFile
# Return      : bool succes
# Description : Annule une session de deploiement precedente
#==========================================================================
function Invoke-Rollback {
    param([string]$RollbackFile)

    Clear-Screen
    Show-Banner
    Show-MenuHeader -Title "Rollback - Annulation"

    if (-not (Test-Path -Path $RollbackFile)) {
        Show-StatusMessage -Message "Fichier de rollback non trouve: $RollbackFile" -Level "ERROR"
        Wait-KeyPress
        return $false
    }

    try {
        $rollbackData = Get-Content -Path $RollbackFile -Raw | ConvertFrom-Json
    }
    catch {
        Show-StatusMessage -Message "Impossible de lire le fichier de rollback" -Level "ERROR"
        Wait-KeyPress
        return $false
    }

    Write-Host "    Session a annuler: $($rollbackData.SessionId)" -ForegroundColor Yellow
    Write-Host "    Date: $($rollbackData.Timestamp)" -ForegroundColor Yellow
    Write-Host "    Objets crees: $($rollbackData.CreatedObjects.Count)" -ForegroundColor Yellow
    Write-Host ""

    # Separer les objets par type pour l'affichage
    $users = @($rollbackData.CreatedObjects | Where-Object { $_.Type -eq "User" })
    $groups = @($rollbackData.CreatedObjects | Where-Object { $_.Type -eq "Group" })
    $computers = @($rollbackData.CreatedObjects | Where-Object { $_.Type -eq "Computer" })
    $ous = @($rollbackData.CreatedObjects | Where-Object { $_.Type -eq "OU" })

    # Afficher la liste des objets qui seront supprimes
    Write-Host "    Objets qui seront supprimes:" -ForegroundColor Cyan
    Write-Host ""

    if ($users.Count -gt 0) {
        Write-Host "    Utilisateurs ($($users.Count)):" -ForegroundColor White
        foreach ($u in $users) {
            Write-Host "      - $($u.Identity)" -ForegroundColor Gray
        }
    }
    if ($groups.Count -gt 0) {
        Write-Host "    Groupes ($($groups.Count)):" -ForegroundColor White
        foreach ($g in $groups) {
            Write-Host "      - $($g.Identity)" -ForegroundColor Gray
        }
    }
    if ($computers.Count -gt 0) {
        Write-Host "    Ordinateurs ($($computers.Count)):" -ForegroundColor White
        foreach ($c in $computers) {
            Write-Host "      - $($c.Identity)" -ForegroundColor Gray
        }
    }
    if ($ous.Count -gt 0) {
        Write-Host "    Unites d'organisation ($($ous.Count)):" -ForegroundColor White
        foreach ($o in $ous) {
            Write-Host "      - $($o.Identity)" -ForegroundColor Gray
        }
    }

    Write-Host ""

    if (-not (Show-Confirmation -Message "Voulez-vous vraiment supprimer ces objets ?")) {
        Write-Host "    Rollback annule." -ForegroundColor Yellow
        Wait-KeyPress
        return $false
    }

    $successCount = 0
    $errorCount = 0

    # Ordre de suppression: Computers -> Users -> Groups -> OUs (plus profondes d'abord)
    $orderedObjects = @()

    # D'abord les objets (Computers, Users, Groups)
    $orderedObjects += $rollbackData.CreatedObjects | Where-Object { $_.Type -ne "OU" } | Sort-Object -Property @{Expression = {
        switch ($_.Type) {
            "Computer" { 1 }
            "User" { 2 }
            "Group" { 3 }
        }
    }}

    # Ensuite les OUs triees par profondeur (les plus profondes d'abord)
    $orderedObjects += $rollbackData.CreatedObjects | Where-Object { $_.Type -eq "OU" } | Sort-Object -Property @{Expression = { ($_.Identity -split ',').Count }} -Descending

    Write-Host ""
    foreach ($obj in $orderedObjects) {
        try {
            switch ($obj.Type) {
                "User" { Remove-ADUser -Identity $obj.Identity -Confirm:$false -ErrorAction Stop }
                "Group" { Remove-ADGroup -Identity $obj.Identity -Confirm:$false -ErrorAction Stop }
                "Computer" { Remove-ADComputer -Identity $obj.Identity -Confirm:$false -ErrorAction Stop }
                "OU" {
                    # Desactiver la protection contre la suppression accidentelle
                    Set-ADOrganizationalUnit -Identity $obj.Identity -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
                    Remove-ADOrganizationalUnit -Identity $obj.Identity -Confirm:$false -ErrorAction Stop
                }
            }
            Show-ProgressItem -Identity $obj.Identity -Action "Suppression $($obj.Type)" -Status "OK"
            $successCount++
        }
        catch {
            Show-ProgressItem -Identity $obj.Identity -Action "Suppression $($obj.Type)" -Status "ERROR"
            $errorCount++
        }
    }

    Write-Host ""
    Write-Host "    Objets supprimes: $successCount" -ForegroundColor Green
    if ($errorCount -gt 0) {
        Write-Host "    Erreurs: $errorCount" -ForegroundColor Red
    }

    # Supprimer le fichier de rollback si tout s'est bien passe
    if ($errorCount -eq 0) {
        try {
            Remove-Item -Path $RollbackFile -Force -ErrorAction Stop
            Write-Host ""
            Write-Host "    Fichier de rollback supprime." -ForegroundColor Gray
        }
        catch {
            Write-Host ""
            Write-Host "    [WARN] Impossible de supprimer le fichier de rollback." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host ""
        Write-Host "    [INFO] Fichier de rollback conserve (erreurs detectees)." -ForegroundColor Yellow
    }

    Wait-KeyPress
    return $true
}

#==========================================================================
# Fonction    : Export-Templates
# Arguments   : aucun
# Return      : bool succes
# Description : Exporte les templates CSV dans le dossier CSV
#==========================================================================
function Export-Templates {
    Clear-Screen
    Show-Banner
    Show-MenuHeader -Title "Export des templates CSV"

    if (-not (Test-Path -Path $script:CSVDefaultPath)) {
        try {
            New-Item -Path $script:CSVDefaultPath -ItemType Directory -Force | Out-Null
            Show-StatusMessage -Message "Dossier CSV cree: $($script:CSVDefaultPath)" -Level "OK"
        }
        catch {
            Show-StatusMessage -Message "Impossible de creer le dossier CSV" -Level "ERROR"
            Wait-KeyPress
            return $false
        }
    }

    $templates = @(
        @{ Name = "Users"; Columns = @('SamAccountName', 'GivenName', 'Surname', 'DisplayName', 'Email', 'Password', 'OU', 'Groups', 'Enabled', 'Description') },
        @{ Name = "Groups"; Columns = @('Name', 'SamAccountName', 'GroupScope', 'GroupCategory', 'OU', 'Description', 'Members') },
        @{ Name = "Computers"; Columns = @('Name', 'SamAccountName', 'OU', 'Description', 'Enabled') }
    )

    foreach ($template in $templates) {
        $filePath = Join-Path -Path $script:CSVDefaultPath -ChildPath "$($template.Name).csv"
        try {
            $template.Columns -join "," | Out-File -FilePath $filePath -Encoding UTF8
            Show-StatusMessage -Message "$($template.Name).csv exporte" -Level "OK"
        }
        catch {
            Show-StatusMessage -Message "Erreur export $($template.Name).csv" -Level "ERROR"
        }
    }

    Write-Host ""
    Write-Host "    Les templates ont ete exportes dans: $($script:CSVDefaultPath)" -ForegroundColor Cyan
    Write-Host "    Editez ces fichiers puis lancez le deploiement." -ForegroundColor Yellow

    Wait-KeyPress
    return $true
}
#endregion

#region Navigation et Menus
#==========================================================================
# Fonction    : Show-MainMenu
# Arguments   : aucun
# Return      : void
# Description : Affiche et gere le menu principal
#==========================================================================
function Show-MainMenu {
    while ($true) {
        Clear-Screen
        Show-Banner
        Show-MenuHeader -Title "Menu Principal"

        Show-MenuItem -Number "1" -Text "Deployer des objets AD"
        Show-MenuItem -Number "2" -Text "Rollback (annuler un deploiement)"
        Show-MenuItem -Number "3" -Text "Dechiffrer archive mots de passe"
        Show-MenuItem -Number "4" -Text "Exporter les templates CSV"
        Show-MenuItem -Number "5" -Text "Aide"
        Show-MenuItem -Number "0" -Text "Quitter" -Color "Red"

        $choice = Get-UserChoice -Prompt "Votre choix" -ValidChoices @("0", "1", "2", "3", "4", "5")

        switch ($choice) {
            "1" { Show-DeployTypeMenu }
            "2" { Show-RollbackMenu }
            "3" { Show-DecryptMenu }
            "4" { Export-Templates }
            "5" { Show-HelpScreen }
            "0" { return }
        }
    }
}

#==========================================================================
# Fonction    : Show-DeployTypeMenu
# Arguments   : aucun
# Return      : void
# Description : Affiche le menu de selection du type de deploiement
#==========================================================================
function Show-DeployTypeMenu {
    while ($true) {
        Clear-Screen
        Show-Banner
        Show-MenuHeader -Title "Deploiement - Type d'objet"

        Show-MenuItem -Number "1" -Text "Utilisateurs"
        Show-MenuItem -Number "2" -Text "Groupes"
        Show-MenuItem -Number "3" -Text "Ordinateurs"
        Show-MenuItem -Number "4" -Text "Tous les types"
        Show-MenuItem -Number "0" -Text "Retour" -Color "Yellow"

        $choice = Get-UserChoice -Prompt "Votre choix" -ValidChoices @("0", "1", "2", "3", "4")

        $deployType = switch ($choice) {
            "1" { "Users" }
            "2" { "Groups" }
            "3" { "Computers" }
            "4" { "All" }
            "0" { return }
            default { $null }
        }

        if ($deployType) {
            Show-DeployConfigMenu -DeployType $deployType
        }
    }
}

#==========================================================================
# Fonction    : Show-DeployConfigMenu
# Arguments   : string deployType
# Return      : void
# Description : Affiche le menu de configuration du deploiement
#==========================================================================
function Show-DeployConfigMenu {
    param([string]$DeployType)

    # Determiner les fichiers CSV
    $csvFiles = @{}
    if ($DeployType -eq "All" -or $DeployType -eq "Users") { $csvFiles["Users"] = $script:DefaultCSVPaths.Users }
    if ($DeployType -eq "All" -or $DeployType -eq "Groups") { $csvFiles["Groups"] = $script:DefaultCSVPaths.Groups }
    if ($DeployType -eq "All" -or $DeployType -eq "Computers") { $csvFiles["Computers"] = $script:DefaultCSVPaths.Computers }

    # Reset config
    $script:DeployConfig.WhatIf = $false
    $script:DeployConfig.Update = $false
    $script:DeployConfig.CreateOUs = $false
    $script:DeployConfig.EncryptPasswords = $false
    $script:DeployConfig.ContinueOnMissingMember = $false

    while ($true) {
        Clear-Screen
        Show-Banner
        Show-MenuHeader -Title "Configuration - $DeployType"

        # Afficher les fichiers CSV
        Write-Host "    Fichiers CSV:" -ForegroundColor Cyan
        foreach ($key in $csvFiles.Keys) {
            $exists = Test-Path -Path $csvFiles[$key] -PathType Leaf
            $status = if ($exists) { "[OK]" } else { "[MANQUANT]" }
            $color = if ($exists) { "Green" } else { "Red" }
            Write-Host "      $key : " -NoNewline
            Write-Host $status -ForegroundColor $color -NoNewline
            Write-Host " $($csvFiles[$key])" -ForegroundColor Gray
        }

        Write-Host ""
        Write-Host "    Options:" -ForegroundColor Cyan
        $whatifMark = if ($script:DeployConfig.WhatIf) { "[X]" } else { "[ ]" }
        $updateMark = if ($script:DeployConfig.Update) { "[X]" } else { "[ ]" }
        $createOUsMark = if ($script:DeployConfig.CreateOUs) { "[X]" } else { "[ ]" }
        $encryptMark = if ($script:DeployConfig.EncryptPasswords) { "[X]" } else { "[ ]" }
        $continueMemberMark = if ($script:DeployConfig.ContinueOnMissingMember) { "[X]" } else { "[ ]" }

        Show-MenuItem -Number "1" -Text "$whatifMark Mode simulation (WhatIf)"
        Show-MenuItem -Number "2" -Text "$updateMark Mettre a jour les existants"
        Show-MenuItem -Number "3" -Text "$createOUsMark Creer les OUs manquantes"
        if ($DeployType -eq "All" -or $DeployType -eq "Users") {
            Show-MenuItem -Number "4" -Text "$encryptMark Chiffrer les mots de passe"
        }
        if ($DeployType -eq "All" -or $DeployType -eq "Groups") {
            Show-MenuItem -Number "5" -Text "$continueMemberMark Continuer si un membre n'existe pas (Groupes)"
        }
        Write-Host ""
        Show-MenuItem -Number "C" -Text "Changer le chemin CSV"
        Show-MenuItem -Number "V" -Text "Valider et continuer" -Color "Green"
        Show-MenuItem -Number "0" -Text "Retour" -Color "Yellow"

        $validChoices = @("0", "1", "2", "3", "C", "c", "V", "v")
        if ($DeployType -eq "All" -or $DeployType -eq "Users") { $validChoices += "4" }
        if ($DeployType -eq "All" -or $DeployType -eq "Groups") { $validChoices += "5" }

        $choice = Get-UserChoice -Prompt "Votre choix" -ValidChoices $validChoices

        switch ($choice.ToUpper()) {
            "1" { $script:DeployConfig.WhatIf = -not $script:DeployConfig.WhatIf }
            "2" { $script:DeployConfig.Update = -not $script:DeployConfig.Update }
            "3" { $script:DeployConfig.CreateOUs = -not $script:DeployConfig.CreateOUs }
            "4" { $script:DeployConfig.EncryptPasswords = -not $script:DeployConfig.EncryptPasswords }
            "5" { $script:DeployConfig.ContinueOnMissingMember = -not $script:DeployConfig.ContinueOnMissingMember }
            "C" {
                Write-Host ""
                Write-Host "    Entrez le nouveau chemin CSV (ou appuyez sur Entree pour annuler):" -ForegroundColor Cyan
                $newPath = Read-Host "    "
                if (-not [string]::IsNullOrWhiteSpace($newPath) -and (Test-Path -Path $newPath -PathType Leaf)) {
                    if ($DeployType -ne "All") {
                        $csvFiles[$DeployType] = $newPath
                    }
                }
            }
            "V" {
                Show-DeployConfirmation -DeployType $DeployType -CSVFiles $csvFiles
                return
            }
            "0" { return }
        }
    }
}

#==========================================================================
# Fonction    : Show-DeployConfirmation
# Arguments   : string deployType, hashtable csvFiles
# Return      : void
# Description : Affiche la confirmation et lance le deploiement
#==========================================================================
function Show-DeployConfirmation {
    param(
        [string]$DeployType,
        [hashtable]$CSVFiles
    )

    Clear-Screen
    Show-Banner
    Show-MenuHeader -Title "Confirmation du deploiement"

    # Compter les entrees
    $totalEntries = 0
    foreach ($key in $CSVFiles.Keys) {
        if (Test-Path -Path $CSVFiles[$key] -PathType Leaf) {
            $data = Import-Csv -Path $CSVFiles[$key] -Encoding UTF8 -ErrorAction SilentlyContinue
            if ($data) { $totalEntries += $data.Count }
        }
    }

    Write-Host "    Type        : $DeployType" -ForegroundColor White
    Write-Host "    Entrees     : $totalEntries objets" -ForegroundColor White
    Write-Host ""
    Write-Host "    Fichiers CSV:" -ForegroundColor Cyan
    foreach ($key in $CSVFiles.Keys) {
        Write-Host "      - $key : $($CSVFiles[$key])" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "    Options actives:" -ForegroundColor Cyan
    if ($script:DeployConfig.WhatIf) { Write-Host "      - Mode simulation (WhatIf)" -ForegroundColor Yellow }
    if ($script:DeployConfig.Update) { Write-Host "      - Mise a jour des existants" -ForegroundColor Cyan }
    if ($script:DeployConfig.CreateOUs) { Write-Host "      - Creation automatique des OUs" -ForegroundColor Cyan }
    if ($script:DeployConfig.EncryptPasswords) { Write-Host "      - Chiffrement des mots de passe" -ForegroundColor Cyan }
    if (-not $script:DeployConfig.WhatIf -and -not $script:DeployConfig.Update -and -not $script:DeployConfig.CreateOUs -and -not $script:DeployConfig.EncryptPasswords) {
        Write-Host "      (aucune)" -ForegroundColor Gray
    }

    Write-Host ""
    if (-not (Show-Confirmation -Message "Lancer le deploiement ?")) {
        Write-Host "    Deploiement annule." -ForegroundColor Yellow
        Wait-KeyPress
        return
    }

    # Lancer le deploiement
    Start-Deployment -DeployType $DeployType -CSVFiles $CSVFiles
}

#==========================================================================
# Fonction    : Test-DeploymentPrerequisites
# Arguments   : hashtable csvFiles, string baseDN, hashtable deployConfig
# Return      : hashtable resultat validation
# Description : Pre-validation complete avant deploiement
#==========================================================================
function Test-DeploymentPrerequisites {
    param(
        [hashtable]$CSVFiles,
        [string]$BaseDN,
        [hashtable]$DeployConfig
    )

    $result = @{
        IsValid        = $true
        Errors         = [System.Collections.ArrayList]::new()
        Warnings       = [System.Collections.ArrayList]::new()
        ValidationData = @{}
    }

    $deployOrder = @("Groups", "Users", "Computers")

    # Phase 1: Valider les CSV et collecter les donnees
    foreach ($objectType in $deployOrder) {
        if (-not $CSVFiles.ContainsKey($objectType)) { continue }
        $csvPath = $CSVFiles[$objectType]
        if (-not (Test-Path -Path $csvPath -PathType Leaf)) { continue }

        $validation = Test-CSVFile -Path $csvPath -ObjectType $objectType

        if (-not $validation.IsValid) {
            foreach ($err in $validation.Errors) {
                [void]$result.Errors.Add(@{ Type = $objectType; Identity = "CSV"; Message = $err })
            }
            $result.IsValid = $false
            continue
        }

        if ($validation.Data.Count -eq 0) {
            continue
        }

        $result.ValidationData[$objectType] = @{
            Data    = $validation.Data
            Valid   = $true
            Errors  = [System.Collections.ArrayList]::new()
        }
    }

    # Si erreurs CSV, arreter ici
    if (-not $result.IsValid) {
        return $result
    }

    # Phase 2: Verifier les OUs (si CreateOUs=false)
    if (-not $DeployConfig.CreateOUs) {
        foreach ($objectType in $deployOrder) {
            if (-not $result.ValidationData.ContainsKey($objectType)) { continue }
            $data = $result.ValidationData[$objectType].Data

            $ouCheck = Test-OUsExist -Data $data -BaseDN $BaseDN
            if (-not $ouCheck.AllExist) {
                foreach ($missingOU in $ouCheck.MissingOUs) {
                    [void]$result.Errors.Add(@{ Type = $objectType; Identity = "OU"; Message = "OU manquante: $missingOU" })
                }
                $result.IsValid = $false
            }
        }
    }

    # Si erreurs OUs, arreter ici
    if (-not $result.IsValid) {
        return $result
    }

    # Phase 3: Collecter les groupes qui seront crees dans cette session
    $groupsToBeCreated = @{}
    if ($result.ValidationData.ContainsKey("Groups")) {
        foreach ($grp in $result.ValidationData["Groups"].Data) {
            $groupsToBeCreated[$grp.SamAccountName] = $true
        }
    }

    # Phase 4: Verifier les objets existants et les dependances
    foreach ($objectType in $deployOrder) {
        if (-not $result.ValidationData.ContainsKey($objectType)) { continue }
        $data = $result.ValidationData[$objectType].Data
        $singularType = $objectType.TrimEnd('s')

        foreach ($item in $data) {
            $identity = $item.SamAccountName
            $exists = Test-ADObjectExists -Identity $identity -ObjectType $singularType

            # Objet existant sans mode Update = Warning
            if ($exists -and -not $DeployConfig.Update) {
                [void]$result.Warnings.Add(@{ Type = $objectType; Identity = $identity; Message = "Ignore (existe deja)" })
                continue
            }

            # Verification specifique pour Users: groupes doivent exister
            if ($objectType -eq "Users" -and -not [string]::IsNullOrWhiteSpace($item.Groups)) {
                $groupNames = $item.Groups -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
                $missingGroups = [System.Collections.ArrayList]::new()

                foreach ($groupName in $groupNames) {
                    # Verifier si le groupe existe OU sera cree dans cette session
                    $groupExists = Test-ADObjectExists -Identity $groupName -ObjectType "Group"
                    if (-not $groupExists -and -not $groupsToBeCreated.ContainsKey($groupName)) {
                        [void]$missingGroups.Add($groupName)
                    }
                }

                if ($missingGroups.Count -gt 0) {
                    [void]$result.Errors.Add(@{
                        Type     = "Users"
                        Identity = $identity
                        Message  = "Groupes inexistants: $($missingGroups -join ', ')"
                    })
                    $result.IsValid = $false
                }
            }

            # Verification specifique pour Groups: membres doivent exister
            if ($objectType -eq "Groups" -and -not [string]::IsNullOrWhiteSpace($item.Members)) {
                $memberNames = $item.Members -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
                $missingMembers = [System.Collections.ArrayList]::new()

                foreach ($memberName in $memberNames) {
                    if (-not (Test-ADObjectExistsAny -Identity $memberName)) {
                        [void]$missingMembers.Add($memberName)
                    }
                }

                if ($missingMembers.Count -gt 0) {
                    if ($DeployConfig.ContinueOnMissingMember) {
                        # Warning seulement
                        [void]$result.Warnings.Add(@{
                            Type     = "Groups"
                            Identity = $identity
                            Message  = "Membres manquants (seront ignores): $($missingMembers -join ', ')"
                        })
                    }
                    else {
                        # Erreur bloquante
                        [void]$result.Errors.Add(@{
                            Type     = "Groups"
                            Identity = $identity
                            Message  = "Membres inexistants: $($missingMembers -join ', ')"
                        })
                        $result.IsValid = $false
                    }
                }
            }
        }
    }

    return $result
}

#==========================================================================
# Fonction    : Show-PreValidationResults
# Arguments   : hashtable preValidation
# Return      : void
# Description : Affiche les resultats de la pre-validation
#==========================================================================
function Show-PreValidationResults {
    param(
        [hashtable]$PreValidation
    )

    Write-Host ""
    Write-Host "  === Pre-validation ===" -ForegroundColor Cyan

    # Grouper les warnings par type
    $warningsByType = @{}
    foreach ($warn in $PreValidation.Warnings) {
        if (-not $warningsByType.ContainsKey($warn.Type)) {
            $warningsByType[$warn.Type] = [System.Collections.ArrayList]::new()
        }
        [void]$warningsByType[$warn.Type].Add($warn)
    }

    # Grouper les erreurs par type
    $errorsByType = @{}
    foreach ($err in $PreValidation.Errors) {
        if (-not $errorsByType.ContainsKey($err.Type)) {
            $errorsByType[$err.Type] = [System.Collections.ArrayList]::new()
        }
        [void]$errorsByType[$err.Type].Add($err)
    }

    # Afficher les warnings
    foreach ($type in $warningsByType.Keys) {
        Write-Host ""
        $typeLabel = $type
        if ($type -eq "Groups" -and $script:DeployConfig.ContinueOnMissingMember) {
            $typeLabel = "$type (option 'Continuer si membre n'existe pas' activee)"
        }
        Write-Host "  [WARN] ${typeLabel}:" -ForegroundColor Yellow
        foreach ($warn in $warningsByType[$type]) {
            if ($warn.Message -eq "Ignore (existe deja)") {
                Write-Host "    Ignore (existe deja): $($warn.Identity)" -ForegroundColor Yellow
            }
            else {
                Write-Host "    $($warn.Identity) : $($warn.Message)" -ForegroundColor Yellow
            }
        }
    }

    # Afficher les erreurs
    foreach ($type in $errorsByType.Keys) {
        Write-Host ""
        Write-Host "  [ERREUR] ${type}:" -ForegroundColor Red
        foreach ($err in $errorsByType[$type]) {
            Write-Host "    Echec pour $($err.Identity) : $($err.Message)" -ForegroundColor Red
        }
    }
}

#==========================================================================
# Fonction    : Start-Deployment
# Arguments   : string deployType, hashtable csvFiles
# Return      : void
# Description : Execute le deploiement des objets AD
#==========================================================================
function Start-Deployment {
    param(
        [string]$DeployType,
        [hashtable]$CSVFiles
    )

    Clear-Screen
    Show-Banner
    Show-MenuHeader -Title "Deploiement en cours"

    $BaseDN = $script:DomainInfo.BaseDN

    # === PRE-VALIDATION COMPLETE ===
    $preValidation = Test-DeploymentPrerequisites -CSVFiles $CSVFiles -BaseDN $BaseDN -DeployConfig $script:DeployConfig

    # Afficher les resultats si warnings ou erreurs
    if ($preValidation.Warnings.Count -gt 0 -or $preValidation.Errors.Count -gt 0) {
        Show-PreValidationResults -PreValidation $preValidation
    }

    # Si erreurs, arreter le deploiement
    if (-not $preValidation.IsValid) {
        Write-Host ""
        Write-Host "  Deploiement annule. Corrigez les erreurs ci-dessus." -ForegroundColor Red
        Wait-KeyPress
        return
    }

    # Verifier si des donnees valides existent
    $hasValidData = $preValidation.ValidationData.Count -gt 0
    if (-not $hasValidData) {
        Write-Host ""
        Write-Host "  Aucune donnee valide a deployer." -ForegroundColor Yellow
        Wait-KeyPress
        return
    }

    # Pre-validation OK
    if ($preValidation.Warnings.Count -gt 0 -or $preValidation.Errors.Count -gt 0) {
        Write-Host ""
    }
    Write-Host ""
    Write-Host "  Pre-validation OK." -ForegroundColor Green

    # Mode WhatIf - pas de logging
    $enableLogging = -not $script:DeployConfig.WhatIf

    if ($script:DeployConfig.WhatIf) {
        Write-Host ""
        Write-Host "    MODE SIMULATION - Aucune modification ne sera effectuee" -ForegroundColor Yellow
    }

    $rollbackData = @{
        SessionId      = [guid]::NewGuid().ToString()
        Timestamp      = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
        CreatedObjects = [System.Collections.ArrayList]::new()
    }

    $globalStats = @{ Created = 0; Updated = 0; Skipped = 0; Errors = 0; OUsCreated = 0 }
    $deployOrder = @("Groups", "Users", "Computers")

    # Initialiser le logging
    if ($enableLogging) {
        $logFolder = Join-Path -Path $script:ExportsPath -ChildPath "Logs"
        Initialize-Logging -LogPath $logFolder -SessionName "Deploy"
        Write-Log -Message "Debut de la session de deploiement" -Level "INFO" -NoConsole
        Write-Log -Message "Type de deploiement: $DeployType" -Level "INFO" -NoConsole
        Write-Log -Message "Options: WhatIf=$($script:DeployConfig.WhatIf), Update=$($script:DeployConfig.Update), CreateOUs=$($script:DeployConfig.CreateOUs), EncryptPasswords=$($script:DeployConfig.EncryptPasswords)" -Level "INFO" -NoConsole
    }

    # Creer les OUs manquantes (toutes en une fois)
    if ($script:DeployConfig.CreateOUs) {
        foreach ($objectType in $deployOrder) {
            if (-not $preValidation.ValidationData.ContainsKey($objectType)) { continue }
            $data = $preValidation.ValidationData[$objectType].Data

            $ouCheck = Test-OUsExist -Data $data -BaseDN $BaseDN
            if (-not $ouCheck.AllExist) {
                $ouResult = New-MissingOUs -MissingOUs $ouCheck.MissingOUs -BaseDN $BaseDN -WhatIfMode:$script:DeployConfig.WhatIf
                foreach ($ou in $ouResult.CreatedOUs) {
                    $prefix = if ($script:DeployConfig.WhatIf) { "[WhatIf]" } else { "[OK]" }
                    Write-Host "    $prefix OU creee: $ou" -ForegroundColor $(if ($script:DeployConfig.WhatIf) { "Yellow" } else { "Green" })
                    $globalStats.OUsCreated++
                    if (-not $script:DeployConfig.WhatIf) {
                        [void]$rollbackData.CreatedObjects.Add(@{ Type = "OU"; Identity = $ou })
                    }
                    if ($enableLogging -and -not $script:DeployConfig.WhatIf) {
                        Write-Log -Message "[OU] Creation: $ou" -Level "SUCCESS" -NoConsole
                    }
                }
                if ($ouResult.Errors.Count -gt 0) {
                    foreach ($err in $ouResult.Errors) {
                        Write-Host "    [ERREUR] $err" -ForegroundColor Red
                        if ($enableLogging) {
                            Write-Log -Message "[OU] $err" -Level "ERROR" -NoConsole
                        }
                    }
                }
            }
        }
    }

    # Deployer chaque type (utiliser les donnees pre-validees)
    foreach ($objectType in $deployOrder) {
        if (-not $preValidation.ValidationData.ContainsKey($objectType)) { continue }
        $validationData = $preValidation.ValidationData[$objectType]

        Write-Host ""
        Write-Host "  === $objectType ===" -ForegroundColor Magenta

        # Deployer les objets
        foreach ($item in $validationData.Data) {
            $identity = $item.SamAccountName
            $singularType = $objectType.TrimEnd('s')
            $exists = Test-ADObjectExists -Identity $identity -ObjectType $singularType

            if ($exists -and -not $script:DeployConfig.Update) {
                Show-ProgressItem -Identity $identity -Action $objectType -Status "SKIP"
                $globalStats.Skipped++
                if ($enableLogging) {
                    Write-Log -Message "[$objectType] Ignore (existe deja): $identity" -Level "INFO" -NoConsole
                }
                continue
            }

            $result = $null
            switch ($objectType) {
                "Users" {
                    if ($exists) {
                        $result = Update-ADUserFromCSV -UserData $item -BaseDN $BaseDN -WhatIfMode:$script:DeployConfig.WhatIf
                    }
                    else {
                        $result = New-ADUserFromCSV -UserData $item -BaseDN $BaseDN -PasswordLength $script:DeployConfig.PasswordLength -WhatIfMode:$script:DeployConfig.WhatIf
                    }
                }
                "Groups" {
                    if ($exists) {
                        $result = Update-ADGroupFromCSV -GroupData $item -BaseDN $BaseDN -WhatIfMode:$script:DeployConfig.WhatIf -ContinueOnMissingMember:$script:DeployConfig.ContinueOnMissingMember
                    }
                    else {
                        $result = New-ADGroupFromCSV -GroupData $item -BaseDN $BaseDN -WhatIfMode:$script:DeployConfig.WhatIf -ContinueOnMissingMember:$script:DeployConfig.ContinueOnMissingMember
                    }
                }
                "Computers" {
                    if ($exists) {
                        $result = Update-ADComputerFromCSV -ComputerData $item -BaseDN $BaseDN -WhatIfMode:$script:DeployConfig.WhatIf
                    }
                    else {
                        $result = New-ADComputerFromCSV -ComputerData $item -BaseDN $BaseDN -WhatIfMode:$script:DeployConfig.WhatIf
                    }
                }
            }

            if ($result.Success) {
                $status = if ($result.Action -eq "Create") { "OK" } else { "UPDATE" }
                Show-ProgressItem -Identity $identity -Action $objectType -Status $status

                if ($result.Action -eq "Create") {
                    $globalStats.Created++
                    if (-not $script:DeployConfig.WhatIf) {
                        [void]$rollbackData.CreatedObjects.Add(@{ Type = $singularType; Identity = $identity })
                    }
                    if ($enableLogging) {
                        Write-Log -Message "[$objectType] Creation reussie: $identity - $($result.Message)" -Level "SUCCESS" -NoConsole
                        # Logs detailles pour les groupes
                        if ($objectType -eq "Groups") {
                            if ($result.MembersAdded -and $result.MembersAdded.Count -gt 0) {
                                Write-Log -Message "[$objectType] Membres ajoutes a $identity : $($result.MembersAdded -join ', ')" -Level "INFO" -NoConsole
                            }
                            if ($result.MembersFailed -and $result.MembersFailed.Count -gt 0) {
                                Write-Log -Message "[$objectType] Membres non ajoutes a $identity : $($result.MembersFailed -join ', ')" -Level "WARNING" -NoConsole
                            }
                        }
                        # Logs detailles pour les utilisateurs
                        if ($objectType -eq "Users") {
                            if ($result.GroupsAdded -and $result.GroupsAdded.Count -gt 0) {
                                Write-Log -Message "[$objectType] Groupes assignes a $identity : $($result.GroupsAdded -join ', ')" -Level "INFO" -NoConsole
                            }
                            if ($result.GroupsFailed -and $result.GroupsFailed.Count -gt 0) {
                                Write-Log -Message "[$objectType] Groupes non assignes a $identity : $($result.GroupsFailed -join ', ')" -Level "WARNING" -NoConsole
                            }
                        }
                    }
                }
                else {
                    $globalStats.Updated++
                    if ($enableLogging) {
                        Write-Log -Message "[$objectType] Mise a jour reussie: $identity - $($result.Message)" -Level "SUCCESS" -NoConsole
                        # Logs detailles pour les mises a jour de groupes
                        if ($objectType -eq "Groups") {
                            if ($result.UpdatedFields -and $result.UpdatedFields.Count -gt 0) {
                                Write-Log -Message "[$objectType] Champs modifies pour $identity : $($result.UpdatedFields -join ', ')" -Level "INFO" -NoConsole
                            }
                            if ($result.MembersAdded -and $result.MembersAdded.Count -gt 0) {
                                Write-Log -Message "[$objectType] Membres ajoutes a $identity : $($result.MembersAdded -join ', ')" -Level "INFO" -NoConsole
                            }
                            if ($result.MembersFailed -and $result.MembersFailed.Count -gt 0) {
                                Write-Log -Message "[$objectType] Membres non ajoutes a $identity : $($result.MembersFailed -join ', ')" -Level "WARNING" -NoConsole
                            }
                        }
                        # Logs detailles pour les mises a jour d'utilisateurs
                        if ($objectType -eq "Users") {
                            if ($result.UpdatedFields -and $result.UpdatedFields.Count -gt 0) {
                                Write-Log -Message "[$objectType] Champs modifies pour $identity : $($result.UpdatedFields -join ', ')" -Level "INFO" -NoConsole
                            }
                            if ($result.GroupsAdded -and $result.GroupsAdded.Count -gt 0) {
                                Write-Log -Message "[$objectType] Groupes assignes a $identity : $($result.GroupsAdded -join ', ')" -Level "INFO" -NoConsole
                            }
                        }
                    }
                }
            }
            else {
                Show-ProgressItem -Identity $identity -Action $objectType -Status "ERROR"
                $globalStats.Errors++
                if ($enableLogging) {
                    Write-Log -Message "[$objectType] Echec pour $identity : $($result.Message)" -Level "ERROR" -NoConsole
                }
            }
        }
    }

    # Sauvegarder le rollback
    if ($rollbackData.CreatedObjects.Count -gt 0 -and -not $script:DeployConfig.WhatIf) {
        $rollbackFolder = Join-Path -Path $script:ExportsPath -ChildPath "Rollbacks"
        if (-not (Test-Path -Path $rollbackFolder)) {
            New-Item -Path $rollbackFolder -ItemType Directory -Force | Out-Null
        }
        $timestamp = Get-Date -Format "ddMMyyyy_HHmm"
        $rollbackFile = Join-Path -Path $rollbackFolder -ChildPath "rollback_${timestamp}.json"
        $rollbackData | ConvertTo-Json -Depth 10 | Out-File -FilePath $rollbackFile -Encoding UTF8
    }

    # Exporter les mots de passe
    if (-not $script:DeployConfig.WhatIf -and $script:GeneratedPasswords.Count -gt 0) {
        $pwdExport = Export-GeneratedPasswords -Encrypt:$script:DeployConfig.EncryptPasswords -ToolsPath $script:ToolsPath
        if ($pwdExport.Success -and $pwdExport.FilePath) {
            Write-Host ""
            Write-Host "    Mots de passe exportes: $($pwdExport.FilePath)" -ForegroundColor Cyan
            if ($pwdExport.Encrypted -and $pwdExport.EncryptionPassword) {
                Write-Host ""
                Write-Host "  ╭─────────────────────────────────────────────────────────╮" -ForegroundColor Red
                Write-Host "  │     MOT DE PASSE DE L'ARCHIVE CHIFFREE                  │" -ForegroundColor Red
                Write-Host "  │                                                         │" -ForegroundColor Red
                Write-Host "  │  $($pwdExport.EncryptionPassword.PadRight(50))     │" -ForegroundColor Green
                Write-Host "  │                                                         │" -ForegroundColor Red
                Write-Host "  │  CONSERVEZ CE MOT DE PASSE ! Il ne sera plus affiche.   │" -ForegroundColor Yellow
                Write-Host "  ╰─────────────────────────────────────────────────────────╯" -ForegroundColor Red
            }
        }
    }

    # Resume
    Write-Host ""
    Write-Host "  ╭─────────────────────────────────────────╮" -ForegroundColor Cyan
    Write-Host "  │         Statistiques globales           │" -ForegroundColor Cyan
    Write-Host "  ╰─────────────────────────────────────────╯" -ForegroundColor Cyan
    Write-Host ""
    if ($globalStats.OUsCreated -gt 0) {
        Write-Host "    OUs creees  : $($globalStats.OUsCreated)" -ForegroundColor Magenta
    }
    Write-Host "    Crees       : $($globalStats.Created)" -ForegroundColor Green
    Write-Host "    Mis a jour  : $($globalStats.Updated)" -ForegroundColor Cyan
    Write-Host "    Ignores     : $($globalStats.Skipped)" -ForegroundColor Yellow
    Write-Host "    Erreurs     : $($globalStats.Errors)" -ForegroundColor $(if ($globalStats.Errors -gt 0) { "Red" } else { "Green" })

    # Log du resume final
    if ($enableLogging) {
        Write-Log -Message "--- Resume de session ---" -Level "INFO" -NoConsole
        if ($globalStats.OUsCreated -gt 0) {
            Write-Log -Message "OUs creees: $($globalStats.OUsCreated)" -Level "INFO" -NoConsole
        }
        Write-Log -Message "Objets crees: $($globalStats.Created)" -Level "INFO" -NoConsole
        Write-Log -Message "Objets mis a jour: $($globalStats.Updated)" -Level "INFO" -NoConsole
        Write-Log -Message "Objets ignores: $($globalStats.Skipped)" -Level "INFO" -NoConsole
        Write-Log -Message "Erreurs: $($globalStats.Errors)" -Level "INFO" -NoConsole
        Write-Log -Message "Fin de la session de deploiement" -Level "INFO" -NoConsole
    }

    if ($script:DeployConfig.WhatIf) {
        Write-Host ""
        Write-Host "    [MODE SIMULATION] Aucune modification n'a ete effectuee." -ForegroundColor Yellow
    }

    # Nettoyer
    $script:GeneratedPasswords.Clear()

    Wait-KeyPress
}

#==========================================================================
# Fonction    : Show-RollbackMenu
# Arguments   : aucun
# Return      : void
# Description : Affiche le menu de selection des fichiers de rollback
#==========================================================================
function Show-RollbackMenu {
    Clear-Screen
    Show-Banner
    Show-MenuHeader -Title "Rollback - Selection"

    $rollbackFolder = Join-Path -Path $script:ExportsPath -ChildPath "Rollbacks"

    if (-not (Test-Path -Path $rollbackFolder)) {
        Write-Host "    Aucun fichier de rollback disponible." -ForegroundColor Yellow
        Wait-KeyPress
        return
    }

    $rollbackFiles = Get-ChildItem -Path $rollbackFolder -Filter "rollback_*.json" | Sort-Object LastWriteTime -Descending

    if ($rollbackFiles.Count -eq 0) {
        Write-Host "    Aucun fichier de rollback disponible." -ForegroundColor Yellow
        Wait-KeyPress
        return
    }

    Write-Host "    Fichiers de rollback disponibles:" -ForegroundColor Cyan
    Write-Host ""

    $index = 1
    foreach ($file in $rollbackFiles) {
        Show-MenuItem -Number "$index" -Text "$($file.Name) ($($file.LastWriteTime.ToString('dd/MM/yyyy HH:mm')))"
        $index++
    }

    Show-MenuItem -Number "0" -Text "Retour" -Color "Yellow"

    $validChoices = @("0") + (1..$rollbackFiles.Count | ForEach-Object { $_.ToString() })
    $choice = Get-UserChoice -Prompt "Votre choix" -ValidChoices $validChoices

    if ($choice -eq "0" -or $null -eq $choice) { return }

    $selectedIndex = [int]$choice - 1
    if ($selectedIndex -ge 0 -and $selectedIndex -lt $rollbackFiles.Count) {
        Invoke-Rollback -RollbackFile $rollbackFiles[$selectedIndex].FullName
    }
}

#==========================================================================
# Fonction    : Show-DecryptMenu
# Arguments   : aucun
# Return      : void
# Description : Affiche le menu de selection des archives de mots de passe
#==========================================================================
function Show-DecryptMenu {
    Clear-Screen
    Show-Banner
    Show-MenuHeader -Title "Dechiffrement - Selection"

    $passwordFolder = Join-Path -Path $script:ExportsPath -ChildPath "Passwords"

    if (-not (Test-Path -Path $passwordFolder)) {
        Write-Host "    Aucune archive de mots de passe disponible." -ForegroundColor Yellow
        Wait-KeyPress
        return
    }

    $archiveFiles = Get-ChildItem -Path $passwordFolder -Filter "Passwords_*.7z" | Sort-Object LastWriteTime -Descending

    if ($archiveFiles.Count -eq 0) {
        Write-Host "    Aucune archive de mots de passe disponible." -ForegroundColor Yellow
        Wait-KeyPress
        return
    }

    Write-Host "    Archives de mots de passe disponibles:" -ForegroundColor Cyan
    Write-Host ""

    $index = 1
    foreach ($file in $archiveFiles) {
        $sizeKB = [math]::Round($file.Length / 1KB, 1)
        Show-MenuItem -Number "$index" -Text "$($file.Name) ($($file.LastWriteTime.ToString('dd/MM/yyyy HH:mm')) - $sizeKB KB)"
        $index++
    }

    Show-MenuItem -Number "0" -Text "Retour" -Color "Yellow"

    $validChoices = @("0") + (1..$archiveFiles.Count | ForEach-Object { $_.ToString() })
    $choice = Get-UserChoice -Prompt "Votre choix" -ValidChoices $validChoices

    if ($choice -eq "0" -or $null -eq $choice) { return }

    $selectedIndex = [int]$choice - 1
    if ($selectedIndex -ge 0 -and $selectedIndex -lt $archiveFiles.Count) {
        Show-DecryptActionMenu -ArchiveFile $archiveFiles[$selectedIndex].FullName
    }
}

#==========================================================================
# Fonction    : Show-DecryptActionMenu
# Arguments   : string archiveFile
# Return      : void
# Description : Affiche le menu d'action pour l'archive selectionnee
#==========================================================================
function Show-DecryptActionMenu {
    param([string]$ArchiveFile)

    $archiveName = Split-Path -Path $ArchiveFile -Leaf

    while ($true) {
        Clear-Screen
        Show-Banner
        Show-MenuHeader -Title "Dechiffrement - Action"

        Write-Host "    Archive selectionnee:" -ForegroundColor Cyan
        Write-Host "    $archiveName" -ForegroundColor White
        Write-Host ""

        Show-MenuItem -Number "1" -Text "Lire le contenu (affichage console)"
        Show-MenuItem -Number "2" -Text "Lire le contenu (grille graphique)"
        Show-MenuItem -Number "3" -Text "Extraire l'archive"
        Show-MenuItem -Number "0" -Text "Retour" -Color "Yellow"

        $choice = Get-UserChoice -Prompt "Votre choix" -ValidChoices @("0", "1", "2", "3")

        switch ($choice) {
            "1" { Read-PasswordArchive -ArchiveFile $ArchiveFile -DisplayMode "Console" }
            "2" { Read-PasswordArchive -ArchiveFile $ArchiveFile -DisplayMode "GridView" }
            "3" { Extract-PasswordArchive -ArchiveFile $ArchiveFile }
            "0" { return }
        }
    }
}

#==========================================================================
# Fonction    : Read-PasswordArchive
# Arguments   : string archiveFile, string displayMode
# Return      : void
# Description : Lit le contenu CSV d'une archive sans l'extraire sur disque
#==========================================================================
function Read-PasswordArchive {
    param(
        [string]$ArchiveFile,
        [ValidateSet("Console", "GridView")]
        [string]$DisplayMode = "Console"
    )

    Clear-Screen
    Show-Banner
    Show-MenuHeader -Title "Lecture de l'archive"

    # Verifier 7za.exe
    $sevenZipPath = Join-Path -Path $script:ToolsPath -ChildPath "7za.exe"
    if (-not (Test-Path -Path $sevenZipPath -PathType Leaf)) {
        Write-Host "    [ERREUR] 7za.exe non trouve dans .\Tools\" -ForegroundColor Red
        Wait-KeyPress
        return
    }

    # Demander le mot de passe
    Write-Host "    Entrez le mot de passe de l'archive:" -ForegroundColor Cyan
    Write-Host "    " -NoNewline
    $passwordPlain = Read-Host

    if ([string]::IsNullOrWhiteSpace($passwordPlain)) {
        Write-Host ""
        Write-Host "    [ERREUR] Mot de passe requis." -ForegroundColor Red
        Wait-KeyPress
        return
    }

    Write-Host ""
    Write-Host "    Lecture en cours..." -ForegroundColor Gray

    try {
        # Utiliser 7za pour extraire vers stdout
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $sevenZipPath
        $psi.Arguments = "e `"$ArchiveFile`" -so -p$passwordPlain"
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true

        $process = [System.Diagnostics.Process]::Start($psi)
        $csvContent = $process.StandardOutput.ReadToEnd()
        $errorOutput = $process.StandardError.ReadToEnd()
        $process.WaitForExit()

        # Effacer le mot de passe de la memoire
        $passwordPlain = $null

        if ($process.ExitCode -ne 0) {
            Write-Host ""
            Write-Host "    [ERREUR] Impossible de dechiffrer l'archive." -ForegroundColor Red
            Write-Host "    Verifiez que le mot de passe est correct." -ForegroundColor Yellow
            Wait-KeyPress
            return
        }

        if ([string]::IsNullOrWhiteSpace($csvContent)) {
            Write-Host ""
            Write-Host "    [ERREUR] L'archive semble vide ou le format est invalide." -ForegroundColor Red
            Wait-KeyPress
            return
        }

        # Convertir le CSV en objets
        $passwords = $csvContent | ConvertFrom-Csv

        if ($null -eq $passwords -or $passwords.Count -eq 0) {
            Write-Host ""
            Write-Host "    [INFO] Aucune donnee dans le fichier CSV." -ForegroundColor Yellow
            Wait-KeyPress
            return
        }

        Write-Host ""
        Write-Host "    [OK] $($passwords.Count) entree(s) trouvee(s)." -ForegroundColor Green
        Write-Host ""

        if ($DisplayMode -eq "GridView") {
            # Affichage dans une grille graphique
            $passwords | Out-GridView -Title "Mots de passe - $(Split-Path -Path $ArchiveFile -Leaf)"
        }
        else {
            # Affichage console avec Format-Table
            $passwords | Format-Table -AutoSize | Out-String | ForEach-Object {
                Write-Host $_
            }
        }
    }
    catch {
        Write-Host ""
        Write-Host "    [ERREUR] $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        # S'assurer que le mot de passe est efface
        $passwordPlain = $null
    }

    Wait-KeyPress
}

#==========================================================================
# Fonction    : Extract-PasswordArchive
# Arguments   : string archiveFile
# Return      : void
# Description : Extrait une archive de mots de passe dans le meme dossier
#==========================================================================
function Extract-PasswordArchive {
    param([string]$ArchiveFile)

    Clear-Screen
    Show-Banner
    Show-MenuHeader -Title "Extraction de l'archive"

    # Verifier 7za.exe
    $sevenZipPath = Join-Path -Path $script:ToolsPath -ChildPath "7za.exe"
    if (-not (Test-Path -Path $sevenZipPath -PathType Leaf)) {
        Write-Host "    [ERREUR] 7za.exe non trouve dans .\Tools\" -ForegroundColor Red
        Wait-KeyPress
        return
    }

    $archiveFolder = Split-Path -Path $ArchiveFile -Parent
    $archiveName = Split-Path -Path $ArchiveFile -Leaf

    Write-Host "    Archive: $archiveName" -ForegroundColor White
    Write-Host "    Destination: $archiveFolder" -ForegroundColor Gray
    Write-Host ""

    # Demander le mot de passe
    Write-Host "    Entrez le mot de passe de l'archive:" -ForegroundColor Cyan
    Write-Host "    " -NoNewline
    $passwordPlain = Read-Host

    if ([string]::IsNullOrWhiteSpace($passwordPlain)) {
        Write-Host ""
        Write-Host "    [ERREUR] Mot de passe requis." -ForegroundColor Red
        Wait-KeyPress
        return
    }

    Write-Host ""
    Write-Host "    Extraction en cours..." -ForegroundColor Gray

    try {
        # Utiliser 7za pour extraire
        $arguments = @("x", "`"$ArchiveFile`"", "-o`"$archiveFolder`"", "-p$passwordPlain", "-y")
        $process = Start-Process -FilePath $sevenZipPath -ArgumentList $arguments -Wait -PassThru -NoNewWindow -RedirectStandardOutput "NUL"

        # Effacer le mot de passe de la memoire
        $passwordPlain = $null

        if ($process.ExitCode -eq 0) {
            # Trouver le fichier extrait
            $csvFileName = [System.IO.Path]::GetFileNameWithoutExtension($archiveName) + ".csv"
            $extractedFile = Join-Path -Path $archiveFolder -ChildPath $csvFileName

            Write-Host ""
            Write-Host "    [OK] Extraction reussie!" -ForegroundColor Green
            Write-Host ""
            Write-Host "    Fichier extrait:" -ForegroundColor Cyan
            Write-Host "    $extractedFile" -ForegroundColor White
            Write-Host ""
            Write-Host "    [SECURITE] Pensez a supprimer ce fichier apres utilisation." -ForegroundColor Yellow
        }
        else {
            Write-Host ""
            Write-Host "    [ERREUR] Impossible d'extraire l'archive (code: $($process.ExitCode))." -ForegroundColor Red
            Write-Host "    Verifiez que le mot de passe est correct." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host ""
        Write-Host "    [ERREUR] $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        # S'assurer que le mot de passe est efface
        $passwordPlain = $null
    }

    Wait-KeyPress
}

#==========================================================================
# Fonction    : Show-HelpScreen
# Arguments   : aucun
# Return      : void
# Description : Affiche l'ecran d'aide
#==========================================================================
function Show-HelpScreen {
    Clear-Screen
    Show-Banner
    Show-MenuHeader -Title "Aide"

    Write-Host "    ADFlow CLI - Deploiement Active Directory" -ForegroundColor Cyan
    Write-Host "    Auteur: Taeckens.M" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    UTILISATION:" -ForegroundColor Yellow
    Write-Host "      Lancez le script sans argument pour acceder a l'interface interactive." -ForegroundColor Gray
    Write-Host "      Toutes les options de deploiement sont configurables via les menus." -ForegroundColor Gray
    Write-Host ""
    Write-Host "    OPTIONS EN LIGNE DE COMMANDE:" -ForegroundColor Yellow
    Write-Host "      -ExportTemplate Exporte les templates CSV dans .\CSV\" -ForegroundColor Gray
    Write-Host "      -Help           Affiche cette aide" -ForegroundColor Gray
    Write-Host "      -Version        Affiche la version" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    FONCTIONNALITES DE L'INTERFACE:" -ForegroundColor Yellow
    Write-Host "      - Deploiement d'utilisateurs, groupes et ordinateurs" -ForegroundColor Gray
    Write-Host "      - Mode simulation (WhatIf)" -ForegroundColor Gray
    Write-Host "      - Mise a jour des objets existants" -ForegroundColor Gray
    Write-Host "      - Creation automatique des OUs manquantes" -ForegroundColor Gray
    Write-Host "      - Export des templates CSV" -ForegroundColor Gray
    Write-Host "      - Rollback des sessions precedentes" -ForegroundColor Gray
    Write-Host "      - Chiffrement des mots de passe (AES-256 avec 7-Zip embarque)" -ForegroundColor Gray
    Write-Host "      - Dechiffrement des archives de mots de passe" -ForegroundColor Gray
    Write-Host "        (lecture directe ou extraction)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    EXEMPLES:" -ForegroundColor Yellow
    Write-Host "      .\Start-ADFlow.ps1                  # Interface interactive" -ForegroundColor Gray
    Write-Host "      .\Start-ADFlow.ps1 -ExportTemplate  # Exporter les templates" -ForegroundColor Gray
    Write-Host ""

    Wait-KeyPress
}
#endregion

#region Point d'entree
#==========================================================================
# Fonction    : Main
# Arguments   : aucun
# Return      : void
# Description : Point d'entree principal du script
#==========================================================================
function Main {
    # Mode Version
    if ($Version) {
        Write-Host "$($script:AppName) - Version $($script:AppVersion)"
        return
    }

    # Mode Help
    if ($Help) {
        Show-Banner
        Show-HelpScreen
        return
    }

    # Mode Export Template
    if ($ExportTemplate) {
        Show-Banner
        $prereqs = Test-Prerequisites -Silent
        Export-Templates
        return
    }

    # Mode CLI Interactif
    Clear-Screen
    Show-Banner

    # Verifier les prerequis au demarrage
    $prereqs = Test-Prerequisites
    if (-not $prereqs.AllPassed) {
        Write-Host ""
        Write-Host "  Corrigez les erreurs ci-dessus avant de continuer." -ForegroundColor Red
        Wait-KeyPress
        return
    }

    Wait-KeyPress -Message "Appuyez sur Entree ou Echap pour continuer..."

    # Lancer le menu principal
    Show-MainMenu

    # Message de sortie
    Clear-Screen
    Show-Banner
    Write-Host "    Merci d'avoir utilise $($script:AppName) !" -ForegroundColor Cyan
    Write-Host ""
}

# Execution
Main
#endregion
