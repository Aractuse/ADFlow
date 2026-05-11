#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Télécharge et installe ADFlow sur le bureau.
.DESCRIPTION
    Ce script télécharge ADFlow depuis GitHub
    et le place dans un dossier sur le bureau de l'utilisateur.
.NOTES
    Usage: irm https://raw.githubusercontent.com/USER/REPO/main/Install-ADFlow.ps1 | iex
#>

# ============================================
# CONFIGURATION - URLs des scripts à modifier
# ============================================
$ADFlow_URL = "https://raw.githubusercontent.com/Aractuse/ADFlow/refs/heads/main/Start-ADFlow.ps1"
$ADFlow_7za_URL = "https://raw.githubusercontent.com/Aractuse/ADFlow/refs/heads/main/Tools/7za.exe"

$ADFlow_Name = "Start-ADFlow.ps1"
$ADFlow_7za_Name = "7za.exe"

# ============================================
# CONFIGURATION - Dossiers de destination
# ============================================
$FolderName = "AD-Tools"
$DestinationPath = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath $FolderName
$ADFlowPath = Join-Path -Path $DestinationPath -ChildPath "ADFlow"
$ADFlowToolsPath = Join-Path -Path $ADFlowPath -ChildPath "Tools"

# ============================================
# SCRIPT PRINCIPAL
# ============================================

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   Installation d'ADFlow" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Création des dossiers de destination
$FoldersToCreate = @($DestinationPath, $ADFlowPath, $ADFlowToolsPath)

foreach ($Folder in $FoldersToCreate) {
    if (-not (Test-Path -Path $Folder)) {
        Write-Host "[+] Création du dossier: $Folder" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
    } else {
        Write-Host "[*] Le dossier existe déjà: $Folder" -ForegroundColor Gray
    }
}

# Fonction de téléchargement
function Download-File {
    param (
        [string]$Url,
        [string]$FileName,
        [string]$DestFolder
    )

    $OutputPath = Join-Path -Path $DestFolder -ChildPath $FileName

    try {
        Write-Host "[>] Téléchargement de $FileName..." -ForegroundColor White
        Invoke-RestMethod -Uri $Url -OutFile $OutputPath -ErrorAction Stop
        Write-Host "[+] $FileName téléchargé avec succès" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "[!] Erreur lors du téléchargement de $FileName : $_" -ForegroundColor Red
        return $false
    }
}

# Téléchargement des scripts
Write-Host ""
Write-Host "Téléchargement des scripts..." -ForegroundColor Cyan
Write-Host ""

$Results = @()

# ADFlow - script principal et outils
$Results += Download-File -Url $ADFlow_URL -FileName $ADFlow_Name -DestFolder $ADFlowPath
$Results += Download-File -Url $ADFlow_7za_URL -FileName $ADFlow_7za_Name -DestFolder $ADFlowToolsPath

# Résumé
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   Résumé de l'installation" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

$SuccessCount = ($Results | Where-Object { $_ -eq $true }).Count
$TotalCount = $Results.Count

if ($SuccessCount -eq $TotalCount) {
    Write-Host "[+] Tous les scripts ont été installés avec succès!" -ForegroundColor Green
} else {
    Write-Host "[!] $SuccessCount/$TotalCount scripts installés" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Emplacement: $DestinationPath" -ForegroundColor Cyan
Write-Host ""

# Ouvrir le dossier dans l'explorateur
explorer.exe $DestinationPath