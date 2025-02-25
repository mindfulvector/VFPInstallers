# Visual FoxPro 9.0 SP2 Hotfix (KB968409) Installation Script
# Assumes script is in the same directory as VFP90SP2-KB968409-ENU.exe

# Stop if an error occurs
$ErrorActionPreference = "Stop"

# Ensure script is run as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "This script must be run as Administrator!"
}

# Request VFP 9 install path from the user
$vfpInstallPath = Read-Host -Prompt "Enter the Visual FoxPro 9.0 installation path (default: ${env:ProgramFiles}\Microsoft Visual FoxPro 9)"
if (-not $vfpInstallPath) {
    $vfpInstallPath = "${env:ProgramFiles(x86)}\Microsoft Visual FoxPro 9"
}

# Define other paths
$extractPath = Join-Path $PSScriptRoot "SP2HF3"
$mergeModulesPath = "${env:ProgramFiles(x86)}\Common Files\Merge Modules"
$runtimesPath = "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\VFP"
$logFile = Join-Path $PSScriptRoot "installation.log"

# Create extraction directory if it doesn't exist
if (-not (Test-Path $extractPath)) {
    New-Item -ItemType Directory -Path $extractPath | Out-Null
}

# Function to log messages
function Log-Message {
    param (
        [string]$Message
    )
    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timeStamp - $Message" | Out-File -FilePath $logFile -Append
    Write-Host $Message
}

# Function to backup and copy files
function Update-File {
    param (
        [string]$SourceFile,
        [string]$DestinationPath,
        [string]$FileName
    )
    
    $destFile = Join-Path $DestinationPath $FileName
    $backupFile = "$destFile.old"
    
    Log-Message "Processing $FileName..."
    
    # Check if destination exists and create backup
    if (Test-Path $destFile) {
        Log-Message "Creating backup: $backupFile"
        try {
            if (Test-Path $backupFile) {
                Remove-Item $backupFile -Force
                Log-Message "Previous backup $backupFile removed."
            }
            Rename-Item -Path $destFile -NewName "$backupFile"
            Log-Message "Backup created for $FileName."
        }
        catch {
            Log-Message "Warning: Unable to remove or create backup for $FileName - $_"
        }
    }
    
    # Copy new file
    try {
        Log-Message "Copying new $FileName"
        Copy-Item -Path $SourceFile -Destination $destFile -Force
        Log-Message "$FileName successfully updated."
    }
    catch {
        Log-Message "Error: Failed to copy $FileName - $_"
        throw "Failed to update $FileName"
    }
}

try {
    # Step 1: Extract files
    Log-Message "Extracting files from VFP90SP2-KB968409-ENU.exe..."
    $extractorPath = Join-Path $PSScriptRoot "VFP90SP2-KB968409-ENU.exe"
    if (-not (Test-Path $extractorPath)) {
        throw "VFP90SP2-KB968409-ENU.exe not found in script directory!"
    }
    Start-Process -FilePath $extractorPath -ArgumentList "/t:`"$extractPath`" /c" -Wait
    Log-Message "Extraction completed."

    # Verify extracted files
    $requiredFiles = @("Vfp9.exe", "Vfp9r.dll", "Vfp9runtime.msm", "Vfp9t.dll")
    foreach ($file in $requiredFiles) {
        if (-not (Test-Path (Join-Path $extractPath $file))) {
            throw "Required file $file not found in extraction directory!"
        }
    }
    Log-Message "All required files verified."

    # Step 2: Update VFP9.exe
    Update-File -SourceFile (Join-Path $extractPath "Vfp9.exe") `
                -DestinationPath $vfpInstallPath `
                -FileName "Vfp9.exe"

    # Step 3: Update VFP9runtime.msm
    Update-File -SourceFile (Join-Path $extractPath "Vfp9runtime.msm") `
                -DestinationPath $mergeModulesPath `
                -FileName "Vfp9runtime.msm"

    # Step 4: Update runtime DLLs
    Update-File -SourceFile (Join-Path $extractPath "Vfp9r.dll") `
                -DestinationPath $runtimesPath `
                -FileName "Vfp9r.dll"

    Update-File -SourceFile (Join-Path $extractPath "Vfp9t.dll") `
                -DestinationPath $runtimesPath `
                -FileName "Vfp9t.dll"

    Log-Message "Installation completed successfully!"
    Write-Host "Installation completed successfully!" -ForegroundColor Green
}
catch {
    Log-Message "An error occurred during installation: $_"
    Write-Host "An error occurred during installation: $_" -ForegroundColor Red
    Write-Host "Installation failed!" -ForegroundColor Red
    exit 1
}
