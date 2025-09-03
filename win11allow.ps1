# Backup your registry before running this script!
# This script attempts to allow Windows 11 installation on unsupported hardware.

# Define the registry paths and values
$MoSetupPath = "HKLM:\SYSTEM\Setup\MoSetup"
$LabConfigPath = "HKLM:\SYSTEM\Setup\LabConfig"
$MoSetupValueName = "AllowUpgradesWithUnsupportedTPMOrCPU"
$MoSetupValue = "1"

# Check if the MoSetup key exists, if not, create it
if (-not (Test-Path $MoSetupPath)) {
    New-Item -Path $MoSetupPath -Force
}

# Set the value for allowing upgrades with unsupported TPM or CPU
New-ItemProperty -Path $MoSetupPath -Name $MoSetupValueName -Value $MoSetupValue -PropertyType DWORD -Force

# Create the LabConfig key and set additional values for further compatibility
if (-not (Test-Path $LabConfigPath)) {
    New-Item -Path $LabConfigPath -Force
}

# Example: Set three additional DWORD values under LabConfig for further compatibility
New-ItemProperty -Path $LabConfigPath -Name "BypassTPMCheck" -Value "1" -PropertyType DWORD -Force
New-ItemProperty -Path $LabConfigPath -Name "BypassSecureBootCheck" -Value "1" -PropertyType DWORD -Force
New-ItemProperty -Path $LabConfigPath -Name "BypassRAMCheck" -Value "1" -PropertyType DWORD -Force
