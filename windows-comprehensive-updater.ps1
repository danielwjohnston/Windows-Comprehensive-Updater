# ============================================================================
# FULLY AUTOMATED WINDOWS COMPREHENSIVE UPDATER WITH SELF-DEPLOYMENT
# Script Version: 2.1.0 - Enhanced Bulletproof Edition
# ============================================================================

# 
# PURPOSE: This script provides a completely hands-off Windows system maintenance experience
# that automatically handles privilege escalation, downloads required tools,
# installs all available updates, performs silent reboots, and deploys itself
# locally for scheduled execution.
#
# SELF-DEPLOYMENT: Script copies itself to C:\Scripts and creates scheduled
# tasks for automated Patch Tuesday execution with retry logic.
#
# DESIGN PHILOSOPHY: Zero user interaction required - the script handles every
# aspect of the update process automatically, including error recovery and
# system continuation after reboots.
#
# ABOUT WIZMO TOOL:
# Wizmo is Steve Gibson's multipurpose Windows utility from GRC.com that provides various
# system control functions via command line. Key features relevant to this script:
# - Silent system reboots without user prompts or dialogs
# - Monitor power control and screen blanking
# - System power management (standby, hibernate, shutdown)
# - Audio control and CD-ROM tray management
# - Created because Windows built-in power management lacked simple command-line control
# - Extremely lightweight (single executable, no installation required)
# - Supports "silent" operations perfect for automated scripts
# The 'reboot' command performs a graceful system restart, and when combined with
# other Wizmo functions, provides completely automated system control capabilities.
# ============================================================================

# DASHBOARD INTEGRATION: Launch monitoring dashboard if requested
# WHY DASHBOARD: Provides real-time visual feedback for administrators
# DESIGN CHOICE: Optional dashboard launch for enhanced monitoring experience
param(
    [switch]$ShowDashboard,        # Launch the HTML dashboard for real-time monitoring
    [string]$DashboardPath = "",   # Custom path to dashboard HTML file
    [switch]$Deploy,               # Force deployment to C:\Scripts even if already local
    [switch]$CreateSchedule,       # Create Patch Tuesday scheduled tasks
    [int]$RetryDays = 7,          # Days to retry if no cumulative updates found
    [switch]$SkipUpdateCheck,      # Skip self-update check (prevents infinite loops)
    [switch]$CheckCumulative,      # Only check for cumulative updates, don't run full process
    [switch]$SkipWin11Upgrade      # Skip Windows 11 feature upgrade pre-check
)

# EVENT ID CONSTANTS: Centralized event ID mapping for consistent logging
# WHY CENTRALIZED: Ensures consistent event IDs across all logging calls
# EVENT ID SCHEMA: 1xxx = Info, 2xxx = Process, 3xxx = Warning, 5xxx = Error
$script:EventIds = @{
    # Information Events (1000-1999)
    SCRIPT_START = 1001
    SCRIPT_SUCCESS = 1002
    WINGET_APP_UPDATE = 1015
    WINGET_COMPLETE = 1016
    WINDOWS_UPDATE_COMPLETE = 1017
    DASHBOARD_LAUNCH = 1020
    
    # Process Events (2000-2999)
    # (Reserved for future use)
    
    # Warning Events (3000-3999)
    SCRIPT_WARNING = 3002
    WINGET_APP_FAILED = 3019
    
    # Error Events (5000-5999)
    LOGGING_ERROR = 5000
    SCRIPT_CRITICAL = 5001
    WINGET_FAILED = 5016
    WINDOWS_UPDATE_FAILED = 5017
    WINDOWS_PROCESS_FAILED = 5018
    DASHBOARD_FAILED = 5020
    WINGET_APP_ERROR = 5022
    WINGET_TIMEOUT = 5021
}

# CONFIGURATION SCHEMA VALIDATION: Validate config file structure and values
# WHY NEEDED: Ensures configuration file is properly formatted and contains valid values
function Test-ConfigurationSchema {
    param(
        [string]$ConfigPath
    )
    
    Write-LogMessage "Validating configuration schema..." "INFO"
    
    try {
        # LOAD CONFIGURATION: Parse JSON configuration
        $config = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        
        # REQUIRED SECTIONS: Check for mandatory configuration sections
        $requiredSections = @("version", "settings")
        foreach ($section in $requiredSections) {
            if (-not $config.PSObject.Properties.Name.Contains($section)) {
                Write-LogMessage "Configuration validation failed: Missing required section '$section'" "ERROR"
                return $false
            }
        }
        
        # VERSION VALIDATION: Ensure version is valid
        if (-not $config.version -or $config.version -notmatch '^\d+\.\d+\.\d+$') {
            Write-LogMessage "Configuration validation failed: Invalid version format '$($config.version)'" "ERROR"
            return $false
        }
        
        # SETTINGS VALIDATION: Validate settings structure
        if (-not $config.settings) {
            Write-LogMessage "Configuration validation failed: Missing settings section" "ERROR"
            return $false
        }
        
        # GENERAL SETTINGS: Validate general configuration
        $generalSettings = $config.settings.general
        if ($generalSettings) {
            # LOG FILE SIZE: Must be positive number
            if ($generalSettings.maxLogFileSizeMB -and ($generalSettings.maxLogFileSizeMB -le 0 -or $generalSettings.maxLogFileSizeMB -gt 1000)) {
                Write-LogMessage "Configuration validation failed: maxLogFileSizeMB must be between 1-1000 MB" "ERROR"
                return $false
            }
            
            # LOG RETENTION: Must be positive number
            if ($generalSettings.logRetentionDays -and ($generalSettings.logRetentionDays -le 0 -or $generalSettings.logRetentionDays -gt 365)) {
                Write-LogMessage "Configuration validation failed: logRetentionDays must be between 1-365 days" "ERROR"
                return $false
            }
        }
        
        # WINGET SETTINGS: Validate Winget configuration
        $wingetSettings = $config.settings.winget
        if ($wingetSettings) {
            # TIMEOUT: Must be reasonable value
            if ($wingetSettings.maxUpdateTimeoutMinutes -and ($wingetSettings.maxUpdateTimeoutMinutes -le 0 -or $wingetSettings.maxUpdateTimeoutMinutes -gt 300)) {
                Write-LogMessage "Configuration validation failed: maxUpdateTimeoutMinutes must be between 1-300 minutes" "ERROR"
                return $false
            }
            
            # RETRY ATTEMPTS: Must be reasonable value
            if ($wingetSettings.maxRetryAttempts -and ($wingetSettings.maxRetryAttempts -lt 0 -or $wingetSettings.maxRetryAttempts -gt 10)) {
                Write-LogMessage "Configuration validation failed: maxRetryAttempts must be between 0-10" "ERROR"
                return $false
            }
        }
        
        # WINDOWS UPDATE SETTINGS: Validate Windows Update configuration
        $windowsSettings = $config.settings.windowsUpdate
        if ($windowsSettings) {
            # UPDATE CYCLES: Must be reasonable value
            if ($windowsSettings.maxUpdateCycles -and ($windowsSettings.maxUpdateCycles -le 0 -or $windowsSettings.maxUpdateCycles -gt 10)) {
                Write-LogMessage "Configuration validation failed: maxUpdateCycles must be between 1-10" "ERROR"
                return $false
            }
        }
        
        # SCHEDULING SETTINGS: Validate scheduling configuration
        $scheduleSettings = $config.settings.scheduling
        if ($scheduleSettings) {
            # HOURS: Must be 0-23
            if ($scheduleSettings.patchTuesdayHour -and ($scheduleSettings.patchTuesdayHour -lt 0 -or $scheduleSettings.patchTuesdayHour -gt 23)) {
                Write-LogMessage "Configuration validation failed: patchTuesdayHour must be between 0-23" "ERROR"
                return $false
            }
            
            # MINUTES: Must be 0-59
            if ($scheduleSettings.patchTuesdayMinute -and ($scheduleSettings.patchTuesdayMinute -lt 0 -or $scheduleSettings.patchTuesdayMinute -gt 59)) {
                Write-LogMessage "Configuration validation failed: patchTuesdayMinute must be between 0-59" "ERROR"
                return $false
            }
            
            # RETRY DAYS: Must be reasonable value
            if ($scheduleSettings.retryDays -and ($scheduleSettings.retryDays -le 0 -or $scheduleSettings.retryDays -gt 30)) {
                Write-LogMessage "Configuration validation failed: retryDays must be between 1-30 days" "ERROR"
                return $false
            }
        }
        
        # DASHBOARD SETTINGS: Validate dashboard configuration
        $dashboardSettings = $config.settings.dashboard
        if ($dashboardSettings) {
            # PORT: Must be valid port number
            if ($dashboardSettings.dashboardPort -and ($dashboardSettings.dashboardPort -le 0 -or $dashboardSettings.dashboardPort -gt 65535)) {
                Write-LogMessage "Configuration validation failed: dashboardPort must be between 1-65535" "ERROR"
                return $false
            }
            
            # REFRESH INTERVAL: Must be reasonable value
            if ($dashboardSettings.dashboardRefreshSeconds -and ($dashboardSettings.dashboardRefreshSeconds -le 0 -or $dashboardSettings.dashboardRefreshSeconds -gt 300)) {
                Write-LogMessage "Configuration validation failed: dashboardRefreshSeconds must be between 1-300 seconds" "ERROR"
                return $false
            }
        }
        
        Write-LogMessage "Configuration schema validation passed" "SUCCESS"
        return $true
        
    } catch {
        Write-LogMessage "Configuration schema validation failed: $_" "ERROR"
        return $false
    }
}

# CONFIGURATION LOADING: Attempt to load optional JSON config for overrides
$global:Config = $null
$configPathCandidates = @(
    Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'windows-update-config.json',
    'C:\Scripts\windows-update-config.json'
)
foreach ($cfg in $configPathCandidates) {
    if (-not $global:Config -and (Test-Path $cfg)) {
        try {
            $raw = Get-Content $cfg -Raw -ErrorAction Stop
            $global:Config = $raw | ConvertFrom-Json -ErrorAction Stop
            
            # VALIDATE CONFIGURATION: Check schema and values
            if (Test-ConfigurationSchema -ConfigPath $cfg) {
                Write-Host "Loaded and validated configuration: $cfg" -ForegroundColor Cyan
            } else {
                Write-Host "Configuration validation failed for $cfg - using defaults" -ForegroundColor Yellow
                $global:Config = $null
            }
        } catch {
            Write-Host "Failed to load configuration from $cfg : $_" -ForegroundColor Yellow
            Write-Host "Using default configuration values" -ForegroundColor Gray
        }
    }
}

# CONFIG OVERRIDES: Apply select settings if config present
if ($global:Config) {
    if ($global:Config.settings.winget.enableWingetUpdates -eq $false) { $script:DisableWinget = $true }
    if ($global:Config.settings.maintenance.enableSystemMaintenance -eq $false) { $script:DisableMaintenance = $true }
    if ($global:Config.settings.windowsUpdate.enableWindowsUpdates -eq $false) { $script:DisableWindowsUpdates = $true }
    if ($global:Config.settings.advanced.skipWin11UpgradeBypass -eq $true) { $SkipWin11Upgrade = $true }
    if ($global:Config.settings.general.enableDetailedLogging -eq $false) { $script:MinimalLogging = $true }
    if ($global:Config.settings.general.maxLogFileSizeMB) { $script:MaxLogMB = [int]$global:Config.settings.general.maxLogFileSizeMB }
    if ($global:Config.settings.general.logRetentionDays) { $script:LogRetentionDays = [int]$global:Config.settings.general.logRetentionDays }
}
if (-not $script:MaxLogMB) { $script:MaxLogMB = 50 }
if (-not $script:LogRetentionDays) { $script:LogRetentionDays = 30 }

# SELF-UPDATE DETECTION: Check if script needs updating before proceeding
# WHY NEEDED: Ensures we're running the latest version with all fixes
function Test-ScriptNeedsUpdate {
    param(
        [string]$CurrentScriptPath
    )
    
    Write-LogMessage "Checking if script needs updating..." "INFO"
    
    try {
        # SCRIPT VERSION DETECTION: Extract version from current script
        $currentContent = Get-Content -Path $CurrentScriptPath -Raw -ErrorAction SilentlyContinue
        $currentVersionMatch = $currentContent | Select-String 'Script Version: ([\d\.]+)'
        $currentVersion = if ($currentVersionMatch) { $currentVersionMatch.Matches[0].Groups[1].Value } else { "1.0" }
        
        # DEPLOYED VERSION CHECK: Check version of deployed script if it exists
        $deployedScriptPath = "C:\Scripts\windows-comprehensive-updater.ps1"
        if (Test-Path $deployedScriptPath) {
            $deployedContent = Get-Content -Path $deployedScriptPath -Raw -ErrorAction SilentlyContinue
            $deployedVersionMatch = $deployedContent | Select-String 'Script Version: ([\d\.]+)'
            $deployedVersion = if ($deployedVersionMatch) { $deployedVersionMatch.Matches[0].Groups[1].Value } else { "1.0" }
            
            # VERSION COMPARISON: Check if current version is newer
            if ([version]$currentVersion -gt [version]$deployedVersion) {
                Write-LogMessage "Script update available: v$deployedVersion → v$currentVersion" "INFO"
                return $true
            } else {
                Write-LogMessage "Script is current version: v$currentVersion" "INFO"
                return $false
            }
        }
        
        # NEW DEPLOYMENT: Always deploy if no local version exists
        Write-LogMessage "No deployed version found - initial deployment needed" "INFO"
        return $true
        
    } catch {
        Write-LogMessage "Error checking script version: $_" "WARNING"
        return $true  # Deploy anyway if version check fails
    }
}

# PROCESS-AWARE WINGET UPDATES: Handle updates that conflict with running processes
# WHY NEEDED: Can't update PowerShell while PowerShell script is running
# SOLUTION: Use external CMD process to perform conflicting updates
function Update-ConflictingApplications {
    param(
        [array]$ConflictingApps
    )
    
    Write-LogMessage "Handling process-conflicting application updates..." "INFO"
    
    foreach ($app in $ConflictingApps) {
        Write-LogMessage "Processing conflicting app: $($app.DisplayName)" "INFO"
        
        try {
            # CONFLICT DETECTION: Check if this app conflicts with current process
            $hasConflict = $false
            switch ($app.Name) {
                "Microsoft.PowerShell" {
                    # PowerShell conflict: Can't update while PowerShell script is running
                    $hasConflict = $true
                    Write-LogMessage "PowerShell update requires external process execution" "WARNING"
                }
                "Microsoft.WindowsTerminal" {
                    # Terminal conflict: Check if Windows Terminal is running
                    $terminalProcess = Get-Process -Name "WindowsTerminal" -ErrorAction SilentlyContinue
                    if ($terminalProcess) {
                        $hasConflict = $true
                        Write-LogMessage "Windows Terminal is running - requires external update" "WARNING"
                    }
                }
            }
            
            if ($hasConflict) {
                # EXTERNAL UPDATE EXECUTION: Use CMD to run winget from outside PowerShell
                Write-LogMessage "Executing $($app.DisplayName) update via external process..." "INFO"
                
                # BATCH FILE CREATION: Create temporary batch file for external execution
                $tempBatchFile = Join-Path $env:TEMP "winget-update-$($app.Name -replace '\.', '-').bat"
                $batchContent = @"
@echo off
echo Updating $($app.DisplayName) via external process...
winget upgrade $($app.Name) --silent --accept-package-agreements --accept-source-agreements
if %ERRORLEVEL% EQU 0 (
    echo $($app.DisplayName) updated successfully
    exit /b 0
) else if %ERRORLEVEL% EQU -1978335212 (
    echo $($app.DisplayName) is already up to date  
    exit /b 0
) else (
    echo $($app.DisplayName) update failed with exit code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
"@
                
                # BATCH FILE EXECUTION: Write and execute batch file
                $batchContent | Out-File -FilePath $tempBatchFile -Encoding ASCII
                
                Write-LogMessage "Created external update batch: $tempBatchFile" "INFO"
                
                # EXTERNAL PROCESS LAUNCH: Run CMD with batch file
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName = "cmd.exe"
                $processInfo.Arguments = "/c `"$tempBatchFile`""
                $processInfo.UseShellExecute = $false
                $processInfo.RedirectStandardOutput = $true
                $processInfo.RedirectStandardError = $true
                $processInfo.CreateNoWindow = $false  # Show window for debugging
                
                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $processInfo
                
                Write-LogMessage "Starting external update process for $($app.DisplayName)..." "INFO"
                $process.Start() | Out-Null
                
                # PROCESS MONITORING: Wait for completion with timeout
                $timeoutSeconds = 300  # 5 minute timeout
                $completed = $process.WaitForExit($timeoutSeconds * 1000)
                
                if ($completed) {
                    $exitCode = $process.ExitCode
                    $stdout = $process.StandardOutput.ReadToEnd()
                    $stderr = $process.StandardError.ReadToEnd()
                    
                    Write-LogMessage "External update completed with exit code: $exitCode" "INFO"
                    if ($stdout) { Write-LogMessage "Output: $stdout" "INFO" }
                    if ($stderr) { Write-LogMessage "Error: $stderr" "WARNING" }
                    
                    # SUCCESS EVALUATION: Determine if update was successful
                    if ($exitCode -eq 0 -or $exitCode -eq -1978335212) {
                        Write-LogMessage "$($app.DisplayName) updated successfully via external process" "SUCCESS"
                        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId $script:EventIds.WINGET_APP_UPDATE -EntryType Information -Message "$($app.DisplayName) updated via external process"
                    } else {
                        Write-LogMessage "$($app.DisplayName) external update failed with exit code: $exitCode" "ERROR"
                        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3019 -EntryType Warning -Message "$($app.DisplayName) external update failed: $exitCode"
                    }
                } else {
                    # TIMEOUT HANDLING: Kill process if it takes too long
                    Write-LogMessage "$($app.DisplayName) update timed out after $timeoutSeconds seconds" "ERROR"
                    try { $process.Kill() } catch { }
                    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId $script:EventIds.WINGET_TIMEOUT -EntryType Error -Message "$($app.DisplayName) update timed out"
                }
                
                # CLEANUP: Remove temporary batch file
                try { Remove-Item -Path $tempBatchFile -Force -ErrorAction SilentlyContinue } catch { }
                
            } else {
                # NO CONFLICT: Update normally via PowerShell
                Write-LogMessage "No process conflict detected for $($app.DisplayName) - updating normally" "INFO"
                $null = & winget upgrade $app.Name --silent --accept-package-agreements --accept-source-agreements 2>&1
                
                # NORMAL UPDATE RESULT HANDLING
                if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335212) {
                    Write-LogMessage "$($app.DisplayName) updated successfully" "SUCCESS"
                } else {
                    Write-LogMessage "$($app.DisplayName) update failed with exit code: $LASTEXITCODE" "WARNING"
                }
            }
            
        } catch {
            Write-LogMessage "Error updating $($app.DisplayName): $_" "ERROR"
            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId $script:EventIds.WINGET_APP_ERROR -EntryType Error -Message "Error updating $($app.DisplayName): $_"
        }
        
        # INTER-APP DELAY: Brief pause between app updates
        Start-Sleep -Seconds 2
    }
}

# SCRIPT UPDATE AND RESTART: Handle self-updating scenario
# WHY NEEDED: If script is updated, need to restart with new version
function Invoke-ScriptUpdateRestart {
    param(
        [string]$UpdatedScriptPath
    )
    
    Write-LogMessage "Script has been updated - restarting with new version..." "INFO"
    
    try {
        # ARGUMENT PRESERVATION: Maintain all original parameters
        $originalArgs = @()
        if ($ShowDashboard) { $originalArgs += "-ShowDashboard" }
        if ($DashboardPath) { $originalArgs += "-DashboardPath `"$DashboardPath`"" }
        if ($CreateSchedule) { $originalArgs += "-CreateSchedule" }
        if ($RetryDays -ne 7) { $originalArgs += "-RetryDays $RetryDays" }
        
        # RESTART MARKER: Add flag to prevent infinite update loops
        $originalArgs += "-SkipUpdateCheck"
        
        $argumentString = $originalArgs -join " "
        $launchArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$UpdatedScriptPath`" $argumentString"
        
        Write-LogMessage "Restarting with updated script: $UpdatedScriptPath" "INFO"
        Write-LogMessage "Arguments: $argumentString" "INFO"
        
        # ELEVATED RESTART: Launch updated script with admin privileges
        Start-Process powershell.exe -Verb RunAs -ArgumentList $launchArgs -WindowStyle Hidden
        
        Write-LogMessage "Updated script launched - current instance exiting" "INFO"
        Exit 0
        
    } catch {
        Write-LogMessage "Failed to restart with updated script: $_" "ERROR"
        Write-LogMessage "Continuing with current version..." "WARNING"
    }
}

function Invoke-SelfDeployment {
    param(
        [string]$SourcePath,
        [bool]$ForceDeployment = $false
    )
    
    # DEPLOYMENT TARGET: Standard location for system scripts
    $targetDir = "C:\Scripts"
    $targetScript = Join-Path $targetDir "windows-comprehensive-updater.ps1"
    $targetDashboard = Join-Path $targetDir "windows-comprehensive-updater-dashboard.html"
    
    # LOCAL EXECUTION CHECK: Determine if we're already running from target location
    $currentPath = $MyInvocation.MyCommand.Path
    $isLocalExecution = $currentPath -and $currentPath.StartsWith($targetDir)
    
    # DEPLOYMENT DECISION: Deploy if not local or forced
    if (-not $isLocalExecution -or $ForceDeployment) {
        Write-Host "Self-deployment required - setting up local system..." -ForegroundColor Cyan
        
        try {
            # CREATE TARGET DIRECTORY: Ensure C:\Scripts exists
            if (-not (Test-Path $targetDir)) {
                Write-Host "Creating directory: $targetDir" -ForegroundColor Green
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            }
            
            # SCRIPT DEPLOYMENT: Copy current script to target location
            if ($currentPath -and (Test-Path $currentPath)) {
                Write-Host "Deploying script to: $targetScript" -ForegroundColor Green
                Copy-Item -Path $currentPath -Destination $targetScript -Force
            } else {
                Write-Host "WARNING: Could not determine current script path for deployment" -ForegroundColor Yellow
            }
            
            # DASHBOARD DEPLOYMENT: Look for dashboard file and deploy it
            $dashboardSources = @(
                # Same directory as current script
                (Join-Path (Split-Path $currentPath -Parent) "windows-comprehensive-updater-dashboard.html"),
                # Custom dashboard path if specified
                $DashboardPath,
                # Current directory
                ".\windows-comprehensive-updater-dashboard.html",
                # Look in common locations
                "C:\Temp\windows-comprehensive-updater-dashboard.html",
                "$env:USERPROFILE\Downloads\windows-comprehensive-updater-dashboard.html"
            ) | Where-Object { $_ -and (Test-Path $_) }
            
            if ($dashboardSources.Count -gt 0) {
                $sourceDashboard = $dashboardSources[0]
                Write-Host "Deploying dashboard from: $sourceDashboard" -ForegroundColor Green
                Copy-Item -Path $sourceDashboard -Destination $targetDashboard -Force
            } else {
                # CREATE EMBEDDED DASHBOARD: Generate dashboard if not found
                Write-Host "Dashboard not found - creating embedded version..." -ForegroundColor Yellow
                New-EmbeddedDashboard -TargetPath $targetDashboard
            }
            
            # RELAUNCH FROM LOCAL: Start new instance from deployed location
            Write-Host "Relaunching from local deployment..." -ForegroundColor Cyan
            
            # ARGUMENT RECONSTRUCTION: Pass through all original parameters
            $newArgs = @()
            if ($ShowDashboard) { $newArgs += "-ShowDashboard" }
            if ($DashboardPath) { $newArgs += "-DashboardPath `"$DashboardPath`"" }
            if ($CreateSchedule) { $newArgs += "-CreateSchedule" }
            if ($RetryDays -ne 7) { $newArgs += "-RetryDays $RetryDays" }
            
            $argumentString = $newArgs -join " "
            $launchArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$targetScript`" $argumentString"
            
            # ELEVATED RELAUNCH: Start from local location with admin privileges
            Start-Process powershell.exe -Verb RunAs -ArgumentList $launchArgs -WindowStyle Hidden
            
            Write-Host "Deployment complete - local instance starting..." -ForegroundColor Green
            Exit 0
            
        } catch {
            Write-Host "Deployment failed: $_" -ForegroundColor Red
            Write-Host "Continuing from current location..." -ForegroundColor Yellow
            return $false
        }
    }
    
    Write-Host "Running from local deployment: $targetDir" -ForegroundColor Green
    return $true
}

# PATCH TUESDAY SCHEDULER: Create automated monthly update tasks
# WHY NEEDED: Ensures regular patching with retry logic for delayed Microsoft releases
function Get-SecondTuesday([datetime]$reference) {
    # Returns the actual second Tuesday date for the month of the reference date
    $firstOfMonth = Get-Date -Year $reference.Year -Month $reference.Month -Day 1 -Hour 0 -Minute 0 -Second 0
    # Find first Tuesday
    $offset = ([int][System.DayOfWeek]::Tuesday - [int]$firstOfMonth.DayOfWeek)
    if ($offset -lt 0) { $offset += 7 }
    $firstTuesday = $firstOfMonth.AddDays($offset)
    return $firstTuesday.AddDays(7) # second Tuesday
}

function New-PatchTuesdaySchedule {
    param(
        [int]$RetryDays = 7,
        [int]$Hour = 2,
        [int]$Minute = 0
    )
    Write-Host "Creating Patch Tuesday automated schedule (precise date logic)..." -ForegroundColor Cyan
    try {
        $scriptPath = Join-Path "C:\Scripts" "windows-comprehensive-updater.ps1"
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 30)
        $actionMain = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ShowDashboard"

        # Remove existing tasks first to avoid duplicates
        'WindowsUpdate-PatchTuesday','WindowsUpdate-Manual' | ForEach-Object { try { Unregister-ScheduledTask -TaskName $_ -Confirm:$false -ErrorAction SilentlyContinue } catch { } }
        Get-ScheduledTask | Where-Object { $_.TaskName -like 'WindowsUpdate-PatchTuesday-Retry*' } | ForEach-Object { try { Unregister-ScheduledTask -TaskName $_.TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { } }

        # Create main task using a OneTime trigger + monthly self-recreation script (lightweight approach)
        $secondTuesday = Get-SecondTuesday (Get-Date)
        if ($secondTuesday -lt (Get-Date)) { $secondTuesday = Get-SecondTuesday ((Get-Date).AddMonths(1)) }
        $runTime = Get-Date -Year $secondTuesday.Year -Month $secondTuesday.Month -Day $secondTuesday.Day -Hour $Hour -Minute $Minute -Second 0
        $mainTrigger = New-ScheduledTaskTrigger -Once -At $runTime
        Register-ScheduledTask -TaskName "WindowsUpdate-PatchTuesday" -Action $actionMain -Trigger $mainTrigger -Principal $principal -Settings $settings -Force | Out-Null
        Write-Host " Main Patch Tuesday task scheduled: $runTime" -ForegroundColor Green

        # Retry tasks: schedule daily after main date for RetryDays using one-time triggers
        for ($d=1; $d -le $RetryDays; $d++) {
            $retryDate = $runTime.AddDays($d)
            $retryAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -CheckCumulative"
            $retryTrigger = New-ScheduledTaskTrigger -Once -At $retryDate
            Register-ScheduledTask -TaskName "WindowsUpdate-PatchTuesday-Retry$d" -Action $retryAction -Trigger $retryTrigger -Principal $principal -Settings $settings -Force | Out-Null
        }
        Write-Host " Retry tasks scheduled for $RetryDays day(s) after Patch Tuesday" -ForegroundColor Green

        # Manual task
        $manualAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ShowDashboard"
        Register-ScheduledTask -TaskName "WindowsUpdate-Manual" -Action $manualAction -Principal $principal -Settings $settings -Force | Out-Null
        Write-Host " Manual run task created" -ForegroundColor Green

# (Future enhancement placeholder for auto-reschedule)

        Write-Host "Patch Tuesday automation configured successfully (one-time triggers)." -ForegroundColor Cyan
        return $true
    } catch {
        Write-Host "Failed to create scheduled tasks: $_" -ForegroundColor Red
        return $false
    }
}

# MONTHLY AUTO-RESCHEDULE: Automatically recreate tasks for next month's Patch Tuesday
# WHY NEEDED: One-time triggers expire after execution, need to recreate for next month
function Update-MonthlySchedule {
    Write-Host "Checking if monthly schedule needs updating..." -ForegroundColor Cyan
    
    try {
        # GET EXISTING TASKS: Check current Patch Tuesday tasks
        $existingTasks = Get-ScheduledTask | Where-Object { $_.TaskName -like 'WindowsUpdate-PatchTuesday*' }
        
        if ($existingTasks.Count -eq 0) {
            Write-Host "No existing Patch Tuesday tasks found - creating new schedule" -ForegroundColor Yellow
            return (New-PatchTuesdaySchedule)
        }
        
        # FIND NEXT PATCH TUESDAY: Calculate next month's date
        $nextMonth = (Get-Date).AddMonths(1)
        $nextPatchTuesday = Get-SecondTuesday $nextMonth
        
        # CHECK IF UPDATE NEEDED: Compare with existing task dates
        $mainTask = $existingTasks | Where-Object { $_.TaskName -eq 'WindowsUpdate-PatchTuesday' }
        if ($mainTask) {
            $currentTrigger = $mainTask.Triggers[0]
            if ($currentTrigger -and $currentTrigger.StartBoundary) {
                $currentDate = [DateTime]::Parse($currentTrigger.StartBoundary)
                
                # UPDATE IF MORE THAN 2 WEEKS OLD: Prevents unnecessary recreation
                if (($nextPatchTuesday - $currentDate).TotalDays -gt 14) {
                    Write-Host "Existing schedule is outdated - recreating for next month" -ForegroundColor Yellow
                    
                    # REMOVE OLD TASKS: Clean up existing tasks
                    $existingTasks | ForEach-Object {
                        try {
                            Unregister-ScheduledTask -TaskName $_.TaskName -Confirm:$false -ErrorAction SilentlyContinue
                        } catch {
                            Write-Host "Warning: Could not remove task $($_.TaskName): $_" -ForegroundColor Yellow
                        }
                    }
                    
                    # CREATE NEW SCHEDULE: For next month's Patch Tuesday
                    Write-Host "Creating new schedule for: $($nextPatchTuesday.ToString('yyyy-MM-dd'))" -ForegroundColor Green
                    return (New-PatchTuesdaySchedule)
                } else {
                    Write-Host "Schedule is current - no update needed" -ForegroundColor Green
                    return $true
                }
            }
        }
        
        # FALLBACK: Recreate if main task missing or invalid
        Write-Host "Main Patch Tuesday task missing or invalid - recreating schedule" -ForegroundColor Yellow
        $existingTasks | ForEach-Object {
            try {
                Unregister-ScheduledTask -TaskName $_.TaskName -Confirm:$false -ErrorAction SilentlyContinue
            } catch { }
        }
        return (New-PatchTuesdaySchedule)
        
    } catch {
        Write-Host "Failed to update monthly schedule: $_" -ForegroundColor Red
        return $false
    }
}

# CUMULATIVE UPDATE DETECTION: Check if monthly cumulative update is available
# WHY NEEDED: Microsoft sometimes delays cumulative updates past Patch Tuesday
function Test-CumulativeUpdateAvailable {
    Write-Host "Checking for monthly cumulative update..." -ForegroundColor Cyan
    
    try {
        # GET AVAILABLE UPDATES: Check for cumulative updates specifically
        $updates = Get-WUList -MicrosoftUpdate | Where-Object { 
            $_.Title -like "*Cumulative Update*" -and 
            ($_.Title -like "*Windows 10*" -or 
            $_.Title -like "*Windows 11*")
        }
        
        if ($updates -and $updates.Count -gt 0) {
            Write-Host " Found $($updates.Count) cumulative update(s) available" -ForegroundColor Green
            $updates | ForEach-Object { Write-Host "  • $($_.Title)" -ForegroundColor Gray }
            return $true
        } else {
            Write-Host " No cumulative updates found - Microsoft may not have released this month's update yet" -ForegroundColor Yellow
            return $false
        }
        
    } catch {
        Write-Host "Error checking for cumulative updates: $_" -ForegroundColor Red
        return $true  # Proceed anyway if check fails
    }
}

# EMBEDDED DASHBOARD CREATION: Create dashboard if not found during deployment
function New-EmbeddedDashboard {
    param([string]$TargetPath)
    
    # MINIMAL DASHBOARD: Basic monitoring interface when full dashboard unavailable
    $dashboardContent = @"
<!DOCTYPE html>
<html><head><title>Windows Comprehensive Updater Monitor</title><style>
body{background:#0b0b0b;color:#e6e6e6;font-family:system-ui;padding:20px;}
.card{background:#171717;border:1px solid #262626;border-radius:16px;padding:20px;margin:10px 0;}
.status{display:inline-block;padding:8px 16px;border-radius:20px;font-weight:bold;}
.running{background:#00ff0040;color:#00ff00;border:1px solid #00ff00;}
.log{background:#1e1e1e;border-radius:8px;padding:15px;height:300px;overflow-y:auto;font-family:monospace;}
</style></head><body>
<div class="card"><h1>Windows Comprehensive Updater Monitor</h1>
<div id="status" class="status running">Monitoring Active</div></div>
<div class="card"><h2>System Status</h2>
<div>Computer: <span id="computer">Loading...</span></div>
<div>Status: <span id="current-status">Waiting for updates...</span></div></div>
<div class="card"><h2>Live Log</h2><div id="log" class="log">Monitoring for script activity...</div></div>
<script>
setInterval(function(){
  fetch('update-status.json?t='+Date.now()).then(r=>r.json()).then(data=>{
    document.getElementById('computer').textContent=data.systemInfo?.computerName||'Unknown';
    document.getElementById('current-status').textContent=data.currentOperation||'Idle';
  }).catch(e=>console.log('No status file'));
  
  fetch('WindowsUpdateLog.txt?t='+Date.now()).then(r=>r.text()).then(data=>{
    const lines=data.split('\n').slice(-20).filter(l=>l.trim());
    document.getElementById('log').innerHTML=lines.map(line=>
      '<div>'+line.replace(/\[(.*?)\]/g,'<span style="color:#00ff00">[$1]</span>')+'</div>'
    ).join('');
  }).catch(e=>console.log('No log file'));
},3000);
</script></body></html>
"@
    
    try {
        $dashboardContent | Out-File -FilePath $TargetPath -Encoding UTF8
        Write-Host "Embedded dashboard created: $TargetPath" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Failed to create embedded dashboard: $_" -ForegroundColor Yellow
        return $false
    }
}

# FUNCTION: Ensure PSWindowsUpdate module is available and Microsoft Update service is added
function Initialize-PSWindowsUpdate {
    try {
        Write-LogMessage "Ensuring PSWindowsUpdate module is available..." "INFO"
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch { }
            Install-Module -Name PSWindowsUpdate -Force -Scope AllUsers -AllowClobber -ErrorAction Stop
            Write-LogMessage "PSWindowsUpdate module installed" "SUCCESS"
        }
        Import-Module PSWindowsUpdate -Force -ErrorAction Stop
        try { Add-WUServiceManager -MicrosoftUpdate -Confirm:$false | Out-Null } catch { }
        return $true
    } catch {
        Write-LogMessage "Failed to ensure PSWindowsUpdate: $_" "WARNING"
        return $false
    }
}

# FUNCTION: Apply registry changes to allow Windows 11 upgrade on unsupported hardware
function Set-Win11UpgradeBypass {
    try {
        $MoSetupPath = "HKLM:\SYSTEM\Setup\MoSetup"
        $LabConfigPath = "HKLM:\SYSTEM\Setup\LabConfig"
        if (-not (Test-Path $MoSetupPath)) { New-Item -Path $MoSetupPath -Force | Out-Null }
        New-ItemProperty -Path $MoSetupPath -Name "AllowUpgradesWithUnsupportedTPMOrCPU" -Value 1 -PropertyType DWORD -Force | Out-Null
        if (-not (Test-Path $LabConfigPath)) { New-Item -Path $LabConfigPath -Force | Out-Null }
        New-ItemProperty -Path $LabConfigPath -Name "BypassTPMCheck" -Value 1 -PropertyType DWORD -Force | Out-Null
        New-ItemProperty -Path $LabConfigPath -Name "BypassSecureBootCheck" -Value 1 -PropertyType DWORD -Force | Out-Null
        New-ItemProperty -Path $LabConfigPath -Name "BypassRAMCheck" -Value 1 -PropertyType DWORD -Force | Out-Null
        Write-LogMessage "Windows 11 upgrade bypass registry keys applied" "INFO"
    } catch {
        Write-LogMessage "Failed to set Windows 11 bypass keys: $_" "WARNING"
    }
}

# FUNCTION: Detect and install Windows 11 Feature Update when applicable
function Invoke-Windows11FeatureUpgradeIfNeeded {
    try {
        $cv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $productName = $cv.ProductName
        $build = [int]$cv.CurrentBuildNumber
        Write-LogMessage "OS detected: $productName (Build $build)" "INFO"
        
        $isWin10 = $productName -like '*Windows 10*'
        $isWin11 = $productName -like '*Windows 11*'
        
        # Ensure Microsoft Update service for Feature Updates
        try { Add-WUServiceManager -MicrosoftUpdate -Confirm:$false | Out-Null } catch { }
        
        # Detect available Windows 11 Feature Update
        $featureUpdates = Get-WUList -MicrosoftUpdate -ErrorAction SilentlyContinue | Where-Object {
            $_.Title -match 'Feature update to Windows 11'
        }
        
        if ($isWin10 -or ($isWin11 -and $featureUpdates)) {
            Write-LogMessage "Evaluating Windows 11 feature upgrade availability" "INFO"
            Set-Win11UpgradeBypass
            
            if ($featureUpdates -and $featureUpdates.Count -gt 0) {
                $titles = ($featureUpdates | Select-Object -ExpandProperty Title)
                Write-LogMessage ("Feature update(s) found: " + ($titles -join '; ')) "INFO"
                try {
                    if (Get-Command Install-WindowsUpdate -ErrorAction SilentlyContinue) {
                        $featureUpdates | Install-WindowsUpdate -AcceptAll -AutoReboot:$false -ErrorAction Stop | Out-Null
                    } elseif (Get-Command Get-WUInstall -ErrorAction SilentlyContinue) {
                        $featureUpdates | Get-WUInstall -AcceptAll -AutoReboot:$false -ErrorAction Stop | Out-Null
                    } else {
                        throw "Neither Install-WindowsUpdate nor Get-WUInstall is available"
                    }
                    Write-LogMessage "Feature update installation initiated; rebooting to continue..." "SUCCESS"
                    Invoke-SilentReboot
                } catch {
                    Write-LogMessage "Feature update installation failed: $_" "ERROR"
                }
            } else {
                # Encourage WU to offer Windows 11 by targeting the product
                try {
                    $wuPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
                    if (-not (Test-Path $wuPolicy)) { New-Item -Path $wuPolicy -Force | Out-Null }
                    New-ItemProperty -Path $wuPolicy -Name 'TargetReleaseVersion' -Value 1 -PropertyType DWORD -Force | Out-Null
                    New-ItemProperty -Path $wuPolicy -Name 'ProductVersion' -Value 'Windows 11' -PropertyType String -Force | Out-Null
                    Write-LogMessage "Configured Windows Update policy to target Windows 11" "INFO"
                } catch {
                    Write-LogMessage "Failed to configure target Windows 11 policy: $_" "WARNING"
                }
            }
        }
    } catch {
        Write-LogMessage "Windows 11 upgrade pre-check failed: $_" "WARNING"
    }
}

# =========================================================================
# DEPLOYMENT AND INITIALIZATION
# =========================================================================

# EARLY INITIALIZATION: logging, event source, elevation, module, OS upgrade
try {
    $logDir = "C:\Scripts"
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $global:logFile = Join-Path $logDir "WindowsUpdateLog.txt"
} catch { $global:logFile = "C:\Scripts\WindowsUpdateLog.txt" }

try {
    if (-not [System.Diagnostics.EventLog]::SourceExists("WindowsUpdateScript")) {
        New-EventLog -LogName Application -Source "WindowsUpdateScript"
    }
} catch { }

# Ensure elevation (required for registry/module operations)
Confirm-RunAsAdmin

# Optionally start dashboard early
if ($ShowDashboard) { Start-UpdateDashboard -HtmlPath $DashboardPath | Out-Null }

# Ensure PSWindowsUpdate is ready and attempt Windows 11 feature upgrade first
Initialize-PSWindowsUpdate | Out-Null
if (-not $SkipWin11Upgrade) { Invoke-Windows11FeatureUpgradeIfNeeded }

# SELF-DEPLOYMENT: Ensure script runs from C:\Scripts for reliability (unless skipped)
if (-not $SkipUpdateCheck) {
    # Invoke self-deployment and continue regardless of result
    Invoke-SelfDeployment -SourcePath $MyInvocation.MyCommand.Path -ForceDeployment $Deploy | Out-Null
} else {
    Write-LogMessage "Skipping update check as requested (post-update restart)" "INFO"
}

# SCHEDULED TASK CREATION: Set up Patch Tuesday automation if requested
if ($CreateSchedule) {
    $scheduleCreated = New-PatchTuesdaySchedule -RetryDays $RetryDays
    if ($scheduleCreated) {
        Write-Host "`nPatch Tuesday automation is now configured!" -ForegroundColor Cyan
        Write-Host "The system will automatically check for and install updates every month." -ForegroundColor Green
        Write-Host "Manual execution: Run 'WindowsUpdate-Manual' task from Task Scheduler" -ForegroundColor Gray
    }
}

# MONTHLY SCHEDULE AUTO-UPDATE: Automatically refresh schedule for next month
# WHY: One-time triggers expire, so we need to recreate them monthly
try {
    $scheduleUpdated = Update-MonthlySchedule
    if ($scheduleUpdated) {
        Write-LogMessage "Monthly schedule check completed successfully" "INFO"
    } else {
        Write-LogMessage "Monthly schedule update failed - manual intervention may be required" "WARNING"
    }
} catch {
    Write-LogMessage "Error during monthly schedule check: $_" "WARNING"
}

# CUMULATIVE UPDATE CHECK: Skip execution if checking for monthly cumulative update
if ($MyInvocation.BoundParameters.ContainsKey('CheckCumulative')) {
    if (-not (Test-CumulativeUpdateAvailable)) {
        Write-Host "No cumulative update available yet - will retry tomorrow" -ForegroundColor Yellow
        Exit 0
    }
    Write-Host "Cumulative update detected - proceeding with full update process" -ForegroundColor Green
}


# FUNCTION: Centralized Logging with Timestamps, Color Coding, and Event Viewer Integration
function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    
    Write-Host $logMessage -ForegroundColor $(
        if ($Type -eq "ERROR") {"Red"}
        elseif ($Type -eq "WARNING") {"Yellow"}
        else {"Green"}
    )
    
    $logMessage | Out-File -FilePath $global:logFile -Append -Encoding UTF8
    
    try {
        $eventType = switch ($Type) {
            "ERROR" { "Error"; $eventId = $script:EventIds.LOGGING_ERROR }
            "WARNING" { "Warning"; $eventId = 3000 }
            default { "Information"; $eventId = 1000 }
        }
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId $eventId -EntryType $eventType -Message $logMessage
    } catch { }
}

# FUNCTION: Launch HTML Dashboard
function Start-UpdateDashboard {
    param([string]$HtmlPath = "")
    
    try {
        if (-not $HtmlPath) {
            $scriptDir = "C:\Scripts"
            if (-not (Test-Path $scriptDir)) {
                $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
            }
            $HtmlPath = Join-Path $scriptDir "windows-comprehensive-updater-dashboard.html"
        }
        
        if (-not (Test-Path $HtmlPath)) {
            Write-LogMessage "Dashboard HTML file not found at: $HtmlPath" "WARNING"
            return $false
        }
        
        Write-LogMessage "Launching Windows Comprehensive Updater Dashboard..."
        Start-Process $HtmlPath
        return $true
    } catch {
        Write-LogMessage "Failed to launch dashboard: $_" "ERROR"
        return $false
    }
}

# FUNCTION: Update Dashboard Status
function Update-DashboardStatus {
    param([string]$Phase, [int]$Progress, [string]$CurrentOperation, [hashtable]$AdditionalData = @{})
    
    try {
        $statusFile = Join-Path (Split-Path $global:logFile) "update-status.json"
        $status = @{
            scriptRunning = $true
            lastUpdate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            phase = $Phase
            progress = $Progress
            currentOperation = $CurrentOperation
            systemInfo = @{
                computerName = $env:COMPUTERNAME
                userName = $env:USERNAME
                psVersion = $PSVersionTable.PSVersion.ToString()
            }
        }
        
        foreach ($key in $AdditionalData.Keys) {
            if ($key -ne 'wingetApps') {
                $status[$key] = $AdditionalData[$key]
            }
        }
        
        $status | ConvertTo-Json -Depth 3 | Out-File -FilePath $statusFile -Encoding UTF8
    } catch { }
}

# FUNCTION: Automatic Administrator Privilege Escalation
function Confirm-RunAsAdmin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "Auto-escalating to Administrator privileges..." -ForegroundColor Yellow
        
        $scriptPath = $MyInvocation.MyCommand.Path
        if (-not $scriptPath) {
            $scriptPath = $MyInvocation.ScriptName
        }
        
        $originalArgs = @()
        if ($ShowDashboard) { $originalArgs += "-ShowDashboard" }
        if ($DashboardPath) { $originalArgs += "-DashboardPath `"$DashboardPath`"" }
        if ($Deploy) { $originalArgs += "-Deploy" }
        if ($CreateSchedule) { $originalArgs += "-CreateSchedule" }
        if ($RetryDays -ne 7) { $originalArgs += "-RetryDays $RetryDays" }
        if ($SkipUpdateCheck) { $originalArgs += "-SkipUpdateCheck" }
        if ($CheckCumulative) { $originalArgs += "-CheckCumulative" }
        $argumentString = $originalArgs -join " "
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" $argumentString"
        
        try {
            Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments -WindowStyle Hidden
            Exit 0
        } catch {
            Write-Host "Failed to escalate privileges automatically. Please run as Administrator." -ForegroundColor Red
            Exit 1
        }
    }
}

# FUNCTION: Wizmo Download and Verification
function Confirm-WizmoAvailability {
    param ([string]$WinDir = $env:WINDIR)
    
    $wizmoPath = Join-Path $WinDir "wizmo.exe"
    
    if (Test-Path $wizmoPath) {
        Write-LogMessage "Wizmo already exists at $wizmoPath"
        return $wizmoPath
    }
    
    Write-LogMessage "Wizmo not found. Downloading from GRC.com..."
    $wizmoUrl = "https://www.grc.com/files/wizmo.exe"
    $expectedSha256 = "7b0b47f936d24686de461bea05a9480179035a5b2b23a74167a7728e95922d5d"
    
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell Windows Comprehensive Updater")
        $webClient.DownloadFile($wizmoUrl, $wizmoPath)
        
        Write-LogMessage "Wizmo downloaded successfully to $wizmoPath"
        
        try {
            $fileHash = (Get-FileHash -Algorithm SHA256 -Path $wizmoPath -ErrorAction Stop).Hash.ToUpper()
            if ($expectedSha256 -and $fileHash -ne $expectedSha256) {
                Write-LogMessage "Wizmo hash mismatch - deleting file" "ERROR"
                Remove-Item $wizmoPath -Force -ErrorAction SilentlyContinue
                return $null
            }
        } catch {
            Write-LogMessage "Hash verification skipped: $_" "WARNING"
        }
        
        if ((Test-Path $wizmoPath) -and ((Get-Item $wizmoPath).Length -gt 10KB)) {
            $sizeKB = [math]::Round((Get-Item $wizmoPath).Length / 1KB, 2)
            Write-LogMessage "Wizmo download verified (Size: $sizeKB KB)"
            
            try {
                Unblock-File -Path $wizmoPath
                Write-LogMessage "Wizmo file unblocked"
            } catch {
                Write-LogMessage "Unblock failed: $_" "WARNING"
            }
            
            return $wizmoPath
        }
    } catch {
        Write-LogMessage "Failed to download Wizmo: $_" "ERROR"
        return $null
    }
}

# FUNCTION: Silent Reboot with Wizmo Integration
function Invoke-SilentReboot {
    $wizmoPath = Confirm-WizmoAvailability
    
    Write-LogMessage "Initiating silent system reboot..."
    
    if ($wizmoPath -and (Test-Path $wizmoPath)) {
        try {
            Write-LogMessage "Using Wizmo for silent reboot"
            & $wizmoPath quiet reboot!
            Start-Sleep -Seconds 5
        } catch {
            Write-LogMessage "Wizmo reboot failed: $_" "WARNING"
            shutdown.exe /r /t 0 /f
        }
    } else {
        Write-LogMessage "Wizmo not available, using Windows reboot" "WARNING"
        shutdown.exe /r /t 0 /f
    }
}

# FUNCTION: Comprehensive Winget Application Updates
function Invoke-WingetUpdates {
    Write-LogMessage "Starting Winget application updates..." "INFO"
    if ($script:DisableWinget) { Write-LogMessage "Winget updates disabled" "WARNING"; return $true }
    Update-DashboardStatus -Phase "winget-updates" -Progress 10 -CurrentOperation "Checking for Winget application updates"
    
    try {
        $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
        if (-not $wingetPath) {
            Write-LogMessage "Winget not found - installing..." "WARNING"
            try {
                Start-Process "https://aka.ms/getwinget" -Wait
                Write-LogMessage "Winget installation initiated" "WARNING"
                return $false
            } catch {
                Write-LogMessage "Failed to install Winget: $_" "ERROR"
                return $false
            }
        }
        
        $null = & winget source update 2>&1
        Write-LogMessage "Checking for available application updates..." "INFO"
        
        $updates = @()
        $upgradeOutput = & winget upgrade 2>&1 | Out-String
        
        if ($upgradeOutput -match "No available upgrades") {
            Write-LogMessage "No Winget application updates available" "INFO"
            return $true
        }
        
        $lines = $upgradeOutput -split "`n"
        $headerFound = $false
        
        foreach ($line in $lines) {
            if ($line -match "^Name\s+Id\s+Version\s+Available") {
                $headerFound = $true
                continue
            }
            
            if ($headerFound -and $line.Trim() -and -not ($line -match "^\-+")) {
                $parts = $line -split '\s{2,}' | Where-Object { $_.Trim() }
                if ($parts.Count -ge 4) {
                    $updates += @{
                        Name = $parts[1].Trim()
                        DisplayName = $parts[0].Trim()
                        CurrentVersion = $parts[2].Trim()
                        AvailableVersion = $parts[3].Trim()
                    }
                }
            }
        }
        
        if ($updates.Count -eq 0) {
            Write-LogMessage "No parseable updates found" "WARNING"
            return $true
        }
        
        Write-LogMessage "Found $($updates.Count) application updates available" "INFO"
        Update-DashboardStatus -Phase "winget-updates" -Progress 20 -CurrentOperation "Processing $($updates.Count) application updates" -AdditionalData @{wingetApps = $updates.Count}
        
        $conflictingApps = $updates | Where-Object { $_.Name -match "Microsoft\.PowerShell|Microsoft\.WindowsTerminal|Microsoft\.VisualStudioCode" }
        $normalApps = $updates | Where-Object { $_.Name -notmatch "Microsoft\.PowerShell|Microsoft\.WindowsTerminal|Microsoft\.VisualStudioCode" }
        
        $totalApps = $updates.Count
        $processedApps = 0
        
        if ($normalApps.Count -gt 0) {
            Write-LogMessage "Updating $($normalApps.Count) standard applications..." "INFO"
            
            foreach ($app in $normalApps) {
                $processedApps++
                $progress = [math]::Round(($processedApps / $totalApps) * 60) + 20
                
                Write-LogMessage "Updating: $($app.DisplayName)" "INFO"
                Update-DashboardStatus -Phase "winget-updates" -Progress $progress -CurrentOperation "Updating $($app.DisplayName)"
                
                try {
                    $result = & winget upgrade $app.Name --silent --accept-package-agreements --accept-source-agreements 2>&1
                    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335212) {
                        Write-LogMessage "$($app.DisplayName) updated successfully" "SUCCESS"
                    } else {
                        Write-LogMessage "$($app.DisplayName) update failed" "WARNING"
                    }
                } catch {
                    Write-LogMessage "Error updating $($app.DisplayName): $_" "ERROR"
                }
                
                Start-Sleep -Seconds 1
            }
        }
        
        if ($conflictingApps.Count -gt 0) {
            Write-LogMessage "Processing $($conflictingApps.Count) conflicting applications..." "WARNING"
            Update-DashboardStatus -Phase "winget-conflicts" -Progress 80 -CurrentOperation "Handling conflicting application updates"
            Update-ConflictingApplications -ConflictingApps $conflictingApps
        }
        
        Write-LogMessage "Winget application updates completed" "SUCCESS"
        Update-DashboardStatus -Phase "winget-complete" -Progress 90 -CurrentOperation "Winget updates completed"
        return $true
        
    } catch {
        Write-LogMessage "Winget update process failed: $_" "ERROR"
        return $false
    }
}

# FUNCTION: Windows Update Installation
function Invoke-WindowsUpdates {
    Write-LogMessage "Starting Windows Update process..." "INFO"
    if ($script:DisableWindowsUpdates) { Write-LogMessage "Windows Updates disabled" "WARNING"; return $true }
    Update-DashboardStatus -Phase "windows-updates" -Progress 0 -CurrentOperation "Initializing Windows Updates"
    
    try {
        if (-not (Get-Module PSWindowsUpdate -ErrorAction SilentlyContinue)) {
            Write-LogMessage "PSWindowsUpdate module not loaded" "WARNING"
            if (-not (Initialize-PSWindowsUpdate)) {
                Write-LogMessage "Failed to initialize PSWindowsUpdate" "ERROR"
                return $false
            }
        }
        
        Write-LogMessage "Scanning for available Windows updates..." "INFO"
        Update-DashboardStatus -Phase "windows-scan" -Progress 10 -CurrentOperation "Scanning for Windows updates"
        
        $availableUpdates = Get-WUList -MicrosoftUpdate -AcceptAll -ErrorAction SilentlyContinue
        
        if (-not $availableUpdates -or $availableUpdates.Count -eq 0) {
            Write-LogMessage "No Windows updates available" "INFO"
            Update-DashboardStatus -Phase "windows-complete" -Progress 100 -CurrentOperation "No Windows updates needed"
            return $true
        }
        
        $updateCount = $availableUpdates.Count
        Write-LogMessage "Found $updateCount Windows update(s) available" "INFO"
        
        foreach ($update in $availableUpdates) {
            Write-LogMessage "Available: $($update.Title)" "INFO"
        }
        
        Update-DashboardStatus -Phase "windows-install" -Progress 20 -CurrentOperation "Installing $updateCount Windows updates"
        Write-LogMessage "Installing Windows updates..." "INFO"
        
        try {
            $installResult = Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -AutoReboot:$false -ErrorAction Stop
            
            $successCount = 0
            $failureCount = 0
            
            foreach ($result in $installResult) {
                if ($result.Result -eq "Installed" -or $result.Result -eq "Downloaded") {
                    Write-LogMessage "Installed: $($result.Title)" "SUCCESS"
                    $successCount++
                } else {
                    Write-LogMessage "Failed: $($result.Title)" "ERROR"
                    $failureCount++
                }
            }
            
            Write-LogMessage "Windows Updates completed: $successCount installed, $failureCount failed" "INFO"
            
            $rebootRequired = Get-WURebootStatus -Silent -ErrorAction SilentlyContinue
            if ($rebootRequired -or (Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue)) {
                Write-LogMessage "System reboot required" "WARNING"
                Update-DashboardStatus -Phase "reboot-required" -Progress 95 -CurrentOperation "System reboot required - restarting..."
                Start-Sleep -Seconds 3
                Invoke-SilentReboot
                return $true
            } else {
                Write-LogMessage "No reboot required" "SUCCESS"
                Update-DashboardStatus -Phase "windows-complete" -Progress 100 -CurrentOperation "Windows updates completed successfully"
                return $true
            }
            
        } catch {
            Write-LogMessage "Windows update installation failed: $_" "ERROR"
            return $false
        }
        
    } catch {
        Write-LogMessage "Windows update process failed: $_" "ERROR"
        return $false
    }
}

# FUNCTION: System Maintenance and Cleanup
function Invoke-SystemMaintenance {
    Write-LogMessage "Starting system maintenance..." "INFO"
    if ($script:DisableMaintenance) { Write-LogMessage "System maintenance disabled" "WARNING"; return $true }
    Update-DashboardStatus -Phase "maintenance" -Progress 0 -CurrentOperation "Starting system maintenance"
    
    try {
        Write-LogMessage "Cleaning temporary files..." "INFO"
        Update-DashboardStatus -Phase "cleanup" -Progress 25 -CurrentOperation "Cleaning temporary files"
        
        $tempPaths = @($env:TEMP, "C:\Windows\Temp", "C:\Windows\SoftwareDistribution\Download")
        
        foreach ($path in $tempPaths) {
            if (Test-Path $path) {
                try {
                    Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue | 
                    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } |
                    Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                    Write-LogMessage "Cleaned: $path" "INFO"
                } catch {
                    Write-LogMessage "Cleanup failed for $path : $_" "WARNING"
                }
            }
        }
        
        Write-LogMessage "Resetting Windows Update components..." "INFO"
        Update-DashboardStatus -Phase "wu-reset" -Progress 50 -CurrentOperation "Resetting Windows Update cache"
        
        try {
            Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
            Stop-Service bits -Force -ErrorAction SilentlyContinue
            Stop-Service cryptsvc -Force -ErrorAction SilentlyContinue
            Stop-Service msiserver -Force -ErrorAction SilentlyContinue
            
            if (Test-Path "C:\Windows\SoftwareDistribution\DataStore") {
                Remove-Item "C:\Windows\SoftwareDistribution\DataStore\*" -Force -Recurse -ErrorAction SilentlyContinue
            }
            
            Start-Service wuauserv -ErrorAction SilentlyContinue
            Start-Service bits -ErrorAction SilentlyContinue
            Start-Service cryptsvc -ErrorAction SilentlyContinue
            
            Write-LogMessage "Windows Update services reset" "SUCCESS"
        } catch {
            Write-LogMessage "Failed to reset Windows Update services: $_" "WARNING"
        }
        
        Write-LogMessage "Running system file integrity check..." "INFO"
        Update-DashboardStatus -Phase "sfc-scan" -Progress 75 -CurrentOperation "Checking system file integrity"
        
        try {
            $null = & sfc /scannow 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-LogMessage "System file check completed" "SUCCESS"
            } else {
                Write-LogMessage "System file check completed with warnings" "WARNING"
            }
        } catch {
            Write-LogMessage "System file check failed: $_" "WARNING"
        }
        
        Write-LogMessage "System maintenance completed" "SUCCESS"
        Update-DashboardStatus -Phase "maintenance-complete" -Progress 100 -CurrentOperation "System maintenance completed"
        return $true
        
    } catch {
        Write-LogMessage "System maintenance failed: $_" "ERROR"
        return $false
    }
}

# =========================================================================
# MAIN SCRIPT EXECUTION
# =========================================================================

try {
    # LOG ROTATION
    try {
        if (Test-Path $global:logFile) {
            $fileInfo = Get-Item $global:logFile -ErrorAction SilentlyContinue
            if ($fileInfo -and ($fileInfo.Length/1MB) -gt $script:MaxLogMB) {
                $archiveName = (Split-Path $global:logFile -Parent) + "\\WindowsUpdateLog_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".txt"
                Move-Item $global:logFile $archiveName -Force
                Write-Host "Log rotated to $archiveName" -ForegroundColor Gray
            }
            Get-ChildItem (Split-Path $global:logFile -Parent) -Filter 'WindowsUpdateLog_*.txt' -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$script:LogRetentionDays) } |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }
    } catch { Write-Host "Log rotation error: $_" -ForegroundColor Yellow }

    # STATE PERSISTENCE
    $stateFile = "C:\Scripts\update-state.json"
    $resumePhase = 'start'
    if (Test-Path $stateFile) {
        try {
            $state = Get-Content $stateFile -Raw | ConvertFrom-Json
            if ($state.rebootPending -and $state.phase -eq 'windows-updates') {
                Write-LogMessage "Resuming post-reboot Windows Updates phase" "INFO"
                $resumePhase = 'post-reboot'
            }
        } catch { }
    }

    # Helper function for state saving
    function Save-UpdateState([string]$phase) {
        try {
            $obj = @{ phase = $phase; timestamp = Get-Date; rebootPending = $true }
            $obj | ConvertTo-Json | Out-File $stateFile -Encoding UTF8
        } catch { }
    }

    # SCRIPT INITIALIZATION
    $global:totalUpdatesInstalled = 0
    $global:cycleCount = 1
    $script:startTime = Get-Date
    
    Write-LogMessage "Windows Update Script v2.1.0 - Enhanced Bulletproof Edition" "INFO"
    Write-LogMessage "Script started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "INFO"
    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId $script:EventIds.SCRIPT_START -EntryType Information -Message "Windows Update Script started"
    
    Update-DashboardStatus -Phase "initialization" -Progress 5 -CurrentOperation "Script initialization completed"
    
    # EXECUTION FLOW
    $wingetSuccess = $true
    $windowsSuccess = $true
    $maintenanceSuccess = $true

    if ($resumePhase -eq 'post-reboot') {
        Write-LogMessage "Skipping Winget phase (already completed before reboot)" "INFO"
        Write-LogMessage "=== RESUMED PHASE: WINDOWS SYSTEM UPDATES CONTINUATION ===" "INFO"
        $windowsSuccess = Invoke-WindowsUpdates
        Write-LogMessage "=== PHASE 3: SYSTEM MAINTENANCE ===" "INFO"
        $maintenanceSuccess = Invoke-SystemMaintenance
        try { Remove-Item $stateFile -Force -ErrorAction SilentlyContinue } catch { }
    } else {
        Write-LogMessage "=== PHASE 1: WINGET APPLICATION UPDATES ===" "INFO"
        $wingetSuccess = Invoke-WingetUpdates
        Write-LogMessage "=== PHASE 2: WINDOWS SYSTEM UPDATES ===" "INFO"
        Save-UpdateState 'windows-updates'
        $windowsSuccess = Invoke-WindowsUpdates
        Write-LogMessage "=== PHASE 3: SYSTEM MAINTENANCE ===" "INFO"
        $maintenanceSuccess = Invoke-SystemMaintenance
        try { if (Test-Path $stateFile) { Remove-Item $stateFile -Force -ErrorAction SilentlyContinue } } catch { }
    }
    
    # COMPLETION SUMMARY
    $endTime = Get-Date
    $duration = $endTime - $script:startTime
    $durationString = "{0:hh\:mm\:ss}" -f $duration
    
    Write-LogMessage "=== UPDATE PROCESS COMPLETED ===" "INFO"
    Write-LogMessage "Total execution time: $durationString" "INFO"
    Write-LogMessage "Winget updates: $(if($wingetSuccess){'SUCCESS'}else{'FAILED'})" "INFO"
    Write-LogMessage "Windows updates: $(if($windowsSuccess){'SUCCESS'}else{'FAILED'})" "INFO"
    Write-LogMessage "System maintenance: $(if($maintenanceSuccess){'SUCCESS'}else{'FAILED'})" "INFO"
    
    $overallSuccess = $wingetSuccess -and $windowsSuccess -and $maintenanceSuccess
    Update-DashboardStatus -Phase "completed" -Progress 100 -CurrentOperation "All updates completed" -AdditionalData @{
        wingetSuccess = $wingetSuccess
        windowsSuccess = $windowsSuccess
        maintenanceSuccess = $maintenanceSuccess
        executionTime = $durationString
        overallSuccess = $overallSuccess
    }
    
    if ($overallSuccess) {
        Write-LogMessage "Windows Update Script completed successfully!" "SUCCESS"
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId $script:EventIds.SCRIPT_SUCCESS -EntryType Information -Message "Windows Update Script completed successfully in $durationString"
    } else {
        Write-LogMessage "Windows Update Script completed with some failures" "WARNING"
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3002 -EntryType Warning -Message "Windows Update Script completed with failures in $durationString"
    }
    
    if ($ShowDashboard) {
        Write-LogMessage "Dashboard will remain open for monitoring - press any key to continue..." "INFO"
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    
} catch {
    Write-LogMessage "CRITICAL ERROR - Script execution failed: $_" "ERROR"
    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId $script:EventIds.SCRIPT_CRITICAL -EntryType Error -Message "Critical script failure: $_"
    
    Update-DashboardStatus -Phase "critical-error" -Progress 0 -CurrentOperation "Script execution failed" -AdditionalData @{
        errorMessage = $_.Exception.Message
        errorDetails = $_.ScriptStackTrace
    }
    
    Exit 1
}

Write-LogMessage "Script execution completed - exiting gracefully" "INFO"
Exit 0
