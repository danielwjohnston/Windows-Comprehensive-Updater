# ============================================================================
# FULLY AUTOMATED WINDOWS UPDATE SCRIPT WITH SELF-DEPLOYMENT
# Script Version: 2.1.0 - Enhanced Bulletproof Edition
# ============================================================================

# 
# PURPOSE: This script provides a completely hands-off Windows update experience
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
    [switch]$CheckCumulative       # Only check for cumulative updates, don't run full process
)

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
        $deployedScriptPath = "C:\Scripts\windows-update-script.ps1"
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
                        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1018 -EntryType Information -Message "$($app.DisplayName) updated via external process"
                    } else {
                        Write-LogMessage "$($app.DisplayName) external update failed with exit code: $exitCode" "ERROR"
                        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3019 -EntryType Warning -Message "$($app.DisplayName) external update failed: $exitCode"
                    }
                } else {
                    # TIMEOUT HANDLING: Kill process if it takes too long
                    Write-LogMessage "$($app.DisplayName) update timed out after $timeoutSeconds seconds" "ERROR"
                    try { $process.Kill() } catch { }
                    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5021 -EntryType Error -Message "$($app.DisplayName) update timed out"
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
            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5022 -EntryType Error -Message "Error updating $($app.DisplayName): $_"
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
    $targetScript = Join-Path $targetDir "windows-update-script.ps1"
    $targetDashboard = Join-Path $targetDir "windows-update-dashboard.html"
    
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
                (Join-Path (Split-Path $currentPath -Parent) "windows-update-dashboard.html"),
                # Custom dashboard path if specified
                $DashboardPath,
                # Current directory
                ".\windows-update-dashboard.html",
                # Look in common locations
                "C:\Temp\windows-update-dashboard.html",
                "$env:USERPROFILE\Downloads\windows-update-dashboard.html"
            ) | Where-Object { $_ -and (Test-Path $_) }
            
            if ($dashboardSources.Count -gt 0) {
                $sourceDashboard = $dashboardSources[0]
                Write-Host "Deploying dashboard from: $sourceDashboard" -ForegroundColor Green
                Copy-Item -Path $sourceDashboard -Destination $targetDashboard -Force
            } else {
                # CREATE EMBEDDED DASHBOARD: Generate dashboard if not found
                Write-Host "Dashboard not found - creating embedded version..." -ForegroundColor Yellow
                $this.CreateEmbeddedDashboard($targetDashboard)
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
function New-PatchTuesdaySchedule {
    param(
        [int]$RetryDays = 7
    )
    
    Write-Host "Creating Patch Tuesday automated schedule..." -ForegroundColor Cyan
    
    try {
        $scriptPath = Join-Path "C:\Scripts" "windows-update-script.ps1"
        
        # MAIN PATCH TUESDAY TASK: Second Tuesday of each month at 2 AM
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ShowDashboard"
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday -WeeksInterval 2 -At 2AM
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 30)
        
        Register-ScheduledTask -TaskName "WindowsUpdate-PatchTuesday" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        Write-Host "✓ Main Patch Tuesday task created (2nd Tuesday, 2:00 AM)" -ForegroundColor Green
        
        # RETRY TASKS: Additional attempts if cumulative update not found
        for ($day = 1; $day -le $RetryDays; $day++) {
            $retryAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -CheckCumulative"
            $retryTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday -WeeksInterval 2 -At (Get-Date "2:00 AM").AddDays($day)
            
            Register-ScheduledTask -TaskName "WindowsUpdate-PatchTuesday-Retry$day" -Action $retryAction -Trigger $retryTrigger -Principal $principal -Settings $settings -Force | Out-Null
            Write-Host "✓ Retry task $day created (+$day days from Patch Tuesday)" -ForegroundColor Green
        }
        
        # EMERGENCY MANUAL TASK: On-demand execution
        $manualAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ShowDashboard"
        Register-ScheduledTask -TaskName "WindowsUpdate-Manual" -Action $manualAction -Principal $principal -Settings $settings -Force | Out-Null
        Write-Host "✓ Manual execution task created (run anytime via Task Scheduler)" -ForegroundColor Green
        
        Write-Host "Patch Tuesday automation configured successfully!" -ForegroundColor Cyan
        Write-Host "Tasks created:" -ForegroundColor White
        Write-Host "  • WindowsUpdate-PatchTuesday (Main monthly execution)" -ForegroundColor Gray
        Write-Host "  • WindowsUpdate-PatchTuesday-Retry1-$RetryDays (Retry if no cumulative update)" -ForegroundColor Gray
        Write-Host "  • WindowsUpdate-Manual (On-demand execution)" -ForegroundColor Gray
        
        return $true
        
    } catch {
        Write-Host "Failed to create scheduled tasks: $_" -ForegroundColor Red
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
            $_.Title -like "*Windows 10*" -or 
            $_.Title -like "*Windows 11*" 
        }
        
        if ($updates -and $updates.Count -gt 0) {
            Write-Host "✓ Found $($updates.Count) cumulative update(s) available" -ForegroundColor Green
            $updates | ForEach-Object { Write-Host "  • $($_.Title)" -ForegroundColor Gray }
            return $true
        } else {
            Write-Host "⚠ No cumulative updates found - Microsoft may not have released this month's update yet" -ForegroundColor Yellow
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
<html><head><title>Windows Update Monitor</title><style>
body{background:#0b0b0b;color:#e6e6e6;font-family:system-ui;padding:20px;}
.card{background:#171717;border:1px solid #262626;border-radius:16px;padding:20px;margin:10px 0;}
.status{display:inline-block;padding:8px 16px;border-radius:20px;font-weight:bold;}
.running{background:#00ff0040;color:#00ff00;border:1px solid #00ff00;}
.log{background:#1e1e1e;border-radius:8px;padding:15px;height:300px;overflow-y:auto;font-family:monospace;}
</style></head><body>
<div class="card"><h1>Windows Update Monitor</h1>
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

# ============================================================================
# DEPLOYMENT AND INITIALIZATION
# ============================================================================

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

# CUMULATIVE UPDATE CHECK: Skip execution if checking for monthly cumulative update
if ($MyInvocation.BoundParameters.ContainsKey('CheckCumulative')) {
    if (-not (Test-CumulativeUpdateAvailable)) {
        Write-Host "No cumulative update available yet - will retry tomorrow" -ForegroundColor Yellow
        Exit 0
    }
    Write-Host "Cumulative update detected - proceeding with full update process" -ForegroundColor Green
}

# Continue with existing script functions...

# FUNCTION: Launch HTML Dashboard
# PURPOSE: Provides real-time visual monitoring of the update process
function Start-UpdateDashboard {
    param(
        [string]$HtmlPath = ""
    )
    
    try {
        # DASHBOARD LOCATION: Look for dashboard in same directory as script
        if (-not $HtmlPath) {
            # DEPLOYMENT-AWARE PATH: Use C:\Scripts if deployed, otherwise same directory
            $scriptDir = "C:\Scripts"
            if (-not (Test-Path $scriptDir)) {
                $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
            }
            $HtmlPath = Join-Path $scriptDir "windows-update-dashboard.html"
        }
        
        # DASHBOARD AVAILABILITY CHECK: Ensure dashboard file exists
        if (-not (Test-Path $HtmlPath)) {
            Write-LogMessage "Dashboard HTML file not found at: $HtmlPath" "WARNING"
            Write-LogMessage "Continuing without dashboard..." "WARNING"
            return $false
        }
        
        # BROWSER LAUNCH: Open dashboard in default browser
        Write-LogMessage "Launching Windows Update Dashboard..."
        Start-Process $HtmlPath
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1020 -EntryType Information -Message "Dashboard launched at: $HtmlPath"
        
        # DASHBOARD STATUS FILE: Create status file for dashboard communication
        $statusFile = Join-Path (Split-Path $global:logFile) "update-status.json"
        $initialStatus = @{
            scriptRunning = $true
            startTime = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            phase = "initialization"
            progress = 0
            totalUpdates = 0
            currentCycle = 0
            wingetApps = 0
            systemInfo = @{
                computerName = $env:COMPUTERNAME
                userName = $env:USERNAME
                psVersion = $PSVersionTable.PSVersion.ToString()
                scriptPath = $MyInvocation.MyCommand.Path
            }
        } | ConvertTo-Json -Depth 3
        
        $initialStatus | Out-File -FilePath $statusFile -Encoding UTF8
        Write-LogMessage "Dashboard status file created: $statusFile"
        
        return $true
        
    } catch {
        Write-LogMessage "Failed to launch dashboard: $_" "ERROR"
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5020 -EntryType Error -Message "Dashboard launch failed: $_"
        return $false
    }
}

# FUNCTION: Update Dashboard Status
# PURPOSE: Provide real-time status updates to the dashboard
function Update-DashboardStatus {
    param(
        [string]$Phase,
        [int]$Progress,
        [string]$CurrentOperation,
        [hashtable]$AdditionalData = @{}
    )
    
    try {
        $statusFile = Join-Path (Split-Path $global:logFile) "update-status.json"
        
        # STATUS OBJECT: Current script status for dashboard
        $status = @{
            scriptRunning = $true
            lastUpdate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            phase = $Phase
            progress = $Progress
            currentOperation = $CurrentOperation
            totalUpdates = $totalUpdatesInstalled
            currentCycle = $cycleCount
            wingetApps = if ($AdditionalData.ContainsKey('wingetApps')) { $AdditionalData.wingetApps } else { 0 }
            systemInfo = @{
                computerName = $env:COMPUTERNAME
                userName = $env:USERNAME
                psVersion = $PSVersionTable.PSVersion.ToString()
                scriptPath = $MyInvocation.MyCommand.Path
            }
        }
        
        # MERGE ADDITIONAL DATA: Include any extra information
        foreach ($key in $AdditionalData.Keys) {
            if ($key -ne 'wingetApps') {  # Already handled above
                $status[$key] = $AdditionalData[$key]
            }
        }
        
        # WRITE STATUS FILE: Update dashboard communication file
        $status | ConvertTo-Json -Depth 3 | Out-File -FilePath $statusFile -Encoding UTF8
        
    } catch {
        # NON-CRITICAL: Dashboard status update failure shouldn't stop script
        Write-LogMessage "Dashboard status update failed: $_" "WARNING"
    }
}
# 
# PURPOSE: This script provides a completely hands-off Windows update experience
# that automatically handles privilege escalation, downloads required tools,
# installs all available updates, and performs silent reboots as needed.
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

# FUNCTION: Automatic Administrator Privilege Escalation
# WHY THIS EXISTS: Many Windows operations (especially updates) require admin rights.
# Rather than forcing users to remember to "Run as Administrator", this function
# detects insufficient privileges and automatically restarts the script with elevation.
# DESIGN CHOICE: Uses hidden window to minimize user disruption during restart.
function Confirm-RunAsAdmin {
    # CHECK: Determine if current process has administrator privileges
    # EXPLANATION: WindowsPrincipal class checks if current user token contains
    # the built-in Administrator role. This is the standard .NET way to check elevation.
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        
        # FEEDBACK: Inform user what's happening (brief message before window closes)
        Write-Host "Auto-escalating to Administrator privileges..." -ForegroundColor Yellow
        
        # PATH DETECTION: Get the current script's full path for restarting
        # WHY BOTH CHECKS: $MyInvocation.MyCommand.Path works in most cases,
        # but $MyInvocation.ScriptName is a fallback for edge cases
        $scriptPath = $MyInvocation.MyCommand.Path
        if (-not $scriptPath) {
            $scriptPath = $MyInvocation.ScriptName
        }
        
        # COMMAND LINE CONSTRUCTION: Build arguments for elevated restart
        # -NoProfile: Skip loading PowerShell profiles (faster startup)
        # -ExecutionPolicy Bypass: Override script execution restrictions temporarily
        # -File: Specify the script file to run (safer than -Command)
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        
        # ESCALATION ATTEMPT: Start new PowerShell process with admin privileges
        try {
            # Start-Process with -Verb RunAs triggers the UAC elevation prompt
            # -WindowStyle Hidden minimizes disruption to user workflow
            Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments -WindowStyle Hidden
            # EXIT: Clean exit from the non-elevated instance
            Exit 0
        } catch {
            # FALLBACK: If automatic elevation fails, provide clear instructions
            Write-Host "Failed to escalate privileges automatically. Please run as Administrator." -ForegroundColor Red
            Exit 1
        }
    }
}

# FUNCTION: Centralized Logging with Timestamps, Color Coding, and Event Viewer Integration
# WHY CENTRALIZED: Having one logging function ensures consistent formatting
# and makes it easy to add features like log levels, file rotation, etc.
# DESIGN CHOICE: Triple output (console + file + event log) for complete audit trail
function Write-LogMessage {
    param (
        [string]$Message,              # The actual log message content
        [string]$Type = "INFO"         # Log level: INFO, WARNING, ERROR
    )
    
    # TIMESTAMP: ISO format for easy parsing and international compatibility
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # MESSAGE FORMATTING: Structured format for easy reading and parsing
    $logMessage = "[$timestamp] [$Type] $Message"
    
    # CONSOLE OUTPUT: Color-coded based on message type for quick visual parsing
    # WHY COLOR CODING: Helps administrators quickly identify issues during execution
    Write-Host $logMessage -ForegroundColor $(
        if ($Type -eq "ERROR") {"Red"}           # Errors in red - critical attention needed
        elseif ($Type -eq "WARNING") {"Yellow"}  # Warnings in yellow - caution advised
        else {"Green"}                           # Normal operations in green - all good
    )
    
    # FILE LOGGING: Persistent record for audit trails and troubleshooting
    # Out-File with -Append: Accumulates log entries without overwriting
    # -Encoding UTF8: Ensures special characters display correctly
    $logMessage | Out-File -FilePath $global:logFile -Append -Encoding UTF8
    
    # EVENT VIEWER LOGGING: Enterprise-grade logging for system administrators
    # WHY EVENT VIEWER: Centralized Windows logging, integrates with monitoring tools
    # EVENT ID SCHEMA: 1xxx = Info, 2xxx = Process, 3xxx = Warning, 5xxx = Error
    try {
        $eventType = switch ($Type) {
            "ERROR" { "Error"; $eventId = 5000 }
            "WARNING" { "Warning"; $eventId = 3000 }
            default { "Information"; $eventId = 1000 }
        }
        
        # WRITE-EVENTLOG: Send to Windows Event Log
        # -LogName Application: Standard application event log
        # -Source: Custom source for filtering (registered during script initialization)
        # -Message: Full formatted message with timestamp
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId $eventId -EntryType $eventType -Message $logMessage
    } catch {
        # EVENT LOG FAILURE: Non-critical, don't interrupt main script flow
        # This can fail if event source isn't registered, but we'll try to register it later
    }
}

# FUNCTION: Wizmo Download and Verification
# WHY NEEDED: Wizmo enables truly silent reboots without user prompts.
# Windows built-in shutdown command shows notifications that can interrupt automation.
# DESIGN PHILOSOPHY: Download only if needed, verify integrity, provide fallback
# POWERSHELL APPROVED VERB: Using 'Confirm' instead of 'Ensure' for PowerShell compliance
function Confirm-WizmoAvailability {
    param (
        [string]$WinDir = $env:WINDIR  # Default to Windows directory for system-wide access
    )
    
    # PATH CONSTRUCTION: Place in Windows directory for system-wide accessibility
    # WHY WINDOWS DIR: It's in the system PATH, so "wizmo command" works from anywhere
    # COLON SAFETY: Use Join-Path to avoid colon parsing issues in string interpolation
    $wizmoPath = Join-Path $WinDir "wizmo.exe"
    
    # EXISTENCE CHECK: Don't download if already present
    if (Test-Path $wizmoPath) {
        Write-LogMessage "Wizmo already exists at $wizmoPath"
        return $wizmoPath
    }
    
    # DOWNLOAD INITIATION: Inform user of download process
    Write-LogMessage "Wizmo not found. Downloading from GRC.com..."
    $wizmoUrl = "https://www.grc.com/files/wizmo.exe"
    
    try {
        # SECURITY PROTOCOL: Force TLS 1.2 for secure HTTPS connections
        # WHY NEEDED: Older .NET versions default to less secure protocols
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        # DOWNLOAD SETUP: Create WebClient with proper user agent
        # USER AGENT: Identifies our script - good practice for web requests
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell Windows Update Script")
        
        # FILE DOWNLOAD: Direct download to target location
        # WHY DownloadFile vs Invoke-WebRequest: More reliable for binary files
        $webClient.DownloadFile($wizmoUrl, $wizmoPath)
        
        Write-LogMessage "Wizmo downloaded successfully to $wizmoPath"
        
        # INTEGRITY VERIFICATION: Ensure download completed properly
        # WHY SIZE CHECK: Corrupted downloads often result in tiny files
        # 10KB threshold: Wizmo is much larger, so this catches obvious failures
        if ((Test-Path $wizmoPath) -and ((Get-Item $wizmoPath).Length -gt 10KB)) {
            # SIZE REPORTING: Convert bytes to KB for human-readable feedback
            $sizeKB = [math]::Round((Get-Item $wizmoPath).Length / 1KB, 2)
            Write-LogMessage "Wizmo download verified (Size: $sizeKB KB)"
            
            # UNBLOCK FILE: Remove "downloaded from internet" restrictions
            # WHY NEEDED: Downloaded files are marked with Zone.Identifier and require "Unblock"
            # This prevents the "Do you want to run this file?" security prompts
            try {
                Unblock-File -Path $wizmoPath
                Write-LogMessage "Wizmo file unblocked for execution"
            } catch {
                # FALLBACK: Manual unblocking if Unblock-File fails
                Write-LogMessage "Automatic unblock failed, attempting manual zone removal..." "WARNING"
                try {
                    # MANUAL ZONE REMOVAL: Delete the Zone.Identifier alternate data stream
                    # This is the low-level approach when Unblock-File doesn't work
                    $zoneStreamPath = "$wizmoPath`:Zone.Identifier"
                    if (Get-Item -Path $zoneStreamPath -Stream Zone.Identifier -ErrorAction SilentlyContinue) {
                        Remove-Item -Path $zoneStreamPath -ErrorAction SilentlyContinue
                        Write-LogMessage "Zone.Identifier stream removed manually"
                    }
                } catch {
                    Write-LogMessage "Manual zone removal also failed: $_" "WARNING"
                    Write-LogMessage "Wizmo may require manual unblocking in Properties dialog" "WARNING"
                }
            }
            
            return $wizmoPath
        } else {
            # FAILURE DETECTION: File exists but is suspiciously small
            throw "Downloaded file appears to be invalid"
        }
    } catch {
        # ERROR HANDLING: Log the failure but don't crash the entire script
        Write-LogMessage "Failed to download Wizmo: $_" "ERROR"
        Write-LogMessage "Will use standard Windows reboot as fallback" "WARNING"
        return $null  # Return null to signal fallback needed
    }
}

# FUNCTION: Silent Reboot with Wizmo Integration
# PURPOSE: Perform system restart without any user prompts or delays
# WHY WIZMO FIRST: Wizmo provides completely silent operation, Windows tools show notifications
function Invoke-SilentReboot {
    # TOOL ACQUISITION: Ensure Wizmo is available for silent reboot
    $wizmoPath = Confirm-WizmoAvailability
    
    Write-LogMessage "Initiating silent system reboot..."
    
    # PRIMARY METHOD: Use Wizmo for completely silent operation
    if ($wizmoPath -and (Test-Path $wizmoPath)) {
        try {
            Write-LogMessage "Using Wizmo for silent reboot"
            # WIZMO EXECUTION: Use correct syntax for truly silent reboot
            # WHY "quiet reboot!": The "quiet" parameter suppresses all notifications and dialogs
            & $wizmoPath quiet reboot!
            # BRIEF PAUSE: Allow the reboot command to initiate before script continues
            Start-Sleep -Seconds 5
        } catch {
            # GRACEFUL DEGRADATION: Fall back to Windows built-in if Wizmo fails
            Write-LogMessage "Wizmo reboot failed, using Windows reboot: $_" "WARNING"
            # SHUTDOWN.EXE: Built-in Windows tool
            # /r = restart, /t 0 = no delay, /f = force close applications
            shutdown.exe /r /t 0 /f
        }
    } else {
        # FALLBACK METHOD: Use standard Windows shutdown when Wizmo unavailable
        Write-LogMessage "Wizmo not available, using Windows reboot" "WARNING"
        shutdown.exe /r /t 0 /f
    }
}

# =========================================================================# ============================================================================
# FULLY AUTOMATED WINDOWS UPDATE SCRIPT WITH SELF-DEPLOYMENT
# Script Version: 2.1.0 - Enhanced Edition
# ============================================================================

# 
# PURPOSE: This script provides a completely hands-off Windows update experience
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
    [switch]$CheckCumulative       # Only check for cumulative updates, don't run full process
)

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
        $deployedScriptPath = "C:\Scripts\windows-update-script.ps1"
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
                        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1018 -EntryType Information -Message "$($app.DisplayName) updated via external process"
                    } else {
                        Write-LogMessage "$($app.DisplayName) external update failed with exit code: $exitCode" "ERROR"
                        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3019 -EntryType Warning -Message "$($app.DisplayName) external update failed: $exitCode"
                    }
                } else {
                    # TIMEOUT HANDLING: Kill process if it takes too long
                    Write-LogMessage "$($app.DisplayName) update timed out after $timeoutSeconds seconds" "ERROR"
                    try { $process.Kill() } catch { }
                    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5021 -EntryType Error -Message "$($app.DisplayName) update timed out"
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
            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5022 -EntryType Error -Message "Error updating $($app.DisplayName): $_"
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
    $targetScript = Join-Path $targetDir "windows-update-script.ps1"
    $targetDashboard = Join-Path $targetDir "windows-update-dashboard.html"
    
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
                (Join-Path (Split-Path $currentPath -Parent) "windows-update-dashboard.html"),
                # Custom dashboard path if specified
                $DashboardPath,
                # Current directory
                ".\windows-update-dashboard.html",
                # Look in common locations
                "C:\Temp\windows-update-dashboard.html",
                "$env:USERPROFILE\Downloads\windows-update-dashboard.html"
            ) | Where-Object { $_ -and (Test-Path $_) }
            
            if ($dashboardSources.Count -gt 0) {
                $sourceDashboard = $dashboardSources[0]
                Write-Host "Deploying dashboard from: $sourceDashboard" -ForegroundColor Green
                Copy-Item -Path $sourceDashboard -Destination $targetDashboard -Force
            } else {
                # CREATE EMBEDDED DASHBOARD: Generate dashboard if not found
                Write-Host "Dashboard not found - creating embedded version..." -ForegroundColor Yellow
                $this.CreateEmbeddedDashboard($targetDashboard)
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
function New-PatchTuesdaySchedule {
    param(
        [int]$RetryDays = 7
    )
    
    Write-Host "Creating Patch Tuesday automated schedule..." -ForegroundColor Cyan
    
    try {
        $scriptPath = Join-Path "C:\Scripts" "windows-update-script.ps1"
        
        # MAIN PATCH TUESDAY TASK: Second Tuesday of each month at 2 AM
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ShowDashboard"
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday -WeeksInterval 2 -At 2AM
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 30)
        
        Register-ScheduledTask -TaskName "WindowsUpdate-PatchTuesday" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        Write-Host "✓ Main Patch Tuesday task created (2nd Tuesday, 2:00 AM)" -ForegroundColor Green
        
        # RETRY TASKS: Additional attempts if cumulative update not found
        for ($day = 1; $day -le $RetryDays; $day++) {
            $retryAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -CheckCumulative"
            $retryTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday -WeeksInterval 2 -At (Get-Date "2:00 AM").AddDays($day)
            
            Register-ScheduledTask -TaskName "WindowsUpdate-PatchTuesday-Retry$day" -Action $retryAction -Trigger $retryTrigger -Principal $principal -Settings $settings -Force | Out-Null
            Write-Host "✓ Retry task $day created (+$day days from Patch Tuesday)" -ForegroundColor Green
        }
        
        # EMERGENCY MANUAL TASK: On-demand execution
        $manualAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ShowDashboard"
        Register-ScheduledTask -TaskName "WindowsUpdate-Manual" -Action $manualAction -Principal $principal -Settings $settings -Force | Out-Null
        Write-Host "✓ Manual execution task created (run anytime via Task Scheduler)" -ForegroundColor Green
        
        Write-Host "Patch Tuesday automation configured successfully!" -ForegroundColor Cyan
        Write-Host "Tasks created:" -ForegroundColor White
        Write-Host "  • WindowsUpdate-PatchTuesday (Main monthly execution)" -ForegroundColor Gray
        Write-Host "  • WindowsUpdate-PatchTuesday-Retry1-$RetryDays (Retry if no cumulative update)" -ForegroundColor Gray
        Write-Host "  • WindowsUpdate-Manual (On-demand execution)" -ForegroundColor Gray
        
        return $true
        
    } catch {
        Write-Host "Failed to create scheduled tasks: $_" -ForegroundColor Red
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
            $_.Title -like "*Windows 10*" -or 
            $_.Title -like "*Windows 11*" 
        }
        
        if ($updates -and $updates.Count -gt 0) {
            Write-Host "✓ Found $($updates.Count) cumulative update(s) available" -ForegroundColor Green
            $updates | ForEach-Object { Write-Host "  • $($_.Title)" -ForegroundColor Gray }
            return $true
        } else {
            Write-Host "⚠ No cumulative updates found - Microsoft may not have released this month's update yet" -ForegroundColor Yellow
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
<html><head><title>Windows Update Monitor</title><style>
body{background:#0b0b0b;color:#e6e6e6;font-family:system-ui;padding:20px;}
.card{background:#171717;border:1px solid #262626;border-radius:16px;padding:20px;margin:10px 0;}
.status{display:inline-block;padding:8px 16px;border-radius:20px;font-weight:bold;}
.running{background:#00ff0040;color:#00ff00;border:1px solid #00ff00;}
.log{background:#1e1e1e;border-radius:8px;padding:15px;height:300px;overflow-y:auto;font-family:monospace;}
</style></head><body>
<div class="card"><h1>Windows Update Monitor</h1>
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

# ============================================================================
# DEPLOYMENT AND INITIALIZATION
# ============================================================================

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

# CUMULATIVE UPDATE CHECK: Skip execution if checking for monthly cumulative update
if ($MyInvocation.BoundParameters.ContainsKey('CheckCumulative')) {
    if (-not (Test-CumulativeUpdateAvailable)) {
        Write-Host "No cumulative update available yet - will retry tomorrow" -ForegroundColor Yellow
        Exit 0
    }
    Write-Host "Cumulative update detected - proceeding with full update process" -ForegroundColor Green
}

# Continue with existing script functions...

# FUNCTION: Launch HTML Dashboard
# PURPOSE: Provides real-time visual monitoring of the update process
function Start-UpdateDashboard {
    param(
        [string]$HtmlPath = ""
    )
    
    try {
        # DASHBOARD LOCATION: Look for dashboard in same directory as script
        if (-not $HtmlPath) {
            # DEPLOYMENT-AWARE PATH: Use C:\Scripts if deployed, otherwise same directory
            $scriptDir = "C:\Scripts"
            if (-not (Test-Path $scriptDir)) {
                $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
            }
            $HtmlPath = Join-Path $scriptDir "windows-update-dashboard.html"
        }
        
        # DASHBOARD AVAILABILITY CHECK: Ensure dashboard file exists
        if (-not (Test-Path $HtmlPath)) {
            Write-LogMessage "Dashboard HTML file not found at: $HtmlPath" "WARNING"
            Write-LogMessage "Continuing without dashboard..." "WARNING"
            return $false
        }
        
        # BROWSER LAUNCH: Open dashboard in default browser
        Write-LogMessage "Launching Windows Update Dashboard..."
        Start-Process $HtmlPath
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1020 -EntryType Information -Message "Dashboard launched at: $HtmlPath"
        
        # DASHBOARD STATUS FILE: Create status file for dashboard communication
        $statusFile = Join-Path (Split-Path $global:logFile) "update-status.json"
        $initialStatus = @{
            scriptRunning = $true
            startTime = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            phase = "initialization"
            progress = 0
            totalUpdates = 0
            currentCycle = 0
            wingetApps = 0
            systemInfo = @{
                computerName = $env:COMPUTERNAME
                userName = $env:USERNAME
                psVersion = $PSVersionTable.PSVersion.ToString()
                scriptPath = $MyInvocation.MyCommand.Path
            }
        } | ConvertTo-Json -Depth 3
        
        $initialStatus | Out-File -FilePath $statusFile -Encoding UTF8
        Write-LogMessage "Dashboard status file created: $statusFile"
        
        return $true
        
    } catch {
        Write-LogMessage "Failed to launch dashboard: $_" "ERROR"
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5020 -EntryType Error -Message "Dashboard launch failed: $_"
        return $false
    }
}

# FUNCTION: Update Dashboard Status
# PURPOSE: Provide real-time status updates to the dashboard
function Update-DashboardStatus {
    param(
        [string]$Phase,
        [int]$Progress,
        [string]$CurrentOperation,
        [hashtable]$AdditionalData = @{}
    )
    
    try {
        $statusFile = Join-Path (Split-Path $global:logFile) "update-status.json"
        
        # STATUS OBJECT: Current script status for dashboard
        $status = @{
            scriptRunning = $true
            lastUpdate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            phase = $Phase
            progress = $Progress
            currentOperation = $CurrentOperation
            totalUpdates = $totalUpdatesInstalled
            currentCycle = $cycleCount
            wingetApps = if ($AdditionalData.ContainsKey('wingetApps')) { $AdditionalData.wingetApps } else { 0 }
            systemInfo = @{
                computerName = $env:COMPUTERNAME
                userName = $env:USERNAME
                psVersion = $PSVersionTable.PSVersion.ToString()
                scriptPath = $MyInvocation.MyCommand.Path
            }
        }
        
        # MERGE ADDITIONAL DATA: Include any extra information
        foreach ($key in $AdditionalData.Keys) {
            if ($key -ne 'wingetApps') {  # Already handled above
                $status[$key] = $AdditionalData[$key]
            }
        }
        
        # WRITE STATUS FILE: Update dashboard communication file
        $status | ConvertTo-Json -Depth 3 | Out-File -FilePath $statusFile -Encoding UTF8
        
    } catch {
        # NON-CRITICAL: Dashboard status update failure shouldn't stop script
        Write-LogMessage "Dashboard status update failed: $_" "WARNING"
    }
}
# 
# PURPOSE: This script provides a completely hands-off Windows update experience
# that automatically handles privilege escalation, downloads required tools,
# installs all available updates, and performs silent reboots as needed.
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

# FUNCTION: Automatic Administrator Privilege Escalation
# WHY THIS EXISTS: Many Windows operations (especially updates) require admin rights.
# Rather than forcing users to remember to "Run as Administrator", this function
# detects insufficient privileges and automatically restarts the script with elevation.
# DESIGN CHOICE: Uses hidden window to minimize user disruption during restart.
function Confirm-RunAsAdmin {
    # CHECK: Determine if current process has administrator privileges
    # EXPLANATION: WindowsPrincipal class checks if current user token contains
    # the built-in Administrator role. This is the standard .NET way to check elevation.
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        
        # FEEDBACK: Inform user what's happening (brief message before window closes)
        Write-Host "Auto-escalating to Administrator privileges..." -ForegroundColor Yellow
        
        # PATH DETECTION: Get the current script's full path for restarting
        # WHY BOTH CHECKS: $MyInvocation.MyCommand.Path works in most cases,
        # but $MyInvocation.ScriptName is a fallback for edge cases
        $scriptPath = $MyInvocation.MyCommand.Path
        if (-not $scriptPath) {
            $scriptPath = $MyInvocation.ScriptName
        }
        
        # COMMAND LINE CONSTRUCTION: Build arguments for elevated restart
        # -NoProfile: Skip loading PowerShell profiles (faster startup)
        # -ExecutionPolicy Bypass: Override script execution restrictions temporarily
        # -File: Specify the script file to run (safer than -Command)
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        
        # ESCALATION ATTEMPT: Start-Process with -Verb RunAs triggers the UAC elevation prompt
            # -WindowStyle Hidden minimizes disruption to user workflow
            Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments -WindowStyle Hidden
            # EXIT: Clean exit from the non-elevated instance
            Exit 0
        } catch {
            # FALLBACK: If automatic elevation fails, provide clear instructions
            Write-Host "Failed to escalate privileges automatically. Please run as Administrator." -ForegroundColor Red
            Exit 1
        }
    }
}

# FUNCTION: Centralized Logging with Timestamps, Color Coding, and Event Viewer Integration
# WHY CENTRALIZED: Having one logging function ensures consistent formatting
# and makes it easy to add features like log levels, file rotation, etc.
# DESIGN CHOICE: Triple output (console + file + event log) for complete audit trail
function Write-LogMessage {
    param (
        [string]$Message,              # The actual log message content
        [string]$Type = "INFO"         # Log level: INFO, WARNING, ERROR
    )
    
    # TIMESTAMP: ISO format for easy parsing and international compatibility
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # MESSAGE FORMATTING: Structured format for easy reading and parsing
    $logMessage = "[$timestamp] [$Type] $Message"
    
    # CONSOLE OUTPUT: Color-coded based on message type for quick visual parsing
    # WHY COLOR CODING: Helps administrators quickly identify issues during execution
    Write-Host $logMessage -ForegroundColor $(
        if ($Type -eq "ERROR") {"Red"}           # Errors in red - critical attention needed
        elseif ($Type -eq "WARNING") {"Yellow"}  # Warnings in yellow - caution advised
        else {"Green"}                           # Normal operations in green - all good
    )
    
    # FILE LOGGING: Persistent record for audit trails and troubleshooting
    # Out-File with -Append: Accumulates log entries without overwriting
    # -Encoding UTF8: Ensures special characters display correctly
    $logMessage | Out-File -FilePath $global:logFile -Append -Encoding UTF8
    
    # EVENT VIEWER LOGGING: Enterprise-grade logging for system administrators
    # WHY EVENT VIEWER: Centralized Windows logging, integrates with monitoring tools
    # EVENT ID SCHEMA: 1xxx = Info, 2xxx = Process, 3xxx = Warning, 5xxx = Error
    try {
        $eventType = switch ($Type) {
            "ERROR" { "Error"; $eventId = 5000 }
            "WARNING" { "Warning"; $eventId = 3000 }
            default { "Information"; $eventId = 1000 }
        }
        
        # WRITE-EVENTLOG: Send to Windows Event Log
        # -LogName Application: Standard application event log
        # -Source: Custom source for filtering (registered during script initialization)
        # -Message: Full formatted message with timestamp
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId $eventId -EntryType $eventType -Message $logMessage
    } catch {
        # EVENT LOG FAILURE: Non-critical, don't interrupt main script flow
        # This can fail if event source isn't registered, but we'll try to register it later
    }
}

# FUNCTION: Wizmo Download and Verification
# WHY NEEDED: Wizmo enables truly silent reboots without user prompts.
# Windows built-in shutdown command shows notifications that can interrupt automation.
# DESIGN PHILOSOPHY: Download only if needed, verify integrity, provide fallback
# POWERSHELL APPROVED VERB: Using 'Confirm' instead of 'Ensure' for PowerShell compliance
function Confirm-WizmoAvailability {
    param (
        [string]$WinDir = $env:WINDIR  # Default to Windows directory for system-wide access
    )
    
    # PATH CONSTRUCTION: Place in Windows directory for system-wide accessibility
    # WHY WINDOWS DIR: It's in the system PATH, so "wizmo command" works from anywhere
    # COLON SAFETY: Use Join-Path to avoid colon parsing issues in string interpolation
    $wizmoPath = Join-Path $WinDir "wizmo.exe"
    
    # EXISTENCE CHECK: Don't download if already present
    if (Test-Path $wizmoPath) {
        Write-LogMessage "Wizmo already exists at $wizmoPath"
        return $wizmoPath
    }
    
    # DOWNLOAD INITIATION: Inform user of download process
    Write-LogMessage "Wizmo not found. Downloading from GRC.com..."
    $wizmoUrl = "https://www.grc.com/files/wizmo.exe"
    
    try {
        # SECURITY PROTOCOL: Force TLS 1.2 for secure HTTPS connections
        # WHY NEEDED: Older .NET versions default to less secure protocols
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        # DOWNLOAD SETUP: Create WebClient with proper user agent
        # USER AGENT: Identifies our script - good practice for web requests
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell Windows Update Script")
        
        # FILE DOWNLOAD: Direct download to target location
        # WHY DownloadFile vs Invoke-WebRequest: More reliable for binary files
        $webClient.DownloadFile($wizmoUrl, $wizmoPath)
        
        Write-LogMessage "Wizmo downloaded successfully to $wizmoPath"
        
        # INTEGRITY VERIFICATION: Ensure download completed properly
        # WHY SIZE CHECK: Corrupted downloads often result in tiny files
        # 10KB threshold: Wizmo is much larger, so this catches obvious failures
        if ((Test-Path $wizmoPath) -and ((Get-Item $wizmoPath).Length -gt 10KB)) {
            # SIZE REPORTING: Convert bytes to KB for human-readable feedback
            $sizeKB = [math]::Round((Get-Item $wizmoPath).Length / 1KB, 2)
            Write-LogMessage "Wizmo download verified (Size: $sizeKB KB)"
            
            # UNBLOCK FILE: Remove "downloaded from internet" restrictions
            # WHY NEEDED: Downloaded files are marked with Zone.Identifier and require "Unblock"
            # This prevents the "Do you want to run this file?" security prompts
            try {
                Unblock-File -Path $wizmoPath
                Write-LogMessage "Wizmo file unblocked for execution"
            } catch {
                # FALLBACK: Manual unblocking if Unblock-File fails
                Write-LogMessage "Automatic unblock failed, attempting manual zone removal..." "WARNING"
                try {
                    # MANUAL ZONE REMOVAL: Delete the Zone.Identifier alternate data stream
                    # This is the low-level approach when Unblock-File doesn't work
                    $zoneStreamPath = "$wizmoPath`:Zone.Identifier"
                    if (Get-Item -Path $zoneStreamPath -Stream Zone.Identifier -ErrorAction SilentlyContinue) {
                        Remove-Item -Path $zoneStreamPath -ErrorAction SilentlyContinue
                        Write-LogMessage "Zone.Identifier stream removed manually"
                    }
                } catch {
                    Write-LogMessage "Manual zone removal also failed: $_" "WARNING"
                    Write-LogMessage "Wizmo may require manual unblocking in Properties dialog" "WARNING"
                }
            }
            
            return $wizmoPath
        } else {
            # FAILURE DETECTION: File exists but is suspiciously small
            throw "Downloaded file appears to be invalid"
        }
    } catch {
        # ERROR HANDLING: Log the failure but don't crash the entire script
        Write-LogMessage "Failed to download Wizmo: $_" "ERROR"
        Write-LogMessage "Will use standard Windows reboot as fallback" "WARNING"
        return $null  # Return null to signal fallback needed
    }
}

# FUNCTION: Silent Reboot with Wizmo Integration
# PURPOSE: Perform system restart without any user prompts or delays
# WHY WIZMO FIRST: Wizmo provides completely silent operation, Windows tools show notifications
function Invoke-SilentReboot {
    # TOOL ACQUISITION: Ensure Wizmo is available for silent reboot
    $wizmoPath = Confirm-WizmoAvailability
    
    Write-LogMessage "Initiating silent system reboot..."
    
    # PRIMARY METHOD: Use Wizmo for completely silent operation
    if ($wizmoPath -and (Test-Path $wizmoPath)) {
        try {
            Write-LogMessage "Using Wizmo for silent reboot"
            # WIZMO EXECUTION: Use correct syntax for truly silent reboot
            # WHY "quiet reboot!": The "quiet" parameter suppresses all notifications and dialogs
            & $wizmoPath quiet reboot!
            # BRIEF PAUSE: Allow the reboot command to initiate before script continues
            Start-Sleep -Seconds 5
        } catch {
            # GRACEFUL DEGRADATION: Fall back to Windows built-in if Wizmo fails
            Write-LogMessage "Wizmo reboot failed, using Windows reboot: $_" "WARNING"
            # SHUTDOWN.EXE: Built-in Windows tool
            # /r = restart, /t 0 = no delay, /f = force close applications
            shutdown.exe /r /t 0 /f
        }
    } else {
        # FALLBACK METHOD: Use standard Windows shutdown when Wizmo unavailable
        Write-LogMessage "Wizmo not available, using Windows reboot" "WARNING"
        shutdown.exe /r /t 0 /f
    }
}