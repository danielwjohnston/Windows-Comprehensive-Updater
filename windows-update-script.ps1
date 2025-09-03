# ============================================================================
# FULLY AUTOMATED WINDOWS UPDATE SCRIPT WITH SELF-DEPLOYMENT
<<<<<<< HEAD
# Script Version: 2.1.0 - Enhanced Bulletproof Edition
# ============================================================================

=======
# ============================================================================
>>>>>>> 812fb49 (Implement code changes to enhance functionality and improve performance)
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
<<<<<<< HEAD
            
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
            
=======
>>>>>>> 812fb49 (Implement code changes to enhance functionality and improve performance)
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
<<<<<<< HEAD
            # WIZMO EXECUTION: Use correct syntax for truly silent reboot
            # WHY "quiet reboot!": The "quiet" parameter suppresses all notifications and dialogs
            & $wizmoPath quiet reboot!
=======
            # WIZMO EXECUTION: Simple command execution - no parameters needed
            # WHY & OPERATOR: Proper way to execute external programs in PowerShell
            & $wizmoPath reboot
>>>>>>> 812fb49 (Implement code changes to enhance functionality and improve performance)
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

<<<<<<< HEAD
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
=======
# ============================================================================
# MAIN SCRIPT EXECUTION BEGINS HERE
# ============================================================================

# PRIVILEGE CHECK: Ensure script runs with necessary permissions
# WHY FIRST: No point in continuing without proper privileges
Confirm-RunAsAdmin

# GLOBAL VARIABLE: Centralized log file location
# WHY GLOBAL: Multiple functions need access, and it ensures consistency
# WHY C:\ ROOT: Accessible regardless of user profile, survives user switches
$global:logFile = "C:\WindowsUpdateLog.txt"

# EVENT VIEWER INTEGRATION: Register custom event source for enterprise logging
# WHY EVENT SOURCE: Allows filtering and monitoring in Event Viewer
# REGISTRATION: Must be done early before any event logging attempts
try {
    # CHECK EXISTING: See if our event source is already registered
    if (-not [System.Diagnostics.EventLog]::SourceExists("WindowsUpdateScript")) {
        Write-Host "Registering Windows Event Log source for enterprise monitoring..." -ForegroundColor Cyan
        # NEW-EVENTLOG: Register our custom event source
        # -LogName Application: Use standard Windows Application log
        # -Source: Our unique identifier for filtering events
        New-EventLog -LogName Application -Source "WindowsUpdateScript"
        Write-Host "Event log source registered successfully" -ForegroundColor Green
    }
} catch {
    # REGISTRATION FAILURE: Non-critical, script can continue without Event Viewer integration
    Write-Host "Event log source registration failed (script will continue): $_" -ForegroundColor Yellow
}

# DASHBOARD LAUNCH: Start monitoring dashboard if requested
if ($ShowDashboard -or $DashboardPath) {
    Write-LogMessage "Dashboard requested - attempting to launch..."
    $dashboardLaunched = Start-UpdateDashboard -HtmlPath $DashboardPath
    if ($dashboardLaunched) {
        Write-LogMessage "Dashboard launched successfully - real-time monitoring available"
        # INITIAL STATUS UPDATE: Set dashboard to initialization phase
        Update-DashboardStatus -Phase "initialization" -Progress 5 -CurrentOperation "Script starting..."
    }
} else {
    Write-LogMessage "Running without dashboard - use -ShowDashboard parameter for visual monitoring"
}

# EXECUTION POLICY: Temporarily bypass script execution restrictions
# WHY NECESSARY: Many systems have restrictive policies that prevent script execution
# SCOPE PROCESS: Only affects this PowerShell session, doesn't change system settings
try {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
    Write-LogMessage "Execution policy set to Bypass for this session"
} catch {
    # NON-FATAL ERROR: Script might still work even if this fails
    Write-LogMessage "Could not set execution policy: $_" "WARNING"
}

# ============================================================================
# WINGET UPDATES: Update core Windows applications before Windows Updates
# ============================================================================
# WHY FIRST: Updating PowerShell, Terminal, and App Installer ensures we have
# the latest tools available for the Windows Update process
# STRATEGY: Use Winget to update Microsoft's core applications silently
# ERROR HANDLING: Multi-pass approach with escalating remediation strategies

# WINGET AVAILABILITY CHECK: Ensure Winget (App Installer) is available
# WINGET: Microsoft's official package manager for Windows
function Update-CoreAppsWithWinget {
    Write-LogMessage "=========================================="
    Write-LogMessage "PHASE 1: Updating Core Windows Applications with Winget"
    Write-LogMessage "=========================================="
    
    try {
        # WINGET DETECTION: Check if winget command is available
        $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
        
        if (-not $wingetPath) {
            Write-LogMessage "Winget not found - attempting to install App Installer..." "WARNING"
            
            # APP INSTALLER INSTALLATION: Download and install latest App Installer
            # WHY NEEDED: Winget comes with App Installer package
            try {
                # DOWNLOAD URL: Official Microsoft Store link for App Installer
                $appInstallerUrl = "https://aka.ms/getwinget"
                # TEMP PATH: Use proper path construction to avoid colon parsing issues
                $tempPath = Join-Path $env:TEMP "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
                
                Write-LogMessage "Downloading App Installer from Microsoft..."
                # WEBCLIENT DOWNLOAD: Reliable download method
                $webClient = New-Object System.Net.WebClient
                $webClient.Headers.Add("User-Agent", "PowerShell Windows Update Script")
                $webClient.DownloadFile($appInstallerUrl, $tempPath)
                
                Write-LogMessage "Installing App Installer..."
                # ADD-APPXPACKAGE: Install MSIX package
                Add-AppxPackage -Path $tempPath -ForceApplicationShutdown
                
                # CLEANUP: Remove temporary download
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                
                Write-LogMessage "App Installer installation completed"
                Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1010 -EntryType Information -Message "App Installer installed successfully"
                
                # PATH REFRESH: Update PATH environment variable
                $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
                
            } catch {
                Write-LogMessage "Failed to install App Installer: $_" "WARNING"
                Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3010 -EntryType Warning -Message "Failed to install App Installer: $_"
                return $false
            }
        
        # VERIFY WINGET: Final check that winget is now available
        $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
        if (-not $wingetPath) {
            Write-LogMessage "Winget still not available after installation attempt" "WARNING"
            return $false
        }
        
        Write-LogMessage "Winget detected at: $($wingetPath.Source)"
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1011 -EntryType Information -Message "Winget available at: $($wingetPath.Source)"
        
        # CORE APPLICATIONS TO UPDATE: Essential Microsoft applications with conflict handling
        # WHY THESE: PowerShell for scripting, Terminal for interface, App Installer for package management
        # CONFLICT AWARENESS: Some apps can't be updated while running (PowerShell, active Terminal)
        $coreApps = @(
            @{ 
                Name = "Microsoft.PowerShell"
                DisplayName = "PowerShell 7+"
                ConflictRisk = $true  # Can't update PowerShell while PowerShell script is running
                RemediationStrategies = @("external-process", "repair", "uninstall-reinstall")
            },
            @{ 
                Name = "Microsoft.WindowsTerminal"
                DisplayName = "Windows Terminal"
                ConflictRisk = $true  # Can't update if Terminal is currently running
                RemediationStrategies = @("external-process", "reset-app", "uninstall-reinstall")
            },
            @{ 
                Name = "Microsoft.DesktopAppInstaller"
                DisplayName = "App Installer (Winget)"
                ConflictRisk = $false  # Usually safe to update
                RemediationStrategies = @("repair", "reset-winget", "reinstall-appinstaller")
            }
        )
        
        Write-LogMessage "Starting intelligent update process with conflict resolution..."
        Update-DashboardStatus -Phase "winget" -Progress 20 -CurrentOperation "Analyzing application conflicts..." -AdditionalData @{wingetApps = 0}
        
        # CONFLICT DETECTION: Identify apps that need special handling
        $conflictingApps = $coreApps | Where-Object { $_.ConflictRisk -eq $true }
        $nonConflictingApps = $coreApps | Where-Object { $_.ConflictRisk -eq $false }
        
        if ($conflictingApps.Count -gt 0) {
            Write-LogMessage "Found $($conflictingApps.Count) applications requiring conflict resolution"
            Write-LogMessage "Conflicting apps: $($conflictingApps.DisplayName -join ', ')"
            
            # HANDLE CONFLICTING APPS FIRST: Use external processes for PowerShell/Terminal
            Update-DashboardStatus -Phase "winget" -Progress 25 -CurrentOperation "Resolving application conflicts..."
            Update-ConflictingApplications -ConflictingApps $conflictingApps
        }
        
        # HANDLE NON-CONFLICTING APPS: Standard PowerShell winget execution
        if ($nonConflictingApps.Count -gt 0) {
            Write-LogMessage "Processing $($nonConflictingApps.Count) non-conflicting applications..."
            Update-DashboardStatus -Phase "winget" -Progress 35 -CurrentOperation "Updating non-conflicting applications..."
            
            foreach ($app in $nonConflictingApps) {
                Write-LogMessage "Updating $($app.DisplayName)..."
                try {
                    $null = & winget upgrade $app.Name --silent --accept-package-agreements --accept-source-agreements 2>&1
                    
                    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335212) {
                        Write-LogMessage "$($app.DisplayName) updated successfully"
                        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1012 -EntryType Information -Message "$($app.DisplayName) updated successfully"
                    } else {
                        Write-LogMessage "$($app.DisplayName) update completed with exit code: $LASTEXITCODE" "WARNING"
                    }
                } catch {
                    Write-LogMessage "Failed to update $($app.DisplayName): $_" "WARNING"
                }
            }
        }
        
        # SKIP MULTI-PASS LOGIC: Since we handled conflicts already, simplified approach
        Write-LogMessage "Core applications update phase completed with conflict resolution"
        
        # MULTI-PASS UPDATE STRATEGY: Up to 3 passes with escalating remediation
        # PASS 1: Normal update attempts
        # PASS 2: Repair and retry failed applications
        # PASS 3: Uninstall/reinstall for stubborn failures
        $maxPasses = 3
        $failedApps = @()
        
        for ($pass = 1; $pass -le $maxPasses; $pass++) {
            Write-LogMessage "--- Winget Update Pass $pass of $maxPasses ---"
            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1015 -EntryType Information -Message "Starting Winget update pass $pass of $maxPasses"
            Update-DashboardStatus -Phase "winget" -Progress (15 + ($pass * 10)) -CurrentOperation "Winget Pass $pass of $maxPasses"
            
            # DETERMINE APPS TO PROCESS: All apps on first pass, only failed apps on subsequent passes
            $appsToProcess = if ($pass -eq 1) { $coreApps } else { $failedApps }
            $currentPassFailed = @()
            
            if ($appsToProcess.Count -eq 0) {
                Write-LogMessage "No applications require processing in pass $pass"
                break
            }
            
            foreach ($app in $appsToProcess) {
                Write-LogMessage "Processing $($app.DisplayName) in pass $pass..."
                $updateSuccessful = $false
                
                try {
                    # PASS-SPECIFIC STRATEGY: Different approaches based on pass number
                    switch ($pass) {
                        1 {
                            # PASS 1: Standard update attempt
                            Write-LogMessage "Attempting standard update for $($app.DisplayName)..."
                            # Execute winget upgrade and log the output
                            $wingetOutput = & winget upgrade $app.Name --silent --accept-package-agreements --accept-source-agreements 2>&1
                            # Log the output for troubleshooting
                            if ($wingetOutput) {
                                Write-LogMessage "Winget output: $($wingetOutput -join '; ')" "INFO"
                            }
                        }
                        
                        2 {
                            # PASS 2: Repair-based remediation
                            Write-LogMessage "Attempting repair-based remediation for $($app.DisplayName)..."
                            
                            # REMEDIATION STRATEGY: Try repair first, then update
                            switch ($app.Name) {
                                "Microsoft.PowerShell" {
                                    Write-LogMessage "Attempting PowerShell repair via winget..."
                                    & winget install $app.Name --force --silent --accept-package-agreements --accept-source-agreements 2>&1 | Out-Null
                                }
                                
                                "Microsoft.WindowsTerminal" {
                                    Write-LogMessage "Resetting Windows Terminal app data..."
                                    # RESET TERMINAL: Clear app data to resolve configuration conflicts
                                    try {
                                        Get-AppxPackage Microsoft.WindowsTerminal | Reset-AppxPackage
                                        Start-Sleep -Seconds 5
                                    } catch {
                                        Write-LogMessage "Terminal reset failed: $_" "WARNING"
                                    }
                                }
                                
                                "Microsoft.DesktopAppInstaller" {
                                    Write-LogMessage "Resetting Winget configuration..."
                                    # WINGET RESET: Clear winget settings and cache
                                    try {
                                        & winget settings --enable LocalManifestFiles 2>&1 | Out-Null
                                        # SAFE PATH CONSTRUCTION: Avoid colon parsing issues with wildcard paths
                                        $appInstallerPath = Join-Path $env:LOCALAPPDATA "Packages"
                                        $appInstallerPath = Join-Path $appInstallerPath "Microsoft.DesktopAppInstaller*"
                                        $localStatePath = Join-Path $appInstallerPath "LocalState"
                                        Remove-Item "$localStatePath\*" -Recurse -Force -ErrorAction SilentlyContinue
                                        Start-Sleep -Seconds 3
                                    } catch {
                                        Write-LogMessage "Winget reset failed: $_" "WARNING"
                                    }
                                }
                            }
                            
                            # RETRY UPDATE: After remediation attempt
                            Start-Sleep -Seconds 5
                            $null = & winget upgrade $app.Name --silent --accept-package-agreements --accept-source-agreements 2>&1
                        }
                        
                        3 {
                            # PASS 3: Nuclear option - uninstall and reinstall
                            Write-LogMessage "Attempting uninstall/reinstall for $($app.DisplayName)..."
                            
                            try {
                                # UNINSTALL FIRST: Remove existing problematic installation
                                Write-LogMessage "Uninstalling $($app.DisplayName)..."
                                & winget uninstall $app.Name --silent --accept-source-agreements 2>&1 | Out-Null
                                Start-Sleep -Seconds 10
                                
                                # CLEAN REINSTALL: Fresh installation
                                Write-LogMessage "Reinstalling $($app.DisplayName)..."
                                $null = & winget install $app.Name --silent --accept-package-agreements --accept-source-agreements 2>&1
                                
                            } catch {
                                Write-LogMessage "Uninstall/reinstall failed for $($app.DisplayName): $_" "ERROR"
                            }
                        }
                    }
                    
                    # RESULT ANALYSIS: Comprehensive exit code interpretation
                    # WINGET EXIT CODES: Microsoft's documented return codes
                    switch ($LASTEXITCODE) {
                        0 {
                            # SUCCESS: Update completed successfully
                            Write-LogMessage "$($app.DisplayName) updated successfully in pass $pass"
                            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1012 -EntryType Information -Message "$($app.DisplayName) updated successfully in pass $pass"
                            $updateSuccessful = $true
                        }
                        
                        -1978335212 {
                            # NO UPDATE AVAILABLE: Already up to date
                            Write-LogMessage "$($app.DisplayName) is already up to date"
                            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1013 -EntryType Information -Message "$($app.DisplayName) is already up to date"
                            $updateSuccessful = $true
                        }
                        
                        -1978335215 {
                            # PACKAGE NOT FOUND: Application not found in repository
                            Write-LogMessage "$($app.DisplayName) not found in winget repository" "WARNING"
                            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3014 -EntryType Warning -Message "$($app.DisplayName) not found in winget repository"
                            # DON'T RETRY: If package doesn't exist, retrying won't help
                            $updateSuccessful = $true  # Mark as "successful" to prevent retries
                        }
                        
                        -1978335216 {
                            # MULTIPLE PACKAGES FOUND: Ambiguous package name
                            Write-LogMessage "$($app.DisplayName) package name is ambiguous, trying with source specification..." "WARNING"
                            try {
                                # RETRY WITH SOURCE: Specify Microsoft Store as source
                                $retryOutput = & winget upgrade $app.Name --source msstore --silent --accept-package-agreements --accept-source-agreements 2>&1
                                if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335212) {
                                    $updateSuccessful = $true
                                    Write-LogMessage "$($app.DisplayName) updated successfully with source specification"
                                    # LOG OUTPUT: Capture successful retry output
                                    if ($retryOutput) {
                                        Write-LogMessage "Source retry output: $($retryOutput -join '; ')" "INFO"
                                    }
                                }
                            } catch {
                                Write-LogMessage "Source specification retry failed: $_" "WARNING"
                            }
                        }
                        
                        -1978335222 {
                            # INSTALLER FAILED: The installer itself failed
                            Write-LogMessage "$($app.DisplayName) installer failed (will retry in next pass if available)" "WARNING"
                            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3015 -EntryType Warning -Message "$($app.DisplayName) installer failed in pass $pass, exit code: $LASTEXITCODE"
                        }
                        
                        -1978335213 {
                            # PACKAGE IN USE: Application currently running
                            Write-LogMessage "$($app.DisplayName) is currently in use, attempting to close and retry..." "WARNING"
                            
                                    # PROCESS TERMINATION: Try to close the application gracefully
                                    try {
                                        switch ($app.Name) {
                                            "Microsoft.PowerShell" {
                                                Get-Process pwsh -ErrorAction SilentlyContinue | Stop-Process -Force
                                            }
                                            "Microsoft.WindowsTerminal" {
                                                Get-Process WindowsTerminal -ErrorAction SilentlyContinue | Stop-Process -Force
                                            }
                                        }
                                        Start-Sleep -Seconds 5
                                
                                        # RETRY AFTER CLOSING: Attempt update again
                                        $retryAfterCloseOutput = & winget upgrade $app.Name --silent --accept-package-agreements --accept-source-agreements 2>&1
                                        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335212) {
                                            $updateSuccessful = $true
                                            Write-LogMessage "$($app.DisplayName) updated successfully after closing processes"
                                            # LOG OUTPUT: Capture retry output
                                            if ($retryAfterCloseOutput) {
                                                Write-LogMessage "Process close retry output: $($retryAfterCloseOutput -join '; ')" "INFO"
                                            }
                                        }
                                    } catch {
                                        Write-LogMessage "Process termination and retry failed: $_" "WARNING"
                                    }
                        }
                        
                        default {
                            # UNKNOWN ERROR: Log the exit code for troubleshooting
                            Write-LogMessage "$($app.DisplayName) update failed with exit code: $LASTEXITCODE (pass $pass)" "WARNING"
                            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3016 -EntryType Warning -Message "$($app.DisplayName) update failed in pass $pass with exit code: $LASTEXITCODE"
                        }
                    }
                    
                } catch {
                    # EXCEPTION HANDLING: PowerShell execution errors
                    # VARIABLE SAFETY: Proper string interpolation to avoid colon parsing issues
                    Write-LogMessage "Exception during $($app.DisplayName) update in pass $($pass): $($_)" "ERROR"
                    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5015 -EntryType Error -Message ("Exception during {0} update in pass {1}: {2}" -f $app.DisplayName, $pass, $_)
                }
                
                # FAILURE TRACKING: Add to failed list if update unsuccessful and more passes available
                if (-not $updateSuccessful -and $pass -lt $maxPasses) {
                    $currentPassFailed += $app
                    Write-LogMessage "$($app.DisplayName) will be retried in pass $($pass + 1)"
                } elseif (-not $updateSuccessful) {
                    Write-LogMessage "$($app.DisplayName) failed all remediation attempts" "ERROR"
                    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5016 -EntryType Error -Message "$($app.DisplayName) failed all $maxPasses remediation attempts"
                }
            }
            
            # UPDATE FAILED LIST: Prepare for next pass
            $failedApps = $currentPassFailed
            
            # PASS COMPLETION: Log pass results
            if ($failedApps.Count -eq 0) {
                Write-LogMessage "All applications updated successfully in pass $pass"
                Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1016 -EntryType Information -Message "All applications updated successfully in pass $pass"
                break
            } else {
                Write-LogMessage "Pass $pass completed with $($failedApps.Count) applications requiring retry"
                Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3017 -EntryType Warning -Message "Pass $pass completed with $($failedApps.Count) applications requiring retry"
            }
            
            # INTER-PASS DELAY: Allow system to stabilize between passes
            if ($pass -lt $maxPasses -and $failedApps.Count -gt 0) {
                Write-LogMessage "Waiting 30 seconds before next remediation pass..."
                Start-Sleep -Seconds 30
            }
        }
        
        # WINGET SOURCE UPDATE: Refresh package source information
        # WHY NEEDED: Ensures we have the latest package information
        try {
            Write-LogMessage "Updating Winget source information..."
            & winget source update --silent 2>&1 | Out-Null
            Write-LogMessage "Winget sources updated successfully"
            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1014 -EntryType Information -Message "Winget sources updated successfully"
        } catch {
            Write-LogMessage "Winget source update: $_" "WARNING"
            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3013 -EntryType Warning -Message "Winget source update failed: $_"
        }
        
        # FINAL RESULTS: Summary of multi-pass update process
        $finalFailedCount = $failedApps.Count
        if ($finalFailedCount -eq 0) {
            Write-LogMessage "All core applications successfully updated with multi-pass remediation"
            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1017 -EntryType Information -Message "All core applications successfully updated using multi-pass remediation strategy"
            return $true
        } else {
            Write-LogMessage "$finalFailedCount applications failed all remediation attempts" "WARNING"
            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3018 -EntryType Warning -Message "$finalFailedCount applications failed all remediation attempts after $maxPasses passes"
            return $true  # Continue with Windows Updates even if some Winget updates failed
        }
        
    } catch {
        Write-LogMessage "Winget update process failed: $_" "ERROR"
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5010 -EntryType Error -Message "Winget update process failed: $_"
        return $false
    }

try {
    $wingetSuccess = Update-CoreAppsWithWinget
} catch {
    Write-LogMessage "Error executing winget updates: $_" "ERROR"
    $wingetSuccess = $false
}

if ($wingetSuccess) {
    Write-LogMessage "Core applications update phase completed successfully"
} else {
    Write-LogMessage "Winget updates completed with some issues - continuing with Windows Updates" "WARNING"
}

Write-LogMessage "=========================================="
Write-LogMessage "PHASE 2: Beginning Windows Update Process"
Write-LogMessage "=========================================="
Update-DashboardStatus -Phase "windows" -Progress 50 -CurrentOperation "Beginning Windows Update scan..."

# POWERSHELL GALLERY TRUST CONFIGURATION: Enable automatic module installation
# WHY NEEDED: First-time users get security prompts when installing from PowerShell Gallery
# SOLUTION: Set the repository as trusted to avoid interactive prompts
try {
    Write-LogMessage "Configuring PowerShell Gallery as trusted repository..."
    # SET-PSREPOSITORY: Configure PowerShell Gallery trust level
    # -Name "PSGallery": The default PowerShell module repository
    # -InstallationPolicy Trusted: Allow automatic installation without prompts
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1001 -EntryType Information -Message "PowerShell Gallery configured as trusted repository"
} catch {
    Write-LogMessage "PowerShell Gallery configuration: $_" "WARNING"
    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3001 -EntryType Warning -Message "PowerShell Gallery trust configuration failed: $_"
}

# NUGET PROVIDER INSTALLATION: Required for PowerShell module installation
# WHY NUGET: PowerShell Gallery (where PSWindowsUpdate lives) requires NuGet provider
# FIRST-TIME USER HANDLING: NuGet provider prompts for confirmation on first install
# SOLUTION: Use -Force and additional trust parameters to handle virgin systems
try {
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-LogMessage "Installing NuGet provider (first-time setup)..."
        # COMPREHENSIVE INSTALLATION PARAMETERS:
        # -Force: Skip all confirmation prompts
        # -ForceBootstrap: Install even if bootstrapping is required
        # -Confirm disabled: Suppress any remaining confirmation dialogs
        # -Scope AllUsers: Install for all users (requires admin, which we have)
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ForceBootstrap -Confirm:$false -Scope AllUsers | Out-Null
        Write-LogMessage "NuGet provider installed successfully (first-time configuration completed)"
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1002 -EntryType Information -Message "NuGet provider installed successfully for first-time user"
    } else {
        Write-LogMessage "NuGet provider already available"
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1003 -EntryType Information -Message "NuGet provider already available"
    }
} catch {
    # CRITICAL ERROR: Can't continue without NuGet provider
    Write-LogMessage "Failed to install NuGet provider: $_" "ERROR"
    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5001 -EntryType Error -Message "Critical failure: NuGet provider installation failed: $_"
    Exit 1
}

# PSWINDOWSUPDATE MODULE: Core functionality for Windows Update management
# WHY THIS MODULE: Provides comprehensive PowerShell cmdlets for update management
# Much more reliable and feature-complete than WUA API directly
# FIRST-TIME INSTALLATION: Handle virgin PowerShell installations with full automation
try {
    # AVAILABILITY CHECK: See if module is already installed
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-LogMessage "Installing PSWindowsUpdate module (with full automation for first-time users)..."
        # COMPREHENSIVE INSTALLATION PARAMETERS FOR VIRGIN SYSTEMS:
        # -Force: Overwrite any existing versions
        # -SkipPublisherCheck: Don't verify digital signatures (module is trusted)
        # -Confirm disabled: Suppress all confirmation prompts
        # -AllowClobber: Overwrite any conflicting commands from other modules
        # -Scope AllUsers: Install for all users system-wide
        # -Repository PSGallery: Explicitly specify the source (now trusted)
        Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -Confirm:$false -AllowClobber -Scope AllUsers -Repository PSGallery | Out-Null
        Write-LogMessage "PSWindowsUpdate module installed successfully"
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1004 -EntryType Information -Message "PSWindowsUpdate module installed successfully"
    }
    
    # IMPORT CHECK: Make sure module is loaded into current session
    if (-not (Get-Module -Name PSWindowsUpdate)) {
        Write-LogMessage "Importing PSWindowsUpdate module..."
        # IMPORT WITH FORCE: Reload module even if already loaded (ensures latest version)
        Import-Module PSWindowsUpdate -Force | Out-Null
        Write-LogMessage "PSWindowsUpdate module imported successfully"
        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1005 -EntryType Information -Message "PSWindowsUpdate module imported successfully"
    }
} catch {
    # CRITICAL ERROR: Cannot perform updates without this module
    Write-LogMessage "Failed to install/import PSWindowsUpdate module: $_" "ERROR"
    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 5002 -EntryType Error -Message "Critical failure: PSWindowsUpdate module installation/import failed: $_"
    Exit 1
}

# MICROSOFT UPDATE SERVICE: Enable updates for Microsoft products (not just Windows)
# WHY NEEDED: By default, Windows Update only updates Windows itself
# This enables Office, .NET, SQL Server, and other Microsoft product updates
try {
    Write-LogMessage "Configuring Microsoft Update Service..."
    # ADD-WUSERVICEMANAGER: Registers Microsoft Update as an update source
    # -MicrosoftUpdate: Specifically adds the Microsoft Update service
    # -Confirm disabled: Suppress confirmation prompts for automation
    Add-WUServiceManager -MicrosoftUpdate -Confirm:$false | Out-Null
    Write-LogMessage "Microsoft Update Service configured"
} catch {
    # NON-CRITICAL: Update process can continue with just Windows updates
    Write-LogMessage "Microsoft Update Service configuration: $_" "WARNING"
}

# WSUS BYPASS CONFIGURATION: Handle corporate WSUS environments
# WHY NEEDED: Many corporate environments use WSUS servers that may not have
# all updates available, or may have delayed update approval processes
# STRATEGY: Temporarily bypass WSUS, then restore original configuration
$wsusBypassed = $false  # Flag to track if we need to restore settings later
$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

try {
    # REGISTRY CHECK: See if WSUS is configured via Group Policy
    if (Test-Path $registryPath) {
        # VALUE RETRIEVAL: Check current UseWUServer setting
        # UseWUServer = 1 means "use WSUS server", 0 means "use Microsoft Update"
        $currentWSUSSetting = Get-ItemProperty -Path $registryPath -Name "UseWUServer" -ErrorAction SilentlyContinue
        
        if ($currentWSUSSetting -and $currentWSUSSetting.UseWUServer -eq 1) {
            Write-LogMessage "Temporarily bypassing WSUS configuration..."
            
            # TEMPORARY BYPASS: Set to 0 to use Microsoft Update directly
            Set-ItemProperty -Path $registryPath -Name "UseWUServer" -Value 0
            $wsusBypassed = $true  # Remember that we changed this
            
            # SERVICE RESTART: Windows Update service must be restarted to recognize change
            Write-LogMessage "Restarting Windows Update service..."
            Restart-Service -Name wuauserv -Force | Out-Null
            
            # STABILIZATION DELAY: Allow service to fully restart and initialize
            Start-Sleep -Seconds 10
            Write-LogMessage "Windows Update service restarted"
        }
    }
} catch {
    # NON-CRITICAL: Updates might still work even if WSUS bypass fails
    Write-LogMessage "WSUS configuration handling: $_" "WARNING"
}

# WIZMO PREPARATION: Download and verify Wizmo for potential reboots
# WHY EARLY: Better to download now while internet is available than during reboot cycle
$wizmoReady = Confirm-WizmoAvailability
Write-LogMessage "Wizmo status: $(if ($wizmoReady) {'Ready for silent reboots'} else {'Will use fallback reboot method'})"

# ============================================================================
# MAIN UPDATE INSTALLATION LOOP
# ============================================================================

# LOOP PARAMETERS: Prevent infinite loops while allowing multiple update cycles
# WHY MULTIPLE CYCLES: Some updates only become available after others are installed
# WHY LIMIT CYCLES: Prevent infinite loops in case of persistent issues
$maxCycles = 10           # Maximum number of update scan/install cycles
$cycleCount = 0           # Current cycle counter
$totalUpdatesInstalled = 0 # Running total of updates installed

Write-LogMessage "Starting automated update installation process..."

# MAIN PROCESSING LOOP: Continue until no more updates available or max cycles reached
do {
    $cycleCount++
    $updatesThisCycle = 0  # Reset counter for this cycle
    
    Write-LogMessage "--- Update Cycle $cycleCount of $maxCycles ---"
    Update-DashboardStatus -Phase "windows" -Progress (50 + ($cycleCount * 5)) -CurrentOperation "Update Cycle $cycleCount of $maxCycles"
    
    # UPDATE SCANNING: Check for available updates
    try {
        Write-LogMessage "Scanning for available updates..."
        
        # GET-WULIST: Retrieve list of available updates
        # -MicrosoftUpdate: Include Microsoft products, not just Windows
        # -IgnoreReboot: Don't exclude updates that require reboot
        $availableUpdates = Get-WUList -MicrosoftUpdate -IgnoreReboot
        
        # AVAILABILITY CHECK: Determine if updates were found
        if ($availableUpdates -and $availableUpdates.Count -gt 0) {
            Write-LogMessage "Found $($availableUpdates.Count) available updates"
            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 2001 -EntryType Information -Message "Found $($availableUpdates.Count) available updates in cycle $cycleCount"
            
            # UPDATE ENUMERATION: Log details of each update for audit trail
            $availableUpdates | ForEach-Object {
                # SIZE CALCULATION: Convert bytes to MB for human readability
                $sizeMB = [math]::Round($_.Size / 1MB, 2)
                $updateDetails = "  - $($_.Title) (Size: $sizeMB MB)"
                Write-LogMessage $updateDetails
                # DETAILED EVENT LOGGING: Individual update information
                Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 2002 -EntryType Information -Message "Available Update: $($_.Title), Size: $sizeMB MB, KB: $($_.KBArticleIDs -join ',')"
            }
            
            $updatesThisCycle = $availableUpdates.Count
        } else {
            # NO UPDATES: Exit the loop as update process is complete
            Write-LogMessage "No updates found in this cycle"
            break
        }
    } catch {
        # SCAN FAILURE: Log error and exit loop
        Write-LogMessage "Failed to scan for updates: $_" "ERROR"
        break
    }
    
    # UPDATE INSTALLATION: Install all available updates
    if ($updatesThisCycle -gt 0) {
        try {
            Write-LogMessage "Installing $updatesThisCycle updates automatically..."
        Update-DashboardStatus -Phase "windows" -Progress (55 + ($cycleCount * 5)) -CurrentOperation "Installing $updatesThisCycle updates..."
            
            # INSTALL-WUUPDATES: Core installation command
            # -MicrosoftUpdate: Include Microsoft products
            # -AcceptAll: Accept all EULAs and terms automatically
            # -IgnoreReboot: Don't automatically reboot (we'll handle this ourselves)
            # -Confirm disabled: Suppress all confirmation prompts
            # -Verbose disabled: Reduce output noise for cleaner logs
            $installOutput = Install-WUUpdates -MicrosoftUpdate -AcceptAll -IgnoreReboot -Confirm:$false -Verbose:$false
            
            # LOG INSTALLATION RESULTS: Capture and log update installation details
            if ($installOutput) {
                Write-LogMessage "Update installation details logged"
                # LOG INDIVIDUAL UPDATES: Record each installed update
                $installOutput | ForEach-Object {
                    if ($_.Title) {
                        Write-LogMessage "Installed: $($_.Title) - Status: $($_.Result)" "INFO"
                        Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 2010 -EntryType Information -Message "Successfully installed update: $($_.Title)"
                    }
                }
            }
            
            # PROGRESS TRACKING: Update running total
            $totalUpdatesInstalled += $updatesThisCycle
            Write-LogMessage "Updates installed successfully (Total so far: $totalUpdatesInstalled)"
            Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 2003 -EntryType Information -Message "Installed $updatesThisCycle updates in cycle $cycleCount. Total installed: $totalUpdatesInstalled"
            
            # REBOOT ASSESSMENT: Check if system restart is required
            # GET-WUREBOOTSTATUS: Determines if pending updates require restart
            # -Silent: Don't display any UI, just return true/false
            $rebootRequired = Get-WURebootStatus -Silent
            
            # REBOOT HANDLING: Manage system restart for update completion
            if ($rebootRequired) {
                Write-LogMessage "System reboot required - initiating silent reboot..."
                Write-LogMessage "Updates will continue after reboot..."
                Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 2004 -EntryType Information -Message "System reboot required after installing $updatesThisCycle updates. Initiating silent reboot and continuation."
                
                # SCHEDULED TASK CREATION: Ensure script continues after reboot
                # WHY NEEDED: PowerShell scripts don't automatically resume after reboot
                # SOLUTION: Create a scheduled task that runs at startup with the same script
                
                # TASK COMPONENTS:
                # Action: What to run (this same PowerShell script)
                $taskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
                
                # Trigger: When to run (at system startup)
                $taskTrigger = New-ScheduledTaskTrigger -AtStartup
                
                # Principal: Run as SYSTEM account with highest privileges
                # WHY SYSTEM: Ensures task runs even if no user is logged in
                # RunLevel Highest: Equivalent to "Run as Administrator"
                $taskPrincipal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                
                # Settings: Task behavior configuration
                # AllowStartIfOnBatteries: Run even on laptop battery power
                # DontStopIfGoingOnBatteries: Continue even if switching to battery
                # StartWhenAvailable: Run as soon as possible if startup time is missed
                $taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
                
                try {
                    # TASK REGISTRATION: Create the scheduled task
                    # -Force: Overwrite any existing task with same name
                    Register-ScheduledTask -TaskName "WindowsUpdateContinuation" -Action $taskAction -Trigger $taskTrigger -Principal $taskPrincipal -Settings $taskSettings -Force | Out-Null
                    Write-LogMessage "Scheduled task created for post-reboot continuation"
                    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 2005 -EntryType Information -Message "Scheduled task 'WindowsUpdateContinuation' created for post-reboot continuation"
                } catch {
                    # TASK CREATION FAILURE: Non-critical, but user might need to manually restart
                    Write-LogMessage "Could not create scheduled task: $_" "WARNING"
                    Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 3020 -EntryType Warning -Message "Failed to create scheduled task for continuation: $_"
                }
                
                # SILENT REBOOT EXECUTION: Restart system without user interaction
                Invoke-SilentReboot
                # SCRIPT TERMINATION: Script ends here, resumes after reboot via scheduled task
                Exit 0
            }
            
        } catch {
            # INSTALLATION FAILURE: Log error and exit loop
            Write-LogMessage "Failed to install updates: $_" "ERROR"
            break
        }
    }
    
    # CYCLE DELAY: Brief pause between update cycles to allow system stabilization
    Start-Sleep -Seconds 5
    
# LOOP CONTINUATION CONDITIONS: Continue while updates exist and within cycle limit
} while ($updatesThisCycle -gt 0 -and $cycleCount -lt $maxCycles)

# ============================================================================
# CLEANUP AND FINALIZATION
# ============================================================================

# SCHEDULED TASK CLEANUP: Remove the continuation task since updates are complete
# WHY CLEANUP: Prevent the task from running unnecessarily on future reboots
try {
    if (Get-ScheduledTask -TaskName "WindowsUpdateContinuation" -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName "WindowsUpdateContinuation" -Confirm:$false
        Write-LogMessage "Cleanup: Removed scheduled continuation task"
    }
} catch {
    # NON-CRITICAL: Task removal failure won't affect system functionality
    Write-LogMessage "Scheduled task cleanup: $_" "WARNING"
}

# WSUS CONFIGURATION RESTORATION: Restore original enterprise settings
# WHY IMPORTANT: Corporate environments need their WSUS settings maintained
if ($wsusBypassed) {
    try {
        Write-LogMessage "Restoring original WSUS configuration..."
        # RESTORE SETTING: Set UseWUServer back to 1 (use WSUS)
        Set-ItemProperty -Path $registryPath -Name "UseWUServer" -Value 1
        # SERVICE RESTART: Apply the restored configuration
        Restart-Service -Name wuauserv -Force | Out-Null
        Write-LogMessage "WSUS configuration restored"
    } catch {
        # RESTORATION FAILURE: Log warning but don't fail entire script
        Write-LogMessage "WSUS restoration: $_" "WARNING"
    }
}

# FINAL STATUS ASSESSMENT: Comprehensive check of update completion status
try {
    # REMAINING UPDATES CHECK: See if any updates are still available
    $remainingUpdates = Get-WUList -MicrosoftUpdate
    
    # FINAL REBOOT CHECK: Determine if one final reboot is needed
    $finalRebootNeeded = Get-WURebootStatus -Silent
    
    # REMAINING UPDATES ANALYSIS: Report any updates that couldn't be installed
    if ($remainingUpdates -and $remainingUpdates.Count -gt 0) {
        Write-LogMessage "Notice: $($remainingUpdates.Count) updates still available (may require user interaction or be optional)"
        # DETAILED REMAINING LIST: Log each remaining update for administrator review
        $remainingUpdates | ForEach-Object {
            Write-LogMessage "  Remaining: $($_.Title)"
        }
    }
    
    # FINAL REBOOT HANDLING: Some updates require a final reboot to complete
    if ($finalRebootNeeded) {
        Write-LogMessage "Final reboot required to complete all updates"
        Write-LogMessage "Performing final silent reboot..."
        # NO SCHEDULED TASK: This is the final reboot, no continuation needed
        Invoke-SilentReboot
        Exit 0
    }
} catch {
    # STATUS CHECK FAILURE: Non-critical, but log for troubleshooting
    Write-LogMessage "Final status check: $_" "WARNING"
}

# ============================================================================
# SCRIPT COMPLETION AND REPORTING
# ============================================================================

# FINAL STATUS UPDATE: Mark script as completed
Update-DashboardStatus -Phase "completed" -Progress 100 -CurrentOperation "Script completed successfully"

# COMPLETION STATUS FILE: Create completion marker
try {
    $completionFile = Join-Path "C:\Scripts" "update-completed.json"
    $completionData = @{
        completedAt = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
        totalUpdates = $totalUpdatesInstalled
        finalStatus = "success"
        message = "Windows Update process completed successfully"
        nextScheduledRun = "2nd Tuesday of next month at 2:00 AM"
        manualRunCommand = "Start-ScheduledTask -TaskName 'WindowsUpdate-Manual'"
    } | ConvertTo-Json
    
    $completionData | Out-File -FilePath $completionFile -Encoding UTF8
    Write-LogMessage "Completion status file created: $completionFile"
} catch {
    Write-LogMessage "Failed to create completion status file: $_" "WARNING"
}

# SUCCESS SUMMARY: Comprehensive completion report
Write-LogMessage "=========================================="
Write-LogMessage "Automated Windows Update Process Completed"
Write-LogMessage "Total Updates Installed: $totalUpdatesInstalled"
Write-LogMessage "System Status: Up to Date"
Write-LogMessage "Log Location: $global:logFile"
Write-LogMessage "Event Viewer: Check Application Log for 'WindowsUpdateScript' events"
Write-LogMessage "Next Scheduled Run: 2nd Tuesday of next month"
Write-LogMessage "Manual Execution: Start-ScheduledTask -TaskName 'WindowsUpdate-Manual'"
Write-LogMessage "=========================================="

# FINAL EVENT LOG ENTRY: Completion summary for monitoring systems
Write-EventLog -LogName Application -Source "WindowsUpdateScript" -EventId 1999 -EntryType Information -Message "Windows Update Script completed successfully. Total updates installed: $totalUpdatesInstalled. System is now up to date."

# CONSOLE SUMMARY: User-friendly completion message with color coding
Write-Host "`nWindows Update Process Completed Successfully!" -ForegroundColor Green
Write-Host "Total updates installed: $totalUpdatesInstalled" -ForegroundColor Green
Write-Host "Check log file: $global:logFile" -ForegroundColor Cyan
Write-Host "Event Viewer: Application Log -> Source: WindowsUpdateScript" -ForegroundColor Cyan

# CLEAN EXIT: Indicate successful completion to the operating system
Exit 0
>>>>>>> 812fb49 (Implement code changes to enhance functionality and improve performance)
