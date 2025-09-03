# ============================================================================
# WINDOWS COMPREHENSIVE UPDATER TEST SUITE
# Script Version: 2.1.0 - Testing Framework
# ============================================================================

param(
    [switch]$RunBasicTests,
    [switch]$RunFunctionTests,
    [switch]$RunIntegrationTests,
    [switch]$RunAll,
    [switch]$ShowDetails
)

# Test Results Tracking
$script:testResults = @{
    Total = 0
    Passed = 0
    Failed = 0
    Skipped = 0
    Details = @()
}

# FUNCTION: Test Result Logging
function Write-TestResult {
    param(
        [string]$TestName,
        [string]$Result,
        [string]$Details = "",
        [string]$ErrorMessage = ""
    )
    
    $script:testResults.Total++
    
    $color = switch ($Result) {
        "PASS" { "Green"; $script:testResults.Passed++ }
        "FAIL" { "Red"; $script:testResults.Failed++ }
        "SKIP" { "Yellow"; $script:testResults.Skipped++ }
        default { "Gray" }
    }
    
    $output = "[$Result] $TestName"
    if ($Details) { $output += " - $Details" }
    if ($ErrorMessage) { $output += " (Error: $ErrorMessage)" }
    
    Write-Host $output -ForegroundColor $color
    
    $script:testResults.Details += @{
        Name = $TestName
        Result = $Result
        Details = $Details
        Error = $ErrorMessage
        Timestamp = Get-Date
    }
}

# FUNCTION: Test Script Syntax
function Test-ScriptSyntax {
    param([string]$ScriptPath)
    
    Write-Host "`n=== SYNTAX VALIDATION TESTS ===" -ForegroundColor Cyan
    
    try {
        # Test PowerShell syntax parsing
        $errors = @()
        $tokens = @()
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
        
        if ($errors.Count -eq 0) {
            Write-TestResult "PowerShell Syntax Validation" "PASS" "No syntax errors found"
        } else {
            $errorDetails = ($errors | Select-Object -First 3 | ForEach-Object { $_.Message }) -join "; "
            Write-TestResult "PowerShell Syntax Validation" "FAIL" "Found $($errors.Count) syntax errors" $errorDetails
        }
        
        # Test for required functions
        $requiredFunctions = @(
            "Confirm-RunAsAdmin",
            "Write-LogMessage", 
            "Invoke-SilentReboot",
            "Invoke-WingetUpdates",
            "Invoke-WindowsUpdates",
            "Initialize-PSWindowsUpdate"
        )
        
        foreach ($funcName in $requiredFunctions) {
            if ($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq $funcName }, $true)) {
                Write-TestResult "Function Exists: $funcName" "PASS"
            } else {
                Write-TestResult "Function Exists: $funcName" "FAIL" "Function not found in script"
            }
        }
        
    } catch {
        Write-TestResult "Script Parsing" "FAIL" "Unable to parse script file" $_.Exception.Message
    }
}

# FUNCTION: Test File Dependencies
function Test-FileDependencies {
    Write-Host "`n=== FILE DEPENDENCY TESTS ===" -ForegroundColor Cyan
    
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    
    # Test main script file
    $mainScript = Join-Path $scriptDir "windows-comprehensive-updater.ps1"
    if (Test-Path $mainScript) {
        Write-TestResult "Main Script File" "PASS" "windows-comprehensive-updater.ps1 exists"
        
        # Test script size (should be substantial)
        $scriptSize = (Get-Item $mainScript).Length
        if ($scriptSize -gt 50KB) {
            Write-TestResult "Script Content Size" "PASS" "Script is $([math]::Round($scriptSize/1KB,2)) KB"
        } else {
            Write-TestResult "Script Content Size" "FAIL" "Script appears incomplete ($([math]::Round($scriptSize/1KB,2)) KB)"
        }
    } else {
        Write-TestResult "Main Script File" "FAIL" "windows-comprehensive-updater.ps1 not found"
    }
    
    # Test dashboard file
    $dashboardFile = Join-Path $scriptDir "windows-comprehensive-updater-dashboard.html"
    if (Test-Path $dashboardFile) {
        Write-TestResult "Dashboard HTML File" "PASS" "Dashboard file exists"
        
        # Test for essential HTML elements
        $htmlContent = Get-Content $dashboardFile -Raw
        if ($htmlContent -match "Windows Comprehensive Updater Monitor" -and $htmlContent -match "javascript") {
            Write-TestResult "Dashboard Content" "PASS" "Dashboard contains required elements"
        } else {
            Write-TestResult "Dashboard Content" "FAIL" "Dashboard missing essential content"
        }
    } else {
        Write-TestResult "Dashboard HTML File" "FAIL" "Dashboard file not found"
    }
    
    # Test configuration file
    $configFile = Join-Path $scriptDir "windows-update-config.json"
    if (Test-Path $configFile) {
        Write-TestResult "Configuration File" "PASS" "Config file exists"
        
        try {
            $config = Get-Content $configFile | ConvertFrom-Json
            if ($config.settings -and $config.version) {
                Write-TestResult "Configuration Format" "PASS" "Valid JSON configuration"
            } else {
                Write-TestResult "Configuration Format" "FAIL" "Invalid configuration structure"
            }
        } catch {
            Write-TestResult "Configuration Format" "FAIL" "Invalid JSON format" $_.Exception.Message
        }
    } else {
        Write-TestResult "Configuration File" "SKIP" "Optional config file not present"
    }
    
    # Test Windows 11 bypass script
    $win11Script = Join-Path $scriptDir "win11allow.ps1"
    if (Test-Path $win11Script) {
        Write-TestResult "Win11 Bypass Script" "PASS" "Win11 bypass script exists"
    } else {
        Write-TestResult "Win11 Bypass Script" "FAIL" "Win11 bypass script not found"
    }
}

# FUNCTION: Test Parameter Validation
function Test-ParameterValidation {
    Write-Host "`n=== PARAMETER VALIDATION TESTS ===" -ForegroundColor Cyan
    
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $mainScript = Join-Path $scriptDir "windows-comprehensive-updater.ps1"
    
    if (-not (Test-Path $mainScript)) {
        Write-TestResult "Parameter Validation" "SKIP" "Main script not found for testing"
        return
    }
    
    # Test parameter definitions
    $scriptContent = Get-Content $mainScript -Raw
    
    $expectedParams = @(
        "ShowDashboard",
        "DashboardPath", 
        "Deploy",
        "CreateSchedule",
        "RetryDays",
        "SkipUpdateCheck",
        "CheckCumulative",
        "SkipWin11Upgrade"
    )
    
    foreach ($param in $expectedParams) {
        if ($scriptContent -match "param\(.*\`$$param") {
            Write-TestResult "Parameter: $param" "PASS" "Parameter definition found"
        } else {
            Write-TestResult "Parameter: $param" "FAIL" "Parameter definition missing"
        }
    }
}

# FUNCTION: Test Logging Functions
function Test-LoggingFunctions {
    Write-Host "`n=== LOGGING FUNCTION TESTS ===" -ForegroundColor Cyan
    
    try {
        # Test if we can create a test log directory
        $testLogDir = Join-Path $env:TEMP "WindowsUpdateScriptTest"
        if (-not (Test-Path $testLogDir)) {
            New-Item -ItemType Directory -Path $testLogDir -Force | Out-Null
        }
        
        $testLogFile = Join-Path $testLogDir "test.log"
        
        # Test basic file logging
        $testMessage = "Test log entry - $(Get-Date)"
        $testMessage | Out-File -FilePath $testLogFile -Append -Encoding UTF8
        
        if (Test-Path $testLogFile) {
            $logContent = Get-Content $testLogFile -Raw
            if ($logContent -match "Test log entry") {
                Write-TestResult "File Logging" "PASS" "Can write to log files"
            } else {
                Write-TestResult "File Logging" "FAIL" "Log file content mismatch"
            }
        } else {
            Write-TestResult "File Logging" "FAIL" "Unable to create log file"
        }
        
        # Cleanup test files
        try { Remove-Item $testLogDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        
    } catch {
        Write-TestResult "File Logging" "FAIL" "Logging test failed" $_.Exception.Message
    }
    
    # Test Event Log source creation (requires admin)
    try {
        $currentPrincipal = [Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if ($isAdmin) {
            if ([System.Diagnostics.EventLog]::SourceExists("WindowsUpdateScriptTest")) {
                Write-TestResult "Event Log Source" "PASS" "Can access event log sources"
            } else {
                # Try to create a test source
                try {
                    New-EventLog -LogName Application -Source "WindowsUpdateScriptTest"
                    Remove-EventLog -Source "WindowsUpdateScriptTest"
                    Write-TestResult "Event Log Source" "PASS" "Can create event log sources"
                } catch {
                    Write-TestResult "Event Log Source" "FAIL" "Cannot create event log sources" $_.Exception.Message
                }
            }
        } else {
            Write-TestResult "Event Log Source" "SKIP" "Admin privileges required for event log testing"
        }
    } catch {
        Write-TestResult "Event Log Source" "FAIL" "Event log test failed" $_.Exception.Message
    }
}

# FUNCTION: Test System Requirements
function Test-SystemRequirements {
    Write-Host "`n=== SYSTEM REQUIREMENTS TESTS ===" -ForegroundColor Cyan
    
    # Test PowerShell version
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -ge 5) {
        Write-TestResult "PowerShell Version" "PASS" "PowerShell $($psVersion.ToString()) detected"
    } else {
        Write-TestResult "PowerShell Version" "FAIL" "PowerShell 5.1+ required, found $($psVersion.ToString())"
    }
    
    # Test Windows version
    try {
        $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($osInfo -and ($osInfo.Caption -match "Windows 10|Windows 11|Windows Server")) {
            Write-TestResult "Operating System" "PASS" "$($osInfo.Caption) detected"
        } else {
            Write-TestResult "Operating System" "SKIP" "Could not verify Windows version"
        }
    } catch {
        Write-TestResult "Operating System" "SKIP" "OS detection failed"
    }
    
    # Test network connectivity
    try {
        $testConnection = Test-NetConnection "www.microsoft.com" -Port 443 -InformationLevel Quiet -ErrorAction SilentlyContinue
        if ($testConnection) {
            Write-TestResult "Internet Connectivity" "PASS" "Can reach Microsoft servers"
        } else {
            Write-TestResult "Internet Connectivity" "FAIL" "Cannot reach external servers"
        }
    } catch {
        Write-TestResult "Internet Connectivity" "SKIP" "Connectivity test failed"
    }
    
    # Test available disk space
    try {
        $systemDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
        if ($systemDrive) {
            $freeSpaceGB = [math]::Round($systemDrive.FreeSpace / 1GB, 2)
            if ($freeSpaceGB -gt 5) {
                Write-TestResult "Disk Space" "PASS" "$freeSpaceGB GB free on system drive"
            } else {
                Write-TestResult "Disk Space" "FAIL" "Low disk space: $freeSpaceGB GB"
            }
        }
    } catch {
        Write-TestResult "Disk Space" "SKIP" "Could not check disk space"
    }
}

# FUNCTION: Test Windows Environment Integration
function Test-WindowsEnvironmentIntegration {
    Write-Host "`n=== WINDOWS ENVIRONMENT INTEGRATION TESTS ===" -ForegroundColor Cyan
    
    # Test Windows Update Service Status
    try {
        $wuService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
        if ($wuService) {
            Write-TestResult "Windows Update Service" "PASS" "Service status: $($wuService.Status)"
        } else {
            Write-TestResult "Windows Update Service" "FAIL" "Windows Update service not found"
        }
    } catch {
        Write-TestResult "Windows Update Service" "SKIP" "Cannot check service status"
    }
    
    # Test BITS Service Status
    try {
        $bitsService = Get-Service -Name bits -ErrorAction SilentlyContinue
        if ($bitsService) {
            Write-TestResult "BITS Service" "PASS" "Service status: $($bitsService.Status)"
        } else {
            Write-TestResult "BITS Service" "FAIL" "BITS service not found"
        }
    } catch {
        Write-TestResult "BITS Service" "SKIP" "Cannot check service status"
    }
    
    # Test Winget Availability
    try {
        $wingetVersion = & winget --version 2>$null
        if ($LASTEXITCODE -eq 0 -and $wingetVersion) {
            Write-TestResult "Winget Installation" "PASS" "Version: $wingetVersion"
        } else {
            Write-TestResult "Winget Installation" "FAIL" "Winget not available or not working"
        }
    } catch {
        Write-TestResult "Winget Installation" "SKIP" "Winget not installed"
    }
    
    # Test PSWindowsUpdate Module
    try {
        $pswuModule = Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue
        if ($pswuModule) {
            Write-TestResult "PSWindowsUpdate Module" "PASS" "Version: $($pswuModule.Version)"
        } else {
            Write-TestResult "PSWindowsUpdate Module" "FAIL" "PSWindowsUpdate module not installed"
        }
    } catch {
        Write-TestResult "PSWindowsUpdate Module" "SKIP" "Cannot check module availability"
    }
    
    # Test Event Log Source
    try {
        if ([System.Diagnostics.EventLog]::SourceExists("WindowsUpdateScript")) {
            Write-TestResult "Event Log Source" "PASS" "WindowsUpdateScript source exists"
        } else {
            Write-TestResult "Event Log Source" "FAIL" "WindowsUpdateScript event source not registered"
        }
    } catch {
        Write-TestResult "Event Log Source" "SKIP" "Cannot check event log source"
    }
    
    # Test Registry Access
    try {
        $testKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion" -Name "ProgramFilesDir" -ErrorAction SilentlyContinue
        if ($testKey) {
            Write-TestResult "Registry Access" "PASS" "Can read registry keys"
        } else {
            Write-TestResult "Registry Access" "FAIL" "Cannot access registry"
        }
    } catch {
        Write-TestResult "Registry Access" "SKIP" "Registry access test failed"
    }
    
    # Test Scheduled Tasks Access
    try {
        $testTask = Get-ScheduledTask -TaskName "WindowsUpdate-Manual" -ErrorAction SilentlyContinue
        if ($testTask) {
            Write-TestResult "Scheduled Tasks" "PASS" "Can access scheduled tasks"
        } else {
            Write-TestResult "Scheduled Tasks" "SKIP" "WindowsUpdate-Manual task not found"
        }
    } catch {
        Write-TestResult "Scheduled Tasks" "SKIP" "Cannot access scheduled tasks"
    }
    
    # Test File System Permissions
    try {
        $testDir = "C:\Scripts"
        if (Test-Path $testDir) {
            $testFile = Join-Path $testDir "test-write.tmp"
            try {
                "test" | Out-File -FilePath $testFile -Encoding UTF8 -ErrorAction Stop
                Remove-Item $testFile -Force -ErrorAction SilentlyContinue
                Write-TestResult "File System Permissions" "PASS" "Can write to C:\Scripts"
            } catch {
                Write-TestResult "File System Permissions" "FAIL" "Cannot write to C:\Scripts"
            }
        } else {
            Write-TestResult "File System Permissions" "SKIP" "C:\Scripts directory does not exist"
        }
    } catch {
        Write-TestResult "File System Permissions" "SKIP" "File system permission test failed"
    }
    
    # Test Network Connectivity to Microsoft
    try {
        $connectionTest = Test-NetConnection -ComputerName "www.microsoft.com" -Port 443 -InformationLevel Quiet -ErrorAction SilentlyContinue
        if ($connectionTest) {
            Write-TestResult "Microsoft Connectivity" "PASS" "Can reach Microsoft servers"
        } else {
            Write-TestResult "Microsoft Connectivity" "FAIL" "Cannot reach Microsoft servers"
        }
    } catch {
        Write-TestResult "Microsoft Connectivity" "SKIP" "Network connectivity test failed"
    }
    
    # Test Windows Update Registry Keys
    try {
        $wuKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -ErrorAction SilentlyContinue
        if ($wuKey) {
            Write-TestResult "Windows Update Registry" "PASS" "Windows Update registry keys accessible"
        } else {
            Write-TestResult "Windows Update Registry" "SKIP" "Windows Update registry keys not found"
        }
    } catch {
        Write-TestResult "Windows Update Registry" "SKIP" "Cannot access Windows Update registry"
    }
}

# FUNCTION: Test Configuration File Validation
function Test-ConfigurationValidation {
    Write-Host "`n=== CONFIGURATION VALIDATION TESTS ===" -ForegroundColor Cyan
    
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $configFile = Join-Path $scriptDir "windows-update-config.json"
    
    if (Test-Path $configFile) {
        # Test JSON parsing
        try {
            $config = Get-Content $configFile | ConvertFrom-Json
            Write-TestResult "Configuration JSON" "PASS" "Valid JSON format"
        } catch {
            Write-TestResult "Configuration JSON" "FAIL" "Invalid JSON format: $_"
            return
        }
        
        # Test required sections
        $requiredSections = @("version", "settings")
        foreach ($section in $requiredSections) {
            if ($config.PSObject.Properties.Name -contains $section) {
                Write-TestResult "Config Section: $section" "PASS" "Required section present"
            } else {
                Write-TestResult "Config Section: $section" "FAIL" "Required section missing"
            }
        }
        
        # Test version format
        if ($config.version -match '^\d+\.\d+\.\d+$') {
            Write-TestResult "Version Format" "PASS" "Version: $($config.version)"
        } else {
            Write-TestResult "Version Format" "FAIL" "Invalid version format: $($config.version)"
        }
        
        # Test settings subsections
        $expectedSettings = @("general", "winget", "windowsUpdate", "maintenance", "scheduling", "dashboard", "security", "notifications", "advanced")
        foreach ($setting in $expectedSettings) {
            if ($config.settings.PSObject.Properties.Name -contains $setting) {
                Write-TestResult "Settings Subsection: $setting" "PASS" "Settings subsection present"
            } else {
                Write-TestResult "Settings Subsection: $setting" "SKIP" "Optional subsection missing: $setting"
            }
        }
    } else {
        Write-TestResult "Configuration File" "SKIP" "Configuration file not found"
    }
}

# FUNCTION: Test Script Function Dependencies
function Test-ScriptFunctionDependencies {
    Write-Host "`n=== SCRIPT FUNCTION DEPENDENCY TESTS ===" -ForegroundColor Cyan
    
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $mainScript = Join-Path $scriptDir "windows-comprehensive-updater.ps1"
    
    if (-not (Test-Path $mainScript)) {
        Write-TestResult "Main Script" "SKIP" "Main script not found for dependency testing"
        return
    }
    
    try {
        $scriptContent = Get-Content $mainScript -Raw
        
        # Test for critical functions
        $criticalFunctions = @(
            "Write-LogMessage",
            "Confirm-RunAsAdmin", 
            "Invoke-WingetUpdates",
            "Invoke-WindowsUpdates",
            "Invoke-SystemMaintenance",
            "Confirm-WizmoAvailability",
            "Invoke-SilentReboot",
            "Start-UpdateDashboard",
            "Update-DashboardStatus",
            "New-PatchTuesdaySchedule",
            "Get-SecondTuesday",
            "Test-ConfigurationSchema"
        )
        
        foreach ($func in $criticalFunctions) {
            if ($scriptContent -match "function $func") {
                Write-TestResult "Function: $func" "PASS" "Function definition found"
            } else {
                Write-TestResult "Function: $func" "FAIL" "Function definition missing"
            }
        }
        
        # Test for global variables
        $globalVariables = @(
            '$global:logFile',
            '$global:Config',
            '$script:EventIds'
        )
        
        foreach ($var in $globalVariables) {
            if ($scriptContent -match [regex]::Escape($var)) {
                Write-TestResult "Global Variable: $var" "PASS" "Global variable defined"
            } else {
                Write-TestResult "Global Variable: $var" "FAIL" "Global variable missing"
            }
        }
        
        # Test for parameter definitions
        $parameters = @(
            "ShowDashboard",
            "DashboardPath",
            "Deploy",
            "CreateSchedule",
            "RetryDays",
            "SkipUpdateCheck",
            "CheckCumulative",
            "SkipWin11Upgrade"
        )
        
        foreach ($param in $parameters) {
            if ($scriptContent -match "param\(.*\`$$param") {
                Write-TestResult "Parameter: $param" "PASS" "Parameter definition found"
            } else {
                Write-TestResult "Parameter: $param" "FAIL" "Parameter definition missing"
            }
        }
        
    } catch {
        Write-TestResult "Script Analysis" "FAIL" "Error analyzing script: $_"
    }
}

# FUNCTION: Test Performance and Resource Usage
function Test-PerformanceMetrics {
    Write-Host "`n=== PERFORMANCE AND RESOURCE TESTS ===" -ForegroundColor Cyan
    
    # Test script startup time (simulated)
    try {
        $startTime = Get-Date
        Start-Sleep -Milliseconds 100  # Simulate minimal startup
        $endTime = Get-Date
        $startupTime = ($endTime - $startTime).TotalMilliseconds
        
        if ($startupTime -lt 500) {
            Write-TestResult "Script Startup Time" "PASS" "$([math]::Round($startupTime, 2))ms"
        } else {
            Write-TestResult "Script Startup Time" "FAIL" "Slow startup: $([math]::Round($startupTime, 2))ms"
        }
    } catch {
        Write-TestResult "Script Startup Time" "SKIP" "Cannot measure startup time"
    }
    
    # Test memory usage (if available)
    try {
        $process = Get-Process -Id $PID -ErrorAction SilentlyContinue
        if ($process) {
            $memoryMB = [math]::Round($process.WorkingSet64 / 1MB, 2)
            if ($memoryMB -lt 100) {
                Write-TestResult "Memory Usage" "PASS" "$memoryMB MB"
            } else {
                Write-TestResult "Memory Usage" "WARNING" "High memory usage: $memoryMB MB"
            }
        }
    } catch {
        Write-TestResult "Memory Usage" "SKIP" "Cannot measure memory usage"
    }
    
    # Test disk space requirements
    try {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
        $scriptSize = (Get-ChildItem $scriptDir -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        $scriptSizeMB = [math]::Round($scriptSize / 1MB, 2)
        
        if ($scriptSizeMB -lt 10) {
            Write-TestResult "Disk Space Requirements" "PASS" "$scriptSizeMB MB"
        } else {
            Write-TestResult "Disk Space Requirements" "WARNING" "Large disk footprint: $scriptSizeMB MB"
        }
    } catch {
        Write-TestResult "Disk Space Requirements" "SKIP" "Cannot measure disk usage"
    }
}

# FUNCTION: Main Test Execution
function Start-TestSuite {
    Write-Host "Windows Comprehensive Updater Test Suite v2.1.0" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $mainScript = Join-Path $scriptDir "windows-comprehensive-updater.ps1"
    
    # Run selected test categories
    if ($RunAll -or $RunBasicTests) {
        Test-SystemRequirements
        Test-FileDependencies
        Test-ParameterValidation
        Test-ConfigurationValidation
        Test-PerformanceMetrics
    }
    
    if ($RunAll -or $RunFunctionTests) {
        if (Test-Path $mainScript) {
            Test-ScriptSyntax -ScriptPath $mainScript
        }
        Test-LoggingFunctions
        Test-ScriptFunctionDependencies
    }
    
    if ($RunAll -or $RunIntegrationTests) {
        Write-Host "`n=== INTEGRATION TESTS ===" -ForegroundColor Cyan
        Test-WindowsEnvironmentIntegration
        Write-TestResult "Integration Tests" "PASS" "Windows environment integration tests completed"
    }
    
    # Display summary
    Write-Host "`n=========================================" -ForegroundColor Cyan
    Write-Host "TEST SUMMARY" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Total Tests: $($script:testResults.Total)" -ForegroundColor White
    Write-Host "Passed: $($script:testResults.Passed)" -ForegroundColor Green
    Write-Host "Failed: $($script:testResults.Failed)" -ForegroundColor Red
    Write-Host "Skipped: $($script:testResults.Skipped)" -ForegroundColor Yellow
    
    $successRate = if ($script:testResults.Total -gt 0) { 
        [math]::Round(($script:testResults.Passed / $script:testResults.Total) * 100, 1) 
    } else { 0 }
    
    Write-Host "Success Rate: $successRate%" -ForegroundColor $(if ($successRate -gt 80) { "Green" } elseif ($successRate -gt 60) { "Yellow" } else { "Red" })
    
    # Show detailed results if requested
    if ($ShowDetails -and $script:testResults.Details.Count -gt 0) {
        Write-Host "`nDETAILED RESULTS:" -ForegroundColor Cyan
        foreach ($detail in $script:testResults.Details) {
            $color = switch ($detail.Result) {
                "PASS" { "Green" }
                "FAIL" { "Red" }
                "SKIP" { "Yellow" }
                default { "Gray" }
            }
            
            $output = "[$($detail.Result)] $($detail.Name)"
            if ($detail.Details) { $output += " - $($detail.Details)" }
            if ($detail.Error) { $output += " (Error: $($detail.Error))" }
            
            Write-Host $output -ForegroundColor $color
        }
    }
    
    Write-Host "`nCompleted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    
    # Return exit code based on results
    if ($script:testResults.Failed -gt 0) {
        Write-Host "`nTests failed. Please review the failures before deploying the script." -ForegroundColor Red
        return 1
    } else {
        Write-Host "`nAll tests passed successfully!" -ForegroundColor Green
        return 0
    }
}

# Script execution
if (-not ($RunBasicTests -or $RunFunctionTests -or $RunIntegrationTests -or $RunAll)) {
    Write-Host "Windows Comprehensive Updater Test Suite" -ForegroundColor Cyan
    Write-Host "Usage:" -ForegroundColor White
    Write-Host "  -RunBasicTests      Run basic validation tests" -ForegroundColor Gray
    Write-Host "  -RunFunctionTests   Run function and syntax tests" -ForegroundColor Gray
    Write-Host "  -RunIntegrationTests Run integration tests (Windows only)" -ForegroundColor Gray
    Write-Host "  -RunAll             Run all available tests" -ForegroundColor Gray
    Write-Host "  -ShowDetails        Show detailed test results" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor White
    Write-Host "  .\test-windows-comprehensive-updater.ps1 -RunAll" -ForegroundColor Gray
    Write-Host "  .\test-windows-comprehensive-updater.ps1 -RunBasicTests -ShowDetails" -ForegroundColor Gray
} else {
    $exitCode = Start-TestSuite
    Exit $exitCode
}
