# ============================================================================
# WINDOWS UPDATE SCRIPT TEST SUITE
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
    $mainScript = Join-Path $scriptDir "windows-update-script.ps1"
    if (Test-Path $mainScript) {
        Write-TestResult "Main Script File" "PASS" "windows-update-script.ps1 exists"
        
        # Test script size (should be substantial)
        $scriptSize = (Get-Item $mainScript).Length
        if ($scriptSize -gt 50KB) {
            Write-TestResult "Script Content Size" "PASS" "Script is $([math]::Round($scriptSize/1KB,2)) KB"
        } else {
            Write-TestResult "Script Content Size" "FAIL" "Script appears incomplete ($([math]::Round($scriptSize/1KB,2)) KB)"
        }
    } else {
        Write-TestResult "Main Script File" "FAIL" "windows-update-script.ps1 not found"
    }
    
    # Test dashboard file
    $dashboardFile = Join-Path $scriptDir "windows-update-dashboard.html"
    if (Test-Path $dashboardFile) {
        Write-TestResult "Dashboard HTML File" "PASS" "Dashboard file exists"
        
        # Test for essential HTML elements
        $htmlContent = Get-Content $dashboardFile -Raw
        if ($htmlContent -match "Windows Update Monitor" -and $htmlContent -match "javascript") {
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
    $mainScript = Join-Path $scriptDir "windows-update-script.ps1"
    
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

# FUNCTION: Main Test Execution
function Start-TestSuite {
    Write-Host "Windows Update Script Test Suite v2.1.0" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $mainScript = Join-Path $scriptDir "windows-update-script.ps1"
    
    # Run selected test categories
    if ($RunAll -or $RunBasicTests) {
        Test-SystemRequirements
        Test-FileDependencies
        Test-ParameterValidation
    }
    
    if ($RunAll -or $RunFunctionTests) {
        if (Test-Path $mainScript) {
            Test-ScriptSyntax -ScriptPath $mainScript
        }
        Test-LoggingFunctions
    }
    
    if ($RunAll -or $RunIntegrationTests) {
        Write-Host "`n=== INTEGRATION TESTS ===" -ForegroundColor Cyan
        Write-TestResult "Integration Tests" "SKIP" "Integration tests require full Windows environment"
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
    Write-Host "Windows Update Script Test Suite" -ForegroundColor Cyan
    Write-Host "Usage:" -ForegroundColor White
    Write-Host "  -RunBasicTests      Run basic validation tests" -ForegroundColor Gray
    Write-Host "  -RunFunctionTests   Run function and syntax tests" -ForegroundColor Gray
    Write-Host "  -RunIntegrationTests Run integration tests (Windows only)" -ForegroundColor Gray
    Write-Host "  -RunAll             Run all available tests" -ForegroundColor Gray
    Write-Host "  -ShowDetails        Show detailed test results" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor White
    Write-Host "  .\test-windows-update-script.ps1 -RunAll" -ForegroundColor Gray
    Write-Host "  .\test-windows-update-script.ps1 -RunBasicTests -ShowDetails" -ForegroundColor Gray
} else {
    $exitCode = Start-TestSuite
    Exit $exitCode
}
