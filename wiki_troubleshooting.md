# Troubleshooting Guide

## Common Issues and Solutions

### Script Execution Issues

#### "Execution Policy" Error

**Error:** `cannot be loaded because running scripts is disabled on this system`

**Solution:**

```powershell
# Set execution policy for current session
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

# Or set it permanently (requires admin)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
```

#### "Access Denied" Error

**Error:** `Access to the path 'C:\Scripts' is denied`

**Solution:**

- Run PowerShell as Administrator
- Check antivirus exclusions for the script directory
- Ensure user has write permissions to C:\Scripts

#### "Winget Not Found" Error

**Error:** `'winget' is not recognized as an internal or external command`

**Solution:**

```powershell
# Install Winget (Windows Package Manager)
# Download from: https://github.com/microsoft/winget-cli/releases
# Or install App Installer from Microsoft Store
```

### Update Installation Issues

#### Windows Update Service Not Running

**Error:** `Windows Update service is not running`

**Solution:**

```powershell
# Start Windows Update service
Start-Service wuauserv

# Enable automatic startup
Set-Service wuauserv -StartupType Automatic
```

#### "Updates Failed" Error

**Error:** `Windows update installation failed`

**Solutions:**

1. Check internet connectivity
2. Restart Windows Update services:

   ```powershell
   Stop-Service wuauserv, bits, cryptsvc, msiserver
   Start-Service wuauserv, bits, cryptsvc
   ```

3. Clear Windows Update cache:

   ```powershell
   Remove-Item "$env:WINDIR\SoftwareDistribution\Download\*" -Recurse -Force
   ```

#### Cumulative Updates Not Found

**Issue:** Script reports no cumulative updates available

**Solution:**

- Microsoft may not have released the monthly cumulative update yet
- The script will automatically retry on subsequent days
- Check Windows Update manually in Settings

### Dashboard Issues

#### Dashboard Won't Load

**Error:** `Cannot access dashboard at http://localhost:8080`

**Solutions:**

1. Check if port 8080 is available:

   ```powershell
   netstat -ano | findstr :8080
   ```

2. Kill conflicting processes
3. Change dashboard port in configuration
4. Check firewall settings

#### Dashboard Shows "No Data"

**Issue:** Dashboard loads but shows no update information

**Solution:**

- Ensure the script is running with `-ShowDashboard` parameter
- Check that `update-status.json` exists in the script directory
- Verify file permissions for dashboard files

### Network and Connectivity Issues

#### SMTP Email Not Working

**Error:** `Failed to send email notification`

**Solutions:**

1. Verify SMTP settings in configuration
2. Check network connectivity to SMTP server
3. Test SMTP credentials
4. Check firewall settings for SMTP port

#### Proxy Server Issues

**Issue:** Updates fail behind corporate proxy

**Solution:**

- Configure proxy settings in PowerShell:

  ```powershell
  $proxy = "http://proxy.company.com:8080"
  $webClient.Proxy = New-Object System.Net.WebProxy($proxy)
  ```

### Performance Issues

#### Script Runs Slowly

**Issue:** Updates take longer than expected

**Solutions:**

1. Check internet connection speed
2. Reduce concurrent update limits in configuration
3. Schedule updates during off-peak hours
4. Exclude large applications from updates

#### High CPU/Memory Usage

**Issue:** Script consumes excessive system resources

**Solutions:**

1. Enable minimal logging mode
2. Reduce update timeout values
3. Schedule during low-usage periods
4. Monitor system resources during execution

### Logging and Diagnostics

#### Enable Debug Logging

```powershell
# In the script, set:
$script:MinimalLogging = $false

# Or in configuration:
{
  "general": {
    "enableDetailedLogging": true
  }
}
```

#### Log File Locations

- Main log: `C:\Scripts\WindowsUpdateLog.txt`
- Archived logs: `C:\Scripts\WindowsUpdateLog_*.txt`
- Event logs: Windows Event Viewer → Application → WindowsUpdateScript

#### Reading Log Files

```powershell
# View recent log entries
Get-Content "C:\Scripts\WindowsUpdateLog.txt" -Tail 50

# Search for specific errors
Select-String -Path "C:\Scripts\WindowsUpdateLog.txt" -Pattern "ERROR"
```

### Advanced Troubleshooting

#### Reset Windows Update Components

```powershell
# Stop services
Stop-Service wuauserv, bits, cryptsvc, msiserver

# Rename folders
Rename-Item "$env:WINDIR\SoftwareDistribution" "$env:WINDIR\SoftwareDistribution.old"
Rename-Item "$env:WINDIR\System32\catroot2" "$env:WINDIR\System32\catroot2.old"

# Start services
Start-Service wuauserv, bits, cryptsvc
```

#### Clean Boot Troubleshooting

1. Perform clean boot to eliminate software conflicts
2. Test script execution
3. Gradually re-enable startup programs to identify conflicts

#### System File Checker

```powershell
# Run system file checker
sfc /scannow

# Check results in CBS.log
Get-Content "$env:WINDIR\Logs\CBS\CBS.log" -Tail 100
```

## Getting Help

### Before Reporting Issues

1. Check this troubleshooting guide
2. Review log files for error messages
3. Test with minimal configuration
4. Verify system requirements

### Information to Include

When reporting issues, please include:

- Script version and PowerShell version
- Windows version and edition
- Complete error messages
- Relevant log file excerpts
- Configuration file (with sensitive data removed)
- Steps to reproduce the issue

### Support Resources

- [GitHub Issues](https://github.com/danielwjohnston/Windows-Comprehensive-Updater/issues)
- [GitHub Wiki](https://github.com/danielwjohnston/Windows-Comprehensive-Updater/wiki)
- Windows Event Viewer logs
- PowerShell error output
