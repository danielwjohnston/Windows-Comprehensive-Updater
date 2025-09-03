# Windows Comprehensive Updater

## Overview

The **Windows Comprehensive Updater** is a fully automated PowerShell script designed to handle Windows system maintenance and updates with zero user interaction. This enterprise-grade solution provides complete hands-off Windows management.

## Key Features

- 🚀 **Zero-Touch Automation** - Complete automation with no user interaction required
- 🔄 **Self-Deployment** - Automatically deploys itself to `C:\Scripts` for system-wide availability
- 📊 **Real-Time Monitoring** - HTML dashboard for live monitoring and status updates
- ⏰ **Scheduled Execution** - Patch Tuesday automation with retry logic
- 📧 **Email Notifications** - Configurable email alerts for update status
- 💾 **System Backup** - Automatic system restore points before updates
- 📈 **Performance Monitoring** - System metrics collection and reporting
- 🛠️ **Enhanced Maintenance** - Disk cleanup, defragmentation, and system file checks

## Quick Start

### Prerequisites

- Windows 10/11
- PowerShell 5.1 or higher
- Administrator privileges
- Internet connection

### Installation

1. Download the script from the repository
2. Run PowerShell as Administrator
3. Execute the script:

   ```powershell
   .\windows-comprehensive-updater.ps1 -Deploy
   ```

### Basic Usage

```powershell
# Run with dashboard
.\windows-comprehensive-updater.ps1 -ShowDashboard

# Create scheduled tasks
.\windows-comprehensive-updater.ps1 -CreateSchedule

# Check for cumulative updates only
.\windows-comprehensive-updater.ps1 -CheckCumulative
```

## Architecture

The script follows a modular architecture with distinct phases:

1. **System Backup** - Creates restore points
2. **Winget Updates** - Updates all applications via Windows Package Manager
3. **Windows Updates** - Installs all available Windows updates
4. **System Maintenance** - Performs cleanup and optimization
5. **Reporting** - Generates logs and notifications

## Configuration

### JSON Configuration File

Create `windows-update-config.json` in the same directory:

```json
{
  "version": "2.1.0",
  "settings": {
    "general": {
      "maxLogFileSizeMB": 50,
      "logRetentionDays": 30,
      "enableDetailedLogging": true
    },
    "winget": {
      "enableWingetUpdates": true,
      "maxUpdateTimeoutMinutes": 30,
      "maxRetryAttempts": 3
    },
    "windowsUpdate": {
      "enableWindowsUpdates": true,
      "maxUpdateCycles": 5
    },
    "notifications": {
      "enableEmailNotifications": false,
      "smtpServer": "smtp.gmail.com",
      "smtpPort": 587,
      "emailFrom": "your-email@gmail.com",
      "emailTo": ["admin@company.com"]
    },
    "advanced": {
      "enableBackupBeforeUpdates": true,
      "backupLocation": "C:\\Backup"
    }
  }
}
```

## Command Line Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-ShowDashboard` | Launch monitoring dashboard | False |
| `-Deploy` | Force deployment to C:\Scripts | False |
| `-CreateSchedule` | Create Patch Tuesday tasks | False |
| `-CheckCumulative` | Only check for cumulative updates | False |
| `-SkipWin11Upgrade` | Skip Windows 11 upgrade checks | False |
| `-RetryDays` | Days to retry if no updates found | 7 |

## Dashboard Features

The monitoring dashboard provides:

- **Real-time Status** - Live update progress
- **System Information** - Computer name, user, PowerShell version
- **Performance Metrics** - CPU, memory, and disk usage
- **Log Viewer** - Live log streaming
- **Progress Tracking** - Visual progress indicators

## Scheduled Tasks

The script creates several scheduled tasks:

- **WindowsUpdate-PatchTuesday** - Main monthly update task
- **WindowsUpdate-PatchTuesday-Retry[1-7]** - Retry tasks for delayed updates
- **WindowsUpdate-Manual** - Manual execution task

All tasks run with SYSTEM privileges and include error recovery.

## Logging and Monitoring

### Log Files

- Main log: `C:\Scripts\WindowsUpdateLog.txt`
- Archived logs: `WindowsUpdateLog_YYYYMMDD_HHMMSS.txt`

### Event Viewer

Events are logged to Windows Event Viewer under:

- **Log**: Application
- **Source**: WindowsUpdateScript
- **Event IDs**: 1000-5999

### Email Notifications

Configure SMTP settings for automated notifications on:

- Update completion
- Errors and failures
- System status changes

## Troubleshooting

### Common Issues

#### Script won't run

- Ensure PowerShell execution policy allows scripts
- Run as Administrator
- Check antivirus exclusions

**Updates failing**

- Verify internet connectivity
- Check Windows Update service status
- Review log files for specific errors

**Dashboard not loading**

- Ensure port 8080 is available
- Check firewall settings
- Verify browser compatibility

### Debug Mode

Enable detailed logging by setting:

```powershell
$script:MinimalLogging = $false
```

## Security Considerations

- Script requires Administrator privileges
- Network access for updates and notifications
- Local file system access for logs and backups
- Registry modifications for Windows 11 bypass (optional)

## Performance Impact

- **CPU Usage**: Minimal during normal operation
- **Memory Usage**: ~50-100MB during execution
- **Disk Space**: ~100MB for logs and temporary files
- **Network**: Required for updates and notifications

## Support and Contributing

### Reporting Issues

1. Check existing issues on GitHub
2. Include log files and system information
3. Describe steps to reproduce

### Contributing

1. Fork the repository
2. Create a feature branch
3. Submit a pull request with detailed description

## License

This project is licensed under the MIT License - see the LICENSE file for details.

---

*For the latest documentation and updates, visit the [GitHub Repository](https://github.com/danielwjohnston/Windows-Comprehensive-Updater)*
