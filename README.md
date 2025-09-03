# Windows Comprehensive Updater

A fully automated Windows system maintenance solution that handles system updates, application updates via Winget, scheduled maintenance, and comprehensive system automation with zero user interaction required.

## 🌟 Features

- **Fully Automated Updates**: Handles Windows and application updates without user intervention
- **Self-Deploying**: Can deploy itself to `C:\Scripts` for scheduled execution
- **Intelligent Update Handling**:
  - Multi-pass update strategy with escalating remediation
  - Handles process conflicts automatically
  - Recovers from failed updates
- **Scheduled Execution**: Automatic Patch Tuesday scheduling with retry logic
- **Comprehensive Logging**: Detailed logs in Event Viewer and console
- **Real-time Dashboard**: Optional web-based monitoring interface
- **Self-Update**: Automatically updates to the latest script version

## 🚀 Quick Start

### Prerequisites

- Windows 10/11 (Server versions may work but are not officially supported)
- PowerShell 5.1 or later
- Administrative privileges

### Basic Usage

```powershell
# Run with default settings (requires admin rights)
.\windows-comprehensive-updater.ps1

# Show the monitoring dashboard
.\windows-comprehensive-updater.ps1 -ShowDashboard

# Deploy script to C:\Scripts and create scheduled tasks
.\windows-comprehensive-updater.ps1 -Deploy -CreateSchedule
```

## ⚙️ Command Line Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-ShowDashboard` | Launch the HTML dashboard for monitoring | `$false` |
| `-DashboardPath` | Custom path to dashboard HTML file | "" |
| `-Deploy` | Force deployment to C:\Scripts | `$false` |
| `-CreateSchedule` | Create Patch Tuesday scheduled tasks | `$false` |
| `-RetryDays` | Days to retry if no cumulative updates found | 7 |
| `-SkipUpdateCheck` | Skip self-update check | `$false` |
| `-CheckCumulative` | Only check for cumulative updates | `$false` |

## 🔄 Update Process

1. **Initialization**
   - Check for admin rights
   - Set up logging and event sources
   - Deploy script if needed

2. **Winget Updates**
   - Update all applications via Winget
   - Handle process conflicts
   - Multi-pass update strategy
   - Automatic retry on failure

3. **Windows Updates**
   - Install all available Windows updates
   - Handle reboots automatically
   - Continue updates after reboot

4. **Cleanup**
   - Remove temporary files
   - Update logs
   - Generate report

## 📊 Dashboard

To monitor the update process in real-time:

```powershell
.\windows-comprehensive-updater.ps1 -ShowDashboard
```

This will start a local web server and open the dashboard in your default browser.

## ⏰ Scheduled Tasks

To create scheduled tasks for automatic updates:

```powershell
.\windows-comprehensive-updater.ps1 -CreateSchedule
```

This will create two scheduled tasks:

- **WindowsUpdate-Maintenance**: Runs on Patch Tuesday at 1:00 AM
- **WindowsUpdate-Retry**: Runs daily at 2:00 AM if the main task fails

## 📝 Logging

All operations are logged to:

- Windows Event Log (Application log, Source: `WindowsUpdateScript`)
- Console output
- Dashboard (if enabled)

## 🔒 Security

- Requires administrative privileges
- Validates all downloads
- Uses secure execution policies
- No external network access required after deployment

<!-- Contributing section moved to bottom to avoid duplication -->

## 🧪 Testing

The script includes a comprehensive test suite to validate functionality:

```powershell
# Run all tests
.\test-windows-comprehensive-updater.ps1 -RunAll

# Run basic validation tests
.\test-windows-comprehensive-updater.ps1 -RunBasicTests -ShowDetails

# Run function and syntax tests
.\test-windows-comprehensive-updater.ps1 -RunFunctionTests
```

The test suite validates:

- PowerShell syntax and structure
- Required function definitions
- File dependencies and integrity
- System requirements
- Parameter validation
- Logging functionality

## ⚙️ Configuration

The script supports an optional configuration file (`windows-update-config.json`) for customizing behavior:

- **Update Settings**: Control which update types to install
- **Scheduling Options**: Customize Patch Tuesday timing
- **Dashboard Settings**: Configure monitoring interface
- **Security Options**: Set download validation rules
- **Notification Settings**: Configure email alerts (if implemented)

## 📁 Project Structure

```text
Windows-Comprehensive-Updater/
├── windows-comprehensive-updater.ps1          # Main update script
├── windows-comprehensive-updater-dashboard.html      # Real-time monitoring dashboard
├── windows-comprehensive-updater-config.json         # Optional configuration file
├── win11allow.ps1                     # Windows 11 upgrade bypass helper
├── test-windows-comprehensive-updater.ps1     # Test suite for validation
└── README.md                          # This documentation
```

## 🔧 Troubleshooting

### Common Issues

1. **PowerShell Execution Policy**: Run `Set-ExecutionPolicy Bypass -Scope CurrentUser`
2. **Missing PSWindowsUpdate**: Script will auto-install the module
3. **Admin Rights**: Script automatically elevates privileges when needed
4. **Network Connectivity**: Ensure access to Windows Update and Winget sources

### Log Locations

- **Main Log**: `C:\Scripts\WindowsUpdateLog.txt`
- **Event Viewer**: Application log, Source: "WindowsUpdateScript"
- **Dashboard Status**: `C:\Scripts\update-status.json`

## 📈 Version History

### v2.1.0 - Enhanced Bulletproof Edition

- Added comprehensive dashboard with real-time monitoring
- Implemented process-conflict resolution for application updates
- Enhanced error handling and logging
- Added Windows 11 feature upgrade support
- Included automated testing framework
- Added configuration file support
- Improved self-deployment and update mechanism

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 📞 Support

For issues and feature requests, please open an issue on the GitHub repository.

## 🤝 Contributing

1. Run the test suite before submitting changes
2. Ensure all tests pass with `./test-windows-comprehensive-updater.ps1 -RunAll`
3. Follow PowerShell best practices and maintain existing code style
4. Update documentation for any new features
