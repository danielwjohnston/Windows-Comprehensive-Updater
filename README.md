# Windows Update Automation Script

A fully automated Windows update solution that handles system updates, application updates via Winget, and scheduled maintenance with zero user interaction required.

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
<<<<<<< HEAD

=======
>>>>>>> 812fb49 (Implement code changes to enhance functionality and improve performance)
- Windows 10/11 (Server versions may work but are not officially supported)
- PowerShell 5.1 or later
- Administrative privileges

### Basic Usage

```powershell
# Run with default settings (requires admin rights)
.\windows-update-script.ps1

# Show the monitoring dashboard
.\windows-update-script.ps1 -ShowDashboard

# Deploy script to C:\Scripts and create scheduled tasks
.\windows-update-script.ps1 -Deploy -CreateSchedule
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
.\windows-update-script.ps1 -ShowDashboard
```

This will start a local web server and open the dashboard in your default browser.

## ⏰ Scheduled Tasks

To create scheduled tasks for automatic updates:

```powershell
.\windows-update-script.ps1 -CreateSchedule
```

This will create two scheduled tasks:
<<<<<<< HEAD

=======
>>>>>>> 812fb49 (Implement code changes to enhance functionality and improve performance)
- **WindowsUpdate-Maintenance**: Runs on Patch Tuesday at 1:00 AM
- **WindowsUpdate-Retry**: Runs daily at 2:00 AM if the main task fails

## 📝 Logging

All operations are logged to:
<<<<<<< HEAD

=======
>>>>>>> 812fb49 (Implement code changes to enhance functionality and improve performance)
- Windows Event Log (Application log, Source: `WindowsUpdateScript`)
- Console output
- Dashboard (if enabled)

## 🔒 Security

- Requires administrative privileges
- Validates all downloads
- Uses secure execution policies
- No external network access required after deployment

## 🤝 Contributing

Contributions are welcome! Please submit issues and pull requests on GitHub.

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 📞 Support

For issues and feature requests, please open an issue on the GitHub repository.
