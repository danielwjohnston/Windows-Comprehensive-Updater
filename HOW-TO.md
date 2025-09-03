# Windows Update Script - Simple How To Guide

## 🚀 Quick Start (3 Easy Steps)

### Step 1: Download & Run
```powershell
# Download the script and run it
.\windows-update-script.ps1
```

### Step 2: First Time Setup
The script will automatically:
- ✅ Check for admin rights (will ask for elevation if needed)
- ✅ Install required PowerShell modules
- ✅ Download necessary tools (like Wizmo for silent reboots)
- ✅ Set up logging and event sources

### Step 3: Choose Your Option
```powershell
# Option A: Just run updates now
.\windows-update-script.ps1

# Option B: Run with monitoring dashboard
.\windows-update-script.ps1 -ShowDashboard

# Option C: Set up automatic monthly updates
.\windows-update-script.ps1 -CreateSchedule
```

## 📋 What The Script Does

### 🔄 Update Process
1. **Winget Updates** - Updates all your applications (Chrome, VS Code, etc.)
2. **Windows Updates** - Installs all Windows security and feature updates
3. **System Cleanup** - Removes temp files and optimizes Windows
4. **Smart Reboot** - Restarts only when needed, silently

### ⏰ Automatic Scheduling
- **Patch Tuesday**: Runs automatically on the 2nd Tuesday of each month
- **Retry Logic**: Tries again for 7 days if updates aren't ready
- **Self-Maintaining**: Updates its own schedule automatically

## 🎯 Common Use Cases

### "I want to update my computer right now"
```powershell
.\windows-update-script.ps1 -ShowDashboard
```
- Opens a web dashboard to watch progress
- Updates everything automatically
- Shows you what's happening in real-time

### "I want automatic monthly updates"
```powershell
.\windows-update-script.ps1 -CreateSchedule
```
- Sets up automatic Patch Tuesday updates
- Creates scheduled tasks in Windows Task Scheduler
- Runs at 2:00 AM every month

### "I want to check if updates are available"
```powershell
.\windows-update-script.ps1 -CheckCumulative
```
- Only checks for updates, doesn't install them
- Useful for testing or scheduling

## 📊 Monitoring Your Updates

### Real-Time Dashboard
When you use `-ShowDashboard`, you'll see:
- ✅ Current operation status
- 📊 Progress percentage
- 📝 Live log of what's happening
- 🖥️ System information

### Check Logs Later
- **Event Viewer**: Search for "WindowsUpdateScript" in Application logs
- **Log File**: `C:\Scripts\WindowsUpdateLog.txt`
- **Status File**: `C:\Scripts\update-status.json`

## ⚙️ Customization (Optional)

### Configuration File
Create `windows-update-config.json` to customize:

```json
{
  "settings": {
    "winget": {
      "enableWingetUpdates": true
    },
    "windowsUpdate": {
      "enableWindowsUpdates": true
    },
    "scheduling": {
      "patchTuesdayHour": 2,
      "patchTuesdayMinute": 0
    }
  }
}
```

### Command Line Options
```powershell
# Skip Winget updates
.\windows-update-script.ps1 -DisableWinget

# Skip Windows updates
.\windows-update-script.ps1 -DisableWindowsUpdates

# Custom retry days
.\windows-update-script.ps1 -RetryDays 14
```

## 🔧 Troubleshooting

### "Script won't run"
```powershell
# Enable PowerShell execution
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### "Permission denied"
- Right-click the script and "Run as Administrator"
- Or the script will ask for elevation automatically

### "Updates failed"
- Check internet connection
- Try running again (the script handles retries automatically)
- Check logs in Event Viewer for details

### "Dashboard won't open"
- Make sure no firewall is blocking local web access
- Try running without `-ShowDashboard` first

## 📞 Need Help?

### Quick Diagnosis
```powershell
# Run the test suite
.\test-windows-update-script.ps1 -RunAll
```
This will check if everything is set up correctly.

### Common Questions

**Q: Does this work on Windows 10/11?**
A: Yes! Works on both Windows 10 and Windows 11.

**Q: Will it restart my computer?**
A: Only if Windows updates require it. Uses silent reboot technology.

**Q: Can I stop it once it starts?**
A: Yes, you can close the PowerShell window at any time.

**Q: Does it update Windows itself?**
A: Yes! It handles both application updates (via Winget) and Windows system updates.

**Q: How long does it take?**
A: Usually 15-45 minutes, depending on how many updates are needed.

## 🎯 Pro Tips

1. **Use the dashboard** for your first few runs to see what it does
2. **Set up scheduling** after testing manually first
3. **Check logs** if something seems wrong
4. **Run monthly** to stay current with security updates
5. **Test after major Windows updates** to ensure everything still works

---

**That's it!** The script is designed to be simple - just run it and let it handle the rest. Most users only need the basic command: `.\windows-update-script.ps1`</content>
<parameter name="filePath">/Users/dwjohnston/Library/Mobile Documents/com~apple~CloudDocs/Windsurf/Windows-Update-Script/Windows-Update-Script/HOW-TO.md
