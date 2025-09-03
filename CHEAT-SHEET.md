# Windows Update Script - Quick Reference

## 🚀 Basic Commands

### Run Updates Now

```powershell
.\windows-update-script.ps1
```

### Run with Dashboard

```powershell
.\windows-update-script.ps1 -ShowDashboard
```

### Set Up Automatic Updates

```powershell
.\windows-update-script.ps1 -CreateSchedule
```

### Just Check for Updates

```powershell
.\windows-update-script.ps1 -CheckCumulative
```

## 📋 What It Does

1. **Winget Updates** → Updates all your apps
2. **Windows Updates** → Installs system updates
3. **Cleanup** → Removes temp files
4. **Smart Reboot** → Restarts only when needed

## 🎯 Quick Options

| Command | What it does |
|---------|-------------|
| `-ShowDashboard` | Opens web dashboard to watch progress |
| `-CreateSchedule` | Sets up monthly automatic updates |
| `-CheckCumulative` | Only checks, doesn't install |
| `-DisableWinget` | Skip app updates |
| `-DisableWindowsUpdates` | Skip Windows updates |

## 🔧 Troubleshooting

### Script won't run?

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Permission issues?

- Right-click → "Run as Administrator"
- Or let script auto-elevate

### Updates failed?

- Check internet connection
- Try again (has retry logic)
- Check Event Viewer logs

## 📊 Monitoring

- **Dashboard**: Real-time progress when using `-ShowDashboard`
- **Logs**: `C:\Scripts\WindowsUpdateLog.txt`
- **Event Viewer**: Search "WindowsUpdateScript"

## ⏰ Automatic Schedule

- **Runs**: 2nd Tuesday of each month at 2:00 AM
- **Retries**: Daily for 7 days if needed
- **Self-updating**: Refreshes schedule automatically

## ❓ Quick FAQ

**Q: Works on Windows 10/11?**  
**A:** Yes!

**Q: Will it restart my PC?**  
**A:** Only if Windows updates require it.

**Q: How long does it take?**  
**A:** 15-45 minutes usually.

**Q: Can I stop it?**  
**A:** Yes, close PowerShell window anytime.

---

**Remember**: Most users just need `.\windows-update-script.ps1`
