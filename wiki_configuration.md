# Configuration Guide

## JSON Configuration File

The Windows Comprehensive Updater supports extensive configuration through a JSON file. Create `windows-update-config.json` in the script directory.

### Complete Configuration Example

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

## Configuration Sections

### General Settings

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `maxLogFileSizeMB` | number | 50 | Maximum log file size in MB |
| `logRetentionDays` | number | 30 | Days to keep archived logs |
| `enableDetailedLogging` | boolean | true | Enable verbose logging |

### Winget Settings

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `enableWingetUpdates` | boolean | true | Enable application updates |
| `maxUpdateTimeoutMinutes` | number | 30 | Timeout for individual updates |
| `maxRetryAttempts` | number | 3 | Retry failed updates |

### Windows Update Settings

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `enableWindowsUpdates` | boolean | true | Enable Windows updates |
| `maxUpdateCycles` | number | 5 | Maximum update installation cycles |

### Notification Settings

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `enableEmailNotifications` | boolean | false | Enable email notifications |
| `smtpServer` | string | - | SMTP server address |
| `smtpPort` | number | - | SMTP server port |
| `emailFrom` | string | - | Sender email address |
| `emailTo` | array | - | Recipient email addresses |

### Advanced Settings

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `enableBackupBeforeUpdates` | boolean | true | Create system restore points |
| `backupLocation` | string | "C:\\Backup" | Backup storage location |

## Email Configuration Examples

### Gmail SMTP

```json
{
  "notifications": {
    "enableEmailNotifications": true,
    "smtpServer": "smtp.gmail.com",
    "smtpPort": 587,
    "emailFrom": "your-email@gmail.com",
    "emailTo": ["admin@company.com", "it@company.com"]
  }
}
```

### Office 365 SMTP

```json
{
  "notifications": {
    "enableEmailNotifications": true,
    "smtpServer": "smtp.office365.com",
    "smtpPort": 587,
    "emailFrom": "noreply@company.com",
    "emailTo": ["helpdesk@company.com"]
  }
}
```

## Validation

The script automatically validates configuration files and falls back to defaults if:

- File is missing
- JSON syntax is invalid
- Required values are missing
- Values are outside acceptable ranges

Check the console output for validation messages when the script starts.
