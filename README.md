# OutlookTools — Open-Source Outlook Add-in

**Free. Local-first. No tracking. No server. No admin rights required.**

A community-built Outlook add-in that adds the features Outlook is missing:
attachment workflows, smart archiving, reminder cleanup, and advanced search.

> **Why does this exist?**
> Because closed-source add-ins that access your email are a security risk.
> This project lets you **read every line of code** before you install it.

---

## Features

### 📎 Reply with Attachments
- **Reply with Attachment(s)** — keeps original files in your reply
- **Reply All with Attachment(s)** — same for all recipients
- **Forward without Attachment(s)** — strips files but preserves inline signature images

### 🔍 Advanced Search
- Filter by From, To, Subject, Body, Date range, Has attachment
- Searches all mailboxes and mounted archives
- Live preview pane, sortable results

### 📦 Smart Archive
- Moves old emails into seasonal PST archives (`2026-Season1.pst` etc.)
- Never deletes — only moves
- Skips protected folders (Calendar, Contacts, Tasks, Drafts, Deleted Items, etc.)
- Configurable age threshold (default: 90 days)
- Runs daily in background, small batches to keep Outlook responsive

### ⏰ Reminder Cleanup
- Automatically dismisses reminders for past meetings
- Runs every 30 minutes (configurable)

### 🔒 Privacy
- **100% local** — search, archive, and index run on your PC only
- **Zero network calls** — no telemetry, no uploads, no accounts
- **Per-user install** — no admin rights needed
- **Open source** — audit the code yourself

---

## Requirements

| Requirement | Value |
|-------------|-------|
| OS | Windows 10 or 11 |
| Outlook | 2016 / 2019 / 2021 / Microsoft 365 Desktop |
| Framework | .NET Framework 4.8 |
| Privileges | None (per-user install) |
| IDE (to build) | Visual Studio 2019/2022 with Office Development |

---

## Building from Source

### Option A: Visual Studio (Recommended)

1. Clone this repository
2. Open `OutlookTools.sln` in Visual Studio
3. Build → Build Solution (Ctrl+Shift+B)
4. The output will be in `OutlookTools/bin/Debug/`

### Option B: Command Line

```bash
# Requires Visual Studio Build Tools with .NET 4.8
msbuild OutlookTools.sln /p:Configuration=Debug
```

---

## Installation

1. Build the project (see above)
2. Run the generated `OutlookTools.vsto` installer
3. Start Outlook — the "OutlookTools" tab will appear

### Per-User (Recommended)
The VSTO installer installs to your user profile only. No admin rights needed.

---

## Project Structure

```
OutlookTools/
├── OutlookTools/
│   ├── ThisAddIn.cs              — Entry point, timers, ribbon creation
│   ├── OutlookToolsRibbon.cs     — Ribbon UI callbacks
│   ├── Commands/
│   │   ├── AttachmentActions.cs  — Reply/Forward with/without attachments
│   │   └── ReminderCleanup.cs    — Dismiss past-due reminders
│   ├── Archive/
│   │   └── SmartArchiveEngine.cs — Seasonal PST archiving
│   ├── Search/
│   │   └── AdvancedSearchForm.cs — Search dialog with preview
│   ├── Settings/
│   │   ├── SettingsForm.cs       — Configuration UI
│   │   └── SettingsManager.cs    — Local JSON settings storage
│   └── Resources/
│       └── Ribbon.xml            — Ribbon UI definition
└── README.md
```

---

## Security Model

| Concern | Our Approach |
|---------|-------------|
| Source code | ✅ Fully open — audit before install |
| Network access | ❌ None — zero network calls |
| Data storage | ✅ Per-user only — no system-wide changes |
| Admin rights | ❌ Not required |
| Telemetry | ❌ None |
| External dependencies | ❌ None — pure .NET Framework |
| Updates | ❌ Manual — no auto-update risk |

---

## How to Verify This Is Safe

1. **Read the code** — every file is in this repository
2. **Search for `HttpClient` or `WebClient`** — you won't find any
3. **Search for `Upload` or `POST`** — none exist
4. **Check `Process.Start`** — only used to open explorer.exe for log folder
5. **VirusTotal** — scan the built DLL (should be clean)
6. **Process Monitor** — monitor network activity while Outlook runs

---

## Settings

Settings are stored in: `%LOCALAPPDATA%\OutlookTools\settings.json`

```json
{
  "archiveAgeDays": 90,
  "archiveHour": 6,
  "autoArchiveEnabled": true,
  "autoReminderEnabled": true,
  "debugLogEnabled": false,
  "startupNotification": false,
  "lastDailyRun": "2026-01-01T00:00:00"
}
```

---

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

All contributions are welcome. Please ensure your code is clean and well-commented.

---

## License

**MIT License** — use it, modify it, distribute it freely.

```
MIT License

Copyright (c) 2026 OutlookTools Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## Acknowledgments

Inspired by [EmailTools](https://github.com/ParhamGhafouri/EmailTools) by Parham Ghafouri.
Built as an open-source alternative so users can verify the code themselves.

---

**Built with ❤️ by the community, for the community.**
