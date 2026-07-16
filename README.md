# OutlookTools — Open-Source Outlook Add-in

**Free. Local-first. No tracking. No server. No admin rights required.**

A community-built Outlook add-in that adds the features Outlook is missing:
attachment workflows, smart archiving, templates, bulk actions, follow-up tracking, snooze, daily digest, and more.

> **Why does this exist?**
> Because closed-source add-ins that access your email are a security risk.
> This project lets you **read every line of code** before you install it.

---

## Features

### 📎 Attachment Actions
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
- Runs daily in background

### ⏰ Reminder Cleanup
- Automatically dismisses reminders for past meetings
- Runs every 30 minutes

### 📧 Email Templates
- Save, edit, and insert email templates
- Placeholders: `{DATE}`, `{TIME}`, `{YEAR}`, `{MONTH}`
- Local storage in `%LOCALAPPDATA%\OutlookTools\Templates\`

### ⚡ Bulk Actions
- Move, Delete, Flag, Categorize, Mark Read/Unread, Export to CSV
- Work on multiple selected emails at once

### 📊 Email Statistics
- Top senders, hourly patterns, category breakdown, size distribution
- Export stats to CSV

### 🔔 Follow-up Tracker (NEW in v1.2.0)
- **Auto-tracks** all sent emails for replies
- Dashboard shows pending/overdue follow-ups
- Snooze, Resolve, or Remove individual items
- Checks every 30 minutes for replies

### 😴 Email Snooze (NEW in v1.2.0)
- Hide emails temporarily and make them reappear later
- Quick options: 1 hour, tomorrow 9AM, Monday, 1 week
- Custom date/time picker
- Checks every 5 minutes for snoozed emails

### 📊 Daily Digest (NEW in v1.2.0)
- Morning summary: new emails, sent, unread, flagged, follow-ups
- Top senders of the day
- Overdue follow-up alerts
- Can be sent as email

### 📝 Quick Notes (NEW in v1.2.0)
- Lightweight note-taking inside Outlook
- Tag notes with email subjects
- Search across all notes
- Local JSON storage

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

---

## Project Structure

```
OutlookTools/
├── OutlookTools/
│   ├── ThisAddIn.cs                    — Entry point, timers, event hooks
│   ├── OutlookToolsRibbon.cs           — Ribbon UI callbacks
│   ├── Commands/
│   │   ├── AttachmentActions.cs        — Reply/Forward with/without attachments
│   │   ├── BulkActions.cs              — Move, delete, flag, categorize, export
│   │   └── ReminderCleanup.cs          — Dismiss past-due reminders
│   ├── Archive/
│   │   └── SmartArchiveEngine.cs       — Seasonal PST archiving
│   ├── Search/
│   │   └── AdvancedSearchForm.cs       — Search dialog with preview
│   ├── Templates/
│   │   └── TemplatesForm.cs            — Email template manager
│   ├── Stats/
│   │   └── EmailStatsForm.cs           — Email analytics dashboard
│   ├── FollowUp/
│   │   ├── FollowUpTracker.cs          — Follow-up tracking logic
│   │   └── FollowUpDashboardForm.cs    — Follow-up dashboard UI
│   ├── Snooze/
│   │   ├── EmailSnooze.cs              — Snooze engine
│   │   └── SnoozePickerForm.cs         — Custom snooze picker
│   ├── Digest/
│   │   ├── DailyDigestGenerator.cs     — Digest generation
│   │   └── DailyDigestForm.cs          — Digest display UI
│   ├── Notes/
│   │   ├── QuickNotesManager.cs        — Notes CRUD
│   │   └── QuickNotesForm.cs           — Notes UI
│   ├── Settings/
│   │   ├── SettingsForm.cs             — Configuration UI
│   │   └── SettingsManager.cs          — Local JSON settings
│   └── Resources/
│       └── Ribbon.xml                  — Ribbon UI definition
├── docs/
│   └── Documentation.html              — Full user manual
├── README.md
└── .gitignore
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

## Settings

Settings are stored in: `%LOCALAPPDATA%\OutlookTools\settings.json`

```json
{
  "archiveAgeDays": 90,
  "archiveHour": 6,
  "followUpDays": 3,
  "autoArchiveEnabled": true,
  "autoReminderEnabled": true,
  "debugLogEnabled": false,
  "startupNotification": false,
  "lastDailyRun": "2026-01-01T00:00:00"
}
```

---

## Acknowledgments

This project was inspired by [**EmailTools**](https://github.com/ParhamGhafouri/EmailTools) by **Parham Ghafouri** ([LinkedIn](https://www.linkedin.com/in/parhaam/)).

EmailTools is a polished Outlook add-in for indexed body search, Smart Archive, bulk mail actions, and attachment workflows. We studied its features and architecture, then rebuilt the concept as an **open-source** project so users can audit the code themselves.

**OutlookTools is NOT a fork of EmailTools.** It is an independent, clean-room implementation written from scratch in C#. EmailTools is closed-source; OutlookTools is fully open-source under the MIT License.

> If you like EmailTools, consider supporting Parham's work as well.

---

## License

**MIT License** — use it, modify it, distribute it freely.

---

**Built with ❤️ by the community, for the community.**
