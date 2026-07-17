using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace OutlookTools
{
    /// <summary>
    /// OutlookTools Ribbon v1.2.0
    /// Provides menu actions for all features.
    /// Without VSTO, functions are called directly.
    /// </summary>
    public class OutlookToolsRibbon
    {
        // HOME TAB
        public void ReplyWithAttachment() => Commands.AttachmentActions.Reply(true, false);
        public void ReplyAllWithAttachment() => Commands.AttachmentActions.Reply(true, true);
        public void ForwardWithoutAttachment() => Commands.AttachmentActions.ForwardWithoutAttachments();

        // SEARCH
        public void AdvancedSearch() => Search.AdvancedSearchForm.ShowForm();

        // ARCHIVE & CLEANUP
        public void SmartArchive() => Archive.SmartArchiveEngine.Run(Globals.ThisAddIn.Application);
        public void ReminderCleanup() => Commands.ReminderCleanup.Run(Globals.ThisAddIn.Application);

        // PRODUCTIVITY
        public void Templates() => Templates.TemplatesForm.ShowForm();
        public void BulkActions() => ShowBulkMenu();
        public void EmailStats() => Stats.EmailStatsForm.ShowForm();

        // SMART TOOLS — NEW in v1.2.0
        public void FollowUp() => FollowUp.FollowUpDashboardForm.ShowForm();
        public void Snooze() => ShowSnoozeMenu();
        public void Digest() => Digest.DailyDigestForm.ShowForm();
        public void Notes() => Notes.QuickNotesForm.ShowForm();

        // SETTINGS
        public void Settings() => Settings.SettingsForm.ShowForm();

        private void ShowBulkMenu()
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("📥 Move to Folder...", null, (s, e) => Commands.BulkActions.MoveToFolder(Globals.ThisAddIn.Application));
            menu.Items.Add("🗑️ Delete Selected", null, (s, e) => Commands.BulkActions.DeleteSelected(Globals.ThisAddIn.Application));
            menu.Items.Add("-");
            menu.Items.Add("🚩 Flag Selected", null, (s, e) => Commands.BulkActions.FlagSelected(Globals.ThisAddIn.Application, true));
            menu.Items.Add("⬜ Unflag Selected", null, (s, e) => Commands.BulkActions.FlagSelected(Globals.ThisAddIn.Application, false));
            menu.Items.Add("-");
            menu.Items.Add("✅ Mark as Read", null, (s, e) => Commands.BulkActions.MarkAsRead(Globals.ThisAddIn.Application, true));
            menu.Items.Add("📩 Mark as Unread", null, (s, e) => Commands.BulkActions.MarkAsRead(Globals.ThisAddIn.Application, false));
            menu.Items.Add("-");
            menu.Items.Add("🏷️ Add Category...", null, (s, e) => Commands.BulkActions.AddCategory(Globals.ThisAddIn.Application));
            menu.Items.Add("📤 Export to CSV...", null, (s, e) => Commands.BulkActions.ExportToCsv(Globals.ThisAddIn.Application));
            menu.Show(Cursor.Position);
        }

        private void ShowSnoozeMenu()
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("😴 Snooze 1 Hour", null, (s, e) => SnoozeSelected(0, 1));
            menu.Items.Add("😴 Snooze Until Tomorrow 9AM", null, (s, e) => SnoozeUntilTomorrow());
            menu.Items.Add("😴 Snooze Until Monday", null, (s, e) => SnoozeUntilMonday());
            menu.Items.Add("😴 Snooze 1 Week", null, (s, e) => SnoozeSelected(7, 0));
            menu.Show(Cursor.Position);
        }

        private void SnoozeSelected(int days, int hours)
        {
            var app = Globals.ThisAddIn.Application;
            var explorer = app?.ActiveExplorer();
            if (explorer == null || explorer.Selection.Count == 0) return;
            var mail = explorer.Selection[1] as MailItem;
            if (mail != null)
                Snooze.EmailSnooze.SnoozeEmail(app, mail, DateTime.Now.AddDays(days).AddHours(hours));
        }

        private void SnoozeUntilTomorrow()
        {
            var tomorrow9am = DateTime.Today.AddDays(1).AddHours(9);
            var app = Globals.ThisAddIn.Application;
            var mail = app?.ActiveExplorer()?.Selection?[1] as MailItem;
            if (mail != null) Snooze.EmailSnooze.SnoozeEmail(app, mail, tomorrow9am);
        }

        private void SnoozeUntilMonday()
        {
            int days = ((int)DayOfWeek.Monday - (int)DateTime.Today.DayOfWeek + 7) % 7;
            if (days == 0) days = 7;
            var monday9am = DateTime.Today.AddDays(days).AddHours(9);
            var app = Globals.ThisAddIn.Application;
            var mail = app?.ActiveExplorer()?.Selection?[1] as MailItem;
            if (mail != null) Snooze.EmailSnooze.SnoozeEmail(app, mail, monday9am);
        }
    }
}
