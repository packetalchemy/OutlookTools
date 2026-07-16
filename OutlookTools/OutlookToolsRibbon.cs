using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace OutlookTools
{
    /// <summary>
    /// OutlookTools Ribbon v1.2.0
    /// HOME TAB: Reply/Forward/Search
    /// OutlookTools TAB: Search, Archive, Productivity, Smart Tools, Settings
    /// </summary>
    [ComVisible(true)]
    public class OutlookToolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Explorer" || ribbonID == "Microsoft.Outlook.Mail.Read")
                return LoadResource("OutlookTools.Resources.Ribbon.xml");
            return string.Empty;
        }

        private static string LoadResource(string name)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (stream == null) return string.Empty;
                using (var reader = new StreamReader(stream)) return reader.ReadToEnd();
            }
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        // HOME TAB
        public void OnReplyWithAttachment(Office.IRibbonControl c) => Commands.AttachmentActions.Reply(true, false);
        public void OnReplyAllWithAttachment(Office.IRibbonControl c) => Commands.AttachmentActions.Reply(true, true);
        public void OnForwardWithoutAttachment(Office.IRibbonControl c) => Commands.AttachmentActions.ForwardWithoutAttachments();

        // SEARCH
        public void OnAdvancedSearch(Office.IRibbonControl c) => Search.AdvancedSearchForm.Show();

        // ARCHIVE & CLEANUP
        public void OnRunSmartArchive(Office.IRibbonControl c) => Archive.SmartArchiveEngine.Run(Globals.ThisAddIn.Application);
        public void OnRunReminderCleanup(Office.IRibbonControl c) => Commands.ReminderCleanup.Run(Globals.ThisAddIn.Application);

        // PRODUCTIVITY
        public void OnTemplates(Office.IRibbonControl c) => Templates.TemplatesForm.Show();
        public void OnBulkActions(Office.IRibbonControl c) => ShowBulkMenu();
        public void OnEmailStats(Office.IRibbonControl c) => Stats.EmailStatsForm.Show();

        // SMART TOOLS — NEW in v1.2.0
        public void OnFollowUp(Office.IRibbonControl c) => FollowUp.FollowUpDashboardForm.Show();
        public void OnSnooze(Office.IRibbonControl c) => ShowSnoozeMenu();
        public void OnDigest(Office.IRibbonControl c) => Digest.DailyDigestForm.Show();
        public void OnNotes(Office.IRibbonControl c) => Notes.QuickNotesForm.Show();

        // SETTINGS
        public void OnSettings(Office.IRibbonControl c) => Settings.SettingsForm.Show();

        private void ShowBulkMenu()
        {
            var menu = new System.Windows.Forms.ContextMenuStrip();
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
            menu.Show(System.Windows.Forms.Cursor.Position);
        }

        private void ShowSnoozeMenu()
        {
            var menu = new System.Windows.Forms.ContextMenuStrip();
            menu.Items.Add("😴 Snooze 1 Hour", null, (s, e) => SnoozeSelected(0, 1));
            menu.Items.Add("😴 Snooze Until Tomorrow 9AM", null, (s, e) => SnoozeUntilTomorrow());
            menu.Items.Add("😴 Snooze Until Monday", null, (s, e) => SnoozeUntilMonday());
            menu.Items.Add("😴 Snooze 1 Week", null, (s, e) => SnoozeSelected(7, 0));
            menu.Items.Add("😴 Snooze Custom...", null, (s, e) => SnoozeCustom());
            menu.Show(System.Windows.Forms.Cursor.Position);
        }

        private void SnoozeSelected(int days, int hours)
        {
            var app = Globals.ThisAddIn.Application;
            var explorer = app.ActiveExplorer();
            if (explorer == null || explorer.Selection.Count == 0) return;

            var mail = explorer.Selection[1] as Microsoft.Office.Interop.Outlook.MailItem;
            if (mail != null)
                Snooze.EmailSnooze.SnoozeEmail(app, mail, DateTime.Now.AddDays(days).AddHours(hours));
        }

        private void SnoozeUntilTomorrow()
        {
            var tomorrow9am = DateTime.Today.AddDays(1).AddHours(9);
            var app = Globals.ThisAddIn.Application;
            var mail = app.ActiveExplorer()?.Selection?[1] as Microsoft.Office.Interop.Outlook.MailItem;
            if (mail != null)
                Snooze.EmailSnooze.SnoozeEmail(app, mail, tomorrow9am);
        }

        private void SnoozeUntilMonday()
        {
            int daysUntilMonday = ((int)DayOfWeek.Monday - (int)DateTime.Today.DayOfWeek + 7) % 7;
            if (daysUntilMonday == 0) daysUntilMonday = 7;
            var monday9am = DateTime.Today.AddDays(daysUntilMonday).AddHours(9);
            var app = Globals.ThisAddIn.Application;
            var mail = app.ActiveExplorer()?.Selection?[1] as Microsoft.Office.Interop.Outlook.MailItem;
            if (mail != null)
                Snooze.EmailSnooze.SnoozeEmail(app, mail, monday9am);
        }

        private void SnoozeCustom()
        {
            using (var form = new Snooze.SnoozePickerForm())
            {
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK && form.SelectedTime.HasValue)
                {
                    var app = Globals.ThisAddIn.Application;
                    var mail = app.ActiveExplorer()?.Selection?[1] as Microsoft.Office.Interop.Outlook.MailItem;
                    if (mail != null)
                        Snooze.EmailSnooze.SnoozeEmail(app, mail, form.SelectedTime.Value);
                }
            }
        }

        public void InvalidateRibbon() => _ribbon?.Invalidate();
    }
}
