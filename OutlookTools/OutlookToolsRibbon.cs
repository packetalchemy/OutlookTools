using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace OutlookTools
{
    /// <summary>
    /// OutlookTools Ribbon — v1.1.0
    /// 
    /// HOME TAB: Reply with Attachment, Reply All with Attachment, Forward without Attachment, Search
    /// OutlookTools TAB: Search, Archive & Cleanup, Productivity (Templates, Bulk, Stats), Settings
    /// </summary>
    [ComVisible(true)]
    public class OutlookToolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Explorer" || ribbonID == "Microsoft.Outlook.Mail.Read")
            {
                return LoadResource("OutlookTools.Resources.Ribbon.xml");
            }
            return string.Empty;
        }

        private static string LoadResource(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return string.Empty;
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        // ===== Ribbon callbacks =====

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        // HOME TAB — Attachment actions
        public void OnReplyWithAttachment(Office.IRibbonControl c) => Commands.AttachmentActions.Reply(true, false);
        public void OnReplyAllWithAttachment(Office.IRibbonControl c) => Commands.AttachmentActions.Reply(true, true);
        public void OnForwardWithoutAttachment(Office.IRibbonControl c) => Commands.AttachmentActions.ForwardWithoutAttachments();

        // SEARCH
        public void OnAdvancedSearch(Office.IRibbonControl c) => Search.AdvancedSearchForm.Show();

        // ARCHIVE & CLEANUP
        public void OnRunSmartArchive(Office.IRibbonControl c) => Archive.SmartArchiveEngine.Run(Globals.ThisAddIn.Application);
        public void OnRunReminderCleanup(Office.IRibbonControl c) => Commands.ReminderCleanup.Run(Globals.ThisAddIn.Application);

        // PRODUCTIVITY — NEW in v1.1.0
        public void OnTemplates(Office.IRibbonControl c) => Templates.TemplatesForm.Show();
        public void OnBulkActions(Office.IRibbonControl c) => ShowBulkMenu();
        public void OnEmailStats(Office.IRibbonControl c) => Stats.EmailStatsForm.Show();

        // SETTINGS
        public void OnSettings(Office.IRibbonControl c) => Settings.SettingsForm.Show();

        /// <summary>
        /// Show a context menu for bulk actions (since there are many options).
        /// </summary>
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

            // Show near cursor
            menu.Show(System.Windows.Forms.Cursor.Position);
        }

        public void InvalidateRibbon() => _ribbon?.Invalidate();
    }
}
