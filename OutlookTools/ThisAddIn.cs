using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Timers;
using Microsoft.Office.Interop.Outlook;
using OutlookTools.Settings;
using OutlookTools.Archive;

namespace OutlookTools
{
    #region COM Interfaces (defined inline to avoid extra PIA references)

    /// <summary>
    /// IDTExtensibility2 — the interface Outlook calls for COM add-in lifecycle.
    /// </summary>
    [ComImport, Guid("B65AD801-ABAA-11D0-BD8D-00C04FD65CBE")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDTExtensibility2
    {
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            int ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddIn,
            ref Array Custom);
        void OnDisconnection(int DisconnectMode, ref Array Custom);
        void OnAddInsUpdate(ref Array Custom);
        void OnStartupComplete(ref Array Custom);
        void OnBeginShutdown(ref Array Custom);
    }

    /// <summary>
    /// IRibbonExtensibility — Outlook calls this to get the Ribbon XML.
    /// </summary>
    [ComImport, Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    #endregion

    /// <summary>
    /// OutlookTools v1.2.1 — COM-visible add-in entry point.
    /// Implements IDTExtensibility2 for lifecycle and IRibbonExtensibility for UI.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("8B8E5F3E-1C2D-4A3B-9E7F-6D5C4B3A2F1E")]
    [ProgId("OutlookTools.AddIn")]
    public class ThisAddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private Application _application;
        private Timer _autoArchiveTimer;
        private Timer _reminderCleanupTimer;
        private Timer _followUpTimer;
        private Timer _snoozeTimer;

        /// <summary>
        /// Public accessor used by all features to get the Outlook Application.
        /// </summary>
        public Application Application => _application;

        #region IDTExtensibility2 — Outlook lifecycle callbacks

        public void OnConnection(object Application, int ConnectMode, object AddIn, ref Array Custom)
        {
            try
            {
                _application = (Application)Application;
                Globals.ThisAddIn = this;
                LogDebug("OutlookTools v1.2.1 starting (OnConnection)...");

                // Hook Outlook events
                _application.ItemSend += Application_ItemSend;

                // Start periodic background tasks
                SetupTimers();

                LogDebug("OutlookTools started — timers & hooks active.");
            }
            catch (System.Exception ex)
            {
                LogDebug($"OnConnection error: {ex.Message}");
            }
        }

        public void OnDisconnection(int DisconnectMode, ref Array Custom)
        {
            try
            {
                LogDebug("OutlookTools shutting down (OnDisconnection)...");

                // Remove event hook
                try { _application.ItemSend -= Application_ItemSend; } catch { }

                // Stop and dispose all timers
                StopTimer(ref _autoArchiveTimer);
                StopTimer(ref _reminderCleanupTimer);
                StopTimer(ref _followUpTimer);
                StopTimer(ref _snoozeTimer);

                LogDebug("OutlookTools stopped.");
            }
            catch (System.Exception ex)
            {
                LogDebug($"OnDisconnection error: {ex.Message}");
            }
        }

        public void OnAddInsUpdate(ref Array Custom) { }

        public void OnStartupComplete(ref Array Custom) { }

        public void OnBeginShutdown(ref Array Custom) { }

        #endregion

        #region IRibbonExtensibility — Ribbon XML loading

        /// <summary>
        /// Called by Outlook to get the Ribbon XML.
        /// Loads Ribbon.xml from embedded resources.
        /// </summary>
        public string GetCustomUI(string RibbonID)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("OutlookTools.Resources.Ribbon.xml"))
                {
                    if (stream == null)
                    {
                        LogDebug("Ribbon.xml not found in embedded resources.");
                        return null;
                    }
                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogDebug($"GetCustomUI error: {ex.Message}");
                return null;
            }
        }

        #endregion

        #region Ribbon Callbacks — called by Outlook when buttons are clicked

        public void OnRibbonLoad(object ribbonUI) { }

        // Home Tab
        public void OnReplyWithAttachment(object control) => Commands.AttachmentActions.Reply(true, false);
        public void OnReplyAllWithAttachment(object control) => Commands.AttachmentActions.Reply(true, true);
        public void OnForwardWithoutAttachment(object control) => Commands.AttachmentActions.ForwardWithoutAttachments();
        public void OnAdvancedSearch(object control) => Search.AdvancedSearchForm.ShowForm();

        // OutlookTools Tab
        public void OnRunSmartArchive(object control) => SmartArchiveEngine.Run(_application);
        public void OnRunReminderCleanup(object control) => Commands.ReminderCleanup.Run(_application);
        public void OnTemplates(object control) => Templates.TemplatesForm.ShowForm();
        public void OnBulkActions(object control) => Commands.BulkActions.MoveToFolder(_application);
        public void OnEmailStats(object control) => Stats.EmailStatsForm.ShowForm();
        public void OnFollowUp(object control) => FollowUp.FollowUpDashboardForm.ShowForm();
        public void OnSnooze(object control) => ShowSnoozeMenu();
        public void OnDigest(object control) => Digest.DailyDigestForm.ShowForm();
        public void OnNotes(object control) => Notes.QuickNotesForm.ShowForm();
        public void OnSettings(object control) => Settings.SettingsForm.ShowForm();

        private void ShowSnoozeMenu()
        {
            try
            {
                var explorer = _application?.ActiveExplorer();
                if (explorer == null || explorer.Selection.Count == 0) return;
                var mail = explorer.Selection[1] as MailItem;
                if (mail != null)
                    Snooze.EmailSnooze.SnoozeEmail(_application, mail, DateTime.Now.AddHours(1));
            }
            catch (System.Exception ex) { LogDebug($"Snooze error: {ex.Message}"); }
        }

        #endregion

        #region Outlook Event Handlers

        private void Application_ItemSend(object item, ref bool cancel)
        {
            try
            {
                if (item is MailItem mail)
                {
                    FollowUp.FollowUpTracker.TrackSentEmail(mail);
                    LogDebug($"Sent mail tracked: {mail.Subject}");
                }
            }
            catch (System.Exception ex)
            {
                LogDebug($"ItemSend error: {ex.Message}");
            }
        }

        #endregion

        #region Timer Setup

        private void SetupTimers()
        {
            // Auto-Archive: runs once per day at the configured hour
            int archiveHour = SettingsManager.GetArchiveHour();
            double minutesUntilArchive = GetMinutesUntilHour(archiveHour);
            _autoArchiveTimer = CreateTimer(minutesUntilArchive, AutoArchiveCallback);
            _autoArchiveTimer.AutoReset = false;

            // Reminder Cleanup: every 30 minutes
            _reminderCleanupTimer = CreateTimer(30, ReminderCleanupCallback);

            // Follow-Up Reply Check: every 30 minutes
            _followUpTimer = CreateTimer(30, FollowUpCallback);

            // Snooze Restore: every 5 minutes
            _snoozeTimer = CreateTimer(5, SnoozeRestoreCallback);

            LogDebug($"Timers set — archive in {minutesUntilArchive:F0} min, others: 30/30/5 min.");
        }

        private void AutoArchiveCallback(object sender, ElapsedEventArgs e)
        {
            try
            {
                if (!SettingsManager.GetAutoArchiveEnabled()) return;
                SmartArchiveEngine.Run(_application);
                LogDebug("Auto-archive completed.");
            }
            catch (System.Exception ex) { LogDebug($"Auto-archive error: {ex.Message}"); }

            double nextRun = GetMinutesUntilHour(SettingsManager.GetArchiveHour());
            _autoArchiveTimer.Interval = nextRun * 60_000;
            _autoArchiveTimer.Start();
        }

        private void ReminderCleanupCallback(object sender, ElapsedEventArgs e)
        {
            try
            {
                if (!SettingsManager.GetAutoReminderEnabled()) return;
                Commands.ReminderCleanup.Run(_application);
            }
            catch (System.Exception ex) { LogDebug($"Reminder cleanup error: {ex.Message}"); }
        }

        private void FollowUpCallback(object sender, ElapsedEventArgs e)
        {
            try { FollowUp.FollowUpTracker.CheckForReplies(_application); }
            catch (System.Exception ex) { LogDebug($"Follow-up check error: {ex.Message}"); }
        }

        private void SnoozeRestoreCallback(object sender, ElapsedEventArgs e)
        {
            try { Snooze.EmailSnooze.CheckAndRestore(_application); }
            catch (System.Exception ex) { LogDebug($"Snooze restore error: {ex.Message}"); }
        }

        #endregion

        #region Timer Helpers

        private Timer CreateTimer(double intervalMinutes, ElapsedEventHandler handler)
        {
            var timer = new Timer(intervalMinutes * 60_000) { AutoReset = true };
            timer.Elapsed += handler;
            timer.Start();
            return timer;
        }

        private void StopTimer(ref Timer timer)
        {
            if (timer == null) return;
            try { timer.Stop(); timer.Dispose(); } catch { }
            timer = null;
        }

        private double GetMinutesUntilHour(int hour)
        {
            DateTime now = DateTime.Now;
            DateTime next = new DateTime(now.Year, now.Month, now.Day, hour, 0, 0);
            if (next <= now) next = next.AddDays(1);
            return (next - now).TotalMinutes;
        }

        #endregion

        #region Debug Logging

        public static void LogDebug(string message)
        {
            try
            {
                if (!SettingsManager.GetDebugLogEnabled()) return;
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OutlookTools");
                Directory.CreateDirectory(dir);
                File.AppendAllText(
                    Path.Combine(dir, "outlook-tools.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {message}{Environment.NewLine}");
            }
            catch { }
        }

        #endregion
    }

    /// <summary>
    /// Static accessor for the Application object.
    /// Set by OnConnection; used by all features.
    /// </summary>
    public static class Globals
    {
        private static ThisAddIn _addIn;

        public static ThisAddIn ThisAddIn
        {
            get => _addIn;
            set => _addIn = value;
        }
    }
}
