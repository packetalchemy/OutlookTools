using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using OutlookTools.Settings;

namespace OutlookTools
{
    /// <summary>
    /// OutlookTools v1.2.0 — Main add-in entry point.
    /// Manages timers for: Follow-up check, Snooze restore, Reminder cleanup, Auto-archive.
    /// </summary>
    public partial class ThisAddIn
    {
        private System.Timers.Timer _followUpTimer;
        private System.Timers.Timer _snoozeTimer;
        private System.Timers.Timer _reminderTimer;
        private System.Timers.Timer _archiveTimer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Hook into mail send event to auto-track follow-ups
                Application.ItemSend += OnItemSend;

                // Follow-up check: every 30 minutes
                _followUpTimer = new System.Timers.Timer { Interval = TimeSpan.FromMinutes(30).TotalMilliseconds, AutoReset = true };
                _followUpTimer.Elapsed += (s, ev) =>
                {
                    try
                    {
                        var resolved = FollowUp.FollowUpTracker.CheckForReplies(Application);
                        if (resolved.Count > 0)
                            LogDebug($"FollowUp: {resolved.Count} follow-up(s) resolved by replies.");

                        var overdue = FollowUp.FollowUpTracker.GetPending()
                            .Where(x => x.Status == FollowUp.FollowUpStatus.Overdue).ToList();
                        if (overdue.Count > 0)
                            LogDebug($"FollowUp: {overdue.Count} follow-up(s) overdue!");
                    }
                    catch (Exception ex) { LogDebug("FollowUp timer: " + ex.Message); }
                };
                _followUpTimer.Start();

                // Snooze check: every 5 minutes
                _snoozeTimer = new System.Timers.Timer { Interval = TimeSpan.FromMinutes(5).TotalMilliseconds, AutoReset = true };
                _snoozeTimer.Elapsed += (s, ev) =>
                {
                    try { Snooze.EmailSnooze.CheckAndRestore(Application); }
                    catch (Exception ex) { LogDebug("Snooze timer: " + ex.Message); }
                };
                _snoozeTimer.Start();

                // Reminder cleanup: every 30 minutes
                _reminderTimer = new System.Timers.Timer { Interval = TimeSpan.FromMinutes(30).TotalMilliseconds, AutoReset = true };
                _reminderTimer.Elapsed += (s, ev) =>
                {
                    try { Commands.ReminderCleanup.Run(Application); }
                    catch (Exception ex) { LogDebug("Reminder timer: " + ex.Message); }
                };
                _reminderTimer.Start();

                // Auto-archive: daily at configured hour
                SetupArchiveTimer();

                // Daily cleanup of old follow-up records
                FollowUp.FollowUpTracker.Cleanup();

                LogDebug("OutlookTools v1.2.0 started successfully.");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "OutlookTools startup error: " + ex.Message,
                    "OutlookTools", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Auto-track sent emails for follow-up.
        /// </summary>
        private void OnItemSend(object item, ref bool cancel)
        {
            try
            {
                if (item is MailItem mail)
                {
                    FollowUp.FollowUpTracker.TrackSentEmail(mail);
                }
            }
            catch (Exception ex) { LogDebug("OnItemSend: " + ex.Message); }
        }

        private void SetupArchiveTimer()
        {
            int hour = SettingsManager.GetArchiveHour();
            var nextRun = DateTime.Today.AddHours(hour);
            if (nextRun < DateTime.Now) nextRun = nextRun.AddDays(1);
            _archiveTimer = new System.Timers.Timer
            {
                Interval = TimeSpan.FromHours(24).TotalMilliseconds,
                AutoReset = true
            };
            _archiveTimer.Elapsed += (s, ev) =>
            {
                try { Archive.SmartArchiveEngine.Run(Application); }
                catch (Exception ex) { LogDebug("Archive: " + ex.Message); }
            };
            _archiveTimer.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _followUpTimer?.Stop();
            _snoozeTimer?.Stop();
            _reminderTimer?.Stop();
            _archiveTimer?.Stop();
            LogDebug("OutlookTools v1.2.0 shut down.");
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OutlookToolsRibbon();
        }

        public static void LogDebug(string message)
        {
            try
            {
                if (!SettingsManager.GetDebugLogEnabled()) return;
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookTools");
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, "outlook-tools.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {message}{Environment.NewLine}");
            }
            catch { }
        }

        #region VSTO generated
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
