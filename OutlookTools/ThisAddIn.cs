using System;
using System.IO;
using System.Timers;
using Microsoft.Office.Interop.Outlook;
using OutlookTools.Settings;
using OutlookTools.Archive;

namespace OutlookTools
{
    /// <summary>
    /// OutlookTools v1.2.0 — Main add-in entry point with VSTO lifecycle.
    /// Hooks Outlook events and runs periodic background tasks.
    /// </summary>
    public partial class ThisAddIn
    {
        private Application _application;
        private Timer _autoArchiveTimer;
        private Timer _reminderCleanupTimer;
        private Timer _followUpTimer;
        private Timer _snoozeTimer;

        public Application Application => _application;

        #region VSTO Lifecycle

        /// <summary>
        /// Called by Outlook when the add-in loads.
        /// </summary>
        public void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _application = this.Application;
            LogDebug("OutlookTools v1.2.0 starting...");

            // Hook Outlook events
            _application.ItemSend += Application_ItemSend;

            // Start periodic background tasks
            SetupTimers();

            LogDebug("OutlookTools started successfully — timers & hooks active.");
        }

        /// <summary>
        /// Called by Outlook when the add-in unloads.
        /// </summary>
        public void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            LogDebug("OutlookTools shutting down...");

            // Remove event hook
            try { _application.ItemSend -= Application_ItemSend; } catch { }

            // Stop and dispose all timers
            StopTimer(ref _autoArchiveTimer);
            StopTimer(ref _reminderCleanupTimer);
            StopTimer(ref _followUpTimer);
            StopTimer(ref _snoozeTimer);

            LogDebug("OutlookTools stopped.");
        }

        #endregion

        #region Outlook Event Handlers

        private void Application_ItemSend(object item, ref bool cancel)
        {
            // Auto-capture sent mail for follow-up tracking
            try
            {
                if (item is MailItem mail)
                {
                    FollowUp.FollowUpTracker.TrackSentMail(mail);
                    LogDebug($"Sent mail tracked: {mail.Subject}");
                }
            }
            catch (Exception ex)
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
            _autoArchiveTimer.AutoReset = false; // will reschedule after each run

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
            catch (Exception ex) { LogDebug($"Auto-archive error: {ex.Message}"); }

            // Reschedule for tomorrow at the same hour
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
            catch (Exception ex) { LogDebug($"Reminder cleanup error: {ex.Message}"); }
        }

        private void FollowUpCallback(object sender, ElapsedEventArgs e)
        {
            try
            {
                FollowUp.FollowUpTracker.CheckForReplies(_application);
            }
            catch (Exception ex) { LogDebug($"Follow-up check error: {ex.Message}"); }
        }

        private void SnoozeRestoreCallback(object sender, ElapsedEventArgs e)
        {
            try
            {
                Snooze.EmailSnooze.CheckAndRestore(_application);
            }
            catch (Exception ex) { LogDebug($"Snooze restore error: {ex.Message}"); }
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
            try
            {
                timer.Stop();
                timer.Dispose();
            }
            catch { }
            timer = null;
        }

        /// <summary>
        /// Calculate minutes until the next occurrence of the given hour.
        /// </summary>
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
    /// Set by ThisAddIn_Startup; used by all features.
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
