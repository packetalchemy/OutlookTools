using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using OutlookTools.Commands;
using OutlookTools.Archive;
using OutlookTools.Settings;
using System.Configuration;

namespace OutlookTools
{
    /// <summary>
    /// OutlookTools — Open-source Outlook add-in
    /// Ribbon integration, attachment actions, smart archive, reminder cleanup.
    /// ALL processing is LOCAL. NO network calls. NO telemetry.
    /// </summary>
    public partial class ThisAddIn
    {
        private OutlookToolsRibbon _ribbon;

        /// <summary>
        /// Auto-cleanup timer for reminders (runs every 30 minutes).
        /// </summary>
        private System.Timers.Timer _reminderCleanupTimer;

        /// <summary>
        /// Auto-archive timer (runs once per day, configurable).
        /// </summary>
        private System.Timers.Timer _archiveTimer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Subscribe to folder switch event (for auto-cleanup)
                Application.ActiveExplorer().FolderSwitch +=
                    new Outlook.ExplorerEvents_10_FolderSwitchEventHandler(OnFolderSwitch);

                // Initialize reminder cleanup timer
                SetupReminderTimer();

                // Initialize archive timer
                SetupArchiveTimer();

                // Show friendly startup notification (configurable in settings)
                if (Settings.Default.ShowStartupNotification)
                {
                    Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                }

                LogDebug("OutlookTools started successfully.");
            }
            catch (Exception ex)
            {
                // Critical error during startup — show user-friendly message
                MessageBox.Show(
                    "OutlookTools encountered an error during startup.\n\n" +
                    "The add-in will continue to be loaded, but some features may not work.\n\n" +
                    "Error: " + ex.Message,
                    "OutlookTools",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                _reminderCleanupTimer?.Stop();
                _archiveTimer?.Stop();
                LogDebug("OutlookTools shut down cleanly.");
            }
            catch { /* ignore shutdown errors */ }
        }

        /// <summary>
        /// Ribbon instance — populated when Outlook loads the Ribbon.
        /// </summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new OutlookToolsRibbon();
            return _ribbon;
        }

        /// <summary>
        /// Setup the reminder cleanup timer — dismisses past reminders.
        /// </summary>
        private void SetupReminderTimer()
        {
            _reminderCleanupTimer = new System.Timers.Timer
            {
                Interval = TimeSpan.FromMinutes(30).TotalMilliseconds,
                AutoReset = true
            };
            _reminderCleanupTimer.Elapsed += (s, e) =>
            {
                try { ReminderCleanup.Run(Application); }
                catch (Exception ex) { LogDebug("ReminderCleanup error: " + ex.Message); }
            };
            _reminderCleanupTimer.Start();
        }

        /// <summary>
        /// Setup the archive timer — runs once per day.
        /// </summary>
        private void SetupArchiveTimer()
        {
            int hour = Settings.Default.ArchiveHourOfDay;
            var nextRun = DateTime.Today.AddHours(hour);
            if (nextRun < DateTime.Now) nextRun = nextRun.AddDays(1);
            var delay = (nextRun - DateTime.Now).TotalMilliseconds;

            _archiveTimer = new System.Timers.Timer
            {
                Interval = TimeSpan.FromHours(24).TotalMilliseconds,
                AutoReset = true
            };
            _archiveTimer.Elapsed += (s, e) =>
            {
                try { SmartArchiveEngine.Run(Application); }
                catch (Exception ex) { LogDebug("Archive error: " + ex.Message); }
            };
            _archiveTimer.Start();

            LogDebug($"Auto-archive scheduled for {nextRun}, then every 24 hours.");
        }

        /// <summary>
        /// Trigger daily cleanup when user opens Outlook in the morning.
        /// </summary>
        private void OnFolderSwitch()
        {
            // Only run on first folder switch of the day
            var lastRun = Settings.Default.LastDailyRun;
            if (lastRun.Date < DateTime.Today)
            {
                try
                {
                    ReminderCleanup.Run(Application);
                    Settings.Default.LastDailyRun = DateTime.Today;
                    Settings.Default.Save();
                }
                catch { }
            }
        }

        /// <summary>
        /// Local debug log. Quiet mode by default. File is per-user only.
        /// </summary>
        public static void LogDebug(string message)
        {
            try
            {
                if (!Settings.Default.EnableDebugLog) return;
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OutlookTools");
                Directory.CreateDirectory(dir);
                var file = Path.Combine(dir, "outlook-tools.log");
                File.AppendAllText(file,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {message}{Environment.NewLine}");
            }
            catch { /* logging must not throw */ }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
