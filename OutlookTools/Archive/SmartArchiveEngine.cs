using System.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookTools.Settings;

namespace OutlookTools.Archive
{
    /// <summary>
    /// OutlookTools — Smart Archive Engine
    /// Moves old emails from Inbox/Sent to seasonal PST archives.
    /// 
    /// Archive structure:
    ///   2026-Season1.pst  (Jan–Mar)
    ///   2026-Season2.pst  (Apr–Jun)
    ///   2026-Season3.pst  (Jul–Sep)
    ///   2026-Season4.pst  (Oct–Dec)
    /// 
    /// Rules:
    /// - Only processes messages older than the configured age threshold.
    /// - Never deletes: only moves.
    /// - Skips protected folders: Calendar, Contacts, Tasks, Notes, Drafts, Outbox,
    ///   Deleted Items, Conversation History.
    /// - Runs as a background task in small batches to keep Outlook responsive.
    /// </summary>
    public static class SmartArchiveEngine
    {
        /// <summary>
        /// Run the archive process for the default mailboxes.
        /// Called from the ribbon button or auto-archive timer.
        /// </summary>
        public static void Run(Outlook.Application app)
        {
            try
            {
                Outlook.NameSpace session = app.Session;
                // Iterate all stores (no IsRootStore filter)

                int processed = 0;
                var dateThreshold = GetArchiveThresholdDate();
                var seasonFolder = EnsureSeasonalArchive(session);

                // Archive Inbox
                Outlook.Folder inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                if (inbox != null)
                    processed += ArchiveFolder(inbox, seasonFolder, dateThreshold);

                // Archive Sent
                Outlook.Folder sent = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail) as Outlook.Folder;
                if (sent != null)
                    processed += ArchiveFolder(sent, seasonFolder, dateThreshold);

                if (processed > 0)
                    ThisAddIn.LogDebug($"SmartArchive: moved {processed} messages to {seasonFolder.Name}.");
            }
            catch (Exception ex)
            {
                ThisAddIn.LogDebug("SmartArchive error: " + ex.Message);
            }
        }

        /// <summary>
        /// Archive messages from a source folder that are older than the threshold.
        /// Returns count of messages moved.
        /// </summary>
        private static int ArchiveFolder(Outlook.Folder source, Outlook.Folder target, DateTime threshold)
        {
            int count = 0;
            // Work in small batches to keep Outlook responsive
            int batchSize = 10;
            int processed = 0;

            try
            {
                Outlook.Items items = source.Items;
                items.Sort("[ReceivedTime]", true); // newest first
                int total = items.Count;

                for (int i = 1; i <= total && processed < batchSize; i++)
                {
                    object obj = null;
                    try
                    {
                        obj = items[i];
                        if (!(obj is Outlook.MailItem mail)) continue;

                        // Check if message is old enough
                        if (mail.ReceivedTime >= threshold) continue;

                        // Skip protected items (flagged for follow-up, categories)
                        if (mail.IsMarkedAsTask) continue;
                        if (mail.Categories != null && mail.Categories.Length > 0) continue;

                        // Move to seasonal archive folder
                        mail.Move(target);
                        count++;
                        processed++;
                    }
                    catch { /* skip individual failures */ }
                    finally
                    {
                        if (obj != null) ReleaseCom(obj);
                    }
                }
            }
            catch { }

            return count;
        }

        /// <summary>
        /// Get the threshold date based on user settings.
        /// Default: 90 days old.
        /// </summary>
        private static DateTime GetArchiveThresholdDate()
        {
            int days = SettingsManager.GetArchiveAgeDays();
            return DateTime.Now.AddDays(-days);
        }

        /// <summary>
        /// Get or create the seasonal archive folder for the current period.
        /// Format: "2026-Season1" → "2026-Season2" etc.
        /// </summary>
        private static Outlook.Folder EnsureSeasonalArchive(Outlook.Application app)
        {
            string seasonName = GetCurrentSeasonName();
            Outlook.NameSpace session = app.Session;

            // Check if the archive store already exists
            Outlook.Stores stores = session.Stores;
            Outlook.Store archiveStore = null;

            foreach (Outlook.Store store in stores)
            {
                string storePath = store.FilePath;
                if (Path.GetFileNameWithoutExtension(storePath).StartsWith(seasonName))
                {
                    archiveStore = store;
                    break;
                }
            }

            // If store doesn't exist, create a new PST for this season
            if (archiveStore == null)
            {
                string archivePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OutlookTools",
                    "Archives",
                    seasonName + ".pst");

                Directory.CreateDirectory(Path.GetDirectoryName(archivePath));

                if (!File.Exists(archivePath))
                {
                    // Create empty PST file via Outlook
                    session.AddStore(archivePath);
                    // Rename the store to our friendly name
                    Outlook.Folder archiveRoot = session.GetDefaultFolder(
                        Outlook.OlDefaultFolders.olFolderInbox).Parent as Outlook.Folder;
                    if (archiveRoot != null)
                    {
                        foreach (Outlook.Folder folder in archiveRoot.Folders)
                        {
                            if (folder.Name.Contains("Personal Folders") ||
                                folder.Name.Contains(seasonName))
                            {
                                folder.Name = seasonName;
                                break;
                            }
                        }
                    }
                }
            }

            // Get or create the Inbox-equivalent inside the archive
            Outlook.Folder archiveInbox = null;
            foreach (Outlook.Store store in stores)
            {
                if (Path.GetFileNameWithoutExtension(store.FilePath).StartsWith(seasonName))
                {
                    Outlook.Folder root = store.GetRootFolder() as Outlook.Folder;
                    archiveInbox = root;
                    break;
                }
            }

            // Create Inbox inside archive if it doesn't exist
            if (archiveInbox != null)
            {
                try { return archiveInbox.Folders["Inbox"] as Outlook.Folder; }
                catch { return archiveInbox.Folders.Add("Inbox") as Outlook.Folder; }
            }

            return null;
        }

        /// <summary>
        /// Get the current season name: 2026-Season1 through 2026-Season4.
        /// </summary>
        private static string GetCurrentSeasonName()
        {
            int year = DateTime.Now.Year;
            int quarter = (DateTime.Now.Month - 1) / 3 + 1;
            return $"{year}-Season{quarter}";
        }

        private static void ReleaseCom(object obj)
        {
            try
            {
                if (obj != null && System.Runtime.InteropServices.Marshal.IsComObject(obj))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch { }
        }
    }
}
