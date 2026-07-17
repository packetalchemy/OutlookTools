using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTools.Snooze
{
    /// <summary>
    /// OutlookTools — Email Snooze
    /// Hide emails temporarily and make them reappear at a specified time.
    /// 
    /// How it works:
    /// 1. Move email to a hidden "Snoozed" folder
    /// 2. Set a flag with due date = snooze until time
    /// 3. A timer checks every 5 minutes for snoozed emails
    /// 4. When time is up, move email back to Inbox
    /// 
    /// Storage: "Snoozed" folder inside the user's mailbox
    /// </summary>
    public class EmailSnooze
    {
        private const string SNOOZE_FOLDER_NAME = "OutlookTools_Snoozed";

        /// <summary>
        /// Snooze selected email until the specified date/time.
        /// </summary>
        public static void SnoozeEmail(Outlook.Application app, Outlook.MailItem mail, DateTime until)
        {
            if (mail == null) return;

            try
            {
                // Create snooze folder if it doesn't exist
                Outlook.Folder snoozeFolder = EnsureSnoozeFolder(app);
                if (snoozeFolder == null) return;

                // Set a reminder flag so we can find it later
                mail.FlagStatus = (Outlook.OlFlagStatus)1; // olFlagForward = 1
                mail.FlagRequest = $"Snoozed until {until:yyyy-MM-dd HH:mm}";
                mail.FlagDueBy = until;
                mail.Save();

                // Move to snooze folder
                mail.Move(snoozeFolder);

                ThisAddIn.LogDebug($"Snooze: moved '{mail.Subject}' until {until}");
            }
            catch (Exception ex)
            {
                ThisAddIn.LogDebug("Snooze error: " + ex.Message);
            }
        }

        /// <summary>
        /// Check for snoozed emails that are ready to return to Inbox.
        /// </summary>
        public static int CheckAndRestore(Outlook.Application app)
        {
            int restored = 0;

            try
            {
                Outlook.Folder snoozeFolder = GetSnoozeFolder(app);
                if (snoozeFolder == null) return 0;

                Outlook.Folder inbox = app.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                if (inbox == null) return 0;

                Outlook.Items items = snoozeFolder.Items;
                var now = DateTime.Now;

                for (int i = items.Count; i >= 1; i--)
                {
                    object obj = null;
                    try
                    {
                        obj = items[i];
                        if (!(obj is Outlook.MailItem mail)) continue;

                        // Check if snooze time has passed
                        DateTime dueBy = mail.FlagDueBy;
                        if (dueBy != DateTime.MinValue && dueBy <= now)
                        {
                            // Clear flag
                            mail.FlagStatus = Outlook.OlFlagStatus.olNoFlag;
                            mail.FlagRequest = "";
                            mail.FlagDueBy = DateTime.MinValue;
                            mail.Save();

                            // Move back to inbox
                            mail.Move(inbox);
                            restored++;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (obj != null) ReleaseCom(obj);
                    }
                }

                if (restored > 0)
                    ThisAddIn.LogDebug($"Snooze: restored {restored} emails to Inbox.");
            }
            catch (Exception ex)
            {
                ThisAddIn.LogDebug("Snooze.CheckAndRestore: " + ex.Message);
            }

            return restored;
        }

        /// <summary>
        /// Get snooze folder. Returns null if it doesn't exist.
        /// </summary>
        private static Outlook.Folder GetSnoozeFolder(Outlook.Application app)
        {
            try
            {
                Outlook.Folder root = app.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox).Parent as Outlook.Folder;
                return root?.Folders[SNOOZE_FOLDER_NAME] as Outlook.Folder;
            }
            catch { return null; }
        }

        /// <summary>
        /// Get or create snooze folder.
        /// </summary>
        private static Outlook.Folder EnsureSnoozeFolder(Outlook.Application app)
        {
            try
            {
                Outlook.Folder existing = GetSnoozeFolder(app);
                if (existing != null) return existing;

                Outlook.Folder root = app.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox).Parent as Outlook.Folder;
                if (root == null) return null;

                Outlook.Folder folder = root.Folders.Add(SNOOZE_FOLDER_NAME) as Outlook.Folder;
                ThisAddIn.LogDebug("Snooze: created snooze folder.");
                return folder;
            }
            catch (Exception ex)
            {
                ThisAddIn.LogDebug("Snooze.EnsureSnoozeFolder: " + ex.Message);
                return null;
            }
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
