using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTools.Commands
{
    /// <summary>
    /// OutlookTools — Reminder Cleanup
    /// Dismisses overdue reminders (past meetings, past tasks).
    /// Uses Calendar folder approach when Reminders collection is unavailable.
    /// </summary>
    public static class ReminderCleanup
    {
        /// <summary>
        /// Run the reminder cleanup — dismiss all past-due reminders.
        /// </summary>
        public static void Run(Outlook.Application app)
        {
            try
            {
                if (app == null) return;

                // Try to dismiss overdue items via Calendar folder
                Outlook.Folder calendar = app.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;
                if (calendar == null) return;

                int dismissed = 0;
                Outlook.Items items = calendar.Items;
                items.IncludeRecurrences = true;
                items.Sort("[Start]");

                var restriction = "[Start] >= '" + DateTime.Now.AddDays(-1).ToString("g") +
                    "' AND [Start] <= '" + DateTime.Now.ToString("g") + "'";
                Outlook.Items restricted = items.Restrict(restriction);

                foreach (object obj in restricted)
                {
                    if (!(obj is Outlook.AppointmentItem apt)) continue;
                    try
                    {
                        // Dismiss reminders for past appointments
                        if (apt.End.AddMinutes(5) < DateTime.Now && apt.ReminderSet)
                        {
                            apt.ReminderSet = false;
                            apt.Save();
                            dismissed++;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (apt != null) ReleaseCom(apt);
                    }
                }

                if (dismissed > 0)
                    ThisAddIn.LogDebug($"ReminderCleanup: dismissed {dismissed} past reminders.");
            }
            catch (Exception ex)
            {
                ThisAddIn.LogDebug("ReminderCleanup error: " + ex.Message);
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
