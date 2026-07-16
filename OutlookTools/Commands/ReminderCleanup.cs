using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTools.Commands
{
    /// <summary>
    /// OutlookTools — Reminder Cleanup
    /// Dismisses overdue reminders (past meetings, past tasks).
    /// Safe: only dismisses items where the due date has passed.
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
                Outlook.Reminders reminders = app.Session.Reminders;
                int dismissed = 0;

                for (int i = reminders.Count; i >= 1; i--)
                {
                    Outlook.Reminder reminder = reminders[i];
                    Outlook.AppointmentItem appointment = null;
                    Outlook.TaskItem task = null;

                    try
                    {
                        if (reminder.Item is Outlook.AppointmentItem ai)
                        {
                            appointment = ai;
                            // Only dismiss if the meeting ended > 5 minutes ago
                            if (ai.End.AddMinutes(5) < DateTime.Now)
                            {
                                reminder.Dismiss();
                                dismissed++;
                            }
                        }
                        else if (reminder.Item is Outlook.TaskItem ti)
                        {
                            task = ti;
                            // Dismiss completed or overdue tasks
                            if (ti.DueDate.Date < DateTime.Today && ti.DueDate != DateTime.MinValue)
                            {
                                reminder.Dismiss();
                                dismissed++;
                            }
                        }
                    }
                    catch { /* skip individual failures */ }
                    finally
                    {
                        if (appointment != null) ReleaseCom(appointment);
                        if (task != null) ReleaseCom(task);
                        if (reminder != null) ReleaseCom(reminder);
                    }
                }

                if (dismissed > 0)
                    ThisAddIn.LogDebug($"ReminderCleanup: dismissed {dismissed} past-due reminders.");
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
