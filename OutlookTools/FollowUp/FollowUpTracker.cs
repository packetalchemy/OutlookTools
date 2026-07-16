using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTools.FollowUp
{
    /// <summary>
    /// OutlookTools — Email Follow-up Tracker
    /// Tracks sent emails that haven't received a reply.
    /// 
    /// Features:
    /// - Auto-tracks all sent emails
    /// - Checks for replies on a schedule (every 30 minutes)
    /// - Shows dashboard of pending follow-ups
    /// - Configurable follow-up threshold (default: 3 days)
    /// - Snooze individual follow-ups
    /// - Mark as resolved when reply received
    /// - Export follow-up list to CSV
    /// 
    /// Storage: %LOCALAPPDATA%\OutlookTools\followup.json
    /// </summary>
    public class FollowUpTracker
    {
        private static readonly string DataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookTools");

        private static readonly string DataFile = Path.Combine(DataDir, "followup.json");

        public static List<FollowUpItem> Items { get; private set; } = new List<FollowUpItem>();

        static FollowUpTracker()
        {
            Load();
        }

        /// <summary>
        /// Track a sent email for follow-up.
        /// </summary>
        public static void TrackSentEmail(Outlook.MailItem mail)
        {
            try
            {
                if (mail == null) return;

                string entryId = mail.EntryID;
                if (Items.Any(x => x.EntryId == entryId)) return;

                string subject = mail.Subject ?? "(no subject)";
                string to = string.Join("; ", mail.Recipients.Cast<Outlook.Recipient>().Select(r => r.Name));
                string conversationId = mail.ConversationID ?? "";

                Items.Add(new FollowUpItem
                {
                    EntryId = entryId,
                    ConversationId = conversationId,
                    Subject = subject,
                    SentTo = to,
                    SentDate = mail.SentOn,
                    FollowUpDate = DateTime.Now.AddDays(
                        Settings.SettingsManager.GetFollowUpDays()),
                    Status = FollowUpStatus.Pending,
                    SnoozedUntil = null
                });

                Save();
                ThisAddIn.LogDebug($"FollowUp: tracking '{subject}' sent to {to}");
            }
            catch (Exception ex)
            {
                ThisAddIn.LogDebug("FollowUp.TrackSentEmail: " + ex.Message);
            }
        }

        /// <summary>
        /// Check for replies to tracked emails.
        /// Returns list of emails that now have replies.
        /// </summary>
        public static List<FollowUpItem> CheckForReplies(Outlook.Application app)
        {
            var resolved = new List<FollowUpItem>();

            try
            {
                Outlook.NameSpace session = app.Session;
                Outlook.Folder inbox = session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                if (inbox == null) return resolved;

                var pending = Items.Where(x =>
                    x.Status == FollowUpStatus.Pending &&
                    (x.SnoozedUntil == null || x.SnoozedUntil <= DateTime.Now))
                    .ToList();

                foreach (var item in pending)
                {
                    try
                    {
                        // Check if there's a reply in inbox matching conversation ID
                        if (HasReplyInInbox(inbox, item))
                        {
                            item.Status = FollowUpStatus.Resolved;
                            item.ResolvedDate = DateTime.Now;
                            resolved.Add(item);
                        }
                        // Check if follow-up date has passed
                        else if (DateTime.Now >= item.FollowUpDate)
                        {
                            item.Status = FollowUpStatus.Overdue;
                        }
                    }
                    catch { }
                }

                if (resolved.Count > 0)
                    Save();
            }
            catch { }

            return resolved;
        }

        /// <summary>
        /// Check if inbox contains a reply to the tracked email.
        /// Uses ConversationID to match threads.
        /// </summary>
        private static bool HasReplyInInbox(Outlook.Folder inbox, FollowUpItem tracked)
        {
            try
            {
                Outlook.Items items = inbox.Items;
                items.Sort("[ReceivedTime]", true);

                // Look for messages in the same conversation, received after the sent date
                foreach (object obj in items)
                {
                    if (!(obj is Outlook.MailItem mail)) continue;
                    try
                    {
                        // Match by conversation thread
                        if (!string.IsNullOrEmpty(tracked.ConversationId) &&
                            mail.ConversationID == tracked.ConversationId &&
                            mail.SentOn > tracked.SentDate &&
                            mail.SenderName != tracked.SentTo) // Not from the same sender
                        {
                            return true;
                        }

                        // Match by subject (Reply: prefix)
                        if (!string.IsNullOrEmpty(tracked.Subject))
                        {
                            string replySubject = tracked.Subject;
                            if (mail.Subject != null &&
                                (mail.Subject.Contains("Re: " + replySubject) ||
                                 mail.Subject.Contains("RE: " + replySubject)) &&
                                mail.SentOn > tracked.SentDate)
                            {
                                return true;
                            }
                        }
                    }
                    finally
                    {
                        if (mail != null && System.Runtime.InteropServices.Marshal.IsComObject(mail))
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                    }
                }
            }
            catch { }

            return false;
        }

        /// <summary>
        /// Snooze a tracked item.
        /// </summary>
        public static void Snooze(string entryId, DateTime until)
        {
            var item = Items.FirstOrDefault(x => x.EntryId == entryId);
            if (item != null)
            {
                item.SnoozedUntil = until;
                item.Status = FollowUpStatus.Snoozed;
                Save();
            }
        }

        /// <summary>
        /// Mark an item as manually resolved.
        /// </summary>
        public static void MarkResolved(string entryId)
        {
            var item = Items.FirstOrDefault(x => x.EntryId == entryId);
            if (item != null)
            {
                item.Status = FollowUpStatus.Resolved;
                item.ResolvedDate = DateTime.Now;
                Save();
            }
        }

        /// <summary>
        /// Remove a tracked item completely.
        /// </summary>
        public static void Remove(string entryId)
        {
            Items.RemoveAll(x => x.EntryId == entryId);
            Save();
        }

        /// <summary>
        /// Get pending follow-ups (overdue + snoozed now ready).
        /// </summary>
        public static List<FollowUpItem> GetPending()
        {
            return Items.Where(x =>
                x.Status == FollowUpStatus.Pending ||
                x.Status == FollowUpStatus.Overdue ||
                (x.Status == FollowUpStatus.Snoozed && x.SnoozedUntil <= DateTime.Now))
                .OrderBy(x => x.FollowUpDate)
                .ToList();
        }

        /// <summary>
        /// Get all active items (not resolved).
        /// </summary>
        public static List<FollowUpItem> GetActive()
        {
            return Items.Where(x => x.Status != FollowUpStatus.Resolved)
                .OrderBy(x => x.FollowUpDate)
                .ToList();
        }

        /// <summary>
        /// Clean up old resolved items (older than 30 days).
        /// </summary>
        public static void Cleanup()
        {
            var cutoff = DateTime.Now.AddDays(-30);
            int removed = Items.RemoveAll(x =>
                x.Status == FollowUpStatus.Resolved &&
                x.ResolvedDate.HasValue &&
                x.ResolvedDate < cutoff);
            if (removed > 0) Save();
        }

        /// <summary>
        /// Export active follow-ups to CSV.
        /// </summary>
        public static void ExportToCsv(string filePath)
        {
            using (var writer = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
            {
                writer.WriteLine("Subject,To,Sent Date,Follow Up Date,Status,Snoozed Until");
                foreach (var item in GetActive())
                {
                    string snooze = item.SnoozedUntil?.ToString("yyyy-MM-dd HH:mm") ?? "";
                    writer.WriteLine($"\"{Escape(item.Subject)}\",\"{Escape(item.SentTo)}\"," +
                        $"{item.SentDate:yyyy-MM-dd HH:mm},{item.FollowUpDate:yyyy-MM-dd}," +
                        $"{item.Status},{snooze}");
                }
            }
        }

        private static string Escape(string s) => s?.Replace("\"", "\"\"") ?? "";

        // ===== Persistence =====

        private static void Load()
        {
            try
            {
                if (!File.Exists(DataFile)) return;
                string json = File.ReadAllText(DataFile);
                Items = ParseJson(json);
            }
            catch
            {
                Items = new List<FollowUpItem>();
            }
        }

        private static void Save()
        {
            try
            {
                Directory.CreateDirectory(DataDir);
                string json = SerializeToJson(Items);
                File.WriteAllText(DataFile, json);
            }
            catch { }
        }

        // Minimal JSON serialization (no external dependency)
        private static string SerializeToJson(List<FollowUpItem> items)
        {
            var lines = new List<string>();
            foreach (var item in items)
            {
                lines.Add($"  {{\"entryId\":\"{Esc(item.EntryId)}\",\"conversationId\":\"{Esc(item.ConversationId)}\"," +
                    $"\"subject\":\"{Esc(item.Subject)}\",\"sentTo\":\"{Esc(item.SentTo)}\"," +
                    $"\"sentDate\":\"{item.SentDate:O}\",\"followUpDate\":\"{item.FollowUpDate:O}\"," +
                    $"\"status\":\"{item.Status}\"," +
                    $"\"snoozedUntil\":\"{(item.SnoozedUntil.HasValue ? item.SnoozedUntil.Value.ToString("O") : "")}\"," +
                    $"\"resolvedDate\":\"{(item.ResolvedDate.HasValue ? item.ResolvedDate.Value.ToString("O") : "")}\"}}");
            }
            return "{\n  \"items\": [\n" + string.Join(",\n", lines) + "\n  ]\n}";
        }

        private static string Esc(string s) => s?.Replace("\\", "\\\\").Replace("\"", "\\\"") ?? "";

        private static List<FollowUpItem> ParseJson(string json)
        {
            var items = new List<FollowUpItem>();
            try
            {
                // Minimal parser: split by "entryId" entries
                var entries = json.Split(new[] { "\"entryId\"" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var entry in entries)
                {
                    try
                    {
                        var item = new FollowUpItem();
                        item.EntryId = ExtractString(entry, "entryId");
                        item.ConversationId = ExtractString(entry, "conversationId");
                        item.Subject = ExtractString(entry, "subject");
                        item.SentTo = ExtractString(entry, "sentTo");
                        item.SentDate = ExtractDateTime(entry, "sentDate");
                        item.FollowUpDate = ExtractDateTime(entry, "followUpDate");
                        item.Status = ParseStatus(ExtractString(entry, "status"));
                        item.SnoozedUntil = ExtractOptionalDateTime(entry, "snoozedUntil");
                        item.ResolvedDate = ExtractOptionalDateTime(entry, "resolvedDate");
                        items.Add(item);
                    }
                    catch { }
                }
            }
            catch { }
            return items;
        }

        private static string ExtractString(string json, string key)
        {
            int idx = json.IndexOf($"\"{key}\":\"");
            if (idx < 0) return "";
            int start = idx + key.Length + 4;
            int end = json.IndexOf("\"", start);
            if (end < 0) return "";
            return json.Substring(start, end - start).Replace("\\\"", "\"").Replace("\\\\", "\\");
        }

        private static DateTime ExtractDateTime(string json, string key)
        {
            string val = ExtractString(json, key);
            if (DateTime.TryParse(val, out DateTime result)) return result;
            return DateTime.MinValue;
        }

        private static DateTime? ExtractOptionalDateTime(string json, string key)
        {
            string val = ExtractString(json, key);
            if (string.IsNullOrEmpty(val)) return null;
            if (DateTime.TryParse(val, out DateTime result)) return result;
            return null;
        }

        private static FollowUpStatus ParseStatus(string s)
        {
            if (Enum.TryParse<FollowUpStatus>(s, out var result)) return result;
            return FollowUpStatus.Pending;
        }
    }

    public enum FollowUpStatus
    {
        Pending,
        Overdue,
        Snoozed,
        Resolved
    }

    public class FollowUpItem
    {
        public string EntryId { get; set; }
        public string ConversationId { get; set; }
        public string Subject { get; set; }
        public string SentTo { get; set; }
        public DateTime SentDate { get; set; }
        public DateTime FollowUpDate { get; set; }
        public FollowUpStatus Status { get; set; }
        public DateTime? SnoozedUntil { get; set; }
        public DateTime? ResolvedDate { get; set; }
    }
}
