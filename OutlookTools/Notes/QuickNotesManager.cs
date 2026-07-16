using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OutlookTools.Notes
{
    /// <summary>
    /// OutlookTools — Quick Notes Manager
    /// A lightweight note-taking system attached to Outlook.
    /// 
    /// Features:
    /// - Create, edit, delete notes
    /// - Tag notes with email subjects
    /// - Search notes
    /// - Color coding
    /// - Markdown-like formatting (bold, headers)
    /// 
    /// Storage: %LOCALAPPDATA%\OutlookTools\notes.json
    /// </summary>
    public class QuickNotesManager
    {
        private static readonly string DataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookTools");
        private static readonly string DataFile = Path.Combine(DataDir, "notes.json");

        public static List<QuickNote> Notes { get; private set; } = new List<QuickNote>();

        static QuickNotesManager()
        {
            Load();
        }

        public static QuickNote Create(string title, string content, string emailTag = "")
        {
            var note = new QuickNote
            {
                Id = Guid.NewGuid().ToString("N").Substring(0, 8),
                Title = title,
                Content = content,
                EmailTag = emailTag,
                Color = "#3b82f6",
                Created = DateTime.Now,
                Modified = DateTime.Now
            };
            Notes.Insert(0, note);
            Save();
            return note;
        }

        public static void Update(QuickNote note)
        {
            note.Modified = DateTime.Now;
            Save();
        }

        public static void Delete(string id)
        {
            Notes.RemoveAll(n => n.Id == id);
            Save();
        }

        public static List<QuickNote> Search(string query)
        {
            string q = query.ToLower();
            return Notes.Where(n =>
                (n.Title?.ToLower().Contains(q) ?? false) ||
                (n.Content?.ToLower().Contains(q) ?? false) ||
                (n.EmailTag?.ToLower().Contains(q) ?? false))
                .ToList();
        }

        // ===== Persistence (minimal JSON) =====

        private static void Load()
        {
            try
            {
                if (!File.Exists(DataFile)) { Notes = new List<QuickNote>(); return; }
                string json = File.ReadAllText(DataFile);
                Notes = ParseNotes(json);
            }
            catch { Notes = new List<QuickNote>(); }
        }

        private static void Save()
        {
            try
            {
                Directory.CreateDirectory(DataDir);
                string json = SerializeNotes(Notes);
                File.WriteAllText(DataFile, json);
            }
            catch { }
        }

        private static string SerializeNotes(List<QuickNote> notes)
        {
            var lines = notes.Select(n =>
                $"  {{\"id\":\"{n.Id}\",\"title\":\"{Esc(n.Title)}\",\"content\":\"{Esc(n.Content)}\"," +
                $"\"emailTag\":\"{Esc(n.EmailTag)}\",\"color\":\"{n.Color}\"," +
                $"\"created\":\"{n.Created:O}\",\"modified\":\"{n.Modified:O}\"}}");
            return "{\n  \"notes\": [\n" + string.Join(",\n", lines) + "\n  ]\n}";
        }

        private static string Esc(string s) => s?.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n") ?? "";

        private static List<QuickNote> ParseNotes(string json)
        {
            var notes = new List<QuickNote>();
            var entries = json.Split(new[] { "\"id\":\"" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var entry in entries.Skip(1)) // skip first split part
            {
                try
                {
                    var note = new QuickNote
                    {
                        Id = Extract(entry, "id"),
                        Title = Extract(entry, "title"),
                        Content = Extract(entry, "content").Replace("\\n", "\n"),
                        EmailTag = Extract(entry, "emailTag"),
                        Color = Extract(entry, "color"),
                        Created = TryParseDate(Extract(entry, "created")),
                        Modified = TryParseDate(Extract(entry, "modified"))
                    };
                    notes.Add(note);
                }
                catch { }
            }
            return notes;
        }

        private static string Extract(string s, string key)
        {
            int idx = s.IndexOf($"\"{key}\":\"");
            if (idx < 0) return "";
            int start = idx + key.Length + 4;
            int end = s.IndexOf("\"", start);
            if (end < 0) return "";
            return s.Substring(start, end - start)
                .Replace("\\\"", "\"")
                .Replace("\\\\", "\\")
                .Replace("\\n", "\n");
        }

        private static DateTime TryParseDate(string s)
        {
            return DateTime.TryParse(s, out DateTime d) ? d : DateTime.MinValue;
        }
    }

    public class QuickNote
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Content { get; set; }
        public string EmailTag { get; set; }
        public string Color { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
    }
}
