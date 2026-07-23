using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTools.Commands
{
    /// <summary>
    /// OutlookTools — Bulk Actions
    /// Perform batch operations on selected emails:
    /// - Move selected to folder
    /// - Delete selected
    /// - Flag/unflag selected
    /// - Mark as read/unread
    /// - Add category
    /// - Export to CSV
    /// 
    /// Works on multi-select in Outlook Explorer.
    /// </summary>
    public static class BulkActions
    {
        /// <summary>
        /// Move selected messages to a specified folder.
        /// </summary>
        public static void MoveToFolder(Outlook.Application app)
        {
            var items = GetSelectedMailItems(app);
            if (items == null || items.Count == 0) return;

            Outlook.Folder target = PickFolder(app, "Select destination folder:");
            if (target == null) return;

            int moved = 0;
            foreach (var mail in items)
            {
                try { mail.Move(target); moved++; }
                catch { }
                finally { ReleaseCom(mail); }
            }

            MessageBox.Show($"✅ Moved {moved} messages to '{target.Name}'.",
                "Bulk Move", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Delete selected messages (with confirmation).
        /// </summary>
        public static void DeleteSelected(Outlook.Application app)
        {
            var items = GetSelectedMailItems(app);
            if (items == null || items.Count == 0) return;

            var result = MessageBox.Show(
                $"Delete {items.Count} selected message(s)?\n\nThis cannot be undone.",
                "Confirm Bulk Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result != DialogResult.Yes) return;

            int deleted = 0;
            foreach (var mail in items)
            {
                try { mail.Delete(); deleted++; }
                catch { }
                finally { ReleaseCom(mail); }
            }

            MessageBox.Show($"🗑️ Deleted {deleted} messages.",
                "Bulk Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Flag selected messages for follow-up.
        /// </summary>
        public static void FlagSelected(Outlook.Application app, bool followUp)
        {
            var items = GetSelectedMailItems(app);
            if (items == null || items.Count == 0) return;

            int count = 0;
            foreach (var mail in items)
            {
                try
                {
                    if (followUp)
                    {
                        mail.FlagStatus = (Outlook.OlFlagStatus)1; // olFlagForward = 1
                        mail.FlagRequest = "Follow up";
                    }
                    else
                    {
                        mail.FlagStatus = Outlook.OlFlagStatus.olNoFlag;
                        mail.FlagRequest = "";
                    }
                    mail.Save();
                    count++;
                }
                catch { }
                finally { ReleaseCom(mail); }
            }

            string action = followUp ? "Flagged" : "Unflagged";
            MessageBox.Show($"✅ {action} {count} messages.",
                "Bulk " + action, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Mark selected messages as read or unread.
        /// </summary>
        public static void MarkAsRead(Outlook.Application app, bool read)
        {
            var items = GetSelectedMailItems(app);
            if (items == null || items.Count == 0) return;

            int count = 0;
            foreach (var mail in items)
            {
                try
                {
                    mail.UnRead = !read;
                    mail.Save();
                    count++;
                }
                catch { }
                finally { ReleaseCom(mail); }
            }

            string action = read ? "read" : "unread";
            MessageBox.Show($"✅ Marked {count} messages as {action}.",
                "Bulk Action", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Add a category to selected messages.
        /// </summary>
        public static void AddCategory(Outlook.Application app)
        {
            var items = GetSelectedMailItems(app);
            if (items == null || items.Count == 0) return;

            using (var dialog = new InputDialog("Enter category name:", "Category"))
            {
                if (dialog.ShowDialog() != DialogResult.OK) return;

                string category = dialog.InputValue;
                if (string.IsNullOrWhiteSpace(category)) return;

                int count = 0;
                foreach (var mail in items)
                {
                    try
                    {
                        string existing = mail.Categories ?? "";
                        if (!existing.Contains(category))
                        {
                            mail.Categories = string.IsNullOrEmpty(existing)
                                ? category
                                : existing + ", " + category;
                            mail.Save();
                        }
                        count++;
                    }
                    catch { }
                    finally { ReleaseCom(mail); }
                }

                MessageBox.Show($"✅ Added category '{category}' to {count} messages.",
                    "Bulk Category", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Export selected messages to CSV file.
        /// </summary>
        public static void ExportToCsv(Outlook.Application app)
        {
            var items = GetSelectedMailItems(app);
            if (items == null || items.Count == 0) return;

            using (var sfd = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                FileName = $"OutlookTools_Export_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                try
                {
                    using (var writer = new System.IO.StreamWriter(sfd.FileName, false, System.Text.Encoding.UTF8))
                    {
                        writer.WriteLine("Date,From,To,Subject,HasAttachment,Size,Category");

                        foreach (var mail in items)
                        {
                            try
                            {
                                string date = mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm");
                                string from = EscapeCsv(mail.SenderName ?? "");
                                string to = EscapeCsv(string.Join("; ",
                                    mail.Recipients.Cast<Outlook.Recipient>().Select(r => r.Name)));
                                string subject = EscapeCsv(mail.Subject ?? "");
                                string hasAtt = mail.Attachments.Count > 0 ? "Yes" : "No";
                                string size = mail.Size.ToString();
                                string cat = EscapeCsv(mail.Categories ?? "");

                                writer.WriteLine($"{date},{from},{to},{subject},{hasAtt},{size},{cat}");
                            }
                            catch { }
                            finally { ReleaseCom(mail); }
                        }
                    }

                    var result = MessageBox.Show(
                        $"✅ Exported {items.Count} messages.\n\nOpen the file?",
                        "Export Complete",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (result == DialogResult.Yes)
                        System.Diagnostics.Process.Start("explorer.exe", sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Export failed: " + ex.Message);
                }
            }
        }

        // === Helpers ===

        private static List<Outlook.MailItem> GetSelectedMailItems(Outlook.Application app)
        {
            try
            {
                var explorer = app.ActiveExplorer();
                var selection = explorer.Selection;
                if (selection == null || selection.Count == 0) return null;

                var items = new List<Outlook.MailItem>();
                foreach (object obj in selection)
                {
                    if (obj is Outlook.MailItem mail)
                        items.Add(mail);
                }
                return items;
            }
            catch
            {
                return null;
            }
        }

        private static Outlook.Folder PickFolder(Outlook.Application app, string title)
        {
            try
            {
                Outlook.NameSpace session = app.Session;
                var picker = session.PickFolder();
                return picker as Outlook.Folder;
            }
            catch { return null; }
        }

        private static string EscapeCsv(string value)
        {
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return value;
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

    /// <summary>
    /// Simple input dialog for bulk category name.
    /// </summary>
    internal class InputDialog : Form
    {
        public string InputValue { get; private set; }
        private TextBox _input;

        public InputDialog(string prompt, string title)
        {
            this.Text = title;
            this.Size = new Size(350, 150);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var lbl = new Label
            {
                Text = prompt,
                Location = new Point(12, 15),
                Width = 310
            };
            this.Controls.Add(lbl);

            _input = new TextBox
            {
                Location = new Point(12, 45),
                Width = 310
            };
            _input.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    InputValue = _input.Text;
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            };
            this.Controls.Add(_input);

            var btnOk = new Button
            {
                Text = "OK",
                Location = new Point(180, 80),
                Width = 70,
                DialogResult = DialogResult.OK
            };
            btnOk.Click += (s, e) => { InputValue = _input.Text; };
            this.Controls.Add(btnOk);

            var btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(260, 80),
                Width = 70,
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
        }
    }
}
