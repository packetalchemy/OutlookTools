using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTools.Stats
{
    /// <summary>
    /// OutlookTools — Email Statistics Dashboard
    /// Analyzes email patterns and generates reports:
    /// - Top senders/receivers
    /// - Emails by hour/day/month
    /// - Busiest periods
    /// - Category breakdown
    /// - Attachment statistics
    /// - Export to CSV
    /// 
    /// Pure local analysis. NO network calls.
    /// </summary>
    public partial class EmailStatsForm : Form
    {
        private TabControl _tabs;
        private ListView _lvTopSenders;
        private ListView _lvByHour;
        private ListView _lvByCategory;
        private ListView _lvBySize;
        private ComboBox _cboPeriod;
        private Button _btnAnalyze;
        private Button _btnExport;

        private Label _lblStatus;
        private Label _lblSummary;

        private DateTime _periodFrom;
        private DateTime _periodTo;

        // Collected data
        private Dictionary<string, int> _senderCounts = new Dictionary<string, int>();
        private Dictionary<int, int> _hourlyCounts = new Dictionary<int, int>();
        private Dictionary<string, int> _categoryCounts = new Dictionary<string, int>();
        private long _totalSize;
        private int _totalEmails;
        private int _totalWithAttachments;
        private int _totalWithBody;

        public EmailStatsForm()
        {
            InitializeComponent();
            SetupUI();
            SetDefaultPeriod();
        }

        public static void ShowForm()
        {
            var form = new EmailStatsForm();
            form.ShowDialog();
        }

        private void SetupUI()
        {
            this.Text = "OutlookTools — Email Statistics";
            this.Size = new Size(900, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9);

            // Top panel: period selection
            var topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                Padding = new Padding(10)
            };

            topPanel.Controls.Add(new Label { Text = "Period:", Location = new Point(10, 18), Width = 50 });

            _cboPeriod = new ComboBox
            {
                Location = new Point(65, 15),
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cboPeriod.Items.AddRange(new object[] {
                "Last 7 days", "Last 30 days", "Last 90 days",
                "Last 6 months", "Last year", "All time"
            });
            _cboPeriod.SelectedIndex = 1; // Default: Last 30 days
            topPanel.Controls.Add(_cboPeriod);

            _btnAnalyze = new Button
            {
                Text = "📊 Analyze",
                Location = new Point(230, 13),
                Width = 100,
                Height = 30,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnAnalyze.Click += (s, e) => Analyze();
            topPanel.Controls.Add(_btnAnalyze);

            _btnExport = new Button
            {
                Text = "📥 Export CSV",
                Location = new Point(340, 13),
                Width = 110,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnExport.Click += (s, e) => ExportStats();
            topPanel.Controls.Add(_btnExport);

            this.Controls.Add(topPanel);

            // Summary panel
            var summaryPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = Color.FromArgb(240, 248, 255),
                Padding = new Padding(10)
            };

            _lblSummary = new Label
            {
                Text = "Click 'Analyze' to generate statistics",
                Location = new Point(10, 15),
                Width = 800,
                Font = new Font("Segoe UI", 10)
            };
            summaryPanel.Controls.Add(_lblSummary);
            this.Controls.Add(summaryPanel);

            // Tab control with different stats views
            _tabs = new TabControl { Dock = DockStyle.Fill };

            // Tab 1: Top Senders
            var tabSenders = new TabPage("📧 Top Senders");
            _lvTopSenders = CreateListView(new[] { "Rank", "Sender", "Count", "% of Total" });
            tabSenders.Controls.Add(_lvTopSenders);
            _tabs.TabPages.Add(tabSenders);

            // Tab 2: By Hour
            var tabHourly = new TabPage("⏰ By Hour");
            _lvByHour = CreateListView(new[] { "Hour", "Count", "Bar" });
            tabHourly.Controls.Add(_lvByHour);
            _tabs.TabPages.Add(tabHourly);

            // Tab 3: By Category
            var tabCategory = new TabPage("🏷️ By Category");
            _lvByCategory = CreateListView(new[] { "Category", "Count", "% of Total" });
            tabCategory.Controls.Add(_lvByCategory);
            _tabs.TabPages.Add(tabCategory);

            // Tab 4: By Size
            var tabSize = new TabPage("📦 By Size");
            _lvBySize = CreateListView(new[] { "Size Range", "Count", "% of Total" });
            tabSize.Controls.Add(_lvBySize);
            _tabs.TabPages.Add(tabSize);

            this.Controls.Add(_tabs);

            // Status bar (Label-based)
            _lblStatus = new Label { Text = "Ready", Dock = DockStyle.Bottom, Height = 25, TextAlign = System.Drawing.ContentAlignment.MiddleLeft, Padding = new Padding(5, 0, 0, 0) };
            this.Controls.Add(_lblStatus);
        }

        private void SetDefaultPeriod()
        {
            _periodFrom = DateTime.Now.AddDays(-30);
            _periodTo = DateTime.Now;
        }

        private void Analyze()
        {
            _lblStatus.Text = "Analyzing...";
            _lblSummary.Text = "Processing...";
            Application.DoEvents();

            try
            {
                // Reset counters
                _senderCounts.Clear();
                _hourlyCounts.Clear();
                _categoryCounts.Clear();
                _totalSize = 0;
                _totalEmails = 0;
                _totalWithAttachments = 0;
                _totalWithBody = 0;

                // Set period
                SetPeriod();

                Outlook.Application app = Globals.ThisAddIn.Application;
                Outlook.NameSpace session = app.Session;

                // Analyze Inbox
                Outlook.Folder inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                if (inbox != null) AnalyzeFolder(inbox);

                // Analyze Sent
                Outlook.Folder sent = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail) as Outlook.Folder;
                if (sent != null) AnalyzeFolder(sent);

                // Display results
                DisplayResults();
                _lblStatus.Text = $"✅ Analyzed {_totalEmails} emails ({_periodFrom:yyyy-MM-dd} to {_periodTo:yyyy-MM-dd})";
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "Error: " + ex.Message;
                _lblSummary.Text = "Analysis failed: " + ex.Message;
            }
        }

        private void SetPeriod()
        {
            switch (_cboPeriod.SelectedIndex)
            {
                case 0: _periodFrom = DateTime.Now.AddDays(-7); break;
                case 1: _periodFrom = DateTime.Now.AddDays(-30); break;
                case 2: _periodFrom = DateTime.Now.AddDays(-90); break;
                case 3: _periodFrom = DateTime.Now.AddMonths(-6); break;
                case 4: _periodFrom = DateTime.Now.AddYears(-1); break;
                case 5: _periodFrom = DateTime.MinValue; break;
                default: _periodFrom = DateTime.Now.AddDays(-30); break;
            }
            _periodTo = DateTime.Now;
        }

        private void AnalyzeFolder(Outlook.Folder folder)
        {
            try
            {
                Outlook.Items items = folder.Items;
                items.Sort("[ReceivedTime]", true);

                foreach (object obj in items)
                {
                    if (!(obj is Outlook.MailItem mail)) continue;

                    try
                    {
                        // Filter by date
                        if (mail.ReceivedTime < _periodFrom || mail.ReceivedTime > _periodTo) continue;

                        _totalEmails++;
                        _totalSize += mail.Size;

                        // Count by sender
                        string sender = mail.SenderName ?? "Unknown";
                        if (_senderCounts.ContainsKey(sender))
                            _senderCounts[sender]++;
                        else
                            _senderCounts[sender] = 1;

                        // Count by hour
                        int hour = mail.ReceivedTime.Hour;
                        if (_hourlyCounts.ContainsKey(hour))
                            _hourlyCounts[hour]++;
                        else
                            _hourlyCounts[hour] = 1;

                        // Count by category
                        string cat = mail.Categories ?? "Uncategorized";
                        foreach (var c in cat.Split(','))
                        {
                            string trimmed = c.Trim();
                            if (string.IsNullOrEmpty(trimmed)) continue;
                            if (_categoryCounts.ContainsKey(trimmed))
                                _categoryCounts[trimmed]++;
                            else
                                _categoryCounts[trimmed] = 1;
                        }

                        // Attachment stats
                        if (mail.Attachments.Count > 0) _totalWithAttachments++;
                        if (!string.IsNullOrEmpty(mail.Body)) _totalWithBody++;
                    }
                    catch { }
                    finally { ReleaseCom(mail); }
                }
            }
            catch { }
        }

        private void DisplayResults()
        {
            // Summary
            string avgSize = _totalEmails > 0 ? FormatSize(_totalSize / _totalEmails) : "0 KB";
            _lblSummary.Text = $"📧 {_totalEmails} emails | " +
                               $"📎 {_totalWithAttachments} with attachments | " +
                               $"💾 Total: {FormatSize(_totalSize)} | " +
                               $"📏 Avg: {avgSize}";

            // Top Senders
            _lvTopSenders.Items.Clear();
            int rank = 1;
            foreach (var kv in _senderCounts.OrderByDescending(x => x.Value).Take(20))
            {
                double pct = _totalEmails > 0 ? (double)kv.Value / _totalEmails * 100 : 0;
                var item = new ListViewItem(rank.ToString());
                item.SubItems.Add(kv.Key);
                item.SubItems.Add(kv.Value.ToString());
                item.SubItems.Add($"{pct:F1}%");
                _lvTopSenders.Items.Add(item);
                rank++;
            }

            // By Hour
            _lvByHour.Items.Clear();
            int maxHourly = _hourlyCounts.Values.DefaultIfEmpty(0).Max();
            for (int h = 0; h < 24; h++)
            {
                int count = _hourlyCounts.ContainsKey(h) ? _hourlyCounts[h] : 0;
                string bar = maxHourly > 0 ? new string('█', (int)((double)count / maxHourly * 30)) : "";
                var item = new ListViewItem($"{h:D2}:00");
                item.SubItems.Add(count.ToString());
                item.SubItems.Add(bar);
                if (h >= 9 && h <= 17)
                    item.BackColor = Color.FromArgb(230, 255, 230); // Work hours highlight
                _lvByHour.Items.Add(item);
            }

            // By Category
            _lvByCategory.Items.Clear();
            foreach (var kv in _categoryCounts.OrderByDescending(x => x.Value).Take(20))
            {
                double pct = _totalEmails > 0 ? (double)kv.Value / _totalEmails * 100 : 0;
                var item = new ListViewItem(kv.Key);
                item.SubItems.Add(kv.Value.ToString());
                item.SubItems.Add($"{pct:F1}%");
                _lvByCategory.Items.Add(item);
            }

            // By Size
            _lvBySize.Items.Clear();
            var sizeRanges = new Dictionary<string, int>
            {
                { "0 - 10 KB", 0 }, { "10 - 50 KB", 0 }, { "50 - 100 KB", 0 },
                { "100 - 500 KB", 0 }, { "500 KB - 1 MB", 0 }, { "1 - 5 MB", 0 },
                { "5 - 10 MB", 0 }, { "> 10 MB", 0 }
            };

            // Re-analyze for size (or store during first pass)
            Outlook.Application app = Globals.ThisAddIn.Application;
            Outlook.Folder inbox = app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
            if (inbox != null)
            {
                try
                {
                    foreach (object obj in inbox.Items)
                    {
                        if (!(obj is Outlook.MailItem mail)) continue;
                        try
                        {
                            if (mail.ReceivedTime < _periodFrom || mail.ReceivedTime > _periodTo) continue;
                            string range = GetSizeRange(mail.Size);
                            sizeRanges[range]++;
                        }
                        catch { }
                        finally { ReleaseCom(mail); }
                    }
                }
                catch { }
            }

            foreach (var kv in sizeRanges.Where(x => x.Value > 0))
            {
                double pct = _totalEmails > 0 ? (double)kv.Value / _totalEmails * 100 : 0;
                var item = new ListViewItem(kv.Key);
                item.SubItems.Add(kv.Value.ToString());
                item.SubItems.Add($"{pct:F1}%");
                _lvBySize.Items.Add(item);
            }
        }

        private string GetSizeRange(long size)
        {
            if (size < 10240) return "0 - 10 KB";
            if (size < 51200) return "10 - 50 KB";
            if (size < 102400) return "50 - 100 KB";
            if (size < 512000) return "100 - 500 KB";
            if (size < 1048576) return "500 KB - 1 MB";
            if (size < 5242880) return "1 - 5 MB";
            if (size < 10485760) return "5 - 10 MB";
            return "> 10 MB";
        }

        private string FormatSize(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1048576) return $"{bytes / 1024.0:F1} KB";
            if (bytes < 1073741824) return $"{bytes / 1048576.0:F1} MB";
            return $"{bytes / 1073741824.0:F2} GB";
        }

        private void ExportStats()
        {
            using (var sfd = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                FileName = $"EmailTools_Stats_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                try
                {
                    using (var writer = new System.IO.StreamWriter(sfd.FileName, false, System.Text.Encoding.UTF8))
                    {
                        writer.WriteLine("=== OutlookTools Email Statistics ===");
                        writer.WriteLine($"Period: {_periodFrom:yyyy-MM-dd} to {_periodTo:yyyy-MM-dd}");
                        writer.WriteLine($"Total Emails: {_totalEmails}");
                        writer.WriteLine($"With Attachments: {_totalWithAttachments}");
                        writer.WriteLine($"Total Size: {FormatSize(_totalSize)}");
                        writer.WriteLine();

                        writer.WriteLine("--- Top Senders ---");
                        writer.WriteLine("Rank,Sender,Count,Percentage");
                        int rank = 1;
                        foreach (var kv in _senderCounts.OrderByDescending(x => x.Value).Take(20))
                        {
                            double pct = _totalEmails > 0 ? (double)kv.Value / _totalEmails * 100 : 0;
                            writer.WriteLine($"{rank},\"{kv.Key}\",{kv.Value},{pct:F1}%");
                            rank++;
                        }

                        writer.WriteLine();
                        writer.WriteLine("--- By Hour ---");
                        writer.WriteLine("Hour,Count");
                        for (int h = 0; h < 24; h++)
                        {
                            int count = _hourlyCounts.ContainsKey(h) ? _hourlyCounts[h] : 0;
                            writer.WriteLine($"{h:D2}:00,{count}");
                        }

                        writer.WriteLine();
                        writer.WriteLine("--- By Category ---");
                        writer.WriteLine("Category,Count");
                        foreach (var kv in _categoryCounts.OrderByDescending(x => x.Value))
                            writer.WriteLine($"\"{kv.Key}\",{kv.Value}");
                    }

                    var result = MessageBox.Show(
                        $"✅ Stats exported!\n\nOpen file?",
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

        private static ListView CreateListView(string[] columns)
        {
            var lv = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true
            };
            foreach (var col in columns)
                lv.Columns.Add(col, 120);
            return lv;
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

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(900, 650);
            this.Name = "EmailStatsForm";
            this.Text = "OutlookTools — Email Statistics";
            this.ResumeLayout(false);
        }
    }
}
