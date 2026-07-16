using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookTools.FollowUp
{
    /// <summary>
    /// OutlookTools — Follow-up Dashboard
    /// Shows pending follow-ups with actions: Snooze, Resolve, Remove, Export.
    /// </summary>
    public partial class FollowUpDashboardForm : Form
    {
        private ListView _listView;
        private Button _btnResolve;
        private Button _btnSnooze;
        private Button _btnRemove;
        private Button _btnExport;
        private Button _btnRefresh;
        private StatusStrip _statusBar;
        private Label _lblStatus;

        public FollowUpDashboardForm()
        {
            InitializeComponent();
            SetupUI();
            LoadItems();
        }

        public static void Show()
        {
            new FollowUpDashboardForm().ShowDialog();
        }

        private void SetupUI()
        {
            this.Text = "OutlookTools — Follow-up Tracker";
            this.Size = new Size(850, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9);

            // Button bar
            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 50,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(8)
            };

            _btnRefresh = Btn("🔄 Refresh", 100);
            _btnRefresh.Click += (s, e) => LoadItems();
            btnPanel.Controls.Add(_btnRefresh);

            _btnExport = Btn("📥 Export CSV", 110);
            _btnExport.Click += (s, e) => Export();
            btnPanel.Controls.Add(_btnExport);

            _btnRemove = Btn("🗑️ Remove", 90);
            _btnRemove.BackColor = Color.FromArgb(220, 53, 69);
            _btnRemove.ForeColor = Color.White;
            _btnRemove.Click += (s, e) => RemoveSelected();
            btnPanel.Controls.Add(_btnRemove);

            _btnSnooze = Btn("😴 Snooze 1 Day", 120);
            _btnSnooze.Click += (s, e) => SnoozeSelected(1);
            btnPanel.Controls.Add(_btnSnooze);

            _btnResolve = Btn("✅ Mark Resolved", 140);
            _btnResolve.BackColor = Color.FromArgb(34, 197, 94);
            _btnResolve.ForeColor = Color.White;
            _btnResolve.Click += (s, e) => ResolveSelected();
            btnPanel.Controls.Add(_btnResolve);

            this.Controls.Add(btnPanel);

            // List
            _listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = false
            };
            _listView.Columns.Add("Subject", 250);
            _listView.Columns.Add("Sent To", 150);
            _listView.Columns.Add("Sent Date", 130);
            _listView.Columns.Add("Follow-up By", 130);
            _listView.Columns.Add("Status", 100);
            _listView.DoubleClick += (s, e) => OpenSelected();
            this.Controls.Add(_listView);

            // Status bar
            _statusBar = new StatusStrip();
            _lblStatus = new ToolStripStatusLabel("Ready");
            _statusBar.Items.Add(_lblStatus);
            this.Controls.Add(_statusBar);
        }

        private void LoadItems()
        {
            _listView.Items.Clear();
            var items = FollowUpTracker.GetActive();

            foreach (var item in items)
            {
                var lvItem = new ListViewItem(item.Subject ?? "(no subject)");
                lvItem.SubItems.Add(item.SentTo);
                lvItem.SubItems.Add(item.SentDate.ToString("yyyy-MM-dd HH:mm"));
                lvItem.SubItems.Add(item.FollowUpDate.ToString("yyyy-MM-dd"));

                string statusText;
                Color statusColor;
                switch (item.Status)
                {
                    case FollowUpStatus.Overdue:
                        statusText = "🔴 OVERDUE";
                        statusColor = Color.FromArgb(255, 230, 230);
                        break;
                    case FollowUpStatus.Snoozed:
                        statusText = $"😴 Snoozed → {item.SnoozedUntil:MM/dd}";
                        statusColor = Color.FromArgb(255, 255, 230);
                        break;
                    default:
                        statusText = "⏳ Pending";
                        statusColor = Color.White;
                        break;
                }
                lvItem.SubItems.Add(statusText);
                lvItem.BackColor = statusColor;
                lvItem.Tag = item;
                _listView.Items.Add(lvItem);
            }

            int overdue = items.Count(x => x.Status == FollowUpStatus.Overdue);
            _lblStatus.Text = $"{items.Count} active follow-ups" +
                (overdue > 0 ? $" ({overdue} overdue)" : "");
        }

        private void ResolveSelected()
        {
            if (_listView.SelectedItems.Count == 0) return;
            var item = (FollowUpItem)_listView.SelectedItems[0].Tag;
            FollowUpTracker.MarkResolved(item.EntryId);
            LoadItems();
        }

        private void SnoozeSelected(int days)
        {
            if (_listView.SelectedItems.Count == 0) return;
            var item = (FollowUpItem)_listView.SelectedItems[0].Tag;
            FollowUpTracker.Snooze(item.EntryId, DateTime.Now.AddDays(days));
            LoadItems();
        }

        private void RemoveSelected()
        {
            if (_listView.SelectedItems.Count == 0) return;
            var item = (FollowUpItem)_listView.SelectedItems[0].Tag;
            FollowUpTracker.Remove(item.EntryId);
            LoadItems();
        }

        private void OpenSelected()
        {
            if (_listView.SelectedItems.Count == 0) return;
            var item = (FollowUpItem)_listView.SelectedItems[0].Tag;
            try
            {
                Outlook.Application app = Globals.ThisAddIn.Application;
                Outlook.MailItem mail = app.Session.GetItemFromID(item.EntryId) as Outlook.MailItem;
                mail?.Display(false);
            }
            catch { }
        }

        private void Export()
        {
            using (var sfd = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                FileName = $"FollowUp_{DateTime.Now:yyyyMMdd}.csv"
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    FollowUpTracker.ExportToCsv(sfd.FileName);
                    _lblStatus.Text = "✅ Exported!";
                    System.Diagnostics.Process.Start("explorer.exe", sfd.FileName);
                }
            }
        }

        private static Button Btn(string text, int w) => new Button
        {
            Text = text, Width = w, Height = 30,
            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(2)
        };

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(850, 550);
            this.Name = "FollowUpDashboardForm";
            this.Text = "OutlookTools — Follow-up Tracker";
            this.ResumeLayout(false);
        }
    }
}
