using System;
using System.Drawing;
using System.Windows.Forms;
using OutlookTools.Settings;

namespace OutlookTools.Settings
{
    /// <summary>
    /// OutlookTools — Settings Window
    /// Configures: archive age threshold, archive schedule, debug logging, startup notification.
    /// All settings are per-user (stored in AppData).
    /// </summary>
    public partial class SettingsForm : Form
    {
        private NumericUpDown _nudArchiveDays;
        private NumericUpDown _nudArchiveHour;
        private CheckBox _chkAutoArchive;
        private CheckBox _chkAutoReminder;
        private CheckBox _chkDebugLog;
        private CheckBox _chkStartupNotification;
        private Button _btnSave;
        private Button _btnCancel;
        private Button _btnOpenLog;
        private Label _lblStatus;

        public SettingsForm()
        {
            InitializeComponent();
            SetupUI();
            LoadSettings();
        }

        public static void Show()
        {
            var form = new SettingsForm();
            form.ShowDialog();
        }

        private void SetupUI()
        {
            this.Text = "OutlookTools — Settings";
            this.Size = new Size(460, 520);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Segoe UI", 9);
            this.RightToLeft = RightToLeft.No;

            int y = 20;

            // Section: Archive
            AddSectionLabel("📅 Archive Settings", ref y);

            AddLabel("Archive messages older than (days):", y);
            _nudArchiveDays = new NumericUpDown
            {
                Location = new Point(250, y - 2),
                Width = 80,
                Minimum = 7,
                Maximum = 365,
                Value = 90
            };
            this.Controls.Add(_nudArchiveDays);
            y += 36;

            _chkAutoArchive = new CheckBox
            {
                Text = "Enable auto-archive (runs daily)",
                Location = new Point(20, y),
                Width = 300,
                Checked = true
            };
            this.Controls.Add(_chkAutoArchive);
            y += 28;

            AddLabel("Archive hour (0–23):", y);
            _nudArchiveHour = new NumericUpDown
            {
                Location = new Point(250, y - 2),
                Width = 80,
                Minimum = 0,
                Maximum = 23,
                Value = 6
            };
            this.Controls.Add(_nudArchiveHour);
            y += 40;

            // Section: Cleanup
            AddSectionLabel("🔔 Reminder Cleanup", ref y);

            _chkAutoReminder = new CheckBox
            {
                Text = "Auto-dismiss past-due reminders (every 30 min)",
                Location = new Point(20, y),
                Width = 380,
                Checked = true
            };
            this.Controls.Add(_chkAutoReminder);
            y += 40;

            // Section: Privacy
            AddSectionLabel("🔒 Privacy & Debug", ref y);

            _chkDebugLog = new CheckBox
            {
                Text = "Enable debug log (LOCAL ONLY — nothing uploaded)",
                Location = new Point(20, y),
                Width = 380,
                Checked = false
            };
            this.Controls.Add(_chkDebugLog);
            y += 28;

            _chkStartupNotification = new CheckBox
            {
                Text = "Show startup notification",
                Location = new Point(20, y),
                Width = 380,
                Checked = false
            };
            this.Controls.Add(_chkStartupNotification);
            y += 40;

            // Buttons
            _btnSave = new Button
            {
                Text = "💾 Save",
                Location = new Point(20, y),
                Width = 100,
                Height = 32,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnSave.Click += (s, e) => SaveSettings();
            this.Controls.Add(_btnSave);

            _btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(130, y),
                Width = 80,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(_btnCancel);

            _btnOpenLog = new Button
            {
                Text = "📄 Open Log Folder",
                Location = new Point(250, y),
                Width = 130,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnOpenLog.Click += (s, e) =>
            {
                string dir = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OutlookTools");
                if (System.IO.Directory.Exists(dir))
                    System.Diagnostics.Process.Start("explorer.exe", dir);
                else
                    MessageBox.Show("No log folder yet.", "OutlookTools");
            };
            this.Controls.Add(_btnOpenLog);

            y += 50;
            _lblStatus = new Label
            {
                Text = "",
                Location = new Point(20, y),
                Width = 400,
                ForeColor = Color.Green
            };
            this.Controls.Add(_lblStatus);
        }

        private void LoadSettings()
        {
            _nudArchiveDays.Value = SettingsManager.GetArchiveAgeDays();
            _nudArchiveHour.Value = SettingsManager.GetArchiveHour();
            _chkAutoArchive.Checked = SettingsManager.GetAutoArchiveEnabled();
            _chkAutoReminder.Checked = SettingsManager.GetAutoReminderEnabled();
            _chkDebugLog.Checked = SettingsManager.GetDebugLogEnabled();
            _chkStartupNotification.Checked = SettingsManager.GetStartupNotification();
        }

        private void SaveSettings()
        {
            SettingsManager.SetArchiveAgeDays((int)_nudArchiveDays.Value);
            SettingsManager.SetArchiveHour((int)_nudArchiveHour.Value);
            SettingsManager.SetAutoArchiveEnabled(_chkAutoArchive.Checked);
            SettingsManager.SetAutoReminderEnabled(_chkAutoReminder.Checked);
            SettingsManager.SetDebugLogEnabled(_chkDebugLog.Checked);
            SettingsManager.SetStartupNotification(_chkStartupNotification.Checked);

            _lblStatus.Text = "✅ Settings saved!";
            _lblStatus.ForeColor = Color.Green;
        }

        private Label AddSectionLabel(string text, ref int y)
        {
            var label = new Label
            {
                Text = text,
                Location = new Point(20, y),
                Width = 400,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 215)
            };
            this.Controls.Add(label);
            y += 28;
            return label;
        }

        private void AddLabel(string text, int y)
        {
            var label = new Label
            {
                Text = text,
                Location = new Point(20, y + 3),
                Width = 230,
                TextAlign = ContentAlignment.MiddleRight
            };
            this.Controls.Add(label);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 481);
            this.Name = "SettingsForm";
            this.Text = "OutlookTools — Settings";
            this.ResumeLayout(false);
        }
    }
}
