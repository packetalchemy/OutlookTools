using System;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookTools.Digest
{
    /// <summary>
    /// OutlookTools — Daily Digest Display Form
    /// Shows the digest in a WebBrowser control with action buttons.
    /// </summary>
    public partial class DailyDigestForm : Form
    {
        private WebBrowser _webBrowser;
        private Button _btnRefresh;
        private Button _btnSendAsEmail;
        private StatusStrip _statusBar;
        private Label _lblStatus;

        public DailyDigestForm()
        {
            InitializeComponent();
            SetupUI();
            LoadDigest();
        }

        public static void Show()
        {
            new DailyDigestForm().ShowDialog();
        }

        private void SetupUI()
        {
            this.Text = "OutlookTools — Daily Digest";
            this.Size = new Size(700, 600);
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

            _btnRefresh = new Button
            {
                Text = "🔄 Refresh", Width = 100, Height = 30,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            _btnRefresh.Click += (s, e) => LoadDigest();
            btnPanel.Controls.Add(_btnRefresh);

            _btnSendAsEmail = new Button
            {
                Text = "📧 Send as Email", Width = 130, Height = 30,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            _btnSendAsEmail.Click += (s, e) => SendAsEmail();
            btnPanel.Controls.Add(_btnSendAsEmail);

            this.Controls.Add(btnPanel);

            // WebBrowser for HTML display
            _webBrowser = new WebBrowser
            {
                Dock = DockStyle.Fill,
                IsWebBrowserContextMenuEnabled = false,
                ScriptErrorsSuppressed = true
            };
            this.Controls.Add(_webBrowser);

            // Status bar
            _statusBar = new StatusStrip();
            _lblStatus = new ToolStripStatusLabel("Ready");
            _statusBar.Items.Add(_lblStatus);
            this.Controls.Add(_statusBar);
        }

        private DigestData _currentData;

        private void LoadDigest()
        {
            try
            {
                _currentData = DailyDigestGenerator.Generate(Globals.ThisAddIn.Application);
                string html = DailyDigestGenerator.FormatAsHtml(_currentData);
                _webBrowser.DocumentText = html;
                _lblStatus.Text = $"Generated at {_currentData.GeneratedAt:HH:mm}";
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "Error: " + ex.Message;
            }
        }

        private void SendAsEmail()
        {
            if (_currentData == null) return;

            try
            {
                var app = Globals.ThisAddIn.Application;
                Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                if (mail != null)
                {
                    mail.Subject = $"📊 Daily Digest — {DateTime.Now:yyyy-MM-dd}";
                    mail.HTMLBody = DailyDigestGenerator.FormatAsHtml(_currentData);
                    mail.Display(false);
                    _lblStatus.Text = "✅ Opened in new message";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed: " + ex.Message);
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(700, 600);
            this.Name = "DailyDigestForm";
            this.Text = "OutlookTools — Daily Digest";
            this.ResumeLayout(false);
        }
    }
}
