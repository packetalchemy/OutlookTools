using System;
using System.Drawing;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTools.Search
{
    /// <summary>
    /// OutlookTools — Advanced Search Window
    /// A WinForms dialog that lets users search through:
    /// - All mailboxes (inbox + mounted archives)
    /// - Filter by: From, To, Subject, Body, Date range, Has attachment
    /// - Uses Outlook's built-in search via Items.Restrict() for reliability.
    /// 
    /// NO local body index (to keep this initial version simple and safe).
    /// The Outlook built-in index IS the index.
    /// </summary>
    public partial class AdvancedSearchForm : Form
    {
        private ListView _resultsList;
        private WebBrowser _previewPane;
        private TextBox _txtFrom;
        private TextBox _txtTo;
        private TextBox _txtSubject;
        private TextBox _txtBody;
        private DateTimePicker _dtpFrom;
        private DateTimePicker _dtpTo;
        private CheckBox _chkHasAttachment;
        private Button _btnSearch;
        private Button _btnSelectAll;
        private Button _btnClear;
        private StatusStrip _statusBar;
        private Label _lblResultCount;
        private Panel _filterPanel;
        private SplitContainer _splitMain;

        public AdvancedSearchForm()
        {
            InitializeComponent();
            SetupUI();
        }

        public static void Show()
        {
            var form = new AdvancedSearchForm();
            form.Show();
        }

        private void SetupUI()
        {
            // Form properties
            this.Text = "OutlookTools — Advanced Search";
            this.Size = new Size(1100, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.RightToLeft = RightToLeft.No;
            this.Font = new Font("Segoe UI", 9);

            // Split container: filter panel on top, results on bottom
            _splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 120,
            };

            // === FILTER PANEL (top) ===
            _filterPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8) };

            int y = 12;
            int labelW = 70;
            int fieldW = 200;
            int spacing = 32;

            AddLabel("From:", 10, y, labelW);
            _txtFrom = AddTextBox(80, y, fieldW);
            y += spacing;

            AddLabel("To:", 10, y, labelW);
            _txtTo = AddTextBox(80, y, fieldW);
            y += spacing;

            AddLabel("Subject:", 10, y, labelW);
            _txtSubject = AddTextBox(80, y, fieldW);
            y += spacing;

            AddLabel("Body:", 10, y, labelW);
            _txtBody = AddTextBox(80, y, fieldW + 100);
            y += spacing;

            AddLabel("Date from:", 320, 12, labelW);
            _dtpFrom = new DateTimePicker
            {
                Location = new Point(400, 12),
                Width = 150,
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Now.AddMonths(-6)
            };
            _filterPanel.Controls.Add(_dtpFrom);

            AddLabel("Date to:", 320, 44, labelW);
            _dtpTo = new DateTimePicker
            {
                Location = new Point(400, 44),
                Width = 150,
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Now
            };
            _filterPanel.Controls.Add(_dtpTo);

            _chkHasAttachment = new CheckBox
            {
                Text = "Has attachment",
                Location = new Point(580, 12),
                Width = 130
            };
            _filterPanel.Controls.Add(_chkHasAttachment);

            _btnSearch = new Button
            {
                Text = "🔍 Search",
                Location = new Point(580, 50),
                Width = 130,
                Height = 30,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnSearch.Click += (s, e) => RunSearch();
            _filterPanel.Controls.Add(_btnSearch);

            _btnSelectAll = new Button
            {
                Text = "Select All",
                Location = new Point(730, 12),
                Width = 100,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _filterPanel.Controls.Add(_btnSelectAll);

            _btnClear = new Button
            {
                Text = "Clear",
                Location = new Point(730, 50),
                Width = 100,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnClear.Click += (s, e) => ClearResults();
            _filterPanel.Controls.Add(_btnClear);

            _splitMain.Panel1.Controls.Add(_filterPanel);

            // === RESULTS PANEL (bottom) ===
            var resultsPanel = new Panel { Dock = DockStyle.Fill };

            _resultsList = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = false
            };
            _resultsList.Columns.Add("Date", 120);
            _resultsList.Columns.Add("From", 180);
            _resultsList.Columns.Add("To", 150);
            _resultsList.Columns.Add("Subject", 350);
            _resultsList.Columns.Add("Has Attach", 80);
            _resultsList.DoubleClick += (s, e) => OpenSelected();
            _resultsList.SelectedIndexChanged += (s, e) => PreviewSelected();
            resultsPanel.Controls.Add(_resultsList);

            // Preview pane (WebBrowser)
            _previewPane = new WebBrowser
            {
                Dock = DockStyle.Right,
                Width = 400,
                IsWebBrowserContextMenuEnabled = false,
                ScriptErrorsSuppressed = true
            };
            resultsPanel.Controls.Add(_previewPane);

            _splitMain.Panel2.Controls.Add(resultsPanel);

            // Status bar
            _statusBar = new StatusStrip();
            _lblResultCount = new ToolStripStatusLabel("Ready");
            _statusBar.Items.Add(_lblResultCount);

            this.Controls.Add(_splitMain);
            this.Controls.Add(_statusBar);
        }

        private void RunSearch()
        {
            _resultsList.Items.Clear();
            _previewPane.DocumentText = "";

            try
            {
                Outlook.Application app = Globals.ThisAddIn.Application;
                Outlook.NameSpace session = app.Session;

                // Build Outlook DASL filter
                string filter = BuildFilter();
                int count = 0;

                // Search Inbox
                Outlook.Folder inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                if (inbox != null)
                    count += SearchFolder(inbox, filter);

                // Search Sent
                Outlook.Folder sent = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail) as Outlook.Folder;
                if (sent != null)
                    count += SearchFolder(sent, filter);

                _lblResultCount.Text = $"{count} results found";
            }
            catch (Exception ex)
            {
                _lblResultCount.Text = "Search error: " + ex.Message;
            }
        }

        private int SearchFolder(Outlook.Folder folder, string filter)
        {
            int count = 0;
            try
            {
                Outlook.Items items = folder.Items;
                Outlook.Items filtered = items.Restrict(filter);
                int total = Math.Min(filtered.Count, 500); // Limit for performance

                for (int i = 1; i <= total; i++)
                {
                    object obj = null;
                    try
                    {
                        obj = filtered[i];
                        if (!(obj is Outlook.MailItem mail)) continue;

                        var item = new ListViewItem(mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm"));
                        item.SubItems.Add(mail.SenderName ?? "");
                        item.SubItems.Add(GetRecipients(mail));
                        item.SubItems.Add(mail.Subject ?? "");
                        item.SubItems.Add(mail.Attachments.Count > 0 ? "✓" : "");
                        item.Tag = mail;
                        _resultsList.Items.Add(item);
                        count++;
                    }
                    catch { }
                    finally
                    {
                        if (obj != null) ReleaseCom(obj);
                    }
                }
            }
            catch { }

            _lblResultCount.Text = $"{count} results";
            return count;
        }

        /// <summary>
        /// Build Outlook DASL filter string from form fields.
        /// </summary>
        private string BuildFilter()
        {
            var parts = new System.Collections.Generic.List<string>();

            if (!string.IsNullOrWhiteSpace(_txtFrom.Text))
                parts.Add($"@SQL=\"urn:schemas:httpmail:from\" LIKE '%{_txtFrom.Text}%'");

            if (!string.IsNullOrWhiteSpace(_txtTo.Text))
                parts.Add($"@SQL=\"urn:schemas:httpmail:to\" LIKE '%{_txtTo.Text}%'");

            if (!string.IsNullOrWhiteSpace(_txtSubject.Text))
                parts.Add($"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{_txtSubject.Text}%'");

            if (!string.IsNullOrWhiteSpace(_txtBody.Text))
                parts.Add($"@SQL=\"urn:schemas:httpmail:textdescription\" LIKE '%{_txtBody.Text}%'");

            // Date range
            parts.Add($"[ReceivedTime] >= '{_dtpFrom.Value:MM/dd/yyyy}'");
            parts.Add($"[ReceivedTime] <= '{_dtpTo.Value:MM/dd/yyyy}'");

            // Has attachment
            if (_chkHasAttachment.Checked)
                parts.Add("[HasAttachments] = true");

            return string.Join(" AND ", parts);
        }

        private string GetRecipients(Outlook.MailItem mail)
        {
            var recipients = new System.Collections.Generic.List<string>();
            foreach (Outlook.Recipient r in mail.Recipients)
                recipients.Add(r.Name);
            return string.Join(", ", recipients);
        }

        private void PreviewSelected()
        {
            if (_resultsList.SelectedItems.Count == 0) return;
            if (_resultsList.SelectedItems[0].Tag is Outlook.MailItem mail)
            {
                try { _previewPane.DocumentText = mail.HTMLBody ?? mail.Body ?? ""; }
                catch { _previewPane.DocumentText = "<p>Preview not available.</p>"; }
            }
        }

        private void OpenSelected()
        {
            if (_resultsList.SelectedItems.Count == 0) return;
            if (_resultsList.SelectedItems[0].Tag is Outlook.MailItem mail)
            {
                try { mail.Display(false); }
                catch { }
            }
        }

        private void ClearResults()
        {
            _resultsList.Items.Clear();
            _previewPane.DocumentText = "";
            _lblResultCount.Text = "Ready";
        }

        private Label AddLabel(string text, int x, int y, int width)
        {
            var label = new Label
            {
                Text = text,
                Location = new Point(x, y + 3),
                Width = width,
                TextAlign = ContentAlignment.MiddleRight
            };
            _filterPanel.Controls.Add(label);
            return label;
        }

        private TextBox AddTextBox(int x, int y, int width)
        {
            var box = new TextBox
            {
                Location = new Point(x, y),
                Width = width
            };
            _filterPanel.Controls.Add(box);
            return box;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1084, 661);
            this.Name = "AdvancedSearchForm";
            this.Text = "OutlookTools — Advanced Search";
            this.ResumeLayout(false);
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
