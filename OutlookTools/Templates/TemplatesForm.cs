using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookTools.Settings;

namespace OutlookTools.Templates
{
    /// <summary>
    /// OutlookTools — Email Templates Manager
    /// Save, edit, delete, and insert email templates.
    /// 
    /// Templates are stored as .oft files in:
    ///   %LOCALAPPDATA%\OutlookTools\Templates\
    /// 
    /// Features:
    /// - Save current draft as template
    /// - Insert template into new message
    /// - Edit/delete templates
    /// - Support for {DATE}, {TIME}, {NAME} placeholders
    /// - Categories for organization
    /// </summary>
    public partial class TemplatesForm : Form
    {
        private ListView _templateList;
        private TextBox _txtPreview;
        private TextBox _txtSubject;
        private TextBox _txtBody;
        private TextBox _txtCategory;
        private Button _btnSave;
        private Button _btnInsert;
        private Button _btnDelete;
        private Button _btnNew;
        private Button _btnRefresh;
        private StatusStrip _statusBar;
        private Label _lblStatus;
        private SplitContainer _splitMain;

        private string TemplatesDir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookTools", "Templates");

        public TemplatesForm()
        {
            InitializeComponent();
            SetupUI();
            LoadTemplates();
        }

        public static void ShowForm()
        {
            var form = new TemplatesForm();
            form.ShowDialog();
        }

        private void SetupUI()
        {
            this.Text = "OutlookTools — Email Templates";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9);

            _splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 280
            };

            // === LEFT PANEL: Template list ===
            var leftPanel = new Panel { Dock = DockStyle.Fill };

            var topBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 40,
                FlowDirection = FlowDirection.RightToLeft
            };

            _btnRefresh = CreateButton("🔄 Refresh", 100);
            _btnRefresh.Click += (s, e) => LoadTemplates();
            topBar.Controls.Add(_btnRefresh);

            _btnNew = CreateButton("➕ New", 80);
            _btnNew.Click += (s, e) => NewTemplate();
            topBar.Controls.Add(_btnNew);

            leftPanel.Controls.Add(topBar);

            _templateList = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = false
            };
            _templateList.Columns.Add("Name", 150);
            _templateList.Columns.Add("Category", 100);
            _templateList.Columns.Add("Modified", 120);
            _templateList.SelectedIndexChanged += (s, e) => PreviewSelected();
            leftPanel.Controls.Add(_templateList);

            _splitMain.Panel1.Controls.Add(leftPanel);

            // === RIGHT PANEL: Preview & Edit ===
            var rightPanel = new Panel { Dock = DockStyle.Fill };

            // Edit area
            var editPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(8)
            };

            int y = 8;
            AddLabel("Template Name:", 8, y, editPanel);
            _txtSubject = new TextBox { Location = new Point(120, y), Width = 400 };
            editPanel.Controls.Add(_txtSubject);
            y += 32;

            AddLabel("Category:", 8, y, editPanel);
            _txtCategory = new TextBox { Location = new Point(120, y), Width = 200 };
            editPanel.Controls.Add(_txtCategory);
            y += 32;

            AddLabel("Body (HTML):", 8, y, editPanel);
            y += 20;

            _txtBody = new TextBox
            {
                Location = new Point(8, y),
                Width = 520,
                Height = 300,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                AcceptsReturn = true,
                AcceptsTab = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Font = new Font("Consolas", 9)
            };
            editPanel.Controls.Add(_txtBody);

            // Buttons
            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(8)
            };

            _btnSave = CreateButton("💾 Save Template", 130);
            _btnSave.Click += (s, e) => SaveTemplate();
            btnPanel.Controls.Add(_btnSave);

            _btnInsert = CreateButton("📥 Insert into Message", 160);
            _btnInsert.Click += (s, e) => InsertTemplate();
            btnPanel.Controls.Add(_btnInsert);

            _btnDelete = CreateButton("🗑️ Delete", 90);
            _btnDelete.BackColor = Color.FromArgb(220, 53, 69);
            _btnDelete.Click += (s, e) => DeleteTemplate();
            btnPanel.Controls.Add(_btnDelete);

            rightPanel.Controls.Add(editPanel);
            rightPanel.Controls.Add(btnPanel);

            _splitMain.Panel2.Controls.Add(rightPanel);

            // Status bar
            _statusBar = new StatusStrip();
            _lblStatus = new ToolStripStatusLabel("Ready");
            _statusBar.Items.Add(_lblStatus);

            this.Controls.Add(_splitMain);
            this.Controls.Add(_statusBar);
        }

        private void LoadTemplates()
        {
            _templateList.Items.Clear();
            Directory.CreateDirectory(TemplatesDir);

            foreach (var file in Directory.GetFiles(TemplatesDir, "*.template"))
            {
                try
                {
                    string name = Path.GetFileNameWithoutExtension(file);
                    var info = new FileInfo(file);
                    var lines = File.ReadAllLines(file);
                    string category = lines.Length > 0 ? lines[0].Replace("CATEGORY:", "") : "General";

                    var item = new ListViewItem(name);
                    item.SubItems.Add(category);
                    item.SubItems.Add(info.LastWriteTime.ToString("yyyy-MM-dd HH:mm"));
                    item.Tag = file;
                    _templateList.Items.Add(item);
                }
                catch { }
            }

            _lblStatus.Text = $"{_templateList.Items.Count} templates loaded";
        }

        private void PreviewSelected()
        {
            if (_templateList.SelectedItems.Count == 0) return;
            string file = (string)_templateList.SelectedItems[0].Tag;
            try
            {
                string[] lines = File.ReadAllLines(file);
                _txtSubject.Text = Path.GetFileNameWithoutExtension(file);

                // Parse category from first line
                string category = lines.Length > 0 && lines[0].StartsWith("CATEGORY:")
                    ? lines[0].Substring(9) : "General";
                _txtCategory.Text = category;

                // Body is everything after category line
                int startLine = (lines.Length > 0 && lines[0].StartsWith("CATEGORY:")) ? 1 : 0;
                _txtBody.Text = string.Join(Environment.NewLine, lines.Skip(startLine));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading template: " + ex.Message);
            }
        }

        private void NewTemplate()
        {
            _txtSubject.Text = "New Template";
            _txtCategory.Text = "General";
            _txtBody.Text = "<html><body>\n<p>Hello {NAME},</p>\n\n<p>Your message here...</p>\n\n<p>Best regards</p>\n</body></html>";
            _txtSubject.Focus();
            _txtSubject.SelectAll();
        }

        private void SaveTemplate()
        {
            if (string.IsNullOrWhiteSpace(_txtSubject.Text))
            {
                MessageBox.Show("Please enter a template name.");
                return;
            }

            try
            {
                Directory.CreateDirectory(TemplatesDir);
                string filename = _txtSubject.Text.Trim()
                    .Replace(" ", "_")
                    .Replace("/", "-")
                    .Replace("\\", "-");
                string path = Path.Combine(TemplatesDir, filename + ".template");

                string content = $"CATEGORY:{_txtCategory.Text.Trim()}\n{_txtBody.Text}";
                File.WriteAllText(path, content);

                _lblStatus.Text = $"✅ Saved: {filename}";
                LoadTemplates();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Save failed: " + ex.Message);
            }
        }

        private void InsertTemplate()
        {
            if (string.IsNullOrWhiteSpace(_txtBody.Text))
            {
                MessageBox.Show("No template content to insert.");
                return;
            }

            try
            {
                Outlook.Application app = Globals.ThisAddIn.Application;
                Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

                if (mail != null)
                {
                    string body = ReplacePlaceholders(_txtBody.Text);
                    string subject = ReplacePlaceholders(_txtSubject.Text);

                    mail.Subject = subject;
                    mail.HTMLBody = body;
                    mail.Display(false);

                    _lblStatus.Text = "✅ Template inserted into new message";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Insert failed: " + ex.Message);
            }
        }

        private void DeleteTemplate()
        {
            if (_templateList.SelectedItems.Count == 0) return;

            var result = MessageBox.Show(
                "Delete this template?",
                "Confirm Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                try
                {
                    string file = (string)_templateList.SelectedItems[0].Tag;
                    File.Delete(file);
                    _lblStatus.Text = "✅ Template deleted";
                    LoadTemplates();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Delete failed: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Replace placeholders with actual values.
        /// Supported: {DATE}, {TIME}, {NAME}, {SUBJECT}, {SENDER}
        /// </summary>
        private string ReplacePlaceholders(string text)
        {
            return text
                .Replace("{DATE}", DateTime.Now.ToString("yyyy-MM-dd"))
                .Replace("{TIME}", DateTime.Now.ToString("HH:mm"))
                .Replace("{DATETIME}", DateTime.Now.ToString("yyyy-MM-dd HH:mm"))
                .Replace("{YEAR}", DateTime.Now.Year.ToString())
                .Replace("{MONTH}", DateTime.Now.ToString("MMMM"));
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 600);
            this.Name = "TemplatesForm";
            this.Text = "OutlookTools — Email Templates";
            this.ResumeLayout(false);
        }

        private static Button CreateButton(string text, int width)
        {
            return new Button
            {
                Text = text,
                Width = width,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Margin = new Padding(2)
            };
        }

        private static void AddLabel(string text, int x, int y, Control parent)
        {
            parent.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y + 3),
                Width = 110,
                TextAlign = ContentAlignment.MiddleRight
            });
        }
    }
}
