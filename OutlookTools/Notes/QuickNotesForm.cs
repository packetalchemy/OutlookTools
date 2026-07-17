using System;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookTools.Notes
{
    /// <summary>
    /// OutlookTools — Quick Notes UI
    /// Simple form for managing notes with search.
    /// </summary>
    public partial class QuickNotesForm : Form
    {
        private ListBox _listNotes;
        private TextBox _txtTitle;
        private TextBox _txtContent;
        private TextBox _txtSearch;
        private TextBox _txtEmailTag;
        private Button _btnNew;
        private Button _btnSave;
        private Button _btnDelete;
        private Label _lblInfo;

        public QuickNotesForm()
        {
            InitializeComponent();
            SetupUI();
            LoadNotes();
        }

        public static void ShowForm()
        {
            new QuickNotesForm().ShowDialog();
        }

        private void SetupUI()
        {
            this.Text = "OutlookTools — Quick Notes";
            this.Size = new Size(800, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9);

            // Search bar
            var topPanel = new Panel { Dock = DockStyle.Top, Height = 40 };
            topPanel.Controls.Add(new Label { Text = "🔍", Location = new Point(12, 10), Width = 20 });
            _txtSearch = new TextBox { Location = new Point(35, 8), Width = 300 };
            _txtSearch.PlaceholderText = "Search notes...";
            _txtSearch.TextChanged += (s, e) => FilterNotes();
            topPanel.Controls.Add(_txtSearch);
            _btnNew = new Button
            {
                Text = "➕ New Note", Location = new Point(360, 7), Width = 110, Height = 28,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            _btnNew.Click += (s, e) => NewNote();
            topPanel.Controls.Add(_btnNew);
            this.Controls.Add(topPanel);

            // Left: notes list
            var leftPanel = new Panel { Dock = DockStyle.Left, Width = 250 };
            _listNotes = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10)
            };
            _listNotes.SelectedIndexChanged += (s, e) => LoadSelectedNote();
            leftPanel.Controls.Add(_listNotes);
            this.Controls.Add(leftPanel);

            // Right: editor
            var rightPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };

            int y = 10;
            rightPanel.Controls.Add(new Label { Text = "Title:", Location = new Point(10, y + 3), Width = 50 });
            _txtTitle = new TextBox { Location = new Point(70, y), Width = 400 };
            rightPanel.Controls.Add(_txtTitle);
            y += 35;

            rightPanel.Controls.Add(new Label { Text = "Email Tag:", Location = new Point(10, y + 3), Width = 60 });
            _txtEmailTag = new TextBox { Location = new Point(70, y), Width = 300, PlaceholderText = "Related email subject..." };
            rightPanel.Controls.Add(_txtEmailTag);
            y += 35;

            rightPanel.Controls.Add(new Label { Text = "Content:", Location = new Point(10, y + 3), Width = 60 });
            y += 25;

            _txtContent = new TextBox
            {
                Location = new Point(10, y),
                Width = 680,
                Height = 300,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                AcceptsReturn = true,
                AcceptsTab = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Font = new Font("Consolas", 10)
            };
            rightPanel.Controls.Add(_txtContent);

            // Buttons at bottom
            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 45,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(8)
            };

            _btnSave = new Button
            {
                Text = "💾 Save", Width = 80, Height = 30,
                BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _btnSave.Click += (s, e) => SaveNote();
            btnPanel.Controls.Add(_btnSave);

            _btnDelete = new Button
            {
                Text = "🗑️ Delete", Width = 80, Height = 30,
                BackColor = Color.FromArgb(220, 53, 69), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _btnDelete.Click += (s, e) => DeleteNote();
            btnPanel.Controls.Add(_btnDelete);

            _lblInfo = new Label
            {
                Text = "",
                AutoSize = true,
                ForeColor = Color.Green
            };
            btnPanel.Controls.Add(_lblInfo);

            rightPanel.Controls.Add(btnPanel);

            this.Controls.Add(rightPanel);
        }

        private QuickNote _currentNote;

        private void LoadNotes()
        {
            _listNotes.Items.Clear();
            foreach (var note in QuickNotesManager.Notes)
            {
                string display = note.Title ?? "(untitled)";
                if (!string.IsNullOrEmpty(note.EmailTag))
                    display += $" [{note.EmailTag}]";
                _listNotes.Items.Add(display);
            }
        }

        private void LoadSelectedNote()
        {
            if (_listNotes.SelectedIndex < 0) return;
            var note = QuickNotesManager.Notes[_listNotes.SelectedIndex];
            _currentNote = note;
            _txtTitle.Text = note.Title;
            _txtContent.Text = note.Content;
            _txtEmailTag.Text = note.EmailTag;
            _lblInfo.Text = $"Created: {note.Created:yyyy-MM-dd HH:mm}";
        }

        private void NewNote()
        {
            _currentNote = null;
            _txtTitle.Text = "";
            _txtContent.Text = "";
            _txtEmailTag.Text = "";
            _txtTitle.Focus();
            _lblInfo.Text = "New note";
        }

        private void SaveNote()
        {
            if (string.IsNullOrWhiteSpace(_txtTitle.Text))
            {
                MessageBox.Show("Please enter a title.");
                return;
            }

            if (_currentNote != null)
            {
                _currentNote.Title = _txtTitle.Text;
                _currentNote.Content = _txtContent.Text;
                _currentNote.EmailTag = _txtEmailTag.Text;
                QuickNotesManager.Update(_currentNote);
            }
            else
            {
                _currentNote = QuickNotesManager.Create(
                    _txtTitle.Text, _txtContent.Text, _txtEmailTag.Text);
            }

            LoadNotes();
            _lblInfo.Text = "✅ Saved!";
            _lblInfo.ForeColor = Color.Green;
        }

        private void DeleteNote()
        {
            if (_currentNote == null) return;

            var result = MessageBox.Show("Delete this note?", "Confirm",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                QuickNotesManager.Delete(_currentNote.Id);
                _currentNote = null;
                _txtTitle.Text = "";
                _txtContent.Text = "";
                _txtEmailTag.Text = "";
                LoadNotes();
                _lblInfo.Text = "🗑️ Deleted";
            }
        }

        private void FilterNotes()
        {
            if (string.IsNullOrWhiteSpace(_txtSearch.Text))
            {
                LoadNotes();
                return;
            }

            var results = QuickNotesManager.Search(_txtSearch.Text);
            _listNotes.Items.Clear();
            foreach (var note in results)
            {
                string display = note.Title ?? "(untitled)";
                if (!string.IsNullOrEmpty(note.EmailTag))
                    display += $" [{note.EmailTag}]";
                _listNotes.Items.Add(display);
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 550);
            this.Name = "QuickNotesForm";
            this.Text = "OutlookTools — Quick Notes";
            this.ResumeLayout(false);
        }
    }
}
