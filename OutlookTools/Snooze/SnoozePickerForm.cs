using System;
using System.Windows.Forms;

namespace OutlookTools.Snooze
{
    /// <summary>
    /// OutlookTools — Snooze Picker Dialog
    /// Lets user pick custom snooze time.
    /// </summary>
    public class SnoozePickerForm : Form
    {
        public DateTime? SelectedTime { get; private set; }

        private DateTimePicker _dtpDate;
        private DateTimePicker _dtpTime;
        private Button _btnOk;
        private Button _btnCancel;

        public SnoozePickerForm()
        {
            this.Text = "Snooze Until...";
            this.Size = new Size(300, 200);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            this.Controls.Add(new Label
            {
                Text = "Date:",
                Location = new Point(20, 20),
                Width = 80
            });
            _dtpDate = new DateTimePicker
            {
                Location = new Point(110, 18),
                Width = 150,
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today.AddDays(1)
            };
            this.Controls.Add(_dtpDate);

            this.Controls.Add(new Label
            {
                Text = "Time:",
                Location = new Point(20, 55),
                Width = 80
            });
            _dtpTime = new DateTimePicker
            {
                Location = new Point(110, 53),
                Width = 150,
                Format = DateTimePickerFormat.Time,
                Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 9, 0, 0)
            };
            this.Controls.Add(_dtpTime);

            _btnOk = new Button
            {
                Text = "OK",
                Location = new Point(70, 110),
                Width = 80,
                DialogResult = DialogResult.OK
            };
            _btnOk.Click += (s, e) =>
            {
                SelectedTime = _dtpDate.Value.Date + _dtpTime.Value.TimeOfDay;
            };
            this.Controls.Add(_btnOk);

            _btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(160, 110),
                Width = 80,
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(_btnCancel);

            this.AcceptButton = _btnOk;
            this.CancelButton = _btnCancel;
        }
    }
}
