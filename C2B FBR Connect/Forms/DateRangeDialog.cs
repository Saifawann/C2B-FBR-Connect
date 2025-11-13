// Add this new file: DateRangeDialog.cs
using System;
using System.Drawing;
using System.Windows.Forms;

namespace C2B_FBR_Connect.Forms
{
    public class DateRangeDialog : Form
    {
        private DateTimePicker dtpFrom;
        private DateTimePicker dtpTo;
        private ComboBox cboQuickSelect;
        private CheckBox chkIncludeInvoices;
        private CheckBox chkIncludeCreditMemos;
        private Button btnOK;
        private Button btnCancel;

        public DateTime DateFrom => dtpFrom.Value.Date;
        public DateTime DateTo => dtpTo.Value.Date.AddDays(1).AddSeconds(-1); // End of day
        public bool IncludeInvoices => chkIncludeInvoices.Checked;
        public bool IncludeCreditMemos => chkIncludeCreditMemos.Checked;

        public DateRangeDialog()
        {
            InitializeComponent();
            SetDefaults();
        }

        private void InitializeComponent()
        {
            Text = "Select Date Range & Transaction Types";
            Size = new Size(450, 320);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            // Quick Select Label
            var lblQuickSelect = new Label
            {
                Text = "Quick Select:",
                Location = new Point(20, 20),
                Size = new Size(100, 23)
            };

            // Quick Select Combo
            cboQuickSelect = new ComboBox
            {
                Location = new Point(130, 20),
                Size = new Size(280, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboQuickSelect.Items.AddRange(new object[] {
                "Custom Range",
                "Today",
                "Yesterday",
                "Last 7 Days",
                "Last 14 Days",
                "Last 30 Days",
                "This Week",
                "Last Week",
                "This Month",
                "Last Month",
                "This Quarter",
                "Last Quarter",
                "This Year",
                "Last Year",
                "All Time"
            });
            cboQuickSelect.SelectedIndex = 0;
            cboQuickSelect.SelectedIndexChanged += CboQuickSelect_SelectedIndexChanged;

            // From Date
            var lblFrom = new Label
            {
                Text = "From Date:",
                Location = new Point(20, 60),
                Size = new Size(100, 23)
            };

            dtpFrom = new DateTimePicker
            {
                Location = new Point(130, 60),
                Size = new Size(280, 23),
                Format = DateTimePickerFormat.Short
            };

            // To Date
            var lblTo = new Label
            {
                Text = "To Date:",
                Location = new Point(20, 95),
                Size = new Size(100, 23)
            };

            dtpTo = new DateTimePicker
            {
                Location = new Point(130, 95),
                Size = new Size(280, 23),
                Format = DateTimePickerFormat.Short
            };

            // Transaction Type Group
            var grpTransactionType = new GroupBox
            {
                Text = "Transaction Types to Fetch",
                Location = new Point(20, 135),
                Size = new Size(390, 80)
            };

            chkIncludeInvoices = new CheckBox
            {
                Text = "Invoices",
                Location = new Point(20, 30),
                Size = new Size(150, 25),
                Checked = true
            };

            chkIncludeCreditMemos = new CheckBox
            {
                Text = "Credit Memos",
                Location = new Point(200, 30),
                Size = new Size(150, 25),
                Checked = true
            };

            grpTransactionType.Controls.Add(chkIncludeInvoices);
            grpTransactionType.Controls.Add(chkIncludeCreditMemos);

            // Buttons
            btnOK = new Button
            {
                Text = "Fetch",
                Location = new Point(235, 235),
                Size = new Size(85, 30),
                DialogResult = DialogResult.OK,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(325, 235),
                Size = new Size(85, 30),
                DialogResult = DialogResult.Cancel,
                BackColor = Color.FromArgb(96, 125, 139),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            // Add validation
            btnOK.Click += (s, e) =>
            {
                if (!chkIncludeInvoices.Checked && !chkIncludeCreditMemos.Checked)
                {
                    MessageBox.Show("Please select at least one transaction type.", "Validation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }

                if (dtpFrom.Value > dtpTo.Value)
                {
                    MessageBox.Show("'From Date' cannot be after 'To Date'.", "Invalid Date Range",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                }
            };

            Controls.AddRange(new Control[] {
                lblQuickSelect, cboQuickSelect,
                lblFrom, dtpFrom,
                lblTo, dtpTo,
                grpTransactionType,
                btnOK, btnCancel
            });

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }

        private void SetDefaults()
        {
            // Default to last 30 days
            dtpTo.Value = DateTime.Today;
            dtpFrom.Value = DateTime.Today.AddDays(-30);
            cboQuickSelect.SelectedIndex = 5; // "Last 30 Days"
        }

        private void CboQuickSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            var today = DateTime.Today;
            var now = DateTime.Now;

            switch (cboQuickSelect.Text)
            {
                case "Today":
                    dtpFrom.Value = today;
                    dtpTo.Value = today;
                    break;

                case "Yesterday":
                    dtpFrom.Value = today.AddDays(-1);
                    dtpTo.Value = today.AddDays(-1);
                    break;

                case "Last 7 Days":
                    dtpFrom.Value = today.AddDays(-6);
                    dtpTo.Value = today;
                    break;

                case "Last 14 Days":
                    dtpFrom.Value = today.AddDays(-13);
                    dtpTo.Value = today;
                    break;

                case "Last 30 Days":
                    dtpFrom.Value = today.AddDays(-29);
                    dtpTo.Value = today;
                    break;

                case "This Week":
                    int daysToMonday = ((int)today.DayOfWeek - 1 + 7) % 7;
                    dtpFrom.Value = today.AddDays(-daysToMonday);
                    dtpTo.Value = today;
                    break;

                case "Last Week":
                    int daysToLastMonday = ((int)today.DayOfWeek - 1 + 7) % 7 + 7;
                    dtpFrom.Value = today.AddDays(-daysToLastMonday);
                    dtpTo.Value = today.AddDays(-daysToLastMonday + 6);
                    break;

                case "This Month":
                    dtpFrom.Value = new DateTime(today.Year, today.Month, 1);
                    dtpTo.Value = today;
                    break;

                case "Last Month":
                    var lastMonth = today.AddMonths(-1);
                    dtpFrom.Value = new DateTime(lastMonth.Year, lastMonth.Month, 1);
                    dtpTo.Value = new DateTime(lastMonth.Year, lastMonth.Month,
                        DateTime.DaysInMonth(lastMonth.Year, lastMonth.Month));
                    break;

                case "This Quarter":
                    int currentQuarter = (today.Month - 1) / 3;
                    dtpFrom.Value = new DateTime(today.Year, currentQuarter * 3 + 1, 1);
                    dtpTo.Value = today;
                    break;

                case "Last Quarter":
                    var lastQuarter = today.AddMonths(-3);
                    int quarter = (lastQuarter.Month - 1) / 3;
                    dtpFrom.Value = new DateTime(lastQuarter.Year, quarter * 3 + 1, 1);
                    var endMonth = quarter * 3 + 3;
                    dtpTo.Value = new DateTime(lastQuarter.Year, endMonth,
                        DateTime.DaysInMonth(lastQuarter.Year, endMonth));
                    break;

                case "This Year":
                    dtpFrom.Value = new DateTime(today.Year, 1, 1);
                    dtpTo.Value = today;
                    break;

                case "Last Year":
                    dtpFrom.Value = new DateTime(today.Year - 1, 1, 1);
                    dtpTo.Value = new DateTime(today.Year - 1, 12, 31);
                    break;

                case "All Time":
                    dtpFrom.Value = new DateTime(2000, 1, 1);
                    dtpTo.Value = today;
                    break;
            }
        }
    }
}