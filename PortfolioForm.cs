using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using ExcelDataReader;

namespace BONDVERSE
{
    public class PortfolioForm : Form
    {
        Dictionary<string, List<PortfolioEntry>> portfolios = new Dictionary<string, List<PortfolioEntry>>();
        int editIndex = -1;

        // UI Controls
        TextBox txtPortfolio = new TextBox();
        TextBox txtInvestor = new TextBox();
        TextBox txtFV = new TextBox();
        TextBox txtQty = new TextBox();
        TextBox txtBondName = new TextBox();
        TextBox txtCoupon = new TextBox();
        TextBox txtCheque = new TextBox();

        ComboBox cmbFreq = new ComboBox();
        ComboBox cmbQuarterStart = new ComboBox();
        ComboBox cmbTDS = new ComboBox();
        ComboBox cmbPortfolioSelect = new ComboBox();
        ComboBox cmbMonthSelect = new ComboBox();

        DateTimePicker dtTrans = new DateTimePicker();
        DateTimePicker dtMaturity = new DateTimePicker();

        DataGridView grid = new DataGridView();

        Panel sidebar = new Panel();
        Panel topbar = new Panel();
        Panel mainPanel = new Panel();

        public PortfolioForm()
        {
            Text = "BONDVERSE";
            WindowState = FormWindowState.Maximized;

            // Sidebar
            sidebar.Width = 200;
            sidebar.Dock = DockStyle.Left;
            sidebar.BackColor = System.Drawing.Color.FromArgb(30, 30, 60);
            Controls.Add(sidebar);

            Button btnImport = CreateSideButton("Import Excel", 50);
            Button btnExport = CreateSideButton("Export PDF", 100);
            Button btnTDS = CreateSideButton("TDS Summary", 150);

            sidebar.Controls.Add(btnImport);
            sidebar.Controls.Add(btnExport);
            sidebar.Controls.Add(btnTDS);

            // Top bar
            topbar.Height = 60;
            topbar.Dock = DockStyle.Top;
            topbar.BackColor = System.Drawing.Color.WhiteSmoke;
            Controls.Add(topbar);

            Label title = new Label()
            {
                Text = "BONDVERSE Dashboard",
                Font = new System.Drawing.Font("Segoe UI", 14, System.Drawing.FontStyle.Bold),
                Left = 220,
                Top = 15
            };
            topbar.Controls.Add(title);

            // Main panel
            mainPanel.Dock = DockStyle.Fill;
            Controls.Add(mainPanel);

            int y = 20;

            AddLabel("Portfolio", y);
            txtPortfolio.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtPortfolio);

            AddLabel("Select Portfolio", y, 300);
            cmbPortfolioSelect.SetBounds(450, y, 150, 25);
            mainPanel.Controls.Add(cmbPortfolioSelect);

            AddLabel("Investor", y += 30);
            txtInvestor.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtInvestor);

            AddLabel("TDS %", y += 30);
            cmbTDS.SetBounds(120, y, 150, 25);
            cmbTDS.Items.AddRange(new string[] { "10.4", "20.8" });
            cmbTDS.SelectedIndex = 0;
            mainPanel.Controls.Add(cmbTDS);

            AddLabel("Date", y += 30);
            dtTrans.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(dtTrans);

            AddLabel("FV", y += 30);
            txtFV.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtFV);

            AddLabel("Qty", y += 30);
            txtQty.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtQty);

            AddLabel("Bond", y += 30);
            txtBondName.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtBondName);

            AddLabel("Coupon", y += 30);
            txtCoupon.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtCoupon);

            AddLabel("Cheque", y += 30);
            txtCheque.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtCheque);

            AddLabel("Frequency", y += 30);
            cmbFreq.SetBounds(120, y, 150, 25);
            cmbFreq.Items.AddRange(new string[] { "Monthly", "Quarterly", "Yearly" });
            cmbFreq.SelectedIndex = 0;
            mainPanel.Controls.Add(cmbFreq);

            AddLabel("Quarter Start", y += 30);
            cmbQuarterStart.SetBounds(120, y, 150, 25);
            cmbQuarterStart.Items.AddRange(new string[]
            {
                "January","February","March","April","May","June",
                "July","August","September","October","November","December"
            });
            cmbQuarterStart.Visible = false;
            mainPanel.Controls.Add(cmbQuarterStart);

            cmbFreq.SelectedIndexChanged += (s, e) =>
            {
                cmbQuarterStart.Visible = cmbFreq.Text == "Quarterly";
            };

            AddLabel("Maturity", y += 30);
            dtMaturity.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(dtMaturity);

            Button btnAdd = new Button() { Text = "Add", Top = y += 40, Left = 10 };
            Button btnEdit = new Button() { Text = "Edit", Top = y, Left = 80 };
            Button btnDelete = new Button() { Text = "Delete", Top = y, Left = 150 };
            Button btnSubmit = new Button() { Text = "Submit", Top = y, Left = 230 };

            mainPanel.Controls.Add(btnAdd);
            mainPanel.Controls.Add(btnEdit);
            mainPanel.Controls.Add(btnDelete);
            mainPanel.Controls.Add(btnSubmit);

            AddLabel("Month", y += 40);
            cmbMonthSelect.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(cmbMonthSelect);

            Button btnMonth = new Button() { Text = "Summary", Top = y, Left = 300 };
            mainPanel.Controls.Add(btnMonth);

            // GRID
            grid.Dock = DockStyle.Right;
            grid.Width = 900;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            mainPanel.Controls.Add(grid);

            // EVENTS
            btnAdd.Click += AddEntry;
            btnEdit.Click += EditEntry;
            btnDelete.Click += DeleteEntry;
            btnSubmit.Click += GenerateTable;
            btnImport.Click += UploadExcel;
            btnExport.Click += ExportPdf;
            btnTDS.Click += (s, e) => ShowTDSQuarterMessage();
            btnMonth.Click += ShowMonthSummary;
        }

        Button CreateSideButton(string text, int top)
        {
            return new Button()
            {
                Text = text,
                Top = top,
                Left = 10,
                Width = 180,
                ForeColor = System.Drawing.Color.White,
                BackColor = System.Drawing.Color.FromArgb(60, 60, 120),
                FlatStyle = FlatStyle.Flat
            };
        }

        void AddLabel(string text, int top, int left = 10)
        {
            mainPanel.Controls.Add(new Label() { Text = text, Top = top, Left = left });
        }

        void AddEntry(object sender, EventArgs e)
        {
            string portfolio = txtPortfolio.Text;

            if (!portfolios.ContainsKey(portfolio))
                portfolios[portfolio] = new List<PortfolioEntry>();

            PortfolioEntry p = new PortfolioEntry()
            {
                PortfolioName = portfolio,
                InvestorName = txtInvestor.Text,
                TransactionDate = dtTrans.Value,
                FV = double.Parse(txtFV.Text),
                Quantity = int.Parse(txtQty.Text),
                BondName = txtBondName.Text,
                CouponRate = double.Parse(txtCoupon.Text),
                ChequeAmount = double.Parse(txtCheque.Text),
                Frequency = cmbFreq.Text,
                QuarterStartMonth = cmbQuarterStart.Text,
                MaturityDate = dtMaturity.Value
            };

            if (editIndex >= 0)
            {
                portfolios[portfolio][editIndex] = p;
                editIndex = -1;
            }
            else
            {
                portfolios[portfolio].Add(p);
            }

            if (!cmbPortfolioSelect.Items.Contains(portfolio))
                cmbPortfolioSelect.Items.Add(portfolio);

            MessageBox.Show("Saved");
        }

        void EditEntry(object sender, EventArgs e)
        {
            if (grid.CurrentRow == null) return;

            string selected = cmbPortfolioSelect.Text;
            editIndex = grid.CurrentRow.Index;

            var p = portfolios[selected][editIndex];

            txtBondName.Text = p.BondName;
            txtFV.Text = p.FV.ToString();
            txtCoupon.Text = p.CouponRate.ToString();
        }

        void DeleteEntry(object sender, EventArgs e)
        {
            if (grid.CurrentRow == null) return;

            string selected = cmbPortfolioSelect.Text;
            int index = grid.CurrentRow.Index;

            portfolios[selected].RemoveAt(index);
            GenerateTable(null, null);
        }

        void GenerateTable(object sender, EventArgs e)
        {
            if (cmbPortfolioSelect.SelectedItem == null) return;

            var data = portfolios[cmbPortfolioSelect.Text];

            DataTable dt = new DataTable();
            dt.Columns.Add("Bond Name");

            DateTime start = DateTime.Today;
            DateTime maxDate = data.Max(x => x.MaturityDate);

            List<DateTime> months = new List<DateTime>();

            while (start <= maxDate)
            {
                dt.Columns.Add(start.ToString("MMM yyyy"));
                months.Add(start);
                start = start.AddMonths(1);
            }

            foreach (var e1 in data)
            {
                var row = dt.NewRow();
                row["Bond Name"] = e1.BondName;

                foreach (var m in months)
                {
                    double interest = 0;

                    if (m <= e1.MaturityDate)
                    {
                        if (e1.Frequency == "Monthly")
                            interest = e1.FV * e1.CouponRate / 100 / 12;
                    }

                    row[m.ToString("MMM yyyy")] = Math.Round(interest);
                }

                dt.Rows.Add(row);
            }

            grid.DataSource = dt;
        }

        void ShowMonthSummary(object sender, EventArgs e)
        {
            if (cmbMonthSelect.SelectedItem == null) return;

            string month = cmbMonthSelect.Text;

            double total = 0;

            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.Cells[month].Value != null)
                    total += Convert.ToDouble(row.Cells[month].Value);
            }

            MessageBox.Show($"Total: {total}");
        }

        void ExportPdf(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "PDF|*.pdf";

            if (save.ShowDialog() == DialogResult.OK)
            {
                using (var writer = new PdfWriter(save.FileName))
                using (var pdf = new PdfDocument(writer))
                using (var doc = new Document(pdf))
                {
                    doc.Add(new Paragraph("Bond Report"));
                }

                MessageBox.Show("PDF Saved");
            }
        }

        void UploadExcel(object sender, EventArgs e)
        {
            MessageBox.Show("Excel Import Ready");
        }

        void ShowTDSQuarterMessage()
        {
            MessageBox.Show("TDS Feature Active");
        }
    }
}
