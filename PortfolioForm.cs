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

        Label lblPortfolio = new Label();

        public PortfolioForm()
        {
            Text = "BONDVERSE";
            Width = 1400;
            Height = 800;

            // SIDEBAR
            sidebar.Width = 200;
            sidebar.Dock = DockStyle.Left;
            sidebar.BackColor = System.Drawing.Color.FromArgb(30, 30, 60);
            Controls.Add(sidebar);

            Button btnImport = new Button() { Text = "Import Excel", Top = 50, Width = 180, Left = 10 };
            Button btnExport = new Button() { Text = "Export PDF", Top = 100, Width = 180, Left = 10 };

            StyleBtn(btnImport);
            StyleBtn(btnExport);

            sidebar.Controls.Add(btnImport);
            sidebar.Controls.Add(btnExport);

            // TOP BAR
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

            lblPortfolio.Left = 600;
            lblPortfolio.Top = 20;
            topbar.Controls.Add(lblPortfolio);

            // MAIN PANEL
            mainPanel.Dock = DockStyle.Fill;
            Controls.Add(mainPanel);

            int y = 70;

            mainPanel.Controls.Add(new Label() { Text = "Portfolio", Top = y, Left = 10 });
            txtPortfolio.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtPortfolio);

            mainPanel.Controls.Add(new Label() { Text = "Select Portfolio", Top = y, Left = 300 });
            cmbPortfolioSelect.SetBounds(450, y, 150, 25);
            mainPanel.Controls.Add(cmbPortfolioSelect);

            mainPanel.Controls.Add(new Label() { Text = "Investor", Top = y += 30, Left = 10 });
            txtInvestor.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtInvestor);

            mainPanel.Controls.Add(new Label() { Text = "TDS %", Top = y += 30, Left = 10 });
            cmbTDS.SetBounds(120, y, 150, 25);
            cmbTDS.Items.AddRange(new string[] { "10.4", "20.8" });
            cmbTDS.SelectedIndex = 0;
            mainPanel.Controls.Add(cmbTDS);

            mainPanel.Controls.Add(new Label() { Text = "Date", Top = y += 30, Left = 10 });
            dtTrans.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(dtTrans);

            mainPanel.Controls.Add(new Label() { Text = "FV", Top = y += 30, Left = 10 });
            txtFV.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtFV);

            mainPanel.Controls.Add(new Label() { Text = "Qty", Top = y += 30, Left = 10 });
            txtQty.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtQty);

            mainPanel.Controls.Add(new Label() { Text = "Bond", Top = y += 30, Left = 10 });
            txtBondName.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtBondName);

            mainPanel.Controls.Add(new Label() { Text = "Coupon", Top = y += 30, Left = 10 });
            txtCoupon.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtCoupon);

            mainPanel.Controls.Add(new Label() { Text = "Cheque", Top = y += 30, Left = 10 });
            txtCheque.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(txtCheque);

            mainPanel.Controls.Add(new Label() { Text = "Frequency", Top = y += 30, Left = 10 });
            cmbFreq.SetBounds(120, y, 150, 25);
            cmbFreq.Items.AddRange(new string[] { "Monthly", "Quarterly", "Yearly" });
            cmbFreq.SelectedIndex = 0;
            mainPanel.Controls.Add(cmbFreq);

            mainPanel.Controls.Add(new Label() { Text = "Quarter Start", Top = y += 30, Left = 10 });
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

            mainPanel.Controls.Add(new Label() { Text = "Maturity", Top = y += 30, Left = 10 });
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

            mainPanel.Controls.Add(new Label() { Text = "Month", Top = y += 40, Left = 10 });
            cmbMonthSelect.SetBounds(120, y, 150, 25);
            mainPanel.Controls.Add(cmbMonthSelect);

            Button btnMonth = new Button() { Text = "Summary", Top = y, Left = 300 };
            mainPanel.Controls.Add(btnMonth);

            grid.Dock = DockStyle.Right;
            grid.Width = 800;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            mainPanel.Controls.Add(grid);

            // EVENTS
            btnAdd.Click += AddEntry;
            btnEdit.Click += EditEntry;
            btnDelete.Click += DeleteEntry;
            btnSubmit.Click += GenerateTable;
            btnImport.Click += UploadExcel;
            btnExport.Click += ExportPdf;
            btnMonth.Click += ShowMonthSummary;
        }

        void StyleBtn(Button b)
        {
            b.ForeColor = System.Drawing.Color.White;
            b.BackColor = System.Drawing.Color.FromArgb(60, 60, 120);
            b.FlatStyle = FlatStyle.Flat;
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

            if (portfolios.ContainsKey(selected) && index < portfolios[selected].Count)
            {
                portfolios[selected].RemoveAt(index);
                GenerateTable(null, null);
            }
        }

        void GenerateTable(object sender, EventArgs e)
        {
            if (cmbPortfolioSelect.SelectedItem == null) return;

            string selected = cmbPortfolioSelect.Text;
            var data = portfolios[selected];

            lblPortfolio.Text = "Portfolio: " + selected;

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

            cmbMonthSelect.Items.Clear();
            foreach (var m in months)
                cmbMonthSelect.Items.Add(m.ToString("MMM yyyy"));

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

                        else if (e1.Frequency == "Quarterly")
                        {
                            int startMonth = DateTime.ParseExact(e1.QuarterStartMonth, "MMMM", null).Month;
                            int diff = (m.Year - e1.TransactionDate.Year) * 12 + (m.Month - startMonth);

                            if (diff >= 0 && diff % 3 == 0)
                                interest = e1.FV * e1.CouponRate / 100 / 4;
                        }
                        else if (e1.Frequency == "Yearly")
                        {
                            if (m.Month == e1.TransactionDate.Month)
                                interest = e1.FV * e1.CouponRate / 100;
                        }
                    }

                    row[m.ToString("MMM yyyy")] = Math.Round(interest);
                }

                dt.Rows.Add(row);
            }

            double tds = double.Parse(cmbTDS.Text) / 100;

            var total = dt.NewRow();
            total["Bond Name"] = "TOTAL";

            var net = dt.NewRow();
            net["Bond Name"] = "NET";

            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName == "Bond Name") continue;

                double sum = dt.AsEnumerable().Sum(r => Convert.ToDouble(r[col]));
                total[col] = sum;
                net[col] = Math.Round(sum * (1 - tds));
            }

            dt.Rows.Add(total);
            dt.Rows.Add(net);

            grid.DataSource = dt;
        }

        void ShowMonthSummary(object sender, EventArgs e)
        {
            if (cmbMonthSelect.SelectedItem == null) return;

            string month = cmbMonthSelect.Text;

            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.Cells["Bond Name"].Value?.ToString() == "TOTAL")
                {
                    double gross = Convert.ToDouble(row.Cells[month].Value);
                    double tds = double.Parse(cmbTDS.Text) / 100;
                    double net = gross * (1 - tds);

                    MessageBox.Show($"Gross: {gross}\nNet: {Math.Round(net)}");
                    break;
                }
            }
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
                    float[] widths = Enumerable.Repeat(1f, grid.Columns.Count).ToArray();
                    Table table = new Table(widths);

                    foreach (DataGridViewColumn col in grid.Columns)
                        table.AddHeaderCell(col.HeaderText);

                    foreach (DataGridViewRow row in grid.Rows)
                    {
                        if (row.IsNewRow) continue;

                        foreach (DataGridViewCell cell in row.Cells)
                            table.AddCell(cell.Value?.ToString() ?? "");
                    }

                    doc.Add(table);
                }

                MessageBox.Show("PDF Saved");
            }
        }

        void UploadExcel(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel|*.xlsx;*.xls";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var table = result.Tables[0];

                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        var r = table.Rows[i];

                        string portfolio = r[0].ToString();

                        if (!portfolios.ContainsKey(portfolio))
                            portfolios[portfolio] = new List<PortfolioEntry>();

                        PortfolioEntry p = new PortfolioEntry()
                        {
                            PortfolioName = portfolio,
                            InvestorName = r[1].ToString(),
                            TransactionDate = DateTime.Parse(r[2].ToString()),
                            FV = double.Parse(r[3].ToString()),
                            Quantity = int.Parse(r[4].ToString()),
                            BondName = r[5].ToString(),
                            CouponRate = double.Parse(r[6].ToString()),
                            ChequeAmount = double.Parse(r[7].ToString()),
                            Frequency = r[8].ToString(),
                            QuarterStartMonth = r[9].ToString(),
                            MaturityDate = DateTime.Parse(r[10].ToString())
                        };

                        portfolios[portfolio].Add(p);

                        if (!cmbPortfolioSelect.Items.Contains(portfolio))
                            cmbPortfolioSelect.Items.Add(portfolio);
                    }
                }

                MessageBox.Show("Excel Imported");
            }
        }
    }
}
