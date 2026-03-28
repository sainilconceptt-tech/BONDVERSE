using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDataReader;

namespace BONDVERSE
{
    public class PortfolioForm : Form
    {
        Dictionary<string, List<PortfolioEntry>> portfolios = new();
        int editIndex = -1;

        ComboBox cmbPortfolio = new ComboBox();
        ComboBox cmbInvestor = new ComboBox();
        ComboBox cmbTDS = new ComboBox();
        ComboBox cmbMonth = new ComboBox();

        TextBox txtBond = new TextBox();
        TextBox txtFV = new TextBox();
        TextBox txtCoupon = new TextBox();

        ComboBox cmbFreq = new ComboBox();
        DateTimePicker dtMat = new DateTimePicker();

        DataGridView grid = new DataGridView();

        public PortfolioForm()
        {
            Text = "BONDVERSE PRO";
            WindowState = FormWindowState.Maximized;

            SplitContainer main = new SplitContainer();
            main.Dock = DockStyle.Fill;
            main.SplitterDistance = 220;
            Controls.Add(main);

            Panel left = new Panel() { Dock = DockStyle.Fill };
            Panel right = new Panel() { Dock = DockStyle.Fill };

            main.Panel1.Controls.Add(left);
            main.Panel2.Controls.Add(right);

            // LEFT PANEL
            Button btnImport = new Button() { Text = "Import Excel", Top = 50, Left = 20, Width = 150 };
            Button btnDelete = new Button() { Text = "Delete Row", Top = 110, Left = 20, Width = 150 };
            Button btnTDS = new Button() { Text = "TDS Summary", Top = 170, Left = 20, Width = 150 };

            left.Controls.Add(btnImport);
            left.Controls.Add(btnDelete);
            left.Controls.Add(btnTDS);

            btnImport.Click += ImportExcel;
            btnDelete.Click += DeleteEntry;
            btnTDS.Click += ShowQuarterTDS;

            // RIGHT SPLIT
            SplitContainer rightSplit = new SplitContainer();
            rightSplit.Dock = DockStyle.Fill;
            rightSplit.Orientation = Orientation.Horizontal;
            rightSplit.SplitterDistance = 240;
            right.Controls.Add(rightSplit);

            Panel form = new Panel() { Dock = DockStyle.Fill };
            Panel table = new Panel() { Dock = DockStyle.Fill };

            rightSplit.Panel1.Controls.Add(form);
            rightSplit.Panel2.Controls.Add(table);

            int y = 10;

            // Portfolio
            form.Controls.Add(new Label() { Text = "Portfolio", Top = y, Left = 10 });
            cmbPortfolio.SetBounds(100, y, 150, 25);
            form.Controls.Add(cmbPortfolio);

            Button btnNew = new Button() { Text = "New", Top = y, Left = 260 };
            form.Controls.Add(btnNew);

            btnNew.Click += (s, e) =>
            {
                string name = Microsoft.VisualBasic.Interaction.InputBox("Portfolio Name:");
                if (!string.IsNullOrEmpty(name) && !portfolios.ContainsKey(name))
                {
                    portfolios[name] = new List<PortfolioEntry>();
                    cmbPortfolio.Items.Add(name);
                    cmbPortfolio.SelectedItem = name;
                }
            };

            // Investor
            form.Controls.Add(new Label() { Text = "Investor", Top = y, Left = 320 });
            cmbInvestor.SetBounds(390, y, 150, 25);
            form.Controls.Add(cmbInvestor);

            // Bond
            form.Controls.Add(new Label() { Text = "Bond", Top = y += 30, Left = 10 });
            txtBond.SetBounds(100, y, 150, 25);
            form.Controls.Add(txtBond);

            // FV
            form.Controls.Add(new Label() { Text = "FV", Top = y += 30, Left = 10 });
            txtFV.SetBounds(100, y, 150, 25);
            form.Controls.Add(txtFV);

            // Coupon
            form.Controls.Add(new Label() { Text = "Coupon %", Top = y += 30, Left = 10 });
            txtCoupon.SetBounds(100, y, 150, 25);
            form.Controls.Add(txtCoupon);

            // Frequency
            form.Controls.Add(new Label() { Text = "Frequency", Top = y += 30, Left = 10 });
            cmbFreq.SetBounds(100, y, 150, 25);
            cmbFreq.Items.AddRange(new[] { "Monthly", "Quarterly", "Yearly" });
            cmbFreq.SelectedIndex = 0;
            form.Controls.Add(cmbFreq);

            // Maturity
            form.Controls.Add(new Label() { Text = "Maturity", Top = y += 30, Left = 10 });
            dtMat.SetBounds(100, y, 150, 25);
            form.Controls.Add(dtMat);

            // TDS
            form.Controls.Add(new Label() { Text = "TDS %", Top = y += 30, Left = 10 });
            cmbTDS.SetBounds(100, y, 150, 25);
            cmbTDS.Items.AddRange(new[] { "10.4", "20.8" });
            cmbTDS.SelectedIndex = 0;
            form.Controls.Add(cmbTDS);

            // Month
            form.Controls.Add(new Label() { Text = "Month", Top = y += 30, Left = 10 });
            cmbMonth.SetBounds(100, y, 150, 25);
            form.Controls.Add(cmbMonth);

            // Buttons
            Button btnAdd = new Button() { Text = "Add/Update", Top = y += 40, Left = 10 };
            Button btnSubmit = new Button() { Text = "Generate", Top = y, Left = 130 };
            Button btnMonth = new Button() { Text = "Month Summary", Top = y, Left = 240 };

            form.Controls.Add(btnAdd);
            form.Controls.Add(btnSubmit);
            form.Controls.Add(btnMonth);

            btnAdd.Click += AddEntry;
            btnSubmit.Click += GenerateTable;
            btnMonth.Click += ShowMonthSummary;

            cmbPortfolio.SelectedIndexChanged += (s, e) => GenerateTable(null, null);

            // GRID
            grid.Dock = DockStyle.Fill;
            table.Controls.Add(grid);
        }

        void AddEntry(object sender, EventArgs e)
        {
            string p = cmbPortfolio.Text;
            if (string.IsNullOrEmpty(p)) { MessageBox.Show("Select Portfolio"); return; }

            var entry = new PortfolioEntry()
            {
                InvestorName = cmbInvestor.Text,
                BondName = txtBond.Text,
                FV = double.Parse(txtFV.Text),
                CouponRate = double.Parse(txtCoupon.Text),
                Frequency = cmbFreq.Text,
                MaturityDate = dtMat.Value
            };

            if (!cmbInvestor.Items.Contains(cmbInvestor.Text))
                cmbInvestor.Items.Add(cmbInvestor.Text);

            if (editIndex >= 0)
            {
                portfolios[p][editIndex] = entry;
                editIndex = -1;
            }
            else
                portfolios[p].Add(entry);

            MessageBox.Show("Saved");
        }

        void DeleteEntry(object sender, EventArgs e)
        {
            if (grid.CurrentRow == null) return;
            string p = cmbPortfolio.Text;
            int i = grid.CurrentRow.Index;

            if (portfolios.ContainsKey(p) && i < portfolios[p].Count)
            {
                portfolios[p].RemoveAt(i);
                GenerateTable(null, null);
            }
        }

        void GenerateTable(object sender, EventArgs e)
        {
            string p = cmbPortfolio.Text;
            if (!portfolios.ContainsKey(p)) return;

            var list = portfolios[p];
            DataTable dt = new DataTable();
            dt.Columns.Add("Bond");

            DateTime start = DateTime.Today;
            DateTime end = list.Max(x => x.MaturityDate);

            List<DateTime> months = new();

            cmbMonth.Items.Clear();

            while (start <= end)
            {
                dt.Columns.Add(start.ToString("MMM yyyy"));
                cmbMonth.Items.Add(start.ToString("MMM yyyy"));
                months.Add(start);
                start = start.AddMonths(1);
            }

            foreach (var b in list)
            {
                var row = dt.NewRow();
                row["Bond"] = b.BondName;

                foreach (var m in months)
                {
                    double interest = b.FV * b.CouponRate / 100 / 12;
                    row[m.ToString("MMM yyyy")] = Math.Round(interest);
                }

                dt.Rows.Add(row);
            }

            double tds = double.Parse(cmbTDS.Text) / 100;

            var total = dt.NewRow();
            total["Bond"] = "TOTAL";

            var net = dt.NewRow();
            net["Bond"] = "NET";

            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName == "Bond") continue;

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
            if (cmbMonth.SelectedItem == null) return;
            string m = cmbMonth.Text;

            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.Cells["Bond"].Value?.ToString() == "TOTAL")
                {
                    double gross = Convert.ToDouble(row.Cells[m].Value);
                    double tds = double.Parse(cmbTDS.Text) / 100;
                    double net = gross * (1 - tds);

                    MessageBox.Show($"Gross: {gross}\nNet: {Math.Round(net)}");
                    break;
                }
            }
        }

        void ShowQuarterTDS(object sender, EventArgs e)
        {
            string p = cmbPortfolio.Text;
            string investor = cmbInvestor.Text;

            if (!portfolios.ContainsKey(p) || string.IsNullOrEmpty(investor))
            {
                MessageBox.Show("Select Portfolio and Investor");
                return;
            }

            var data = portfolios[p].Where(x => x.InvestorName == investor).ToList();

            double tdsRate = double.Parse(cmbTDS.Text) / 100;

            double q1 = 0, q2 = 0, q3 = 0, q4 = 0;

            DateTime fyStart = new DateTime(DateTime.Today.Month >= 4 ? DateTime.Today.Year : DateTime.Today.Year - 1, 4, 1);
            DateTime fyEnd = fyStart.AddYears(1).AddDays(-1);

            foreach (var bond in data)
            {
                DateTime m = fyStart;

                while (m <= fyEnd)
                {
                    double interest = bond.FV * bond.CouponRate / 100 / 12;

                    if (m.Month >= 4 && m.Month <= 6) q1 += interest;
                    else if (m.Month >= 7 && m.Month <= 9) q2 += interest;
                    else if (m.Month >= 10 && m.Month <= 12) q3 += interest;
                    else q4 += interest;

                    m = m.AddMonths(1);
                }
            }

            MessageBox.Show(
                $"Q1: ₹{Math.Round(q1 * tdsRate)}\n" +
                $"Q2: ₹{Math.Round(q2 * tdsRate)}\n" +
                $"Q3: ₹{Math.Round(q3 * tdsRate)}\n" +
                $"Q4: ₹{Math.Round(q4 * tdsRate)}",
                "Quarterly TDS"
            );
        }

        void ImportExcel(object sender, EventArgs e)
        {
            string p = cmbPortfolio.Text;
            if (string.IsNullOrEmpty(p)) { MessageBox.Show("Select Portfolio"); return; }

            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);
                using var reader = ExcelReaderFactory.CreateReader(stream);

                var dt = reader.AsDataSet().Tables[0];

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        portfolios[p].Add(new PortfolioEntry()
                        {
                            InvestorName = row["Investor"].ToString(),
                            BondName = row["Bond Name"].ToString(),
                            FV = Convert.ToDouble(row["FV"]),
                            CouponRate = Convert.ToDouble(row["CouponRate"]),
                            Frequency = row["Frequency"].ToString(),
                            MaturityDate = Convert.ToDateTime(row["MaturityDate"])
                        });
                    }
                    catch { }
                }

                GenerateTable(null, null);
                MessageBox.Show("Excel Imported");
            }
        }
    }
}
