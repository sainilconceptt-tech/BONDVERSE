using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace BONDVERSE
{
    public class PortfolioForm : Form
    {
        Dictionary<string, List<PortfolioEntry>> portfolios = new();

        TabControl tabs = new TabControl();
        DataGridView grid = new DataGridView();
        Chart chart = new Chart();

        // Inputs
        TextBox txtPortfolio = new TextBox();
        TextBox txtInvestor = new TextBox();
        TextBox txtFV = new TextBox();
        TextBox txtQty = new TextBox();
        TextBox txtBond = new TextBox();
        TextBox txtCoupon = new TextBox();
        TextBox txtCheque = new TextBox();

        ComboBox cmbFreq = new ComboBox();
        ComboBox cmbTDS = new ComboBox();

        DateTimePicker dtTrans = new DateTimePicker();
        DateTimePicker dtMat = new DateTimePicker();

        public PortfolioForm()
        {
            Text = "BONDVERSE ENTERPRISE";
            WindowState = FormWindowState.Maximized;

            BuildUI();
        }

        void BuildUI()
        {
            tabs.Dock = DockStyle.Fill;
            Controls.Add(tabs);

            tabs.TabPages.Add(BuildDashboard());
            tabs.TabPages.Add(BuildPortfolio());
            tabs.TabPages.Add(BuildReports());
        }

        // ================= DASHBOARD =================
        TabPage BuildDashboard()
        {
            TabPage tab = new TabPage("Dashboard");

            chart.Dock = DockStyle.Fill;

            chart.ChartAreas.Add(new ChartArea());
            chart.Series.Add("Interest");
            chart.Series["Interest"].ChartType = SeriesChartType.Line;

            tab.Controls.Add(chart);

            return tab;
        }

        // ================= PORTFOLIO =================
        TabPage BuildPortfolio()
        {
            TabPage tab = new TabPage("Portfolio");

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 2;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 350));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // LEFT PANEL
            Panel left = new Panel();
            left.Dock = DockStyle.Fill;
            left.Padding = new Padding(10);

            int y = 10;

            AddField(left, "Portfolio", txtPortfolio, ref y);
            AddField(left, "Investor", txtInvestor, ref y);
            AddField(left, "Date", dtTrans, ref y);
            AddField(left, "FV", txtFV, ref y);
            AddField(left, "Qty", txtQty, ref y);
            AddField(left, "Bond", txtBond, ref y);
            AddField(left, "Coupon", txtCoupon, ref y);
            AddField(left, "Cheque", txtCheque, ref y);

            cmbFreq.Items.AddRange(new[] { "Monthly", "Quarterly", "Yearly" });
            cmbFreq.SelectedIndex = 0;
            AddField(left, "Frequency", cmbFreq, ref y);

            AddField(left, "Maturity", dtMat, ref y);

            Button btnAdd = new Button() { Text = "Add", Width = 100, Top = y + 10 };
            Button btnSubmit = new Button() { Text = "Submit", Width = 100, Left = 120, Top = y + 10 };

            left.Controls.Add(btnAdd);
            left.Controls.Add(btnSubmit);

            btnAdd.Click += AddEntry;
            btnSubmit.Click += GenerateTable;

            // RIGHT PANEL (GRID)
            grid.Dock = DockStyle.Fill;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            layout.Controls.Add(left, 0, 0);
            layout.Controls.Add(grid, 1, 0);

            tab.Controls.Add(layout);

            return tab;
        }

        // ================= REPORT =================
        TabPage BuildReports()
        {
            TabPage tab = new TabPage("Reports");

            Button btnTDS = new Button()
            {
                Text = "Quarter TDS",
                Top = 20,
                Left = 20
            };

            cmbTDS.Items.AddRange(new[] { "10.4", "20.8" });
            cmbTDS.SelectedIndex = 0;
            cmbTDS.Left = 150;
            cmbTDS.Top = 20;

            btnTDS.Click += ShowTDS;

            tab.Controls.Add(btnTDS);
            tab.Controls.Add(cmbTDS);

            return tab;
        }

        void AddField(Panel p, string label, Control ctrl, ref int y)
        {
            Label lbl = new Label() { Text = label, Top = y, Left = 10 };
            ctrl.SetBounds(120, y, 180, 25);

            p.Controls.Add(lbl);
            p.Controls.Add(ctrl);

            y += 35;
        }

        // ================= LOGIC =================

        void AddEntry(object sender, EventArgs e)
        {
            string name = txtPortfolio.Text;

            if (!portfolios.ContainsKey(name))
                portfolios[name] = new List<PortfolioEntry>();

            portfolios[name].Add(new PortfolioEntry()
            {
                PortfolioName = name,
                InvestorName = txtInvestor.Text,
                FV = double.Parse(txtFV.Text),
                Quantity = int.Parse(txtQty.Text),
                BondName = txtBond.Text,
                CouponRate = double.Parse(txtCoupon.Text),
                ChequeAmount = double.Parse(txtCheque.Text),
                Frequency = cmbFreq.Text,
                TransactionDate = dtTrans.Value,
                MaturityDate = dtMat.Value
            });

            MessageBox.Show("Added");
        }

        void GenerateTable(object sender, EventArgs e)
        {
            var list = portfolios[txtPortfolio.Text];

            DataTable dt = new DataTable();
            dt.Columns.Add("Bond");

            DateTime start = DateTime.Today;
            DateTime end = list.Max(x => x.MaturityDate);

            List<DateTime> months = new();

            while (start <= end)
            {
                dt.Columns.Add(start.ToString("MMM yyyy"));
                months.Add(start);
                start = start.AddMonths(1);
            }

            foreach (var e1 in list)
            {
                var row = dt.NewRow();
                row["Bond"] = e1.BondName;

                foreach (var m in months)
                {
                    double interest = e1.FV * e1.CouponRate / 100 / 12;
                    row[m.ToString("MMM yyyy")] = Math.Round(interest);
                }

                dt.Rows.Add(row);
            }

            grid.DataSource = dt;

            UpdateChart(dt);
        }

        void UpdateChart(DataTable dt)
        {
            chart.Series["Interest"].Points.Clear();

            for (int i = 1; i < dt.Columns.Count; i++)
            {
                double total = 0;

                foreach (DataRow row in dt.Rows)
                    total += Convert.ToDouble(row[i]);

                chart.Series["Interest"].Points.AddXY(dt.Columns[i].ColumnName, total);
            }
        }

        void ShowTDS(object sender, EventArgs e)
        {
            double rate = double.Parse(cmbTDS.Text) / 100;

            double total = 0;

            foreach (DataGridViewRow row in grid.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ColumnIndex > 0)
                        total += Convert.ToDouble(cell.Value ?? 0);
                }
            }

            double tds = total * rate;

            MessageBox.Show($"Gross: {total}\nTDS: {tds}\nNet: {total - tds}");
        }
    }
}
