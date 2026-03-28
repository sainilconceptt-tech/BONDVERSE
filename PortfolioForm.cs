using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BONDVERSE
{
    public class PortfolioForm : Form
    {
        Dictionary<string, List<PortfolioEntry>> portfolios = new();

        // UI
        TabControl tabs = new TabControl();
        TabPage tabPortfolio = new TabPage("Portfolio");
        TabPage tabReports = new TabPage("Reports");

        Panel sidebar = new Panel();
        Panel topbar = new Panel();
        Panel leftPanel = new Panel();
        Panel rightPanel = new Panel();

        DataGridView grid = new DataGridView();

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
        ComboBox cmbMonth = new ComboBox();

        DateTimePicker dtTrans = new DateTimePicker();
        DateTimePicker dtMat = new DateTimePicker();

        public PortfolioForm()
        {
            Text = "BONDVERSE PRO";
            WindowState = FormWindowState.Maximized;

            BuildUI();
        }

        void BuildUI()
        {
            // Sidebar
            sidebar.Dock = DockStyle.Left;
            sidebar.Width = 200;
            sidebar.BackColor = Color.FromArgb(25, 35, 70);
            Controls.Add(sidebar);

            sidebar.Controls.Add(CreateMenu("Import Excel", 50));
            sidebar.Controls.Add(CreateMenu("Export PDF", 100));
            sidebar.Controls.Add(CreateMenu("TDS Summary", 150));

            // Topbar
            topbar.Dock = DockStyle.Top;
            topbar.Height = 60;
            topbar.BackColor = Color.WhiteSmoke;
            Controls.Add(topbar);

            topbar.Controls.Add(new Label()
            {
                Text = "BONDVERSE PRO DASHBOARD",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                Left = 220,
                Top = 18
            });

            // Tabs
            tabs.Dock = DockStyle.Fill;
            Controls.Add(tabs);

            tabs.TabPages.Add(tabPortfolio);
            tabs.TabPages.Add(tabReports);

            BuildPortfolioTab();
            BuildReportsTab();
        }

        Button CreateMenu(string text, int top)
        {
            return new Button()
            {
                Text = text,
                Top = top,
                Left = 10,
                Width = 180,
                Height = 40,
                BackColor = Color.FromArgb(45, 65, 120),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
        }

        void BuildPortfolioTab()
        {
            leftPanel.Dock = DockStyle.Left;
            leftPanel.Width = 350;
            leftPanel.Padding = new Padding(20);

            rightPanel.Dock = DockStyle.Fill;

            tabPortfolio.Controls.Add(rightPanel);
            tabPortfolio.Controls.Add(leftPanel);

            int y = 10;

            AddField("Portfolio", txtPortfolio, ref y);
            AddField("Investor", txtInvestor, ref y);

            cmbTDS.Items.AddRange(new string[] { "10.4", "20.8" });
            cmbTDS.SelectedIndex = 0;
            AddField("TDS %", cmbTDS, ref y);

            AddField("Date", dtTrans, ref y);
            AddField("FV", txtFV, ref y);
            AddField("Qty", txtQty, ref y);
            AddField("Bond", txtBond, ref y);
            AddField("Coupon", txtCoupon, ref y);
            AddField("Cheque", txtCheque, ref y);

            cmbFreq.Items.AddRange(new string[] { "Monthly", "Quarterly", "Yearly" });
            cmbFreq.SelectedIndex = 0;
            AddField("Frequency", cmbFreq, ref y);

            AddField("Maturity", dtMat, ref y);

            Button btnAdd = CreateAction("Add", y);
            Button btnSubmit = CreateAction("Submit", y);
            btnSubmit.Left = 120;

            leftPanel.Controls.Add(btnAdd);
            leftPanel.Controls.Add(btnSubmit);

            btnAdd.Click += AddEntry;
            btnSubmit.Click += GenerateTable;

            // GRID
            grid.Dock = DockStyle.Fill;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            rightPanel.Controls.Add(grid);
        }

        void BuildReportsTab()
        {
            cmbMonth.Items.AddRange(new string[]
            {
                "Apr","May","Jun","Jul","Aug","Sep",
                "Oct","Nov","Dec","Jan","Feb","Mar"
            });

            cmbMonth.Left = 20;
            cmbMonth.Top = 20;

            Button btnMonth = new Button()
            {
                Text = "Monthly Summary",
                Left = 150,
                Top = 20
            };

            Button btnTDS = new Button()
            {
                Text = "Quarter TDS",
                Left = 300,
                Top = 20
            };

            tabReports.Controls.Add(cmbMonth);
            tabReports.Controls.Add(btnMonth);
            tabReports.Controls.Add(btnTDS);

            btnMonth.Click += ShowMonthSummary;
            btnTDS.Click += ShowTDS;
        }

        void AddField(string label, Control ctrl, ref int y)
        {
            leftPanel.Controls.Add(new Label()
            {
                Text = label,
                Top = y + 5,
                Left = 10
            });

            ctrl.SetBounds(120, y, 180, 25);
            leftPanel.Controls.Add(ctrl);

            y += 35;
        }

        Button CreateAction(string text, int y)
        {
            return new Button()
            {
                Text = text,
                Width = 100,
                Height = 35,
                Top = y + 10,
                Left = 10,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White
            };
        }

        void AddEntry(object sender, EventArgs e)
        {
            string p = txtPortfolio.Text;

            if (!portfolios.ContainsKey(p))
                portfolios[p] = new List<PortfolioEntry>();

            portfolios[p].Add(new PortfolioEntry()
            {
                PortfolioName = p,
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
                    double val = e1.FV * e1.CouponRate / 100 / 12;
                    row[m.ToString("MMM yyyy")] = Math.Round(val);
                }

                dt.Rows.Add(row);
            }

            grid.DataSource = dt;
        }

        void ShowMonthSummary(object sender, EventArgs e)
        {
            if (cmbMonth.SelectedItem == null) return;

            double total = 0;

            foreach (DataGridViewRow row in grid.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.OwningColumn.HeaderText.Contains(cmbMonth.Text))
                        total += Convert.ToDouble(cell.Value ?? 0);
                }
            }

            MessageBox.Show($"Total Interest: {total}");
        }

        void ShowTDS(object sender, EventArgs e)
        {
            double tdsRate = double.Parse(cmbTDS.Text) / 100;

            double total = 0;

            foreach (DataGridViewRow row in grid.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ColumnIndex > 0)
                        total += Convert.ToDouble(cell.Value ?? 0);
                }
            }

            double tds = total * tdsRate;

            MessageBox.Show($"Gross: {total}\nTDS: {tds}\nNet: {total - tds}");
        }
    }
}
