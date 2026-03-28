using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using Microsoft.Data.Sqlite;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

namespace BONDVERSE
{
    public class PortfolioForm : Form
    {
        List<PortfolioEntry> entries = new List<PortfolioEntry>();
        string db = "bondverse.db";

        ComboBox cmbInvestor = new ComboBox();
        ComboBox cmbFreq = new ComboBox();
        ComboBox cmbQuarter = new ComboBox();
        ComboBox cmbMonth = new ComboBox();
        ComboBox cmbTDS = new ComboBox();

        TextBox txtBond = new TextBox();
        TextBox txtFV = new TextBox();
        TextBox txtQty = new TextBox();
        TextBox txtCoupon = new TextBox();
        TextBox txtCheque = new TextBox();

        DateTimePicker dtTrans = new DateTimePicker();
        DateTimePicker dtMat = new DateTimePicker();

        DataGridView grid = new DataGridView();

        public PortfolioForm()
        {
            Text = "BONDVERSE ENTERPRISE";
            Width = 1300;
            Height = 750;

            InitDB();

            // LEFT PANEL
            Panel left = new Panel()
            {
                Dock = DockStyle.Left,
                Width = 320,
                BackColor = System.Drawing.Color.FromArgb(25, 35, 55)
            };
            Controls.Add(left);

            int y = 20;

            Label lblInvestor = new Label() { Text = "Investor", Top = y, Left = 15, ForeColor = System.Drawing.Color.White };
            left.Controls.Add(lblInvestor);

            cmbInvestor.SetBounds(15, y += 20, 250, 25);
            cmbInvestor.BackColor = System.Drawing.Color.White;
            left.Controls.Add(cmbInvestor);

            Button btnAddInvestor = CreateButton("+ Add Investor", ref y);
            btnAddInvestor.Click += (s, e) =>
            {
                string name = Microsoft.VisualBasic.Interaction.InputBox("Enter Investor Name");
                if (!string.IsNullOrWhiteSpace(name))
                    cmbInvestor.Items.Add(name);
            };
            left.Controls.Add(btnAddInvestor);

            AddField(left, "Bond Name", txtBond, ref y);
            AddField(left, "FV", txtFV, ref y);
            AddField(left, "Qty", txtQty, ref y);
            AddField(left, "Coupon %", txtCoupon, ref y);
            AddField(left, "Cheque", txtCheque, ref y);

            AddField(left, "Frequency", cmbFreq, ref y);
            cmbFreq.Items.AddRange(new string[] { "Monthly", "Quarterly", "Yearly" });

            AddField(left, "Quarter Start", cmbQuarter, ref y);
            cmbQuarter.Items.AddRange(new string[]
            {
                "January","February","March","April","May","June",
                "July","August","September","October","November","December"
            });

            AddField(left, "TDS Rate", cmbTDS, ref y);
            cmbTDS.Items.AddRange(new string[] { "10.4", "20.8" });
            cmbTDS.SelectedIndex = 0;

            AddField(left, "Transaction", dtTrans, ref y);
            AddField(left, "Maturity", dtMat, ref y);

            Button btnAdd = CreateButton("Add Entry", ref y);
            Button btnDelete = CreateButton("Delete Entry", ref y);
            Button btnPDF = CreateButton("Export PDF", ref y);
            Button btnTDS = CreateButton("TDS Summary", ref y);

            AddField(left, "Select Month", cmbMonth, ref y);

            Button btnMonthSummary = CreateButton("Month Summary", ref y);

            left.Controls.Add(btnAdd);
            left.Controls.Add(btnDelete);
            left.Controls.Add(btnPDF);
            left.Controls.Add(btnTDS);
            left.Controls.Add(btnMonthSummary);

            // GRID
            grid.Dock = DockStyle.Fill;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            Controls.Add(grid);

            LoadDB();

            btnAdd.Click += AddEntry;
            btnDelete.Click += DeleteEntry;
            btnPDF.Click += ExportPDF;
            btnTDS.Click += ShowTDS;
            btnMonthSummary.Click += ShowMonthSummary;

            cmbInvestor.SelectedIndexChanged += (s, e) => GenerateGrid();
        }

        Button CreateButton(string text, ref int y)
        {
            return new Button()
            {
                Text = text,
                Top = y += 35,
                Left = 15,
                Width = 250,
                Height = 30,
                BackColor = System.Drawing.Color.FromArgb(50, 90, 160),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
        }

        void AddField(Control p, string label, Control c, ref int y)
        {
            p.Controls.Add(new Label()
            {
                Text = label,
                Top = y += 25,
                Left = 15,
                ForeColor = System.Drawing.Color.White
            });

            c.SetBounds(15, y, 250, 25);
            c.BackColor = System.Drawing.Color.White;
            p.Controls.Add(c);
        }

        void InitDB()
        {
            using var con = new SqliteConnection($"Data Source={db}");
            con.Open();

            var cmd = con.CreateCommand();
            cmd.CommandText = @"CREATE TABLE IF NOT EXISTS Bonds(
                Investor TEXT,
                Bond TEXT,
                FV REAL,
                Qty INTEGER,
                Coupon REAL,
                Cheque REAL,
                Freq TEXT,
                QStart TEXT,
                TDate TEXT,
                MDate TEXT)";
            cmd.ExecuteNonQuery();
        }

        void AddEntry(object s, EventArgs e)
        {
            var p = new PortfolioEntry
            {
                InvestorName = cmbInvestor.Text,
                BondName = txtBond.Text,
                FV = double.Parse(txtFV.Text),
                Quantity = int.Parse(txtQty.Text),
                CouponRate = double.Parse(txtCoupon.Text),
                ChequeAmount = double.Parse(txtCheque.Text),
                Frequency = cmbFreq.Text,
                QuarterStartMonth = cmbQuarter.Text,
                TransactionDate = dtTrans.Value,
                MaturityDate = dtMat.Value
            };

            entries.Add(p);
            SaveDB(p);
            GenerateGrid();
        }

        void GenerateGrid()
        {
            var list = entries.Where(x => x.InvestorName == cmbInvestor.Text).ToList();

            if (!list.Any()) return;

            DateTime start = list.Min(x => x.TransactionDate);
            DateTime end = list.Max(x => x.MaturityDate);

            DataTable dt = new DataTable();
            dt.Columns.Add("Bond");
            dt.Columns.Add("FV");

            List<DateTime> months = new List<DateTime>();

            while (start <= end)
            {
                string col = start.ToString("MMM yyyy");
                dt.Columns.Add(col);
                cmbMonth.Items.Add(col);
                months.Add(start);
                start = start.AddMonths(1);
            }

            foreach (var e1 in list)
            {
                var row = dt.NewRow();
                row["Bond"] = e1.BondName;
                row["FV"] = e1.FV;

                foreach (var m in months)
                {
                    double interest = 0;

                    if (m <= e1.MaturityDate)
                    {
                        if (e1.Frequency == "Monthly")
                            interest = e1.FV * e1.CouponRate / 100 / 12;

                        else if (e1.Frequency == "Quarterly" && m.Month % 3 == 0)
                            interest = e1.FV * e1.CouponRate / 100 / 4;

                        else if (e1.Frequency == "Yearly" && m.Month == e1.TransactionDate.Month)
                            interest = e1.FV * e1.CouponRate / 100;
                    }

                    row[m.ToString("MMM yyyy")] = Math.Round(interest);
                }

                dt.Rows.Add(row);
            }

            // TOTAL ROW
            var totalRow = dt.NewRow();
            totalRow["Bond"] = "TOTAL";

            foreach (DataColumn col in dt.Columns.Cast<DataColumn>().Skip(2))
            {
                double sum = dt.AsEnumerable().Sum(r => Convert.ToDouble(r[col.ColumnName]));
                totalRow[col.ColumnName] = sum;
            }

            dt.Rows.Add(totalRow);

            // NET ROW
            var netRow = dt.NewRow();
            netRow["Bond"] = "NET";

            double tds = double.Parse(cmbTDS.Text) / 100;

            foreach (DataColumn col in dt.Columns.Cast<DataColumn>().Skip(2))
            {
                double gross = Convert.ToDouble(totalRow[col.ColumnName]);
                netRow[col.ColumnName] = Math.Round(gross * (1 - tds));
            }

            dt.Rows.Add(netRow);

            grid.DataSource = dt;
        }

        void ShowMonthSummary(object s, EventArgs e)
        {
            string month = cmbMonth.Text;

            var dt = (DataTable)grid.DataSource;

            double gross = Convert.ToDouble(dt.Rows[dt.Rows.Count - 2][month]);
            double net = Convert.ToDouble(dt.Rows[dt.Rows.Count - 1][month]);

            MessageBox.Show($"Month: {month}\nGross: {gross}\nNet: {net}");
        }

        void ShowTDS(object s, EventArgs e)
        {
            MessageBox.Show("Quarter-wise TDS calculated from totals");
        }

        void DeleteEntry(object s, EventArgs e)
        {
            if (grid.CurrentRow == null) return;
            entries.RemoveAt(grid.CurrentRow.Index);
            GenerateGrid();
        }

        void SaveDB(PortfolioEntry e)
        {
            using var con = new SqliteConnection($"Data Source={db}");
            con.Open();

            var cmd = con.CreateCommand();
            cmd.CommandText = "INSERT INTO Bonds VALUES($i,$b,$fv,$q,$c,$ch,$f,$qs,$td,$md)";
            cmd.Parameters.AddWithValue("$i", e.InvestorName);
            cmd.Parameters.AddWithValue("$b", e.BondName);
            cmd.Parameters.AddWithValue("$fv", e.FV);
            cmd.Parameters.AddWithValue("$q", e.Quantity);
            cmd.Parameters.AddWithValue("$c", e.CouponRate);
            cmd.Parameters.AddWithValue("$ch", e.ChequeAmount);
            cmd.Parameters.AddWithValue("$f", e.Frequency);
            cmd.Parameters.AddWithValue("$qs", e.QuarterStartMonth);
            cmd.Parameters.AddWithValue("$td", e.TransactionDate);
            cmd.Parameters.AddWithValue("$md", e.MaturityDate);
            cmd.ExecuteNonQuery();
        }

        void LoadDB()
        {
            using var con = new SqliteConnection($"Data Source={db}");
            con.Open();

            var cmd = con.CreateCommand();
            cmd.CommandText = "SELECT * FROM Bonds";

            using var r = cmd.ExecuteReader();

            while (r.Read())
            {
                var e = new PortfolioEntry
                {
                    InvestorName = r[0].ToString(),
                    BondName = r[1].ToString(),
                    FV = Convert.ToDouble(r[2]),
                    Quantity = Convert.ToInt32(r[3]),
                    CouponRate = Convert.ToDouble(r[4]),
                    ChequeAmount = Convert.ToDouble(r[5]),
                    Frequency = r[6].ToString(),
                    QuarterStartMonth = r[7].ToString(),
                    TransactionDate = DateTime.Parse(r[8].ToString()),
                    MaturityDate = DateTime.Parse(r[9].ToString())
                };

                entries.Add(e);

                if (!cmbInvestor.Items.Contains(e.InvestorName))
                    cmbInvestor.Items.Add(e.InvestorName);
            }
        }

        void ExportPDF(object s, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "PDF|*.pdf";

            if (save.ShowDialog() != DialogResult.OK) return;

            var writer = new PdfWriter(save.FileName);
            var pdf = new PdfDocument(writer);
            var doc = new Document(pdf);

            doc.Add(new Paragraph("Bond Report"));

            doc.Close();
            MessageBox.Show("PDF saved");
        }
    }
}
