using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms.DataVisualization.Charting;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

namespace BONDVERSE
{
    public class PortfolioForm : Form
    {
        List<PortfolioEntry> entries = new List<PortfolioEntry>();
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
        ComboBox cmbMonthSelect = new ComboBox();

        DateTimePicker dtTrans = new DateTimePicker();
        DateTimePicker dtMaturity = new DateTimePicker();

        DataGridView grid = new DataGridView();

        Button btnAdd, btnEdit, btnDelete, btnSubmit, btnPdf, btnSave, btnLoad, btnChart, btnMonthSummary;

        public PortfolioForm()
        {
            Text = "BONDVERSE";
            Width = 1300;
            Height = 750;

            int y = 10;

            Controls.Add(new Label() { Text = "Portfolio Name", Top = y, Left = 10 });
            txtPortfolio.SetBounds(150, y, 150, 25);
            Controls.Add(txtPortfolio);

            Controls.Add(new Label() { Text = "Investor Name", Top = y += 30, Left = 10 });
            txtInvestor.SetBounds(150, y, 150, 25);
            Controls.Add(txtInvestor);

            Controls.Add(new Label() { Text = "TDS Rate (%)", Top = y += 30, Left = 10 });
            cmbTDS.SetBounds(150, y, 150, 25);
            cmbTDS.Items.AddRange(new string[] { "10.4", "20.8" });
            cmbTDS.SelectedIndex = 0;
            Controls.Add(cmbTDS);

            Controls.Add(new Label() { Text = "Transaction Date", Top = y += 30, Left = 10 });
            dtTrans.SetBounds(150, y, 150, 25);
            Controls.Add(dtTrans);

            Controls.Add(new Label() { Text = "Bond FV", Top = y += 30, Left = 10 });
            txtFV.SetBounds(150, y, 150, 25);
            Controls.Add(txtFV);

            Controls.Add(new Label() { Text = "Quantity", Top = y += 30, Left = 10 });
            txtQty.SetBounds(150, y, 150, 25);
            Controls.Add(txtQty);

            Controls.Add(new Label() { Text = "Bond Name", Top = y += 30, Left = 10 });
            txtBondName.SetBounds(150, y, 150, 25);
            Controls.Add(txtBondName);

            Controls.Add(new Label() { Text = "Coupon Rate (%)", Top = y += 30, Left = 10 });
            txtCoupon.SetBounds(150, y, 150, 25);
            Controls.Add(txtCoupon);

            Controls.Add(new Label() { Text = "Cheque Amount", Top = y += 30, Left = 10 });
            txtCheque.SetBounds(150, y, 150, 25);
            Controls.Add(txtCheque);

            Controls.Add(new Label() { Text = "Frequency", Top = y += 30, Left = 10 });
            cmbFreq.SetBounds(150, y, 150, 25);
            cmbFreq.Items.AddRange(new string[] { "Monthly", "Quarterly", "Yearly" });
            cmbFreq.SelectedIndex = 0;
            Controls.Add(cmbFreq);

            Controls.Add(new Label() { Text = "Quarter Start", Top = y += 30, Left = 10 });
            cmbQuarterStart.SetBounds(150, y, 150, 25);
            cmbQuarterStart.Items.AddRange(new string[]
            {
                "January","February","March","April","May","June",
                "July","August","September","October","November","December"
            });
            cmbQuarterStart.Visible = false;
            Controls.Add(cmbQuarterStart);

            cmbFreq.SelectedIndexChanged += (s, e) =>
            {
                cmbQuarterStart.Visible = cmbFreq.Text == "Quarterly";
            };

            Controls.Add(new Label() { Text = "Maturity Date", Top = y += 30, Left = 10 });
            dtMaturity.SetBounds(150, y, 150, 25);
            Controls.Add(dtMaturity);

            btnAdd = new Button() { Text = "Add", Top = y += 40, Left = 10 };
            btnEdit = new Button() { Text = "Edit", Top = y, Left = 80 };
            btnDelete = new Button() { Text = "Delete", Top = y, Left = 150 };
            btnSubmit = new Button() { Text = "Submit", Top = y, Left = 230 };
            btnPdf = new Button() { Text = "PDF", Top = y, Left = 320 };
            btnSave = new Button() { Text = "Save", Top = y, Left = 400 };
            btnLoad = new Button() { Text = "Load", Top = y, Left = 480 };
            btnChart = new Button() { Text = "Chart", Top = y, Left = 560 };

            Controls.Add(btnAdd);
            Controls.Add(btnEdit);
            Controls.Add(btnDelete);
            Controls.Add(btnSubmit);
            Controls.Add(btnPdf);
            Controls.Add(btnSave);
            Controls.Add(btnLoad);
            Controls.Add(btnChart);

            Controls.Add(new Label() { Text = "Select Month", Top = y += 40, Left = 10 });
            cmbMonthSelect.SetBounds(150, y, 150, 25);
            Controls.Add(cmbMonthSelect);

            btnMonthSummary = new Button() { Text = "Month Summary", Top = y, Left = 320 };
            Controls.Add(btnMonthSummary);

            grid.Dock = DockStyle.Right;
            grid.Width = 800;
            Controls.Add(grid);

            btnAdd.Click += AddEntry;
            btnEdit.Click += EditEntry;
            btnDelete.Click += DeleteEntry;
            btnSubmit.Click += GenerateTable;
            btnPdf.Click += ExportPdf;
            btnSave.Click += SavePortfolio;
            btnLoad.Click += LoadPortfolio;
            btnChart.Click += ShowChart;
            btnMonthSummary.Click += ShowMonthSummary;
        }

        void AddEntry(object sender, EventArgs e)
        {
            PortfolioEntry p = new PortfolioEntry()
            {
                PortfolioName = txtPortfolio.Text,
                InvestorName = txtInvestor.Text,
                TransactionDate = dtTrans.Value,
                FV = double.Parse(txtFV.Text),
                Quantity = int.Parse(txtQty.Text),
                BondName = txtBondName.Text,
                CouponRate = double.Parse(txtCoupon.Text),
                ChequeAmount = double.Parse(txtCheque.Text),
                Frequency = cmbFreq.Text,
                MaturityDate = dtMaturity.Value,
                QuarterStartMonth = cmbQuarterStart.Text
            };

            entries.Add(p);
            MessageBox.Show("Added");
        }

        void EditEntry(object sender, EventArgs e)
        {
            if (grid.CurrentRow == null) return;

            int index = grid.CurrentRow.Index;
            if (index >= entries.Count) return;

            var e1 = entries[index];

            txtBondName.Text = e1.BondName;
            txtFV.Text = e1.FV.ToString();
            txtCoupon.Text = e1.CouponRate.ToString();

            entries.RemoveAt(index);
        }

        void DeleteEntry(object sender, EventArgs e)
        {
            if (grid.CurrentRow == null) return;

            int index = grid.CurrentRow.Index;
            if (index < entries.Count)
            {
                entries.RemoveAt(index);
                GenerateTable(null, null);
            }
        }

        void GenerateTable(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("Bond Name");
            dt.Columns.Add("FV");

            DateTime start = DateTime.Today;
            DateTime maxDate = entries.Max(x => x.MaturityDate);

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

            foreach (var e1 in entries)
            {
                var row = dt.NewRow();
                row["Bond Name"] = e1.BondName;
                row["FV"] = e1.FV;

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

            double tdsRate = double.Parse(cmbTDS.Text) / 100;

            var totalRow = dt.NewRow();
            totalRow["Bond Name"] = "TOTAL";

            var netRow = dt.NewRow();
            netRow["Bond Name"] = "NET";

            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName == "Bond Name" || col.ColumnName == "FV") continue;

                double sum = dt.AsEnumerable().Sum(r => Convert.ToDouble(r[col]));
                totalRow[col] = sum;
                netRow[col] = Math.Round(sum * (1 - tdsRate));
            }

            dt.Rows.Add(totalRow);
            dt.Rows.Add(netRow);

            grid.DataSource = dt;
        }

        void ShowMonthSummary(object sender, EventArgs e)
        {
            if (cmbMonthSelect.SelectedItem == null) return;

            string month = cmbMonthSelect.SelectedItem.ToString();

            double gross = 0;

            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.Cells["Bond Name"].Value?.ToString() == "TOTAL")
                {
                    gross = Convert.ToDouble(row.Cells[month].Value);
                    break;
                }
            }

            double tdsRate = double.Parse(cmbTDS.Text) / 100;
            double net = gross * (1 - tdsRate);

            MessageBox.Show($"Month: {month}\nGross: {gross}\nNet: {Math.Round(net)}");
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

        void SavePortfolio(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Portfolio|*.dat";

            if (save.ShowDialog() == DialogResult.OK)
            {
                BinaryFormatter bf = new BinaryFormatter();
                using (FileStream fs = new FileStream(save.FileName, FileMode.Create))
                {
                    bf.Serialize(fs, entries);
                }
            }
        }

        void LoadPortfolio(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Portfolio|*.dat";

            if (open.ShowDialog() == DialogResult.OK)
            {
                BinaryFormatter bf = new BinaryFormatter();
                using (FileStream fs = new FileStream(open.FileName, FileMode.Open))
                {
                    entries = (List<PortfolioEntry>)bf.Deserialize(fs);
                }

                GenerateTable(null, null);
            }
        }

        void ShowChart(object sender, EventArgs e)
        {
            Form f = new Form();
            Chart chart = new Chart();
            chart.Dock = DockStyle.Fill;
            chart.ChartAreas.Add(new ChartArea());

            Series s = new Series();
            s.ChartType = SeriesChartType.Line;

            foreach (DataGridViewColumn col in grid.Columns)
            {
                if (col.Name == "Bond Name" || col.Name == "FV") continue;

                double total = 0;

                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.Cells["Bond Name"].Value?.ToString() == "TOTAL")
                    {
                        total = Convert.ToDouble(row.Cells[col.Name].Value);
                        break;
                    }
                }

                s.Points.AddXY(col.Name, total);
            }

            chart.Series.Add(s);
            f.Controls.Add(chart);
            f.Show();
        }
    }
}
