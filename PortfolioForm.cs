using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

namespace BONDVERSE
{
    public class PortfolioForm : Form
    {
        List<PortfolioEntry> entries = new List<PortfolioEntry>();

        TextBox txtPortfolio = new TextBox();
        TextBox txtInvestor = new TextBox();
        TextBox txtFV = new TextBox();
        TextBox txtQty = new TextBox();
        TextBox txtBondName = new TextBox();
        TextBox txtCoupon = new TextBox();
        TextBox txtCheque = new TextBox();

        ComboBox cmbFreq = new ComboBox();
        ComboBox cmbQuarterStart = new ComboBox();

        DateTimePicker dtTrans = new DateTimePicker();
        DateTimePicker dtMaturity = new DateTimePicker();

        DataGridView grid = new DataGridView();

        Button btnAdd;
        Button btnSubmit;
        Button btnPdf;

        public PortfolioForm()
        {
            Text = "Create Portfolio";
            Width = 1000;
            Height = 600;

            int y = 10;

            Controls.Add(new Label() { Text = "Portfolio Name", Top = y, Left = 10 });
            txtPortfolio.SetBounds(150, y, 150, 25);
            Controls.Add(txtPortfolio);

            Controls.Add(new Label() { Text = "Investor Name", Top = y += 30, Left = 10 });
            txtInvestor.SetBounds(150, y, 150, 25);
            Controls.Add(txtInvestor);

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

            Controls.Add(new Label() { Text = "Interest Frequency", Top = y += 30, Left = 10 });
            cmbFreq.SetBounds(150, y, 150, 25);
            cmbFreq.Items.AddRange(new string[] { "Monthly", "Quarterly", "Yearly" });
            cmbFreq.SelectedIndex = 0;
            Controls.Add(cmbFreq);

            Controls.Add(new Label() { Text = "Quarter Start Month", Top = y += 30, Left = 10 });
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

            btnAdd = new Button() { Text = "Add Entry", Top = y += 40, Left = 10 };
            btnSubmit = new Button() { Text = "Submit", Top = y, Left = 120 };
            btnPdf = new Button() { Text = "Export PDF", Top = 460, Left = 300 };

            Controls.Add(btnAdd);
            Controls.Add(btnSubmit);
            Controls.Add(btnPdf);

            grid.SetBounds(350, 10, 600, 500);
            Controls.Add(grid);

            btnAdd.Click += AddEntry;
            btnSubmit.Click += GenerateTable;
            btnPdf.Click += ExportPdf;
        }

        void AddEntry(object sender, EventArgs e)
        {
            try
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
                MessageBox.Show("Entry Added");
                ClearForm();
            }
            catch
            {
                MessageBox.Show("Invalid input");
            }
        }

        void GenerateTable(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("Bond Name");
            dt.Columns.Add("FV");

            DateTime start = DateTime.Today;
            DateTime end = DateTime.Today.AddYears(2);

            List<DateTime> months = new List<DateTime>();

            while (start <= end)
            {
                dt.Columns.Add(start.ToString("MMM yyyy"));
                months.Add(start);
                start = start.AddMonths(1);
            }

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
                            if (!string.IsNullOrEmpty(e1.QuarterStartMonth))
                            {
                                int startMonth = DateTime.ParseExact(e1.QuarterStartMonth, "MMMM", null).Month;
                                int diff = (m.Year - e1.TransactionDate.Year) * 12 + (m.Month - startMonth);

                                if (diff >= 0 && diff % 3 == 0)
                                    interest = e1.FV * e1.CouponRate / 100 / 4;
                            }
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

            grid.DataSource = dt;
        }

        void ExportPdf(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "PDF files (*.pdf)|*.pdf";
                saveFile.FileName = "BondReport.pdf";

                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFile.FileName;

                    var writer = new PdfWriter(filePath);
                    var pdf = new PdfDocument(writer);
                    var document = new Document(pdf);

                    Table table = new Table(grid.Columns.Count);

                    foreach (DataGridViewColumn col in grid.Columns)
                        table.AddHeaderCell(col.HeaderText);

                    foreach (DataGridViewRow row in grid.Rows)
                    {
                        foreach (DataGridViewCell cell in row.Cells)
                            table.AddCell(cell.Value?.ToString() ?? "");
                    }

                    document.Add(table);
                    document.Close();

                    MessageBox.Show("PDF saved successfully!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void ClearForm()
        {
            txtFV.Text = "";
            txtQty.Text = "";
            txtBondName.Text = "";
            txtCoupon.Text = "";
            txtCheque.Text = "";
        }
    }
}
