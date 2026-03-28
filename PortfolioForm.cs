using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using ExcelDataReader;

namespace BONDVERSE
{
    public class PortfolioForm : Form
    {
        List<PortfolioEntry> entries = new List<PortfolioEntry>();

        ComboBox cmbInvestor = new ComboBox();
        ComboBox cmbFreq = new ComboBox();
        ComboBox cmbQuarterStart = new ComboBox();
        ComboBox cmbMonth = new ComboBox();
        ComboBox cmbTDS = new ComboBox();

        TextBox txtFV = new TextBox();
        TextBox txtQty = new TextBox();
        TextBox txtBond = new TextBox();
        TextBox txtCoupon = new TextBox();
        TextBox txtCheque = new TextBox();

        DateTimePicker dtTrans = new DateTimePicker();
        DateTimePicker dtMat = new DateTimePicker();

        DataGridView grid = new DataGridView();

        public PortfolioForm()
        {
            Text = "BONDVERSE ENTERPRISE";
            Width = 1200;
            Height = 700;

            // LEFT PANEL
            Panel left = new Panel() { Width = 300, Dock = DockStyle.Left };
            left.BackColor = System.Drawing.Color.FromArgb(30, 30, 60);
            Controls.Add(left);

            int y = 20;

            left.Controls.Add(new Label() { Text = "Investor", Top = y, Left = 20, ForeColor = System.Drawing.Color.White });
            cmbInvestor.SetBounds(20, y += 20, 200, 25);
            left.Controls.Add(cmbInvestor);

            Button btnAddInvestor = new Button() { Text = "+ Add Investor", Top = y += 30, Left = 20, Width = 200 };
            left.Controls.Add(btnAddInvestor);

            btnAddInvestor.Click += (s, e) =>
            {
                string name = Microsoft.VisualBasic.Interaction.InputBox("Enter Investor Name");
                if (!string.IsNullOrWhiteSpace(name) && !cmbInvestor.Items.Contains(name))
                    cmbInvestor.Items.Add(name);
            };

            // INPUTS
            AddLabel(left, "Bond Name", ref y);
            txtBond.SetBounds(20, y, 200, 25); left.Controls.Add(txtBond);

            AddLabel(left, "FV", ref y);
            txtFV.SetBounds(20, y, 200, 25); left.Controls.Add(txtFV);

            AddLabel(left, "Qty", ref y);
            txtQty.SetBounds(20, y, 200, 25); left.Controls.Add(txtQty);

            AddLabel(left, "Coupon %", ref y);
            txtCoupon.SetBounds(20, y, 200, 25); left.Controls.Add(txtCoupon);

            AddLabel(left, "Cheque", ref y);
            txtCheque.SetBounds(20, y, 200, 25); left.Controls.Add(txtCheque);

            AddLabel(left, "Frequency", ref y);
            cmbFreq.SetBounds(20, y, 200, 25);
            cmbFreq.Items.AddRange(new string[] { "Monthly", "Quarterly", "Yearly" });
            left.Controls.Add(cmbFreq);

            AddLabel(left, "Quarter Start", ref y);
            cmbQuarterStart.SetBounds(20, y, 200, 25);
            cmbQuarterStart.Items.AddRange(new string[]
            {
                "January","February","March","April","May","June",
                "July","August","September","October","November","December"
            });
            left.Controls.Add(cmbQuarterStart);

            AddLabel(left, "Transaction", ref y);
            dtTrans.SetBounds(20, y, 200, 25); left.Controls.Add(dtTrans);

            AddLabel(left, "Maturity", ref y);
            dtMat.SetBounds(20, y, 200, 25); left.Controls.Add(dtMat);

            // BUTTONS
            Button btnAdd = new Button() { Text = "Add", Top = y += 40, Left = 20, Width = 90 };
            Button btnDelete = new Button() { Text = "Delete", Top = y, Left = 130, Width = 90 };
            left.Controls.Add(btnAdd);
            left.Controls.Add(btnDelete);

            Button btnImport = new Button() { Text = "Import Excel", Top = y += 40, Left = 20, Width = 200 };
            Button btnPDF = new Button() { Text = "Export PDF", Top = y += 40, Left = 20, Width = 200 };
            Button btnTDS = new Button() { Text = "TDS Summary", Top = y += 40, Left = 20, Width = 200 };

            left.Controls.Add(btnImport);
            left.Controls.Add(btnPDF);
            left.Controls.Add(btnTDS);

            // RIGHT GRID
            grid.Dock = DockStyle.Fill;
            Controls.Add(grid);

            // EVENTS
            btnAdd.Click += AddEntry;
            btnDelete.Click += DeleteEntry;
            btnImport.Click += ImportExcel;
            btnPDF.Click += ExportPDF;
            btnTDS.Click += ShowTDS;

            cmbInvestor.SelectedIndexChanged += (s, e) => RefreshGrid();
        }

        void AddLabel(Control c, string text, ref int y)
        {
            c.Controls.Add(new Label() { Text = text, Top = y += 25, Left = 20, ForeColor = System.Drawing.Color.White });
        }

        void AddEntry(object s, EventArgs e)
        {
            try
            {
                entries.Add(new PortfolioEntry
                {
                    InvestorName = cmbInvestor.Text,
                    BondName = txtBond.Text,
                    FV = double.Parse(txtFV.Text),
                    Quantity = int.Parse(txtQty.Text),
                    CouponRate = double.Parse(txtCoupon.Text),
                    ChequeAmount = double.Parse(txtCheque.Text),
                    Frequency = cmbFreq.Text,
                    QuarterStartMonth = cmbQuarterStart.Text,
                    TransactionDate = dtTrans.Value,
                    MaturityDate = dtMat.Value
                });

                RefreshGrid();
            }
            catch { MessageBox.Show("Invalid data"); }
        }

        void DeleteEntry(object s, EventArgs e)
        {
            if (grid.CurrentRow == null) return;
            entries.RemoveAt(grid.CurrentRow.Index);
            RefreshGrid();
        }

        void RefreshGrid()
        {
            var data = entries.Where(x => x.InvestorName == cmbInvestor.Text).ToList();
            grid.DataSource = data;
        }

        void ImportExcel(object s, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() != DialogResult.OK) return;

            using (var stream = File.Open(ofd.FileName, FileMode.Open))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var ds = reader.AsDataSet();
                var dt = ds.Tables[0];

                foreach (DataRow r in dt.Rows)
                {
                    entries.Add(new PortfolioEntry
                    {
                        InvestorName = cmbInvestor.Text,
                        BondName = r[0].ToString(),
                        FV = Convert.ToDouble(r[1]),
                        CouponRate = Convert.ToDouble(r[2]),
                        Frequency = r[3].ToString(),
                        MaturityDate = Convert.ToDateTime(r[4])
                    });
                }
            }

            RefreshGrid();
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

            foreach (var e1 in entries.Where(x => x.InvestorName == cmbInvestor.Text))
            {
                doc.Add(new Paragraph($"{e1.BondName} - {e1.FV}"));
            }

            doc.Close();
            MessageBox.Show("PDF saved");
        }

        void ShowTDS(object s, EventArgs e)
        {
            var inv = cmbInvestor.Text;
            var list = entries.Where(x => x.InvestorName == inv);

            double total = list.Sum(x => x.FV * x.CouponRate / 100);

            string msg = $"Investor: {inv}\n\nQ1: {total/4}\nQ2: {total/4}\nQ3: {total/4}\nQ4: {total/4}";
            MessageBox.Show(msg);
        }
    }
}
