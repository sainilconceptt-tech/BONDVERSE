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
        TextBox txtBond = new TextBox();
        TextBox txtFV = new TextBox();
        TextBox txtCoupon = new TextBox();

        ComboBox cmbFreq = new ComboBox();
        DateTimePicker dtMat = new DateTimePicker();

        DataGridView grid = new DataGridView();

        public PortfolioForm()
        {
            Text = "BONDVERSE";
            WindowState = FormWindowState.Maximized;

            // MAIN LAYOUT
            SplitContainer main = new SplitContainer();
            main.Dock = DockStyle.Fill;
            main.SplitterDistance = 220;
            Controls.Add(main);

            Panel left = new Panel() { Dock = DockStyle.Fill };
            Panel right = new Panel() { Dock = DockStyle.Fill };

            main.Panel1.Controls.Add(left);
            main.Panel2.Controls.Add(right);

            // LEFT BUTTONS
            Button btnImport = new Button() { Text = "Import Excel", Top = 50, Left = 20, Width = 150 };
            Button btnDelete = new Button() { Text = "Delete", Top = 100, Left = 20, Width = 150 };

            left.Controls.Add(btnImport);
            left.Controls.Add(btnDelete);

            btnImport.Click += ImportExcel;
            btnDelete.Click += DeleteEntry;

            // RIGHT SIDE SPLIT
            SplitContainer rightSplit = new SplitContainer();
            rightSplit.Dock = DockStyle.Fill;
            rightSplit.Orientation = Orientation.Horizontal;
            rightSplit.SplitterDistance = 200;

            right.Controls.Add(rightSplit);

            Panel form = new Panel() { Dock = DockStyle.Fill };
            Panel table = new Panel() { Dock = DockStyle.Fill };

            rightSplit.Panel1.Controls.Add(form);
            rightSplit.Panel2.Controls.Add(table);

            // ===== FORM =====
            int y = 10;

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

            form.Controls.Add(new Label() { Text = "Bond", Top = y += 30, Left = 10 });
            txtBond.SetBounds(100, y, 150, 25);
            form.Controls.Add(txtBond);

            form.Controls.Add(new Label() { Text = "FV", Top = y += 30, Left = 10 });
            txtFV.SetBounds(100, y, 150, 25);
            form.Controls.Add(txtFV);

            form.Controls.Add(new Label() { Text = "Coupon %", Top = y += 30, Left = 10 });
            txtCoupon.SetBounds(100, y, 150, 25);
            form.Controls.Add(txtCoupon);

            form.Controls.Add(new Label() { Text = "Frequency", Top = y += 30, Left = 10 });
            cmbFreq.SetBounds(100, y, 150, 25);
            cmbFreq.Items.AddRange(new[] { "Monthly", "Quarterly", "Yearly" });
            cmbFreq.SelectedIndex = 0;
            form.Controls.Add(cmbFreq);

            form.Controls.Add(new Label() { Text = "Maturity", Top = y += 30, Left = 10 });
            dtMat.SetBounds(100, y, 150, 25);
            form.Controls.Add(dtMat);

            Button btnAdd = new Button() { Text = "Add / Update", Top = y += 40, Left = 10 };
            Button btnSubmit = new Button() { Text = "Show Table", Top = y, Left = 130 };

            form.Controls.Add(btnAdd);
            form.Controls.Add(btnSubmit);

            btnAdd.Click += AddEntry;
            btnSubmit.Click += GenerateTable;

            cmbPortfolio.SelectedIndexChanged += (s, e) => GenerateTable(null, null);

            // ===== GRID =====
            grid.Dock = DockStyle.Fill;
            table.Controls.Add(grid);
        }

        // ADD / UPDATE
        void AddEntry(object sender, EventArgs e)
        {
            string p = cmbPortfolio.Text;
            if (string.IsNullOrEmpty(p))
            {
                MessageBox.Show("Select Portfolio");
                return;
            }

            PortfolioEntry entry = new PortfolioEntry()
            {
                BondName = txtBond.Text,
                FV = double.Parse(txtFV.Text),
                CouponRate = double.Parse(txtCoupon.Text),
                Frequency = cmbFreq.Text,
                MaturityDate = dtMat.Value
            };

            if (editIndex >= 0)
            {
                portfolios[p][editIndex] = entry;
                editIndex = -1;
            }
            else
            {
                portfolios[p].Add(entry);
            }

            MessageBox.Show("Saved");
        }

        // DELETE
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

        // TABLE GENERATION
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

            while (start <= end)
            {
                dt.Columns.Add(start.ToString("MMM yyyy"));
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

            grid.DataSource = dt;
        }

        // EXCEL IMPORT
        void ImportExcel(object sender, EventArgs e)
        {
            string p = cmbPortfolio.Text;
            if (string.IsNullOrEmpty(p))
            {
                MessageBox.Show("Select Portfolio first");
                return;
            }

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel|*.xlsx;*.xls";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var dt = result.Tables[0];

                    foreach (DataRow row in dt.Rows)
                    {
                        try
                        {
                            portfolios[p].Add(new PortfolioEntry()
                            {
                                BondName = row["Bond Name"].ToString(),
                                FV = Convert.ToDouble(row["FV"]),
                                CouponRate = Convert.ToDouble(row["CouponRate"]),
                                Frequency = row["Frequency"].ToString(),
                                MaturityDate = Convert.ToDateTime(row["MaturityDate"])
                            });
                        }
                        catch { }
                    }
                }

                GenerateTable(null, null);
                MessageBox.Show("Excel Imported");
            }
        }
    }
}
