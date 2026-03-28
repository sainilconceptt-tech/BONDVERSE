using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace BONDVERSE
{
    public class PortfolioForm : Form
    {
        SplitContainer mainSplit = new SplitContainer();
        Panel leftPanel = new Panel();
        Panel rightPanel = new Panel();

        DataGridView grid = new DataGridView();

        Button btnImport = new Button();
        Button btnPdf = new Button();
        Button btnTds = new Button();

        public PortfolioForm()
        {
            Text = "BONDVERSE";
            WindowState = FormWindowState.Maximized;

            // Layout (NO OVERLAP)
            mainSplit.Dock = DockStyle.Fill;
            mainSplit.SplitterDistance = 250;

            Controls.Add(mainSplit);

            mainSplit.Panel1.Controls.Add(leftPanel);
            mainSplit.Panel2.Controls.Add(rightPanel);

            BuildLeftPanel();
            BuildRightPanel();
        }

        void BuildLeftPanel()
        {
            leftPanel.Dock = DockStyle.Fill;
            leftPanel.BackColor = System.Drawing.Color.FromArgb(30, 35, 50);

            btnImport.Text = "Import Excel";
            btnPdf.Text = "Export PDF";
            btnTds.Text = "TDS Summary";

            Button[] buttons = { btnImport, btnPdf, btnTds };

            int y = 50;

            foreach (var btn in buttons)
            {
                btn.Width = 200;
                btn.Height = 40;
                btn.Left = 20;
                btn.Top = y;

                btn.BackColor = System.Drawing.Color.FromArgb(50, 60, 90);
                btn.ForeColor = System.Drawing.Color.White;

                leftPanel.Controls.Add(btn);
                y += 60;
            }

            btnImport.Click += (s, e) => MessageBox.Show("Excel Import Ready");
            btnPdf.Click += (s, e) => MessageBox.Show("PDF Export Ready");
            btnTds.Click += (s, e) => ShowTds();
        }

        void BuildRightPanel()
        {
            rightPanel.Dock = DockStyle.Fill;

            grid.Dock = DockStyle.Fill;
            rightPanel.Controls.Add(grid);

            LoadSampleData();
        }

        void LoadSampleData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Bond Name");
            dt.Columns.Add("FV");
            dt.Columns.Add("Apr 2026");
            dt.Columns.Add("May 2026");

            dt.Rows.Add("Bond A", "100000", "2000", "2500");
            dt.Rows.Add("Bond B", "50000", "1000", "1200");

            grid.DataSource = dt;
        }

        void ShowTds()
        {
            MessageBox.Show("TDS Summary will be shown here");
        }
    }
}
