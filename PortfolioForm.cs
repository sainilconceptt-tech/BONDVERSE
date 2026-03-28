using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.WinForms;

namespace BONDVERSE
{
    public class PortfolioForm : Form
    {
        SplitContainer mainSplit = new SplitContainer();
        Panel leftPanel = new Panel();
        Panel rightPanel = new Panel();

        DataGridView grid = new DataGridView();
        CartesianChart chart = new CartesianChart();

        Button btnImport = new Button();
        Button btnPdf = new Button();
        Button btnTds = new Button();

        public PortfolioForm()
        {
            Text = "BONDVERSE ENTERPRISE";
            WindowState = FormWindowState.Maximized;

            // SPLIT CONTAINER (NO OVERLAP GUARANTEE)
            mainSplit.Dock = DockStyle.Fill;
            mainSplit.SplitterDistance = 250;

            Controls.Add(mainSplit);

            leftPanel.Dock = DockStyle.Fill;
            rightPanel.Dock = DockStyle.Fill;

            mainSplit.Panel1.Controls.Add(leftPanel);
            mainSplit.Panel2.Controls.Add(rightPanel);

            BuildLeftPanel();
            BuildRightPanel();
        }

        void BuildLeftPanel()
        {
            leftPanel.BackColor = System.Drawing.Color.FromArgb(20, 25, 40);

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
                btn.BackColor = System.Drawing.Color.FromArgb(40, 50, 80);
                btn.ForeColor = System.Drawing.Color.White;

                leftPanel.Controls.Add(btn);
                y += 60;
            }

            btnImport.Click += ImportExcel;
            btnPdf.Click += ExportPdf;
            btnTds.Click += ShowTds;
        }

        void BuildRightPanel()
        {
            // GRID
            grid.Dock = DockStyle.Top;
            grid.Height = 300;

            // CHART
            chart.Dock = DockStyle.Fill;

            rightPanel.Controls.Add(chart);
            rightPanel.Controls.Add(grid);

            LoadDummyData();
            LoadChart();
        }

        void LoadDummyData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Month");
            dt.Columns.Add("Interest");

            dt.Rows.Add("Apr", 2000);
            dt.Rows.Add("May", 3000);
            dt.Rows.Add("Jun", 2500);

            grid.DataSource = dt;
        }

        void LoadChart()
        {
            chart.Series = new ISeries[]
            {
                new ColumnSeries<double>
                {
                    Values = new double[] { 2000, 3000, 2500 }
                }
            };
        }

        void ImportExcel(object sender, EventArgs e)
        {
            MessageBox.Show("Excel Import Ready");
        }

        void ExportPdf(object sender, EventArgs e)
        {
            MessageBox.Show("PDF Export Ready");
        }

        void ShowTds(object sender, EventArgs e)
        {
            MessageBox.Show("TDS Summary Ready");
        }
    }
}
