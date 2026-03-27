using System;
using System.Windows.Forms;

namespace BONDVERSE
{
    public class Dashboard : Form
    {
        public Dashboard()
        {
            Text = "BONDVERSE Dashboard";
            Width = 400; Height = 300;

            Button create = new Button() { Text = "Create Portfolio", Top = 20, Width = 200 };
            Button reports = new Button() { Text = "Reports", Top = 60, Width = 200 };
            Button logout = new Button() { Text = "Logout", Top = 100, Width = 200 };

            create.Click += (s, e) => new PortfolioForm().Show();
            reports.Click += (s, e) => MessageBox.Show("Report feature below");
            logout.Click += (s, e) => Application.Exit();

            Controls.AddRange(new Control[] { create, reports, logout });
        }
    }
}
