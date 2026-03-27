using System;
using System.Windows.Forms;

namespace BONDVERSE
{
    public class LoginForm : Form
    {
        TextBox txtUser = new TextBox();
        TextBox txtPin = new TextBox();

        public LoginForm()
        {
            Text = "BONDVERSE Login";
            Width = 300; Height = 200;

            Label l1 = new Label() { Text = "Username", Top = 20, Left = 20 };
            txtUser.Top = 40; txtUser.Left = 20;

            Label l2 = new Label() { Text = "6-digit PIN", Top = 70, Left = 20 };
            txtPin.Top = 90; txtPin.Left = 20; txtPin.PasswordChar = '*';

            Button btn = new Button() { Text = "Login", Top = 120, Left = 20 };

            btn.Click += (s, e) =>
            {
                if (txtUser.Text == "admin" && txtPin.Text == "123456")
                {
                    new Dashboard().Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Invalid Login");
                }
            };

            Controls.AddRange(new Control[] { l1, txtUser, l2, txtPin, btn });
        }
    }
}
