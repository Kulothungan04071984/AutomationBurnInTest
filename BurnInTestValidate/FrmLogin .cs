using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BurnInTestValidate
{
    public partial class FrmLogin : Form
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly UserSession _userSession;
        private TextBox txtUser;
        private TextBox txtPassword;
        private Button btnLogin;
        private Label lblMessage;
        public FrmLogin(IServiceProvider serviceProvider, UserSession userSession)
        {
            InitializeComponent();
            _serviceProvider = serviceProvider;
            SetupUI();
            _userSession = userSession;
        }

        private void SetupUI()
        {
            this.Text = "Secure Login";
            this.Size = new Size(420, 300);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.White;

            Label lblTitle = new Label
            {
                Text = "Burn In Test System",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(80, 20)
            };
            this.Controls.Add(lblTitle);

            txtUser = new TextBox
            {
               Text = "Username",
                Font = new Font("Segoe UI", 10),
                Location = new Point(80, 80),
                Width = 250
            };
            txtUser.GotFocus += RemoveUserPlaceholder;
            txtUser.LostFocus += SetUserPlaceholder;
            this.Controls.Add(txtUser);

            txtPassword = new TextBox
            {
                Text = "Password",
                Font = new Font("Segoe UI", 10),
                Location = new Point(80, 120),
                Width = 250,
                UseSystemPasswordChar = true
            };
            txtPassword.GotFocus += RemovePassPlaceholder;
            txtPassword.LostFocus += SetPassPlaceholder;
            this.Controls.Add(txtPassword);

            btnLogin = new Button
            {
                Text = "LOGIN",
                Width = 250,
                Height = 35,
                Location = new Point(80, 170),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnLogin.FlatAppearance.BorderSize = 0;
            btnLogin.Click += BtnLogin_Click;
            this.Controls.Add(btnLogin);

            lblMessage = new Label
            {
                ForeColor = Color.Red,
                AutoSize = true,
                Location = new Point(80, 215)
            };
            this.Controls.Add(lblMessage);

            this.AcceptButton = btnLogin;
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            Login();
        }
        private void RemoveUserPlaceholder(object sender, EventArgs e)
        {
            if (txtUser.Text == "Username")
            {
                txtUser.Text = "";
                txtUser.ForeColor = Color.Black;
            }
        }

        private void SetUserPlaceholder(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUser.Text))
            {
                txtUser.Text = "Username";
                txtUser.ForeColor = Color.Gray;
            }
        }

        private void RemovePassPlaceholder(object sender, EventArgs e)
        {
            if (txtPassword.Text == "Password")
            {
                txtPassword.Text = "";
                txtPassword.ForeColor = Color.Black;
                txtPassword.UseSystemPasswordChar = true;
            }
        }

        private void SetPassPlaceholder(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                txtPassword.UseSystemPasswordChar = false;
                txtPassword.Text = "Password";
                txtPassword.ForeColor = Color.Gray;
            }
        }
        private void Login()
        {
            if (string.IsNullOrWhiteSpace(txtUser.Text) ||
                string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                lblMessage.Text = "Please enter username and password";
                return;
            }
            _userSession.UserName = txtUser.Text.Trim();
            // Temporary check (replace with DB)
            if (txtUser.Text.Trim() == "admin" && txtPassword.Text.Trim() == "1234")
            {
                this.Hide();
                var burnInForm = _serviceProvider.GetRequiredService<FrmProductSelection>();
                burnInForm.Show();
            }
            else
            {
                lblMessage.Text = "Invalid username or password";
            }
        }
    }
}
