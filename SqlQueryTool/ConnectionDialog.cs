using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace SqlQueryTool
{
    public class ConnectionDialog : Form
    {
        public string ConnectionString { get; private set; } = string.Empty;

        private TabControl tabControl;
        private TabPage tabSimple;
        private TabPage tabAdvanced;

        // Simple tab
        private TextBox txtServer;
        private TextBox txtDatabase;
        private RadioButton rdoWindowsAuth;
        private RadioButton rdoSqlAuth;
        private TextBox txtUser;
        private TextBox txtPassword;
        private CheckBox chkSavePassword;
        private CheckBox chkTrustCert;

        // Advanced tab
        private TextBox txtConnectionString;

        // Buttons
        private Button btnTest;
        private Button btnOk;
        private Button btnCancel;

        private Label lblTestResult;

        public ConnectionDialog(string existingConnectionString)
        {
            InitializeComponent();
            if (!string.IsNullOrEmpty(existingConnectionString))
                LoadFromConnectionString(existingConnectionString);
        }

        private void LoadFromConnectionString(string cs)
        {
            txtConnectionString.Text = cs;
            try
            {
                var b = new SqlConnectionStringBuilder(cs);
                txtServer.Text = b.DataSource;
                txtDatabase.Text = b.InitialCatalog;

                if (b.IntegratedSecurity)
                {
                    rdoWindowsAuth.Checked = true;
                }
                else
                {
                    rdoSqlAuth.Checked = true;
                    txtUser.Text = b.UserID;
                    txtPassword.Text = b.Password;
                }
            }
            catch { /* ignore parse errors */ }
        }

        private void rdoAuth_CheckedChanged(object sender, EventArgs e)
        {
            bool sql = rdoSqlAuth.Checked;
            txtUser.Enabled = sql;
            txtPassword.Enabled = sql;
            chkSavePassword.Enabled = sql;
        }

        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            // When switching to Advanced tab, build string from simple fields
            if (tabControl.SelectedTab == tabAdvanced)
                txtConnectionString.Text = BuildConnectionString();
        }

        private string BuildConnectionString()
        {
            // If user is on Advanced tab just return what they typed
            if (tabControl.SelectedTab == tabAdvanced && !string.IsNullOrWhiteSpace(txtConnectionString.Text))
                return txtConnectionString.Text.Trim();

            var b = new SqlConnectionStringBuilder
            {
                DataSource = txtServer.Text.Trim(),
                InitialCatalog = txtDatabase.Text.Trim(),
                TrustServerCertificate = chkTrustCert.Checked
            };

            if (rdoWindowsAuth.Checked)
            {
                b.IntegratedSecurity = true;
            }
            else
            {
                b.IntegratedSecurity = false;
                b.UserID = txtUser.Text.Trim();
                b.Password = txtPassword.Text;
                b.PersistSecurityInfo = chkSavePassword.Checked;
            }

            return b.ConnectionString;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            string cs = BuildConnectionString();
            lblTestResult.Text = "Testing...";
            lblTestResult.ForeColor = Color.Gray;

            try
            {
                using var conn = new SqlConnection(cs);
                conn.Open();
                lblTestResult.Text = "Connection successful!";
                lblTestResult.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                lblTestResult.Text = $"Failed: {ex.Message}";
                lblTestResult.ForeColor = Color.Red;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            ConnectionString = BuildConnectionString();
            if (string.IsNullOrWhiteSpace(ConnectionString))
            {
                MessageBox.Show("Please fill in server and database fields.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            DialogResult = DialogResult.OK;
            Close();
        }

        private void InitializeComponent()
        {
            Text = "Connect to SQL Server";
            Size = new Size(480, 420);
            MinimumSize = new Size(440, 380);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(37, 37, 38);
            ForeColor = Color.FromArgb(212, 212, 212);

            var font = new Font("Segoe UI", 9.5f);

            Label MakeLabel(string text, int x, int y) => new Label
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                Font = font,
                ForeColor = Color.FromArgb(200, 200, 200)
            };

            TextBox MakeTextBox(int x, int y, int w) => new TextBox
            {
                Location = new Point(x, y),
                Width = w,
                Font = font,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            // ── Simple tab ───────────────────────────────────────────────────────
            tabSimple = new TabPage("Connection");
            tabSimple.BackColor = Color.FromArgb(37, 37, 38);

            var lblServer = MakeLabel("Server:", 16, 20);
            txtServer = MakeTextBox(130, 17, 300);
            txtServer.PlaceholderText = "localhost\\SQLEXPRESS";

            var lblDb = MakeLabel("Database:", 16, 55);
            txtDatabase = MakeTextBox(130, 52, 300);
            txtDatabase.PlaceholderText = "master";

            var lblAuth = MakeLabel("Authentication:", 16, 92);
            rdoWindowsAuth = new RadioButton
            {
                Text = "Windows Authentication",
                Location = new Point(130, 90),
                AutoSize = true,
                Font = font,
                ForeColor = Color.FromArgb(200, 200, 200),
                Checked = true
            };
            rdoWindowsAuth.CheckedChanged += rdoAuth_CheckedChanged;

            rdoSqlAuth = new RadioButton
            {
                Text = "SQL Server Authentication",
                Location = new Point(130, 115),
                AutoSize = true,
                Font = font,
                ForeColor = Color.FromArgb(200, 200, 200)
            };
            rdoSqlAuth.CheckedChanged += rdoAuth_CheckedChanged;

            var lblUser = MakeLabel("Login:", 16, 148);
            txtUser = MakeTextBox(130, 145, 200);
            txtUser.Enabled = false;

            var lblPass = MakeLabel("Password:", 16, 180);
            txtPassword = MakeTextBox(130, 177, 200);
            txtPassword.PasswordChar = '●';
            txtPassword.Enabled = false;

            chkSavePassword = new CheckBox
            {
                Text = "Remember password",
                Location = new Point(130, 208),
                AutoSize = true,
                Font = font,
                ForeColor = Color.FromArgb(180, 180, 180),
                Enabled = false
            };

            chkTrustCert = new CheckBox
            {
                Text = "Trust Server Certificate",
                Location = new Point(130, 235),
                AutoSize = true,
                Font = font,
                ForeColor = Color.FromArgb(180, 180, 180),
                Checked = true
            };

            tabSimple.Controls.AddRange(new Control[]
            {
                lblServer, txtServer, lblDb, txtDatabase,
                lblAuth, rdoWindowsAuth, rdoSqlAuth,
                lblUser, txtUser, lblPass, txtPassword,
                chkSavePassword, chkTrustCert
            });

            // ── Advanced tab ─────────────────────────────────────────────────────
            tabAdvanced = new TabPage("Connection String");
            tabAdvanced.BackColor = Color.FromArgb(37, 37, 38);

            var lblCs = MakeLabel("Connection String:", 16, 20);
            txtConnectionString = new TextBox
            {
                Location = new Point(16, 45),
                Width = 415,
                Height = 180,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9.5f),
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            var lblExample = new Label
            {
                Text = "Example: Server=localhost;Database=master;Integrated Security=true;TrustServerCertificate=true",
                Location = new Point(16, 235),
                Width = 415,
                AutoSize = false,
                Height = 40,
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                ForeColor = Color.FromArgb(130, 130, 130)
            };

            tabAdvanced.Controls.AddRange(new Control[] { lblCs, txtConnectionString, lblExample });

            // ── Tab control ──────────────────────────────────────────────────────
            tabControl = new TabControl
            {
                Location = new Point(10, 10),
                Width = 445,
                Height = 290,
                Font = font
            };
            tabControl.TabPages.AddRange(new[] { tabSimple, tabAdvanced });
            tabControl.SelectedIndexChanged += tabControl_SelectedIndexChanged;

            // ── Bottom buttons ───────────────────────────────────────────────────
            lblTestResult = new Label
            {
                Text = string.Empty,
                Location = new Point(14, 310),
                Width = 300,
                AutoSize = false,
                Font = new Font("Segoe UI", 9f),
                ForeColor = Color.Green
            };

            btnTest = new Button
            {
                Text = "Test Connection",
                Location = new Point(14, 334),
                Width = 130,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 72),
                ForeColor = Color.White,
                Font = font
            };
            btnTest.FlatAppearance.BorderSize = 0;
            btnTest.Click += btnTest_Click;

            btnOk = new Button
            {
                Text = "Connect",
                Location = new Point(260, 334),
                Width = 90,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                Font = font
            };
            btnOk.FlatAppearance.BorderSize = 0;
            btnOk.Click += btnOk_Click;

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(358, 334),
                Width = 80,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 72),
                ForeColor = Color.White,
                Font = font,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Controls.AddRange(new Control[]
            {
                tabControl, lblTestResult, btnTest, btnOk, btnCancel
            });
        }
    }
}
