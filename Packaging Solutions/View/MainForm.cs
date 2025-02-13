using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PackagingSolutions.Model;
using PackagingSolutions.Services;
using Aras.IOM;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace PackagingSolutions.View
{
    public partial class MainForm : Form
    {
        private ArasApiService _apiService;
        private ExcelReader _excelReader;
        private List<PackageElement> _elements;
        private TextBox txtFilePath;
        private Label lblServerUrl;
        private TextBox txtServer;
        private TextBox txtUser;
        private TextBox txtDbName;
        private Label lblUser;
        private Label lblPassword;
        private Label lblDbName;
        private MaskedTextBox mskTxtPassword;
        private Button btnLogin;
        private Label lblFilePath;
        private Label lblPackageName;
        private TextBox txtPackageName;
        private Button btnOk;
        private Button btnCancel;
        private Splitter splitter1;
        private CheckBox chkShowPassword;
        private Button btnBrowse;


        public MainForm()
        {
            InitializeComponent();
            _excelReader = new ExcelReader();
        }

        private void InitializeComponent()
        {
            lblServerUrl = new Label();
            txtServer = new TextBox();
            txtUser = new TextBox();
            txtDbName = new TextBox();
            lblUser = new Label();
            lblPassword = new Label();
            lblDbName = new Label();
            mskTxtPassword = new MaskedTextBox();
            btnLogin = new Button();
            lblFilePath = new Label();
            lblPackageName = new Label();
            txtFilePath = new TextBox();
            txtPackageName = new TextBox();
            btnOk = new Button();
            btnCancel = new Button();
            splitter1 = new Splitter();
            btnBrowse = new Button();
            chkShowPassword = new CheckBox();
            SuspendLayout();
            // 
            // lblServerUrl
            // 
            lblServerUrl.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblServerUrl.AutoSize = true;
            lblServerUrl.Location = new Point(10, 31);
            lblServerUrl.Name = "lblServerUrl";
            lblServerUrl.Size = new Size(45, 15);
            lblServerUrl.TabIndex = 0;
            lblServerUrl.Text = "Server :";
            // 
            // txtServer
            // 
            txtServer.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtServer.Location = new Point(81, 28);
            txtServer.Name = "txtServer";
            txtServer.Size = new Size(538, 23);
            txtServer.TabIndex = 1;
            txtServer.TextChanged += txtServer_TextChanged;
            // 
            // txtUser
            // 
            txtUser.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtUser.Location = new Point(81, 70);
            txtUser.Name = "txtUser";
            txtUser.Size = new Size(220, 23);
            txtUser.TabIndex = 2;
            txtUser.TextChanged += txtUser_TextChanged;
            // 
            // txtDbName
            // 
            txtDbName.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            txtDbName.Location = new Point(81, 117);
            txtDbName.Name = "txtDbName";
            txtDbName.Size = new Size(220, 23);
            txtDbName.TabIndex = 4;
            txtDbName.TextChanged += txtDbName_TextChanged;
            // 
            // lblUser
            // 
            lblUser.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblUser.AutoSize = true;
            lblUser.Location = new Point(9, 73);
            lblUser.Name = "lblUser";
            lblUser.Size = new Size(36, 15);
            lblUser.TabIndex = 5;
            lblUser.Text = "User :";
            // 
            // lblPassword
            // 
            lblPassword.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblPassword.AutoSize = true;
            lblPassword.Location = new Point(332, 73);
            lblPassword.Name = "lblPassword";
            lblPassword.Size = new Size(63, 15);
            lblPassword.TabIndex = 6;
            lblPassword.Text = "Password :";
            lblPassword.Click += lblPassword_Click;
            // 
            // lblDbName
            // 
            lblDbName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblDbName.AutoSize = true;
            lblDbName.Location = new Point(9, 120);
            lblDbName.Name = "lblDbName";
            lblDbName.Size = new Size(61, 15);
            lblDbName.TabIndex = 7;
            lblDbName.Text = "Database :";
            // 
            // mskTxtPassword
            // 
            mskTxtPassword.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            mskTxtPassword.Location = new Point(401, 70);
            mskTxtPassword.Name = "mskTxtPassword";
            mskTxtPassword.Size = new Size(197, 23);
            mskTxtPassword.TabIndex = 3;
            mskTxtPassword.UseSystemPasswordChar = true;
            mskTxtPassword.MaskInputRejected += mskTxtPassword_MaskInputRejected;
            // 
            // btnLogin
            // 
            btnLogin.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnLogin.AutoSize = true;
            btnLogin.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            btnLogin.Location = new Point(332, 115);
            btnLogin.Name = "btnLogin";
            btnLogin.Size = new Size(47, 25);
            btnLogin.TabIndex = 9;
            btnLogin.Text = "Login";
            btnLogin.UseVisualStyleBackColor = true;
            btnLogin.Click += btnLogin_Click;
            // 
            // lblFilePath
            // 
            lblFilePath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblFilePath.AutoSize = true;
            lblFilePath.Location = new Point(9, 187);
            lblFilePath.Name = "lblFilePath";
            lblFilePath.Size = new Size(58, 15);
            lblFilePath.TabIndex = 10;
            lblFilePath.Text = "File Path :";
            // 
            // lblPackageName
            // 
            lblPackageName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblPackageName.AutoSize = true;
            lblPackageName.Location = new Point(10, 228);
            lblPackageName.Name = "lblPackageName";
            lblPackageName.Size = new Size(57, 15);
            lblPackageName.TabIndex = 11;
            lblPackageName.Text = "Package :";
            // 
            // txtFilePath
            // 
            txtFilePath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtFilePath.Location = new Point(81, 184);
            txtFilePath.Name = "txtFilePath";
            txtFilePath.Size = new Size(457, 23);
            txtFilePath.TabIndex = 12;
            txtFilePath.TextChanged += txtFilePath_TextChanged;
            // 
            // txtPackageName
            // 
            txtPackageName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtPackageName.Location = new Point(81, 225);
            txtPackageName.Name = "txtPackageName";
            txtPackageName.Size = new Size(220, 23);
            txtPackageName.TabIndex = 14;
            txtPackageName.TextChanged += txtPackageName_TextChanged;
            // 
            // btnOk
            // 
            btnOk.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnOk.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            btnOk.Location = new Point(491, 301);
            btnOk.Name = "btnOk";
            btnOk.Size = new Size(47, 25);
            btnOk.TabIndex = 15;
            btnOk.Text = "OK";
            btnOk.UseVisualStyleBackColor = true;
            btnOk.Click += btnOk_Click;
            // 
            // btnCancel
            // 
            btnCancel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnCancel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            btnCancel.Location = new Point(545, 301);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(74, 25);
            btnCancel.TabIndex = 16;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            // 
            // splitter1
            // 
            splitter1.Location = new Point(0, 0);
            splitter1.Name = "splitter1";
            splitter1.Size = new Size(8, 399);
            splitter1.TabIndex = 17;
            splitter1.TabStop = false;
            // 
            // btnBrowse
            // 
            btnBrowse.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            btnBrowse.Location = new Point(545, 183);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(74, 23);
            btnBrowse.TabIndex = 13;
            btnBrowse.Tag = "";
            btnBrowse.Text = "Browse";
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // chkShowPassword
            // 
            chkShowPassword.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            chkShowPassword.Location = new Point(605, 76);
            chkShowPassword.Name = "chkShowPassword";
            chkShowPassword.Size = new Size(14, 14);
            chkShowPassword.TabIndex = 18;
            chkShowPassword.UseVisualStyleBackColor = true;
            chkShowPassword.CheckedChanged += chkShowPassword_CheckedChanged;
            // 
            // MainForm
            // 
            ClientSize = new Size(664, 399);
            Controls.Add(chkShowPassword);
            Controls.Add(btnBrowse);
            Controls.Add(splitter1);
            Controls.Add(btnCancel);
            Controls.Add(btnOk);
            Controls.Add(txtPackageName);
            Controls.Add(txtFilePath);
            Controls.Add(lblPackageName);
            Controls.Add(lblFilePath);
            Controls.Add(btnLogin);
            Controls.Add(mskTxtPassword);
            Controls.Add(lblDbName);
            Controls.Add(lblPassword);
            Controls.Add(lblUser);
            Controls.Add(txtDbName);
            Controls.Add(txtUser);
            Controls.Add(txtServer);
            Controls.Add(lblServerUrl);
            Name = "MainForm";
            Text = "Packaging Solution";
            Load += MainForm_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void txtServer_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtUser_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtDbName_TextChanged(object sender, EventArgs e)
        {

        }

        private void mskTxtPassword_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {


        }


        private void btnLogin_Click(object sender, EventArgs e)
        {
            string serverUrl = txtServer.Text.Trim();
            string database = txtDbName.Text.Trim();
            string username = txtUser.Text.Trim();
            string password = mskTxtPassword.Text.Trim();

            if (string.IsNullOrWhiteSpace(serverUrl) || string.IsNullOrWhiteSpace(database) ||
                string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                MessageBox.Show("Please fill all fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _apiService = new ArasApiService(serverUrl, database, username, password);

            if (_apiService.Authenticate())
            {
                MessageBox.Show("Authentication successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Authentication failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void txtFilePath_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select a File";
                openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = openFileDialog.FileName;  // Set selected file path in textbox
                }
            }
        }

        private void lblPassword_Click(object sender, EventArgs e)
        {

        }

        private void txtPackageName_TextChanged(object sender, EventArgs e)
        {

        }

        private void chkShowPassword_CheckedChanged(object sender, EventArgs e)
        {
            mskTxtPassword.UseSystemPasswordChar = !chkShowPassword.Checked;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            string filePath = txtFilePath.Text;
            string packageName = txtPackageName.Text;

            if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(packageName))
            {
                MessageBox.Show("Please select an Excel file and enter a package name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (_apiService == null || !_apiService.Authenticate())
            {
                MessageBox.Show("Authentication failed. Please check your credentials.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ProcessPackage(filePath, packageName);
        }

        private void ProcessPackage(string filePath, string packageName)
        {
            try
            {
                if (_apiService == null)
                {
                    MessageBox.Show("Please login first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                // Read elements from Excel
                List<PackageElement> elements = _excelReader.ReadExcel(filePath);
                if (elements.Count == 0)
                {
                    MessageBox.Show("No valid elements found in the Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Check if package exists
                Item package = _apiService.GetPackageByName(packageName);

                // If package doesn't exist, create it
                if (package == null || package.isError() || package.isEmpty())
                {
                    package = _apiService.CreatePackage(packageName);
                    if (package.isError())
                    {
                        MessageBox.Show("Failed to create package.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // Add elements to package
                foreach (var element in elements)
                {
                    _apiService.AddElementToPackage(package, element);
                }

                MessageBox.Show("Package processing completed successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error processing package: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

    }
}
