namespace zlShortMsg
{
    partial class frmLogin
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmLogin));
            this.pctForm = new System.Windows.Forms.PictureBox();
            this.lblLogin = new System.Windows.Forms.Label();
            this.lblTip3 = new System.Windows.Forms.Label();
            this.lblTip2 = new System.Windows.Forms.Label();
            this.lblTip1 = new System.Windows.Forms.Label();
            this.pctLogin = new System.Windows.Forms.PictureBox();
            this.chkPwd = new System.Windows.Forms.CheckBox();
            this.lblMin = new System.Windows.Forms.Label();
            this.lblMax = new System.Windows.Forms.Label();
            this.lblTip = new System.Windows.Forms.Label();
            this.pctLeft = new System.Windows.Forms.PictureBox();
            this.cboRole = new System.Windows.Forms.ComboBox();
            this.lblRole = new System.Windows.Forms.Label();
            this.cboList = new System.Windows.Forms.ComboBox();
            this.lblService = new System.Windows.Forms.Label();
            this.txtPwd = new System.Windows.Forms.TextBox();
            this.lblPwd = new System.Windows.Forms.Label();
            this.lblUser = new System.Windows.Forms.Label();
            this.txtUser = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pctForm)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pctLogin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pctLeft)).BeginInit();
            this.SuspendLayout();
            // 
            // pctForm
            // 
            this.pctForm.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pctForm.Location = new System.Drawing.Point(0, 0);
            this.pctForm.Margin = new System.Windows.Forms.Padding(0);
            this.pctForm.Name = "pctForm";
            this.pctForm.Size = new System.Drawing.Size(486, 224);
            this.pctForm.TabIndex = 0;
            this.pctForm.TabStop = false;
            this.pctForm.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pctForm_MouseMove);
            // 
            // lblLogin
            // 
            this.lblLogin.AutoSize = true;
            this.lblLogin.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.World, ((byte)(134)));
            this.lblLogin.Location = new System.Drawing.Point(292, 26);
            this.lblLogin.Name = "lblLogin";
            this.lblLogin.Size = new System.Drawing.Size(84, 18);
            this.lblLogin.TabIndex = 59;
            this.lblLogin.Text = "用户登录";
            // 
            // lblTip3
            // 
            this.lblTip3.AutoSize = true;
            this.lblTip3.BackColor = System.Drawing.Color.Transparent;
            this.lblTip3.CausesValidation = false;
            this.lblTip3.ForeColor = System.Drawing.Color.White;
            this.lblTip3.Location = new System.Drawing.Point(22, 148);
            this.lblTip3.Name = "lblTip3";
            this.lblTip3.Size = new System.Drawing.Size(149, 12);
            this.lblTip3.TabIndex = 58;
            this.lblTip3.Text = "√通过参数配置，简单灵活";
            // 
            // lblTip2
            // 
            this.lblTip2.AutoSize = true;
            this.lblTip2.BackColor = System.Drawing.Color.Transparent;
            this.lblTip2.CausesValidation = false;
            this.lblTip2.ForeColor = System.Drawing.Color.White;
            this.lblTip2.Location = new System.Drawing.Point(22, 122);
            this.lblTip2.Name = "lblTip2";
            this.lblTip2.Size = new System.Drawing.Size(197, 12);
            this.lblTip2.TabIndex = 57;
            this.lblTip2.Text = "√采用数据变动触发，避免轮询查询";
            // 
            // lblTip1
            // 
            this.lblTip1.AutoSize = true;
            this.lblTip1.BackColor = System.Drawing.Color.Transparent;
            this.lblTip1.CausesValidation = false;
            this.lblTip1.ForeColor = System.Drawing.Color.White;
            this.lblTip1.Location = new System.Drawing.Point(22, 96);
            this.lblTip1.Name = "lblTip1";
            this.lblTip1.Size = new System.Drawing.Size(125, 12);
            this.lblTip1.TabIndex = 56;
            this.lblTip1.Text = "√支持常见短信云平台";
            // 
            // pctLogin
            // 
            this.pctLogin.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pctLogin.Image = global::zlShortMsg.Properties.Resources.normal;
            this.pctLogin.Location = new System.Drawing.Point(409, 189);
            this.pctLogin.Name = "pctLogin";
            this.pctLogin.Size = new System.Drawing.Size(55, 24);
            this.pctLogin.TabIndex = 55;
            this.pctLogin.TabStop = false;
            this.pctLogin.Click += new System.EventHandler(this.pctLogin_Click);
            this.pctLogin.MouseEnter += new System.EventHandler(this.picLogin_MouseEnter);
            this.pctLogin.MouseLeave += new System.EventHandler(this.picLogin_MouseLeave);
            // 
            // chkPwd
            // 
            this.chkPwd.AutoSize = true;
            this.chkPwd.Location = new System.Drawing.Point(392, 153);
            this.chkPwd.Name = "chkPwd";
            this.chkPwd.Size = new System.Drawing.Size(72, 16);
            this.chkPwd.TabIndex = 54;
            this.chkPwd.Text = "记住密码";
            this.chkPwd.UseVisualStyleBackColor = true;
            // 
            // lblMin
            // 
            this.lblMin.AutoSize = true;
            this.lblMin.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblMin.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblMin.Location = new System.Drawing.Point(436, 8);
            this.lblMin.Name = "lblMin";
            this.lblMin.Size = new System.Drawing.Size(17, 16);
            this.lblMin.TabIndex = 53;
            this.lblMin.Text = "-";
            this.lblMin.Click += new System.EventHandler(this.lblMin_Click);
            // 
            // lblMax
            // 
            this.lblMax.AutoSize = true;
            this.lblMax.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblMax.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblMax.Location = new System.Drawing.Point(455, 8);
            this.lblMax.Name = "lblMax";
            this.lblMax.Size = new System.Drawing.Size(25, 16);
            this.lblMax.TabIndex = 52;
            this.lblMax.Text = "×";
            this.lblMax.Click += new System.EventHandler(this.lblMax_Click);
            // 
            // lblTip
            // 
            this.lblTip.AutoSize = true;
            this.lblTip.BackColor = System.Drawing.Color.Transparent;
            this.lblTip.Font = new System.Drawing.Font("宋体", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblTip.ForeColor = System.Drawing.Color.White;
            this.lblTip.Location = new System.Drawing.Point(20, 37);
            this.lblTip.Name = "lblTip";
            this.lblTip.Size = new System.Drawing.Size(194, 22);
            this.lblTip.TabIndex = 51;
            this.lblTip.Text = "中联短信发送服务";
            // 
            // pctLeft
            // 
            this.pctLeft.BackColor = System.Drawing.Color.CornflowerBlue;
            this.pctLeft.Location = new System.Drawing.Point(1, 1);
            this.pctLeft.Margin = new System.Windows.Forms.Padding(0);
            this.pctLeft.Name = "pctLeft";
            this.pctLeft.Size = new System.Drawing.Size(234, 222);
            this.pctLeft.TabIndex = 50;
            this.pctLeft.TabStop = false;
            // 
            // cboRole
            // 
            this.cboRole.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.cboRole.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboRole.BackColor = System.Drawing.Color.White;
            this.cboRole.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboRole.FormattingEnabled = true;
            this.cboRole.Items.AddRange(new object[] {
            "Normal",
            "SYSDBA",
            "SYSOPER"});
            this.cboRole.Location = new System.Drawing.Point(312, 151);
            this.cboRole.Name = "cboRole";
            this.cboRole.Size = new System.Drawing.Size(66, 20);
            this.cboRole.TabIndex = 47;
            // 
            // lblRole
            // 
            this.lblRole.AutoSize = true;
            this.lblRole.Location = new System.Drawing.Point(265, 154);
            this.lblRole.Name = "lblRole";
            this.lblRole.Size = new System.Drawing.Size(41, 12);
            this.lblRole.TabIndex = 49;
            this.lblRole.Text = "连接为";
            // 
            // cboList
            // 
            this.cboList.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.cboList.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboList.FormattingEnabled = true;
            this.cboList.Location = new System.Drawing.Point(312, 121);
            this.cboList.Name = "cboList";
            this.cboList.Size = new System.Drawing.Size(152, 20);
            this.cboList.TabIndex = 46;
            // 
            // lblService
            // 
            this.lblService.AutoSize = true;
            this.lblService.Location = new System.Drawing.Point(265, 125);
            this.lblService.Name = "lblService";
            this.lblService.Size = new System.Drawing.Size(41, 12);
            this.lblService.TabIndex = 48;
            this.lblService.Text = "服务器";
            // 
            // txtPwd
            // 
            this.txtPwd.Location = new System.Drawing.Point(312, 94);
            this.txtPwd.Name = "txtPwd";
            this.txtPwd.PasswordChar = '*';
            this.txtPwd.Size = new System.Drawing.Size(152, 21);
            this.txtPwd.TabIndex = 45;
            // 
            // lblPwd
            // 
            this.lblPwd.AutoSize = true;
            this.lblPwd.Location = new System.Drawing.Point(277, 97);
            this.lblPwd.Name = "lblPwd";
            this.lblPwd.Size = new System.Drawing.Size(29, 12);
            this.lblPwd.TabIndex = 44;
            this.lblPwd.Text = "密码";
            // 
            // lblUser
            // 
            this.lblUser.AutoSize = true;
            this.lblUser.Location = new System.Drawing.Point(265, 70);
            this.lblUser.Name = "lblUser";
            this.lblUser.Size = new System.Drawing.Size(41, 12);
            this.lblUser.TabIndex = 43;
            this.lblUser.Text = "用户名";
            // 
            // txtUser
            // 
            this.txtUser.Location = new System.Drawing.Point(312, 67);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(152, 21);
            this.txtUser.TabIndex = 42;
            // 
            // frmLogin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(487, 224);
            this.ControlBox = false;
            this.Controls.Add(this.lblLogin);
            this.Controls.Add(this.lblTip3);
            this.Controls.Add(this.lblTip2);
            this.Controls.Add(this.lblTip1);
            this.Controls.Add(this.pctLogin);
            this.Controls.Add(this.chkPwd);
            this.Controls.Add(this.lblMin);
            this.Controls.Add(this.lblMax);
            this.Controls.Add(this.lblTip);
            this.Controls.Add(this.pctLeft);
            this.Controls.Add(this.cboRole);
            this.Controls.Add(this.lblRole);
            this.Controls.Add(this.cboList);
            this.Controls.Add(this.lblService);
            this.Controls.Add(this.txtPwd);
            this.Controls.Add(this.lblPwd);
            this.Controls.Add(this.lblUser);
            this.Controls.Add(this.txtUser);
            this.Controls.Add(this.pctForm);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmLogin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmLogin";
            this.Load += new System.EventHandler(this.frmLogin_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pctForm)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pctLogin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pctLeft)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pctForm;
        private System.Windows.Forms.Label lblLogin;
        private System.Windows.Forms.Label lblTip3;
        private System.Windows.Forms.Label lblTip2;
        private System.Windows.Forms.Label lblTip1;
        private System.Windows.Forms.PictureBox pctLogin;
        private System.Windows.Forms.CheckBox chkPwd;
        private System.Windows.Forms.Label lblMin;
        private System.Windows.Forms.Label lblMax;
        private System.Windows.Forms.Label lblTip;
        private System.Windows.Forms.PictureBox pctLeft;
        private System.Windows.Forms.ComboBox cboRole;
        private System.Windows.Forms.Label lblRole;
        private System.Windows.Forms.ComboBox cboList;
        private System.Windows.Forms.Label lblService;
        private System.Windows.Forms.TextBox txtPwd;
        private System.Windows.Forms.Label lblPwd;
        private System.Windows.Forms.Label lblUser;
        private System.Windows.Forms.TextBox txtUser;
    }
}