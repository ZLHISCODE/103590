namespace zlShortMsg
{
    partial class frmAddSetting
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAddSetting));
            this.grpPlat = new System.Windows.Forms.GroupBox();
            this.txtSecret = new System.Windows.Forms.TextBox();
            this.lblSecret = new System.Windows.Forms.Label();
            this.txtKey = new System.Windows.Forms.TextBox();
            this.lblKey = new System.Windows.Forms.Label();
            this.cbxPlat = new System.Windows.Forms.ComboBox();
            this.lblPlat = new System.Windows.Forms.Label();
            this.grpTemp = new System.Windows.Forms.GroupBox();
            this.txtKind = new System.Windows.Forms.TextBox();
            this.lblKind = new System.Windows.Forms.Label();
            this.txtText = new System.Windows.Forms.TextBox();
            this.lblText = new System.Windows.Forms.Label();
            this.txtSign = new System.Windows.Forms.TextBox();
            this.lblSign = new System.Windows.Forms.Label();
            this.txtCode = new System.Windows.Forms.TextBox();
            this.lblCode = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.grpPlat.SuspendLayout();
            this.grpTemp.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpPlat
            // 
            this.grpPlat.Controls.Add(this.txtSecret);
            this.grpPlat.Controls.Add(this.lblSecret);
            this.grpPlat.Controls.Add(this.txtKey);
            this.grpPlat.Controls.Add(this.lblKey);
            this.grpPlat.Controls.Add(this.cbxPlat);
            this.grpPlat.Controls.Add(this.lblPlat);
            this.grpPlat.Location = new System.Drawing.Point(12, 12);
            this.grpPlat.Name = "grpPlat";
            this.grpPlat.Size = new System.Drawing.Size(300, 117);
            this.grpPlat.TabIndex = 0;
            this.grpPlat.TabStop = false;
            this.grpPlat.Text = "平台信息";
            // 
            // txtSecret
            // 
            this.txtSecret.Location = new System.Drawing.Point(90, 83);
            this.txtSecret.Name = "txtSecret";
            this.txtSecret.Size = new System.Drawing.Size(192, 21);
            this.txtSecret.TabIndex = 11;
            // 
            // lblSecret
            // 
            this.lblSecret.AutoSize = true;
            this.lblSecret.Location = new System.Drawing.Point(22, 86);
            this.lblSecret.Name = "lblSecret";
            this.lblSecret.Size = new System.Drawing.Size(59, 12);
            this.lblSecret.TabIndex = 10;
            this.lblSecret.Text = "AppSecret";
            // 
            // txtKey
            // 
            this.txtKey.Location = new System.Drawing.Point(90, 56);
            this.txtKey.Name = "txtKey";
            this.txtKey.Size = new System.Drawing.Size(192, 21);
            this.txtKey.TabIndex = 9;
            // 
            // lblKey
            // 
            this.lblKey.AutoSize = true;
            this.lblKey.Location = new System.Drawing.Point(40, 59);
            this.lblKey.Name = "lblKey";
            this.lblKey.Size = new System.Drawing.Size(41, 12);
            this.lblKey.TabIndex = 8;
            this.lblKey.Text = "AppKey";
            // 
            // cbxPlat
            // 
            this.cbxPlat.FormattingEnabled = true;
            this.cbxPlat.Items.AddRange(new object[] {
            "阿里云",
            "腾讯云",
            "华为云"});
            this.cbxPlat.Location = new System.Drawing.Point(90, 30);
            this.cbxPlat.Name = "cbxPlat";
            this.cbxPlat.Size = new System.Drawing.Size(192, 20);
            this.cbxPlat.TabIndex = 7;
            this.cbxPlat.SelectedIndexChanged += new System.EventHandler(this.cbxPlat_SelectedIndexChanged);
            // 
            // lblPlat
            // 
            this.lblPlat.AutoSize = true;
            this.lblPlat.Location = new System.Drawing.Point(52, 33);
            this.lblPlat.Name = "lblPlat";
            this.lblPlat.Size = new System.Drawing.Size(29, 12);
            this.lblPlat.TabIndex = 6;
            this.lblPlat.Text = "平台";
            // 
            // grpTemp
            // 
            this.grpTemp.Controls.Add(this.txtKind);
            this.grpTemp.Controls.Add(this.lblKind);
            this.grpTemp.Controls.Add(this.txtText);
            this.grpTemp.Controls.Add(this.lblText);
            this.grpTemp.Controls.Add(this.txtSign);
            this.grpTemp.Controls.Add(this.lblSign);
            this.grpTemp.Controls.Add(this.txtCode);
            this.grpTemp.Controls.Add(this.lblCode);
            this.grpTemp.Location = new System.Drawing.Point(12, 145);
            this.grpTemp.Name = "grpTemp";
            this.grpTemp.Size = new System.Drawing.Size(300, 195);
            this.grpTemp.TabIndex = 1;
            this.grpTemp.TabStop = false;
            this.grpTemp.Text = "模版信息";
            // 
            // txtKind
            // 
            this.txtKind.Location = new System.Drawing.Point(88, 85);
            this.txtKind.Name = "txtKind";
            this.txtKind.Size = new System.Drawing.Size(194, 21);
            this.txtKind.TabIndex = 16;
            // 
            // lblKind
            // 
            this.lblKind.AutoSize = true;
            this.lblKind.Location = new System.Drawing.Point(28, 88);
            this.lblKind.Name = "lblKind";
            this.lblKind.Size = new System.Drawing.Size(53, 12);
            this.lblKind.TabIndex = 18;
            this.lblKind.Text = "模版分类";
            // 
            // txtText
            // 
            this.txtText.Location = new System.Drawing.Point(87, 113);
            this.txtText.Multiline = true;
            this.txtText.Name = "txtText";
            this.txtText.Size = new System.Drawing.Size(195, 64);
            this.txtText.TabIndex = 17;
            // 
            // lblText
            // 
            this.lblText.AutoSize = true;
            this.lblText.Location = new System.Drawing.Point(28, 116);
            this.lblText.Name = "lblText";
            this.lblText.Size = new System.Drawing.Size(53, 12);
            this.lblText.TabIndex = 16;
            this.lblText.Text = "模版文本";
            // 
            // txtSign
            // 
            this.txtSign.Location = new System.Drawing.Point(87, 59);
            this.txtSign.Name = "txtSign";
            this.txtSign.Size = new System.Drawing.Size(195, 21);
            this.txtSign.TabIndex = 15;
            // 
            // lblSign
            // 
            this.lblSign.AutoSize = true;
            this.lblSign.Location = new System.Drawing.Point(28, 62);
            this.lblSign.Name = "lblSign";
            this.lblSign.Size = new System.Drawing.Size(53, 12);
            this.lblSign.TabIndex = 14;
            this.lblSign.Text = "模版签名";
            // 
            // txtCode
            // 
            this.txtCode.Location = new System.Drawing.Point(87, 28);
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(195, 21);
            this.txtCode.TabIndex = 13;
            // 
            // lblCode
            // 
            this.lblCode.AutoSize = true;
            this.lblCode.Location = new System.Drawing.Point(28, 31);
            this.lblCode.Name = "lblCode";
            this.lblCode.Size = new System.Drawing.Size(53, 12);
            this.lblCode.TabIndex = 0;
            this.lblCode.Text = "模版Code";
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(245, 347);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(67, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAdd.Location = new System.Drawing.Point(177, 347);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(67, 23);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "确定";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // frmAddSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(324, 376);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.grpTemp);
            this.Controls.Add(this.grpPlat);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmAddSetting";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Load += new System.EventHandler(this.frmAddSetting_Load);
            this.grpPlat.ResumeLayout(false);
            this.grpPlat.PerformLayout();
            this.grpTemp.ResumeLayout(false);
            this.grpTemp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpPlat;
        private System.Windows.Forms.TextBox txtSecret;
        private System.Windows.Forms.Label lblSecret;
        private System.Windows.Forms.TextBox txtKey;
        private System.Windows.Forms.Label lblKey;
        private System.Windows.Forms.ComboBox cbxPlat;
        private System.Windows.Forms.Label lblPlat;
        private System.Windows.Forms.GroupBox grpTemp;
        private System.Windows.Forms.TextBox txtText;
        private System.Windows.Forms.Label lblText;
        private System.Windows.Forms.TextBox txtSign;
        private System.Windows.Forms.Label lblSign;
        private System.Windows.Forms.TextBox txtCode;
        private System.Windows.Forms.Label lblCode;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.TextBox txtKind;
        private System.Windows.Forms.Label lblKind;
    }
}