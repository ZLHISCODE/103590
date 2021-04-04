namespace zlShortMsg
{
    partial class frmLog
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
            this.txtLog = new System.Windows.Forms.TextBox();
            this.pnlTip = new System.Windows.Forms.Panel();
            this.lblTip = new System.Windows.Forms.Label();
            this.chkSave = new System.Windows.Forms.CheckBox();
            this.btnClear = new System.Windows.Forms.Button();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
            this.btnChange = new System.Windows.Forms.Button();
            this.pnlTip.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtLog
            // 
            this.txtLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLog.BackColor = System.Drawing.SystemColors.Window;
            this.txtLog.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtLog.Location = new System.Drawing.Point(3, 41);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtLog.Size = new System.Drawing.Size(640, 275);
            this.txtLog.TabIndex = 1;
            this.txtLog.TabStop = false;
            // 
            // pnlTip
            // 
            this.pnlTip.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlTip.BackColor = System.Drawing.Color.LightGray;
            this.pnlTip.Controls.Add(this.lblTip);
            this.pnlTip.Location = new System.Drawing.Point(0, 0);
            this.pnlTip.Name = "pnlTip";
            this.pnlTip.Size = new System.Drawing.Size(649, 39);
            this.pnlTip.TabIndex = 2;
            // 
            // lblTip
            // 
            this.lblTip.AutoSize = true;
            this.lblTip.Location = new System.Drawing.Point(10, 12);
            this.lblTip.Name = "lblTip";
            this.lblTip.Size = new System.Drawing.Size(77, 12);
            this.lblTip.TabIndex = 1;
            this.lblTip.Text = "短信发送日志";
            // 
            // chkSave
            // 
            this.chkSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkSave.AutoSize = true;
            this.chkSave.Checked = true;
            this.chkSave.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSave.Location = new System.Drawing.Point(12, 324);
            this.chkSave.Name = "chkSave";
            this.chkSave.Size = new System.Drawing.Size(108, 16);
            this.chkSave.TabIndex = 2;
            this.chkSave.Text = "日志同步保存到";
            this.chkSave.UseVisualStyleBackColor = true;
            this.chkSave.CheckedChanged += new System.EventHandler(this.chkSave_CheckedChanged);
            // 
            // btnClear
            // 
            this.btnClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClear.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnClear.Location = new System.Drawing.Point(570, 321);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(67, 25);
            this.btnClear.TabIndex = 5;
            this.btnClear.Text = "清空";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // txtPath
            // 
            this.txtPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtPath.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.txtPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtPath.Location = new System.Drawing.Point(123, 322);
            this.txtPath.Name = "txtPath";
            this.txtPath.ReadOnly = true;
            this.txtPath.Size = new System.Drawing.Size(306, 21);
            this.txtPath.TabIndex = 3;
            this.txtPath.TextChanged += new System.EventHandler(this.txtPath_TextChanged);
            // 
            // btnChange
            // 
            this.btnChange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnChange.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnChange.Location = new System.Drawing.Point(435, 321);
            this.btnChange.Name = "btnChange";
            this.btnChange.Size = new System.Drawing.Size(67, 25);
            this.btnChange.TabIndex = 6;
            this.btnChange.Text = "更改";
            this.btnChange.UseVisualStyleBackColor = true;
            this.btnChange.Click += new System.EventHandler(this.btnChange_Click);
            // 
            // frmLog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.ClientSize = new System.Drawing.Size(645, 357);
            this.ControlBox = false;
            this.Controls.Add(this.btnChange);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.chkSave);
            this.Controls.Add(this.pnlTip);
            this.Controls.Add(this.txtLog);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frmLog";
            this.Text = "frmLog";
            this.Load += new System.EventHandler(this.frmLog_Load);
            this.pnlTip.ResumeLayout(false);
            this.pnlTip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.Panel pnlTip;
        private System.Windows.Forms.Label lblTip;
        private System.Windows.Forms.CheckBox chkSave;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.FolderBrowserDialog folderBrowser;
        private System.Windows.Forms.Button btnChange;
    }
}