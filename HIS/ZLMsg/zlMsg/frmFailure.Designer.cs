namespace zlShortMsg
{
    partial class frmFailure
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
            this.components = new System.ComponentModel.Container();
            this.pnlTop = new System.Windows.Forms.Panel();
            this.lblTip = new System.Windows.Forms.Label();
            this.dgvList = new System.Windows.Forms.DataGridView();
            this.chkSend = new System.Windows.Forms.CheckBox();
            this.btnResend = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.Timer = new System.Windows.Forms.Timer(this.components);
            this.btnRefresh = new System.Windows.Forms.Button();
            this.pnlTop.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvList)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlTop
            // 
            this.pnlTop.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlTop.BackColor = System.Drawing.Color.LightGray;
            this.pnlTop.Controls.Add(this.lblTip);
            this.pnlTop.Location = new System.Drawing.Point(0, 0);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Size = new System.Drawing.Size(514, 39);
            this.pnlTop.TabIndex = 3;
            // 
            // lblTip
            // 
            this.lblTip.AutoSize = true;
            this.lblTip.Location = new System.Drawing.Point(10, 14);
            this.lblTip.Name = "lblTip";
            this.lblTip.Size = new System.Drawing.Size(77, 12);
            this.lblTip.TabIndex = 1;
            this.lblTip.Text = "发送失败清单";
            // 
            // dgvList
            // 
            this.dgvList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvList.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvList.GridColor = System.Drawing.SystemColors.Window;
            this.dgvList.Location = new System.Drawing.Point(0, 40);
            this.dgvList.MultiSelect = false;
            this.dgvList.Name = "dgvList";
            this.dgvList.RowTemplate.Height = 23;
            this.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvList.Size = new System.Drawing.Size(510, 226);
            this.dgvList.TabIndex = 4;
            // 
            // chkSend
            // 
            this.chkSend.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkSend.AutoSize = true;
            this.chkSend.Checked = true;
            this.chkSend.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSend.Location = new System.Drawing.Point(12, 282);
            this.chkSend.Name = "chkSend";
            this.chkSend.Size = new System.Drawing.Size(72, 16);
            this.chkSend.TabIndex = 5;
            this.chkSend.Text = "自动重发";
            this.chkSend.UseVisualStyleBackColor = true;
            this.chkSend.CheckedChanged += new System.EventHandler(this.chkSend_CheckedChanged);
            // 
            // btnResend
            // 
            this.btnResend.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnResend.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnResend.Location = new System.Drawing.Point(439, 276);
            this.btnResend.Name = "btnResend";
            this.btnResend.Size = new System.Drawing.Size(67, 27);
            this.btnResend.TabIndex = 6;
            this.btnResend.Text = "重发";
            this.btnResend.UseVisualStyleBackColor = true;
            this.btnResend.Click += new System.EventHandler(this.btnResend_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnDelete.Location = new System.Drawing.Point(366, 276);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(67, 27);
            this.btnDelete.TabIndex = 7;
            this.btnDelete.Text = "删除";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // Timer
            // 
            this.Timer.Enabled = true;
            this.Timer.Interval = 10000;
            this.Timer.Tick += new System.EventHandler(this.Timer_Tick);
            // 
            // btnRefresh
            // 
            this.btnRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefresh.Location = new System.Drawing.Point(293, 276);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(67, 27);
            this.btnRefresh.TabIndex = 8;
            this.btnRefresh.Text = "刷新";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // frmFailure
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.ClientSize = new System.Drawing.Size(513, 315);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnResend);
            this.Controls.Add(this.chkSend);
            this.Controls.Add(this.dgvList);
            this.Controls.Add(this.pnlTop);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frmFailure";
            this.Text = "frmFailure";
            this.Load += new System.EventHandler(this.frmFailure_Load);
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvList)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Label lblTip;
        private System.Windows.Forms.DataGridView dgvList;
        private System.Windows.Forms.CheckBox chkSend;
        private System.Windows.Forms.Button btnResend;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Timer Timer;
        private System.Windows.Forms.Button btnRefresh;
    }
}