namespace zlShortMsg
{
    partial class frmMsgSet
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
            this.pnlTop = new System.Windows.Forms.Panel();
            this.lblTop = new System.Windows.Forms.Label();
            this.dgvSetting = new System.Windows.Forms.DataGridView();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnInsert = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.pnlTop.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSetting)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlTop
            // 
            this.pnlTop.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlTop.BackColor = System.Drawing.Color.LightGray;
            this.pnlTop.Controls.Add(this.lblTop);
            this.pnlTop.Location = new System.Drawing.Point(0, 0);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Size = new System.Drawing.Size(511, 39);
            this.pnlTop.TabIndex = 4;
            // 
            // lblTop
            // 
            this.lblTop.AutoSize = true;
            this.lblTop.Location = new System.Drawing.Point(10, 14);
            this.lblTop.Name = "lblTop";
            this.lblTop.Size = new System.Drawing.Size(77, 12);
            this.lblTop.TabIndex = 1;
            this.lblTop.Text = "短信接口配置";
            // 
            // dgvSetting
            // 
            this.dgvSetting.AllowUserToResizeRows = false;
            this.dgvSetting.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvSetting.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dgvSetting.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Raised;
            this.dgvSetting.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSetting.GridColor = System.Drawing.SystemColors.Window;
            this.dgvSetting.Location = new System.Drawing.Point(0, 40);
            this.dgvSetting.MultiSelect = false;
            this.dgvSetting.Name = "dgvSetting";
            this.dgvSetting.RowHeadersWidth = 25;
            this.dgvSetting.RowTemplate.Height = 23;
            this.dgvSetting.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvSetting.Size = new System.Drawing.Size(510, 226);
            this.dgvSetting.TabIndex = 5;
            // 
            // btnUpdate
            // 
            this.btnUpdate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnUpdate.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnUpdate.Location = new System.Drawing.Point(360, 277);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(67, 27);
            this.btnUpdate.TabIndex = 9;
            this.btnUpdate.Text = "修改";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnInsert
            // 
            this.btnInsert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnInsert.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnInsert.Location = new System.Drawing.Point(433, 277);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(67, 27);
            this.btnInsert.TabIndex = 8;
            this.btnInsert.Text = "新增";
            this.btnInsert.UseVisualStyleBackColor = true;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnDelete.Location = new System.Drawing.Point(287, 277);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(67, 27);
            this.btnDelete.TabIndex = 10;
            this.btnDelete.Text = "删除";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // frmMsgSet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.ClientSize = new System.Drawing.Size(512, 316);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.btnInsert);
            this.Controls.Add(this.dgvSetting);
            this.Controls.Add(this.pnlTop);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frmMsgSet";
            this.Text = "frmMsgSet";
            this.Load += new System.EventHandler(this.frmMsgSet_Load);
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSetting)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Label lblTop;
        private System.Windows.Forms.DataGridView dgvSetting;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.Button btnDelete;
    }
}