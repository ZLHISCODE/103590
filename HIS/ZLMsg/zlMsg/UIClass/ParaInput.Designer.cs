namespace zlShortMsg
{
    partial class ParaInput
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.lblAdd = new System.Windows.Forms.Label();
            this.txtPara = new System.Windows.Forms.TextBox();
            this.lblPara = new System.Windows.Forms.Label();
            this.lblDelete = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblAdd
            // 
            this.lblAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblAdd.AutoSize = true;
            this.lblAdd.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblAdd.Font = new System.Drawing.Font("宋体", 12F);
            this.lblAdd.Location = new System.Drawing.Point(218, 4);
            this.lblAdd.Name = "lblAdd";
            this.lblAdd.Size = new System.Drawing.Size(16, 16);
            this.lblAdd.TabIndex = 23;
            this.lblAdd.Text = "+";
            this.lblAdd.Click += new System.EventHandler(this.lblAdd_Click);
            // 
            // txtPara
            // 
            this.txtPara.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPara.Location = new System.Drawing.Point(69, 2);
            this.txtPara.Name = "txtPara";
            this.txtPara.Size = new System.Drawing.Size(142, 21);
            this.txtPara.TabIndex = 22;
            this.txtPara.Enter += new System.EventHandler(this.txtPara_Enter);
            // 
            // lblPara
            // 
            this.lblPara.AutoSize = true;
            this.lblPara.Location = new System.Drawing.Point(10, 6);
            this.lblPara.Name = "lblPara";
            this.lblPara.Size = new System.Drawing.Size(53, 12);
            this.lblPara.TabIndex = 21;
            this.lblPara.Text = "业务参数";
            // 
            // lblDelete
            // 
            this.lblDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblDelete.AutoSize = true;
            this.lblDelete.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblDelete.Font = new System.Drawing.Font("宋体", 12F);
            this.lblDelete.Location = new System.Drawing.Point(236, 4);
            this.lblDelete.Name = "lblDelete";
            this.lblDelete.Size = new System.Drawing.Size(16, 16);
            this.lblDelete.TabIndex = 24;
            this.lblDelete.Text = "-";
            this.lblDelete.Click += new System.EventHandler(this.lblDelete_Click);
            // 
            // ParaInput
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblDelete);
            this.Controls.Add(this.lblAdd);
            this.Controls.Add(this.txtPara);
            this.Controls.Add(this.lblPara);
            this.Name = "ParaInput";
            this.Size = new System.Drawing.Size(255, 25);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblAdd;
        private System.Windows.Forms.TextBox txtPara;
        private System.Windows.Forms.Label lblPara;
        private System.Windows.Forms.Label lblDelete;
    }
}
