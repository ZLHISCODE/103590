namespace ZLSOFT.HIS.ZyEdit
{
    partial class TextBoxEx
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TextBoxEx));
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.button = new System.Windows.Forms.Button();
            this.textBoxBack = new System.Windows.Forms.TextBox();
            this.textBoxFront = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button
            // 
            this.button.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button.BackColor = System.Drawing.Color.White;
            this.button.FlatAppearance.BorderSize = 0;
            this.button.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button.Image = ((System.Drawing.Image)(resources.GetObject("button.Image")));
            this.button.Location = new System.Drawing.Point(136, 1);
            this.button.Name = "button";
            this.button.Size = new System.Drawing.Size(21, 21);
            this.button.TabIndex = 0;
            this.button.UseVisualStyleBackColor = false;
            this.button.Click += new System.EventHandler(this.button_Click);
            this.button.Enter += new System.EventHandler(this.button_Enter);
            // 
            // textBoxBack
            // 
            this.textBoxBack.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxBack.BackColor = System.Drawing.Color.White;
            this.textBoxBack.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxBack.Enabled = false;
            this.textBoxBack.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBoxBack.Location = new System.Drawing.Point(0, 0);
            this.textBoxBack.Name = "textBoxBack";
            this.textBoxBack.Size = new System.Drawing.Size(159, 23);
            this.textBoxBack.TabIndex = 6;
            // 
            // textBoxFront
            // 
            this.textBoxFront.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxFront.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBoxFront.Location = new System.Drawing.Point(4, 5);
            this.textBoxFront.Name = "textBoxFront";
            this.textBoxFront.Size = new System.Drawing.Size(133, 14);
            this.textBoxFront.TabIndex = 1;
            this.textBoxFront.Enter += new System.EventHandler(this.textBoxFront_Enter);
            this.textBoxFront.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBoxFront_KeyPress);
            // 
            // TextBoxEx
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button);
            this.Controls.Add(this.textBoxFront);
            this.Controls.Add(this.textBoxBack);
            this.Name = "TextBoxEx";
            this.Size = new System.Drawing.Size(160, 23);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button button;
        private System.Windows.Forms.TextBox textBoxBack;
        private System.Windows.Forms.TextBox textBoxFront;
    }
}
