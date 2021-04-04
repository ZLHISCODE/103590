namespace ZLSOFT.HIS.ZyEdit
{
    partial class HsComboBox
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HsComboBox));
            this.picDown = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.picDown)).BeginInit();
            this.SuspendLayout();
            // 
            // picDown
            // 
            this.picDown.BackColor = System.Drawing.SystemColors.Control;
            this.picDown.Image = ((System.Drawing.Image)(resources.GetObject("picDown.Image")));
            this.picDown.Location = new System.Drawing.Point(0, 0);
            this.picDown.Name = "picDown";
            this.picDown.Size = new System.Drawing.Size(12, 12);
            this.picDown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picDown.TabIndex = 0;
            this.picDown.TabStop = false;
            ((System.ComponentModel.ISupportInitialize)(this.picDown)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox picDown;
    }
}
