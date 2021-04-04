namespace ZLSOFT.HIS.ZyEdit.From
{
    partial class frmDrugUseEdit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDrugUseEdit));
            this.gpbDrug = new System.Windows.Forms.GroupBox();
            this.txt草药 = new ZLSOFT.HIS.ZyEdit.TextBoxEx();
            this.lbl单位 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txt用量 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.gpbDrug.SuspendLayout();
            this.SuspendLayout();
            // 
            // gpbDrug
            // 
            this.gpbDrug.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(247)))));
            this.gpbDrug.Controls.Add(this.txt草药);
            this.gpbDrug.Controls.Add(this.lbl单位);
            this.gpbDrug.Controls.Add(this.label2);
            this.gpbDrug.Controls.Add(this.txt用量);
            this.gpbDrug.Controls.Add(this.label1);
            this.gpbDrug.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.gpbDrug.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.gpbDrug.Location = new System.Drawing.Point(4, 4);
            this.gpbDrug.Name = "gpbDrug";
            this.gpbDrug.Size = new System.Drawing.Size(307, 79);
            this.gpbDrug.TabIndex = 29;
            this.gpbDrug.TabStop = false;
            this.gpbDrug.Text = "加症用药";
            // 
            // txt草药
            // 
            this.txt草药.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txt草药.ButtonImage = ((System.Drawing.Image)(resources.GetObject("txt草药.ButtonImage")));
            this.txt草药.ButtonToolTipText = "";
            this.txt草药.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt草药.Location = new System.Drawing.Point(81, 18);
            this.txt草药.Name = "txt草药";
            this.txt草药.Size = new System.Drawing.Size(220, 23);
            this.txt草药.TabIndex = 1;
            this.txt草药.TextBoxBackColor = System.Drawing.Color.White;
            this.txt草药.TextBoxReadOnly = false;
            this.txt草药.TextBoxText = "";
            this.txt草药.TextBoxToolTipText = "";
            this.txt草药.ButtonClick += new ZLSOFT.HIS.ZyEdit.TextBoxEx.ButtonClickEventHandler(this.txt草药_ButtonClick);
            this.txt草药.EditBoxKeyPress += new ZLSOFT.HIS.ZyEdit.TextBoxEx.EditBoxKeyPressEventHandler(this.txt草药_EditBoxKeyPress);
            this.txt草药.Validating += new System.ComponentModel.CancelEventHandler(this.txt草药_Validating);
            // 
            // lbl单位
            // 
            this.lbl单位.AutoSize = true;
            this.lbl单位.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lbl单位.ForeColor = System.Drawing.Color.Black;
            this.lbl单位.Location = new System.Drawing.Point(283, 52);
            this.lbl单位.Name = "lbl单位";
            this.lbl单位.Size = new System.Drawing.Size(11, 12);
            this.lbl单位.TabIndex = 37;
            this.lbl单位.Text = "g";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(46, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 36;
            this.label2.Text = "用量";
            // 
            // txt用量
            // 
            this.txt用量.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txt用量.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt用量.ForeColor = System.Drawing.Color.Maroon;
            this.txt用量.Location = new System.Drawing.Point(81, 49);
            this.txt用量.MaxLength = 25;
            this.txt用量.Name = "txt用量";
            this.txt用量.Size = new System.Drawing.Size(196, 21);
            this.txt用量.TabIndex = 2;
            this.txt用量.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txt用量.Enter += new System.EventHandler(this.txt用量_Enter);
            this.txt用量.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt用量_KeyPress);
            this.txt用量.Validating += new System.ComponentModel.CancelEventHandler(this.txt用量_Validating);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(22, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 34;
            this.label1.Text = "选择草药";
            // 
            // btnExit
            // 
            this.btnExit.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnExit.BackColor = System.Drawing.Color.White;
            this.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExit.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExit.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
            this.btnExit.ImageAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.btnExit.Location = new System.Drawing.Point(227, 90);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(81, 24);
            this.btnExit.TabIndex = 4;
            this.btnExit.Text = "退出(&E)";
            this.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnExit.UseVisualStyleBackColor = false;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnSave
            // 
            this.btnSave.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnSave.BackColor = System.Drawing.Color.White;
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSave.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSave.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
            this.btnSave.ImageAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.btnSave.Location = new System.Drawing.Point(128, 90);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(81, 24);
            this.btnSave.TabIndex = 3;
            this.btnSave.Text = "保存(&S)";
            this.btnSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // frmDrugUseEdit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(247)))));
            this.ClientSize = new System.Drawing.Size(316, 118);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.gpbDrug);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmDrugUseEdit";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "加症用药新增";
            this.Load += new System.EventHandler(this.frmDiagEdit_Load);
            this.Shown += new System.EventHandler(this.frmDrugUseEdit_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDiagEdit_KeyDown);
            this.gpbDrug.ResumeLayout(false);
            this.gpbDrug.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gpbDrug;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt用量;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lbl单位;
        private TextBoxEx txt草药;
    }
}