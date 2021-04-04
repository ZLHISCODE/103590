namespace ZLSOFT.HIS.ZyEdit.From
{
    partial class frmZxEdit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmZxEdit));
            this.gpbDrug = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txt证型描述 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txt证型治法 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txt症状表现 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.txt简码 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.gpbDrug.SuspendLayout();
            this.SuspendLayout();
            // 
            // gpbDrug
            // 
            this.gpbDrug.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(247)))));
            this.gpbDrug.Controls.Add(this.label5);
            this.gpbDrug.Controls.Add(this.txt证型描述);
            this.gpbDrug.Controls.Add(this.label7);
            this.gpbDrug.Controls.Add(this.txt证型治法);
            this.gpbDrug.Controls.Add(this.label8);
            this.gpbDrug.Controls.Add(this.txt症状表现);
            this.gpbDrug.Controls.Add(this.label2);
            this.gpbDrug.Controls.Add(this.txtName);
            this.gpbDrug.Controls.Add(this.txt简码);
            this.gpbDrug.Controls.Add(this.label1);
            this.gpbDrug.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.gpbDrug.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.gpbDrug.Location = new System.Drawing.Point(4, 4);
            this.gpbDrug.Name = "gpbDrug";
            this.gpbDrug.Size = new System.Drawing.Size(380, 270);
            this.gpbDrug.TabIndex = 29;
            this.gpbDrug.TabStop = false;
            this.gpbDrug.Text = "中医证型";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(9, 128);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 40;
            this.label5.Text = "证型描述";
            // 
            // txt证型描述
            // 
            this.txt证型描述.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txt证型描述.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt证型描述.Location = new System.Drawing.Point(64, 121);
            this.txt证型描述.MaxLength = 250;
            this.txt证型描述.Multiline = true;
            this.txt证型描述.Name = "txt证型描述";
            this.txt证型描述.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txt证型描述.Size = new System.Drawing.Size(310, 65);
            this.txt证型描述.TabIndex = 4;
            this.txt证型描述.Enter += new System.EventHandler(this.txt证型描述_Enter);
            this.txt证型描述.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtName_KeyPress);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.ForeColor = System.Drawing.Color.Black;
            this.label7.Location = new System.Drawing.Point(9, 54);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 12);
            this.label7.TabIndex = 41;
            this.label7.Text = "证型治法";
            // 
            // txt证型治法
            // 
            this.txt证型治法.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txt证型治法.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt证型治法.Location = new System.Drawing.Point(64, 50);
            this.txt证型治法.MaxLength = 50;
            this.txt证型治法.Multiline = true;
            this.txt证型治法.Name = "txt证型治法";
            this.txt证型治法.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txt证型治法.Size = new System.Drawing.Size(310, 65);
            this.txt证型治法.TabIndex = 3;
            this.txt证型治法.Enter += new System.EventHandler(this.txt证型治法_Enter);
            this.txt证型治法.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtName_KeyPress);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.ForeColor = System.Drawing.Color.Black;
            this.label8.Location = new System.Drawing.Point(9, 193);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(53, 12);
            this.label8.TabIndex = 42;
            this.label8.Text = "症状表现";
            // 
            // txt症状表现
            // 
            this.txt症状表现.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txt症状表现.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt症状表现.Location = new System.Drawing.Point(64, 193);
            this.txt症状表现.MaxLength = 250;
            this.txt症状表现.Multiline = true;
            this.txt症状表现.Name = "txt症状表现";
            this.txt症状表现.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txt症状表现.Size = new System.Drawing.Size(310, 71);
            this.txt症状表现.TabIndex = 5;
            this.txt症状表现.Enter += new System.EventHandler(this.txt症状表现_Enter);
            this.txt症状表现.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtName_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(207, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 36;
            this.label2.Text = "简码";
            // 
            // txtName
            // 
            this.txtName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtName.ForeColor = System.Drawing.Color.Maroon;
            this.txtName.Location = new System.Drawing.Point(64, 20);
            this.txtName.MaxLength = 25;
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(129, 21);
            this.txtName.TabIndex = 1;
            this.txtName.Enter += new System.EventHandler(this.txtName_Enter);
            this.txtName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtName_KeyPress);
            this.txtName.Validating += new System.ComponentModel.CancelEventHandler(this.txtName_Validating);
            // 
            // txt简码
            // 
            this.txt简码.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txt简码.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt简码.Location = new System.Drawing.Point(242, 20);
            this.txt简码.MaxLength = 25;
            this.txt简码.Name = "txt简码";
            this.txt简码.Size = new System.Drawing.Size(129, 21);
            this.txt简码.TabIndex = 2;
            this.txt简码.Enter += new System.EventHandler(this.txt简码_Enter);
            this.txt简码.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtName_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(5, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 34;
            this.label1.Text = "证型名称";
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
            this.btnExit.Location = new System.Drawing.Point(301, 276);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(81, 24);
            this.btnExit.TabIndex = 7;
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
            this.btnSave.Location = new System.Drawing.Point(202, 276);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(81, 24);
            this.btnSave.TabIndex = 6;
            this.btnSave.Text = "保存(&S)";
            this.btnSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // frmZxEdit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(247)))));
            this.ClientSize = new System.Drawing.Size(392, 303);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.gpbDrug);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmZxEdit";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "中医证型新增";
            this.Load += new System.EventHandler(this.frmZxEdit_Load);
            this.Shown += new System.EventHandler(this.frmZxEdit_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmZxEdit_KeyDown);
            this.gpbDrug.ResumeLayout(false);
            this.gpbDrug.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gpbDrug;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.TextBox txt简码;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt证型描述;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txt证型治法;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt症状表现;
    }
}