using System;
using DevExpress.XtraWaitForm;

namespace ZLSOFT.HIS.PreTriage
{
    public class MyWaitFrom : WaitForm
    {
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Timer tirTimer;
        private System.ComponentModel.IContainer components;
        private ProgressPanel progressPanel1;
        private int intCloseTime = 0;
        private int intTime = 0;

        public MyWaitFrom()
        {
            InitializeComponent();
            this.progressPanel1.AutoHeight = true;
        }

        #region Overrides

        public override void SetCaption(string caption)
        {
            base.SetCaption(caption);
            this.progressPanel1.Caption = caption;
        }
        public override void SetDescription(string description)
        {
            base.SetDescription(description);
            this.progressPanel1.Description = description;
        }
        public override void ProcessCommand(Enum cmd, object arg)
        {
            base.ProcessCommand(cmd, arg);
        }
        public  void setCloseTime(int intTime1)
        {
            intCloseTime = intTime1;
        }


        #endregion

        public enum WaitFormCommand
        {
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.progressPanel1 = new DevExpress.XtraWaitForm.ProgressPanel();
            this.tirTimer = new System.Windows.Forms.Timer(this.components);
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel1.BackColor = System.Drawing.Color.Transparent;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.progressPanel1, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.Padding = new System.Windows.Forms.Padding(0, 13, 0, 13);
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(419, 77);
            this.tableLayoutPanel1.TabIndex = 2;
            this.tableLayoutPanel1.Paint += new System.Windows.Forms.PaintEventHandler(this.tableLayoutPanel1_Paint);
            // 
            // progressPanel1
            // 
            this.progressPanel1.Appearance.BackColor = System.Drawing.Color.Transparent;
            this.progressPanel1.Appearance.Options.UseBackColor = true;
            this.progressPanel1.AppearanceCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.progressPanel1.AppearanceCaption.ForeColor = System.Drawing.Color.DimGray;
            this.progressPanel1.AppearanceCaption.Options.UseFont = true;
            this.progressPanel1.AppearanceCaption.Options.UseForeColor = true;
            this.progressPanel1.AppearanceDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.progressPanel1.AppearanceDescription.ForeColor = System.Drawing.Color.DimGray;
            this.progressPanel1.AppearanceDescription.Options.UseFont = true;
            this.progressPanel1.AppearanceDescription.Options.UseForeColor = true;
            this.progressPanel1.BarAnimationElementThickness = 2;
            this.progressPanel1.Caption = "123";
            this.progressPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressPanel1.ImageHorzOffset = 20;
            this.progressPanel1.Location = new System.Drawing.Point(0, 16);
            this.progressPanel1.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3);
            this.progressPanel1.Name = "progressPanel1";
            this.progressPanel1.Size = new System.Drawing.Size(419, 45);
            this.progressPanel1.TabIndex = 0;
            this.progressPanel1.Text = "1";
            // 
            // tirTimer
            // 
            this.tirTimer.Interval = 500;
            this.tirTimer.Tick += new System.EventHandler(this.tirTimer_Tick);
            // 
            // MyWaitFrom
            // 
            this.ClientSize = new System.Drawing.Size(419, 77);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "MyWaitFrom";
            this.Load += new System.EventHandler(this.MyWaitFrom_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void tableLayoutPanel1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
        {

        }

        private void MyWaitFrom_Load(object sender, EventArgs e)
        {
            this.TopLevel = true;

            intTime = 0;
            if (intCloseTime > 0) {
                tirTimer.Enabled = true;
            }
        }

        private void tirTimer_Tick(object sender, EventArgs e)
        {
            if (intCloseTime <= intTime) {
                this.Close();            }
            intTime = intTime + 1;
        }
    }
}
