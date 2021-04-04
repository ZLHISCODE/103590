namespace zlShortMsg
{
    partial class frmMain
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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.menuTop = new System.Windows.Forms.MenuStrip();
            this.服务ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.开启短信服务ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.退出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.最小化ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.退出ToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.pnlNavi = new System.Windows.Forms.Panel();
            this.pnlSetting = new System.Windows.Forms.Panel();
            this.lblSetting = new System.Windows.Forms.Label();
            this.picSetting = new System.Windows.Forms.PictureBox();
            this.pnlFailure = new System.Windows.Forms.Panel();
            this.lblFailure = new System.Windows.Forms.Label();
            this.picFailure = new System.Windows.Forms.PictureBox();
            this.pnlLog = new System.Windows.Forms.Panel();
            this.lblLog = new System.Windows.Forms.Label();
            this.picLog = new System.Windows.Forms.PictureBox();
            this.pnlTip = new System.Windows.Forms.Panel();
            this.lblTip = new System.Windows.Forms.Label();
            this.staBar = new System.Windows.Forms.StatusStrip();
            this.staState = new System.Windows.Forms.ToolStripStatusLabel();
            this.staSpilit = new System.Windows.Forms.ToolStripStatusLabel();
            this.staTime = new System.Windows.Forms.ToolStripStatusLabel();
            this.Timer = new System.Windows.Forms.Timer(this.components);
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.pnlBoder = new System.Windows.Forms.Panel();
            this.pnlContainer = new System.Windows.Forms.Panel();
            this.menuTop.SuspendLayout();
            this.pnlNavi.SuspendLayout();
            this.pnlSetting.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picSetting)).BeginInit();
            this.pnlFailure.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picFailure)).BeginInit();
            this.pnlLog.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picLog)).BeginInit();
            this.pnlTip.SuspendLayout();
            this.staBar.SuspendLayout();
            this.pnlBoder.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuTop
            // 
            this.menuTop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(83)))), ((int)(((byte)(83)))), ((int)(((byte)(83)))));
            this.menuTop.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.服务ToolStripMenuItem,
            this.退出ToolStripMenuItem});
            this.menuTop.Location = new System.Drawing.Point(0, 0);
            this.menuTop.Name = "menuTop";
            this.menuTop.Size = new System.Drawing.Size(861, 24);
            this.menuTop.TabIndex = 0;
            // 
            // 服务ToolStripMenuItem
            // 
            this.服务ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.开启短信服务ToolStripMenuItem});
            this.服务ToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.服务ToolStripMenuItem.Name = "服务ToolStripMenuItem";
            this.服务ToolStripMenuItem.Size = new System.Drawing.Size(42, 20);
            this.服务ToolStripMenuItem.Text = "服务";
            // 
            // 开启短信服务ToolStripMenuItem
            // 
            this.开启短信服务ToolStripMenuItem.Name = "开启短信服务ToolStripMenuItem";
            this.开启短信服务ToolStripMenuItem.Size = new System.Drawing.Size(142, 22);
            this.开启短信服务ToolStripMenuItem.Text = "关闭短信服务";
            this.开启短信服务ToolStripMenuItem.Click += new System.EventHandler(this.开启短信服务ToolStripMenuItem_Click);
            // 
            // 退出ToolStripMenuItem
            // 
            this.退出ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.最小化ToolStripMenuItem,
            this.退出ToolStripMenuItem1});
            this.退出ToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            this.退出ToolStripMenuItem.Size = new System.Drawing.Size(42, 20);
            this.退出ToolStripMenuItem.Text = "退出";
            // 
            // 最小化ToolStripMenuItem
            // 
            this.最小化ToolStripMenuItem.Name = "最小化ToolStripMenuItem";
            this.最小化ToolStripMenuItem.Size = new System.Drawing.Size(109, 22);
            this.最小化ToolStripMenuItem.Text = "最小化";
            this.最小化ToolStripMenuItem.Click += new System.EventHandler(this.最小化ToolStripMenuItem_Click);
            // 
            // 退出ToolStripMenuItem1
            // 
            this.退出ToolStripMenuItem1.Name = "退出ToolStripMenuItem1";
            this.退出ToolStripMenuItem1.Size = new System.Drawing.Size(109, 22);
            this.退出ToolStripMenuItem1.Text = "退出";
            this.退出ToolStripMenuItem1.Click += new System.EventHandler(this.退出ToolStripMenuItem1_Click);
            // 
            // pnlNavi
            // 
            this.pnlNavi.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.pnlNavi.Controls.Add(this.pnlSetting);
            this.pnlNavi.Controls.Add(this.pnlFailure);
            this.pnlNavi.Controls.Add(this.pnlLog);
            this.pnlNavi.Controls.Add(this.pnlTip);
            this.pnlNavi.Location = new System.Drawing.Point(9, 27);
            this.pnlNavi.Name = "pnlNavi";
            this.pnlNavi.Size = new System.Drawing.Size(197, 470);
            this.pnlNavi.TabIndex = 1;
            this.pnlNavi.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlNavi_Paint);
            // 
            // pnlSetting
            // 
            this.pnlSetting.BackColor = System.Drawing.Color.Transparent;
            this.pnlSetting.Controls.Add(this.lblSetting);
            this.pnlSetting.Controls.Add(this.picSetting);
            this.pnlSetting.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pnlSetting.Location = new System.Drawing.Point(2, 135);
            this.pnlSetting.Name = "pnlSetting";
            this.pnlSetting.Size = new System.Drawing.Size(193, 51);
            this.pnlSetting.TabIndex = 3;
            this.pnlSetting.Click += new System.EventHandler(this.pnlSetting_Click);
            // 
            // lblSetting
            // 
            this.lblSetting.AutoSize = true;
            this.lblSetting.BackColor = System.Drawing.Color.Transparent;
            this.lblSetting.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblSetting.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(56)))), ((int)(((byte)(56)))));
            this.lblSetting.Location = new System.Drawing.Point(67, 20);
            this.lblSetting.Name = "lblSetting";
            this.lblSetting.Size = new System.Drawing.Size(91, 14);
            this.lblSetting.TabIndex = 5;
            this.lblSetting.Text = "短信接口配置";
            this.lblSetting.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // picSetting
            // 
            this.picSetting.Image = global::zlShortMsg.Properties.Resources.设置;
            this.picSetting.Location = new System.Drawing.Point(16, 12);
            this.picSetting.Name = "picSetting";
            this.picSetting.Size = new System.Drawing.Size(39, 38);
            this.picSetting.TabIndex = 4;
            this.picSetting.TabStop = false;
            // 
            // pnlFailure
            // 
            this.pnlFailure.BackColor = System.Drawing.Color.Transparent;
            this.pnlFailure.Controls.Add(this.lblFailure);
            this.pnlFailure.Controls.Add(this.picFailure);
            this.pnlFailure.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pnlFailure.Location = new System.Drawing.Point(2, 82);
            this.pnlFailure.Name = "pnlFailure";
            this.pnlFailure.Size = new System.Drawing.Size(193, 51);
            this.pnlFailure.TabIndex = 2;
            this.pnlFailure.Click += new System.EventHandler(this.pnlFailure_Click);
            // 
            // lblFailure
            // 
            this.lblFailure.AutoSize = true;
            this.lblFailure.BackColor = System.Drawing.Color.Transparent;
            this.lblFailure.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblFailure.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(56)))), ((int)(((byte)(56)))));
            this.lblFailure.Location = new System.Drawing.Point(68, 19);
            this.lblFailure.Name = "lblFailure";
            this.lblFailure.Size = new System.Drawing.Size(91, 14);
            this.lblFailure.TabIndex = 3;
            this.lblFailure.Text = "发送失败清单";
            this.lblFailure.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // picFailure
            // 
            this.picFailure.Image = global::zlShortMsg.Properties.Resources.错误;
            this.picFailure.Location = new System.Drawing.Point(16, 10);
            this.picFailure.Name = "picFailure";
            this.picFailure.Size = new System.Drawing.Size(39, 38);
            this.picFailure.TabIndex = 2;
            this.picFailure.TabStop = false;
            // 
            // pnlLog
            // 
            this.pnlLog.BackColor = System.Drawing.Color.Transparent;
            this.pnlLog.Controls.Add(this.lblLog);
            this.pnlLog.Controls.Add(this.picLog);
            this.pnlLog.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pnlLog.Location = new System.Drawing.Point(2, 30);
            this.pnlLog.Name = "pnlLog";
            this.pnlLog.Size = new System.Drawing.Size(193, 51);
            this.pnlLog.TabIndex = 1;
            this.pnlLog.Click += new System.EventHandler(this.pnlLog_Click);
            // 
            // lblLog
            // 
            this.lblLog.AutoSize = true;
            this.lblLog.BackColor = System.Drawing.Color.Transparent;
            this.lblLog.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblLog.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(56)))), ((int)(((byte)(56)))));
            this.lblLog.Location = new System.Drawing.Point(68, 21);
            this.lblLog.Name = "lblLog";
            this.lblLog.Size = new System.Drawing.Size(91, 14);
            this.lblLog.TabIndex = 1;
            this.lblLog.Text = "短信发送日志";
            this.lblLog.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // picLog
            // 
            this.picLog.Image = global::zlShortMsg.Properties.Resources.消息;
            this.picLog.Location = new System.Drawing.Point(16, 12);
            this.picLog.Name = "picLog";
            this.picLog.Size = new System.Drawing.Size(39, 38);
            this.picLog.TabIndex = 0;
            this.picLog.TabStop = false;
            // 
            // pnlTip
            // 
            this.pnlTip.Controls.Add(this.lblTip);
            this.pnlTip.Location = new System.Drawing.Point(0, 0);
            this.pnlTip.Margin = new System.Windows.Forms.Padding(0);
            this.pnlTip.Name = "pnlTip";
            this.pnlTip.Size = new System.Drawing.Size(197, 28);
            this.pnlTip.TabIndex = 0;
            this.pnlTip.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlTip_Paint);
            // 
            // lblTip
            // 
            this.lblTip.AutoSize = true;
            this.lblTip.BackColor = System.Drawing.Color.Transparent;
            this.lblTip.Font = new System.Drawing.Font("宋体", 10F);
            this.lblTip.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(56)))), ((int)(((byte)(56)))));
            this.lblTip.Location = new System.Drawing.Point(14, 8);
            this.lblTip.Name = "lblTip";
            this.lblTip.Size = new System.Drawing.Size(49, 14);
            this.lblTip.TabIndex = 0;
            this.lblTip.Text = "导  航";
            this.lblTip.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // staBar
            // 
            this.staBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.staState,
            this.staSpilit,
            this.staTime});
            this.staBar.Location = new System.Drawing.Point(0, 513);
            this.staBar.Name = "staBar";
            this.staBar.Size = new System.Drawing.Size(861, 22);
            this.staBar.TabIndex = 3;
            this.staBar.Text = "statusStrip1";
            // 
            // staState
            // 
            this.staState.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.staState.Name = "staState";
            this.staState.Size = new System.Drawing.Size(99, 17);
            this.staState.Text = "短信服务状态:开启";
            // 
            // staSpilit
            // 
            this.staSpilit.Name = "staSpilit";
            this.staSpilit.Size = new System.Drawing.Size(747, 17);
            this.staSpilit.Spring = true;
            // 
            // staTime
            // 
            this.staTime.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.staTime.Name = "staTime";
            this.staTime.Size = new System.Drawing.Size(0, 17);
            // 
            // Timer
            // 
            this.Timer.Enabled = true;
            this.Timer.Interval = 1000;
            this.Timer.Tick += new System.EventHandler(this.Timer_Tick);
            // 
            // notifyIcon
            // 
            this.notifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon.Icon")));
            this.notifyIcon.Text = "中联短信发送服务";
            this.notifyIcon.Visible = true;
            this.notifyIcon.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon_MouseDoubleClick);
            // 
            // pnlBoder
            // 
            this.pnlBoder.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlBoder.Controls.Add(this.pnlContainer);
            this.pnlBoder.Location = new System.Drawing.Point(216, 27);
            this.pnlBoder.Name = "pnlBoder";
            this.pnlBoder.Size = new System.Drawing.Size(635, 470);
            this.pnlBoder.TabIndex = 4;
            this.pnlBoder.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlBoder_Paint);
            // 
            // pnlContainer
            // 
            this.pnlContainer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlContainer.Location = new System.Drawing.Point(3, 3);
            this.pnlContainer.Name = "pnlContainer";
            this.pnlContainer.Size = new System.Drawing.Size(629, 464);
            this.pnlContainer.TabIndex = 3;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(861, 535);
            this.Controls.Add(this.pnlBoder);
            this.Controls.Add(this.staBar);
            this.Controls.Add(this.pnlNavi);
            this.Controls.Add(this.menuTop);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuTop;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "中联短信发送服务";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.Resize += new System.EventHandler(this.frmMain_Resize);
            this.menuTop.ResumeLayout(false);
            this.menuTop.PerformLayout();
            this.pnlNavi.ResumeLayout(false);
            this.pnlSetting.ResumeLayout(false);
            this.pnlSetting.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picSetting)).EndInit();
            this.pnlFailure.ResumeLayout(false);
            this.pnlFailure.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picFailure)).EndInit();
            this.pnlLog.ResumeLayout(false);
            this.pnlLog.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picLog)).EndInit();
            this.pnlTip.ResumeLayout(false);
            this.pnlTip.PerformLayout();
            this.staBar.ResumeLayout(false);
            this.staBar.PerformLayout();
            this.pnlBoder.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuTop;
        private System.Windows.Forms.ToolStripMenuItem 服务ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 退出ToolStripMenuItem;
        private System.Windows.Forms.Panel pnlNavi;
        private System.Windows.Forms.Panel pnlTip;
        private System.Windows.Forms.Label lblTip;
        private System.Windows.Forms.Panel pnlSetting;
        private System.Windows.Forms.Panel pnlFailure;
        private System.Windows.Forms.Panel pnlLog;
        private System.Windows.Forms.Label lblLog;
        private System.Windows.Forms.PictureBox picLog;
        private System.Windows.Forms.Label lblSetting;
        private System.Windows.Forms.PictureBox picSetting;
        private System.Windows.Forms.Label lblFailure;
        private System.Windows.Forms.PictureBox picFailure;
        private System.Windows.Forms.StatusStrip staBar;
        private System.Windows.Forms.ToolStripStatusLabel staState;
        private System.Windows.Forms.ToolStripStatusLabel staTime;
        private System.Windows.Forms.ToolStripStatusLabel staSpilit;
        private System.Windows.Forms.Timer Timer;
        private System.Windows.Forms.ToolStripMenuItem 开启短信服务ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 最小化ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 退出ToolStripMenuItem1;
        private System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.Panel pnlBoder;
        private System.Windows.Forms.Panel pnlContainer;
    }
}

