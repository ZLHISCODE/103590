namespace ZLSOFT.HIS.PreTriage
{
    partial class frmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.staBar = new DevExpress.XtraBars.Ribbon.RibbonStatusBar();
            this.lblUser = new DevExpress.XtraBars.BarStaticItem();
            this.lblUserName = new DevExpress.XtraBars.BarStaticItem();
            this.rbControl = new DevExpress.XtraBars.Ribbon.RibbonControl();
            this.btnSend = new DevExpress.XtraBars.BarButtonItem();
            this.btnAdd = new DevExpress.XtraBars.BarButtonItem();
            this.btnView = new DevExpress.XtraBars.BarButtonItem();
            this.btnHelp = new DevExpress.XtraBars.BarButtonItem();
            this.btnExit = new DevExpress.XtraBars.BarButtonItem();
            this.btnSet = new DevExpress.XtraBars.BarButtonItem();
            this.btnReadCard = new DevExpress.XtraBars.BarButtonItem();
            this.btnReadIDCard = new DevExpress.XtraBars.BarButtonItem();
            this.btn报表查看 = new DevExpress.XtraBars.BarButtonItem();
            this.popupReport = new DevExpress.XtraBars.PopupMenu(this.components);
            this.btn体征采集 = new DevExpress.XtraBars.BarButtonItem();
            this.btnReportSet = new DevExpress.XtraBars.BarButtonItem();
            this.ribbon开始 = new DevExpress.XtraBars.Ribbon.RibbonPage();
            this.PageGroup接口读取 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
            this.PageGroup分诊操作 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
            this.PageGroup历史分诊 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
            this.PageGroup报表 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
            this.PageGroup系统操作 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
            this.repositoryItemComboBox1 = new DevExpress.XtraEditors.Repository.RepositoryItemComboBox();
            this.palEdit = new System.Windows.Forms.Panel();
            this.tablePanel = new System.Windows.Forms.TableLayoutPanel();
            this.palPati = new System.Windows.Forms.Panel();
            this.grpEdit = new DevExpress.XtraEditors.GroupControl();
            this.imgCol16 = new DevExpress.Utils.ImageCollection(this.components);
            this.btnTimeOut = new DevExpress.XtraEditors.SimpleButton();
            this.timeIcon = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.rbControl)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.popupReport)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemComboBox1)).BeginInit();
            this.palEdit.SuspendLayout();
            this.tablePanel.SuspendLayout();
            this.palPati.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grpEdit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.imgCol16)).BeginInit();
            this.SuspendLayout();
            // 
            // staBar
            // 
            this.staBar.ItemLinks.Add(this.lblUser);
            this.staBar.ItemLinks.Add(this.lblUserName);
            this.staBar.Location = new System.Drawing.Point(0, 622);
            this.staBar.Name = "staBar";
            this.staBar.Ribbon = this.rbControl;
            this.staBar.Size = new System.Drawing.Size(1029, 23);
            // 
            // lblUser
            // 
            this.lblUser.Caption = "当前用户:";
            this.lblUser.Id = 23;
            this.lblUser.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("lblUser.ImageOptions.Image")));
            this.lblUser.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("lblUser.ImageOptions.LargeImage")));
            this.lblUser.Name = "lblUser";
            // 
            // lblUserName
            // 
            this.lblUserName.Caption = "管理员";
            this.lblUserName.Id = 24;
            this.lblUserName.Name = "lblUserName";
            // 
            // rbControl
            // 
            this.rbControl.AutoSizeItems = true;
            this.rbControl.ExpandCollapseItem.Id = 0;
            this.rbControl.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.rbControl.ExpandCollapseItem,
            this.btnSend,
            this.btnAdd,
            this.btnView,
            this.btnHelp,
            this.btnExit,
            this.btnSet,
            this.lblUser,
            this.lblUserName,
            this.btnReadCard,
            this.btnReadIDCard,
            this.btn报表查看,
            this.btn体征采集,
            this.btnReportSet});
            this.rbControl.Location = new System.Drawing.Point(0, 0);
            this.rbControl.MaxItemId = 36;
            this.rbControl.Name = "rbControl";
            this.rbControl.Pages.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPage[] {
            this.ribbon开始});
            this.rbControl.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemComboBox1});
            this.rbControl.ShowApplicationButton = DevExpress.Utils.DefaultBoolean.False;
            this.rbControl.Size = new System.Drawing.Size(1029, 119);
            this.rbControl.StatusBar = this.staBar;
            this.rbControl.ToolbarLocation = DevExpress.XtraBars.Ribbon.RibbonQuickAccessToolbarLocation.Hidden;
            this.rbControl.TransparentEditorsMode = DevExpress.Utils.DefaultBoolean.False;
            // 
            // btnSend
            // 
            this.btnSend.Caption = "发送";
            this.btnSend.Id = 16;
            this.btnSend.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnSend.ImageOptions.Image")));
            this.btnSend.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnSend.ImageOptions.LargeImage")));
            this.btnSend.Name = "btnSend";
            this.btnSend.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btnSend_ItemClick);
            // 
            // btnAdd
            // 
            this.btnAdd.Caption = "新增";
            this.btnAdd.Id = 17;
            this.btnAdd.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.ImageOptions.Image")));
            this.btnAdd.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnAdd.ImageOptions.LargeImage")));
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btnAdd_ItemClick);
            // 
            // btnView
            // 
            this.btnView.Caption = "分诊列表";
            this.btnView.Id = 18;
            this.btnView.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnView.ImageOptions.Image")));
            this.btnView.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnView.ImageOptions.LargeImage")));
            this.btnView.Name = "btnView";
            this.btnView.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btnView_ItemClick);
            // 
            // btnHelp
            // 
            this.btnHelp.Caption = "关于";
            this.btnHelp.Id = 19;
            this.btnHelp.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.ImageOptions.Image")));
            this.btnHelp.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnHelp.ImageOptions.LargeImage")));
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btnHelp_ItemClick);
            // 
            // btnExit
            // 
            this.btnExit.Caption = "退出";
            this.btnExit.Id = 20;
            this.btnExit.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.ImageOptions.Image")));
            this.btnExit.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnExit.ImageOptions.LargeImage")));
            this.btnExit.Name = "btnExit";
            this.btnExit.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btnExit_ItemClick);
            // 
            // btnSet
            // 
            this.btnSet.Caption = "主诉维护";
            this.btnSet.Id = 22;
            this.btnSet.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnSet.ImageOptions.Image")));
            this.btnSet.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnSet.ImageOptions.LargeImage")));
            this.btnSet.Name = "btnSet";
            // 
            // btnReadCard
            // 
            this.btnReadCard.Caption = "读取医保卡";
            this.btnReadCard.Id = 27;
            this.btnReadCard.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnReadCard.ImageOptions.Image")));
            this.btnReadCard.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnReadCard.ImageOptions.LargeImage")));
            this.btnReadCard.Name = "btnReadCard";
            this.btnReadCard.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btnReadCard_ItemClick);
            // 
            // btnReadIDCard
            // 
            this.btnReadIDCard.Caption = "读取身份证";
            this.btnReadIDCard.Id = 28;
            this.btnReadIDCard.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnReadIDCard.ImageOptions.Image")));
            this.btnReadIDCard.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnReadIDCard.ImageOptions.LargeImage")));
            this.btnReadIDCard.Name = "btnReadIDCard";
            this.btnReadIDCard.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btnReadIDCard_ItemClick);
            // 
            // btn报表查看
            // 
            this.btn报表查看.ActAsDropDown = true;
            this.btn报表查看.ButtonStyle = DevExpress.XtraBars.BarButtonStyle.DropDown;
            this.btn报表查看.Caption = "报表查看";
            this.btn报表查看.DropDownControl = this.popupReport;
            this.btn报表查看.Id = 30;
            this.btn报表查看.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn报表查看.ImageOptions.Image")));
            this.btn报表查看.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn报表查看.ImageOptions.LargeImage")));
            this.btn报表查看.Name = "btn报表查看";
            // 
            // popupReport
            // 
            this.popupReport.Name = "popupReport";
            this.popupReport.Ribbon = this.rbControl;
            // 
            // btn体征采集
            // 
            this.btn体征采集.Caption = "采集生命体征";
            this.btn体征采集.Id = 33;
            this.btn体征采集.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn体征采集.ImageOptions.LargeImage")));
            this.btn体征采集.Name = "btn体征采集";
            this.btn体征采集.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn体征采集_ItemClick);
            // 
            // btnReportSet
            // 
            this.btnReportSet.Caption = "打印设置";
            this.btnReportSet.Id = 35;
            this.btnReportSet.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnReportSet.ImageOptions.Image")));
            this.btnReportSet.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btnReportSet.ImageOptions.LargeImage")));
            this.btnReportSet.Name = "btnReportSet";
            this.btnReportSet.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btnReportSet_ItemClick);
            // 
            // ribbon开始
            // 
            this.ribbon开始.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.PageGroup接口读取,
            this.PageGroup分诊操作,
            this.PageGroup历史分诊,
            this.PageGroup报表,
            this.PageGroup系统操作});
            this.ribbon开始.Name = "ribbon开始";
            this.ribbon开始.Text = "开始";
            // 
            // PageGroup接口读取
            // 
            this.PageGroup接口读取.ItemLinks.Add(this.btnReadCard);
            this.PageGroup接口读取.ItemLinks.Add(this.btnReadIDCard);
            this.PageGroup接口读取.ItemLinks.Add(this.btn体征采集);
            this.PageGroup接口读取.Name = "PageGroup接口读取";
            this.PageGroup接口读取.Text = "信息读取";
            // 
            // PageGroup分诊操作
            // 
            this.PageGroup分诊操作.ItemLinks.Add(this.btnAdd);
            this.PageGroup分诊操作.ItemLinks.Add(this.btnSend);
            this.PageGroup分诊操作.Name = "PageGroup分诊操作";
            this.PageGroup分诊操作.Text = "分诊操作";
            // 
            // PageGroup历史分诊
            // 
            this.PageGroup历史分诊.ItemLinks.Add(this.btnView);
            this.PageGroup历史分诊.Name = "PageGroup历史分诊";
            this.PageGroup历史分诊.State = DevExpress.XtraBars.Ribbon.RibbonPageGroupState.Expanded;
            this.PageGroup历史分诊.Text = "历史";
            // 
            // PageGroup报表
            // 
            this.PageGroup报表.ItemLinks.Add(this.btn报表查看);
            this.PageGroup报表.ItemLinks.Add(this.btnReportSet);
            this.PageGroup报表.Name = "PageGroup报表";
            this.PageGroup报表.Text = "报表";
            // 
            // PageGroup系统操作
            // 
            this.PageGroup系统操作.ItemLinks.Add(this.btnHelp);
            this.PageGroup系统操作.ItemLinks.Add(this.btnExit);
            this.PageGroup系统操作.Name = "PageGroup系统操作";
            this.PageGroup系统操作.Text = "系统操作";
            // 
            // repositoryItemComboBox1
            // 
            this.repositoryItemComboBox1.AutoHeight = false;
            this.repositoryItemComboBox1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemComboBox1.Name = "repositoryItemComboBox1";
            // 
            // palEdit
            // 
            this.palEdit.Controls.Add(this.tablePanel);
            this.palEdit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.palEdit.Location = new System.Drawing.Point(0, 119);
            this.palEdit.Name = "palEdit";
            this.palEdit.Size = new System.Drawing.Size(1029, 503);
            this.palEdit.TabIndex = 2;
            // 
            // tablePanel
            // 
            this.tablePanel.ColumnCount = 1;
            this.tablePanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tablePanel.Controls.Add(this.palPati, 0, 0);
            this.tablePanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tablePanel.Location = new System.Drawing.Point(0, 0);
            this.tablePanel.Margin = new System.Windows.Forms.Padding(1);
            this.tablePanel.Name = "tablePanel";
            this.tablePanel.RowCount = 1;
            this.tablePanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tablePanel.Size = new System.Drawing.Size(1029, 503);
            this.tablePanel.TabIndex = 2;
            // 
            // palPati
            // 
            this.palPati.Controls.Add(this.grpEdit);
            this.palPati.Dock = System.Windows.Forms.DockStyle.Fill;
            this.palPati.Location = new System.Drawing.Point(1, 1);
            this.palPati.Margin = new System.Windows.Forms.Padding(1);
            this.palPati.Name = "palPati";
            this.palPati.Size = new System.Drawing.Size(1027, 501);
            this.palPati.TabIndex = 1;
            // 
            // grpEdit
            // 
            this.grpEdit.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.grpEdit.CaptionImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("grpEdit.CaptionImageOptions.Image")));
            this.grpEdit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpEdit.GroupStyle = DevExpress.Utils.GroupStyle.Light;
            this.grpEdit.Location = new System.Drawing.Point(0, 0);
            this.grpEdit.Name = "grpEdit";
            this.grpEdit.Size = new System.Drawing.Size(1027, 501);
            this.grpEdit.TabIndex = 0;
            this.grpEdit.Text = "分诊信息";
            // 
            // imgCol16
            // 
            this.imgCol16.ImageStream = ((DevExpress.Utils.ImageCollectionStreamer)(resources.GetObject("imgCol16.ImageStream")));
            this.imgCol16.Images.SetKeyName(0, "contact_16x16.png");
            this.imgCol16.Images.SetKeyName(1, "card_16x16.png");
            this.imgCol16.Images.SetKeyName(2, "contact_16x16.png");
            this.imgCol16.Images.SetKeyName(3, "boreport2_16x16.png");
            // 
            // btnTimeOut
            // 
            this.btnTimeOut.Appearance.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btnTimeOut.Appearance.BackColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.btnTimeOut.Appearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnTimeOut.Appearance.Font = new System.Drawing.Font("Tahoma", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTimeOut.Appearance.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnTimeOut.Appearance.Options.UseBackColor = true;
            this.btnTimeOut.Appearance.Options.UseBorderColor = true;
            this.btnTimeOut.Appearance.Options.UseFont = true;
            this.btnTimeOut.Appearance.Options.UseForeColor = true;
            this.btnTimeOut.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnTimeOut.ImageOptions.Image")));
            this.btnTimeOut.Location = new System.Drawing.Point(805, 6);
            this.btnTimeOut.Name = "btnTimeOut";
            this.btnTimeOut.Size = new System.Drawing.Size(217, 49);
            this.btnTimeOut.TabIndex = 5;
            this.btnTimeOut.Text = "当前有9999个病人候诊超时";
            this.btnTimeOut.Click += new System.EventHandler(this.btnTimeOut_Click);
            // 
            // timeIcon
            // 
            this.timeIcon.Interval = 60000;
            this.timeIcon.Tick += new System.EventHandler(this.timeIcon_Tick);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1029, 645);
            this.Controls.Add(this.btnTimeOut);
            this.Controls.Add(this.palEdit);
            this.Controls.Add(this.staBar);
            this.Controls.Add(this.rbControl);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.LookAndFeel.SkinName = "Office 2007 Pink";
            this.LookAndFeel.UseDefaultLookAndFeel = false;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "急诊预检分诊工作站";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmMain_FormClosed);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.Resize += new System.EventHandler(this.frmMain_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.rbControl)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.popupReport)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemComboBox1)).EndInit();
            this.palEdit.ResumeLayout(false);
            this.tablePanel.ResumeLayout(false);
            this.palPati.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grpEdit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.imgCol16)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private DevExpress.XtraBars.Ribbon.RibbonStatusBar staBar;
        private DevExpress.XtraBars.Ribbon.RibbonControl rbControl;
        private DevExpress.XtraBars.Ribbon.RibbonPage ribbon开始;
        private DevExpress.XtraBars.Ribbon.RibbonPageGroup PageGroup分诊操作;
        private DevExpress.XtraBars.BarButtonItem btnSend;
        private DevExpress.XtraBars.BarButtonItem btnAdd;
        private DevExpress.XtraBars.BarButtonItem btnView;
        private DevExpress.XtraBars.BarButtonItem btnHelp;
        private DevExpress.XtraBars.BarButtonItem btnExit;
        private DevExpress.XtraBars.Ribbon.RibbonPageGroup PageGroup系统操作;
        private DevExpress.XtraBars.BarButtonItem btnSet;
        private DevExpress.XtraBars.Ribbon.RibbonPageGroup PageGroup历史分诊;
        private DevExpress.XtraBars.BarStaticItem lblUser;
        private DevExpress.XtraBars.BarStaticItem lblUserName;
        private System.Windows.Forms.Panel palEdit;
        private DevExpress.Utils.ImageCollection imgCol16;
        private System.Windows.Forms.TableLayoutPanel tablePanel;
        private System.Windows.Forms.Panel palPati;
        private DevExpress.XtraEditors.GroupControl grpEdit;
        private DevExpress.XtraBars.BarButtonItem btnReadCard;
        private DevExpress.XtraBars.BarButtonItem btnReadIDCard;
        private DevExpress.XtraEditors.Repository.RepositoryItemComboBox repositoryItemComboBox1;
        private DevExpress.XtraBars.Ribbon.RibbonPageGroup PageGroup报表;
        private DevExpress.XtraBars.BarButtonItem btn报表查看;
        private DevExpress.XtraBars.PopupMenu popupReport;
        private DevExpress.XtraBars.BarButtonItem btn体征采集;
        private DevExpress.XtraBars.Ribbon.RibbonPageGroup PageGroup接口读取;
        private DevExpress.XtraBars.BarButtonItem btnReportSet;
        private DevExpress.XtraEditors.SimpleButton btnTimeOut;
        private System.Windows.Forms.Timer timeIcon;
    }
}

