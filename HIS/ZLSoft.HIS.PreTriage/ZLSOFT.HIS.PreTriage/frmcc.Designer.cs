namespace ZLSOFT.HIS.PreTriage
{
    partial class frmcc
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmcc));
            this.txtSearch = new DevExpress.XtraEditors.TextEdit();
            this.btnSearch = new DevExpress.XtraEditors.SimpleButton();
            this.btnOK = new DevExpress.XtraEditors.SimpleButton();
            this.btnCancel = new DevExpress.XtraEditors.SimpleButton();
            this.gctMain = new DevExpress.XtraGrid.GridControl();
            this.gvcc = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colChoose = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemCheckEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
            this.colType = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colCode = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colSCode = new DevExpress.XtraGrid.Columns.GridColumn();
            ((System.ComponentModel.ISupportInitialize)(this.txtSearch.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gctMain)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvcc)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtSearch
            // 
            this.txtSearch.Location = new System.Drawing.Point(14, 14);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(374, 20);
            this.txtSearch.TabIndex = 0;
            this.txtSearch.EditValueChanged += new System.EventHandler(this.txtSearch_EditValueChanged);
            this.txtSearch.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSearch_KeyPress);
            // 
            // btnSearch
            // 
            this.btnSearch.Location = new System.Drawing.Point(395, 12);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(79, 26);
            this.btnSearch.TabIndex = 1;
            this.btnSearch.Text = "查找";
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(486, 425);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(80, 35);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "确定(&S)";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(603, 425);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(80, 35);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "取消(&C)";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // gctMain
            // 
            this.gctMain.Location = new System.Drawing.Point(14, 44);
            this.gctMain.MainView = this.gvcc;
            this.gctMain.Name = "gctMain";
            this.gctMain.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemCheckEdit1});
            this.gctMain.Size = new System.Drawing.Size(670, 373);
            this.gctMain.TabIndex = 2;
            this.gctMain.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvcc});
            this.gctMain.DoubleClick += new System.EventHandler(this.gctMain_DoubleClick);
            this.gctMain.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.gctMain_KeyPress);
            // 
            // gvcc
            // 
            this.gvcc.Appearance.FocusedRow.BackColor = System.Drawing.Color.LightSteelBlue;
            this.gvcc.Appearance.FocusedRow.Options.UseBackColor = true;
            this.gvcc.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colChoose,
            this.colType,
            this.colCode,
            this.colName,
            this.colSCode});
            this.gvcc.DetailHeight = 408;
            this.gvcc.GridControl = this.gctMain;
            this.gvcc.GroupCount = 1;
            this.gvcc.Name = "gvcc";
            this.gvcc.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.colType, DevExpress.Data.ColumnSortOrder.Ascending)});
            // 
            // colChoose
            // 
            this.colChoose.AppearanceHeader.FontStyleDelta = System.Drawing.FontStyle.Bold;
            this.colChoose.AppearanceHeader.Options.UseTextOptions = true;
            this.colChoose.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colChoose.ColumnEdit = this.repositoryItemCheckEdit1;
            this.colChoose.FieldName = "选择";
            this.colChoose.MinWidth = 23;
            this.colChoose.Name = "colChoose";
            this.colChoose.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;
            this.colChoose.OptionsColumn.AllowMove = false;
            this.colChoose.OptionsColumn.AllowSize = false;
            this.colChoose.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.False;
            this.colChoose.Visible = true;
            this.colChoose.VisibleIndex = 0;
            this.colChoose.Width = 70;
            // 
            // repositoryItemCheckEdit1
            // 
            this.repositoryItemCheckEdit1.AutoHeight = false;
            this.repositoryItemCheckEdit1.Name = "repositoryItemCheckEdit1";
            this.repositoryItemCheckEdit1.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked;
            // 
            // colType
            // 
            this.colType.AppearanceHeader.FontStyleDelta = System.Drawing.FontStyle.Bold;
            this.colType.AppearanceHeader.Options.UseFont = true;
            this.colType.AppearanceHeader.Options.UseTextOptions = true;
            this.colType.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colType.Caption = "分类";
            this.colType.FieldName = "分类";
            this.colType.MinWidth = 23;
            this.colType.Name = "colType";
            this.colType.OptionsColumn.AllowEdit = false;
            this.colType.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.True;
            this.colType.OptionsColumn.AllowSize = false;
            this.colType.Visible = true;
            this.colType.VisibleIndex = 1;
            this.colType.Width = 70;
            // 
            // colCode
            // 
            this.colCode.AppearanceHeader.FontStyleDelta = System.Drawing.FontStyle.Bold;
            this.colCode.AppearanceHeader.Options.UseFont = true;
            this.colCode.AppearanceHeader.Options.UseTextOptions = true;
            this.colCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colCode.Caption = "编码";
            this.colCode.FieldName = "编码";
            this.colCode.MinWidth = 23;
            this.colCode.Name = "colCode";
            this.colCode.OptionsColumn.AllowEdit = false;
            this.colCode.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;
            this.colCode.OptionsColumn.AllowSize = false;
            this.colCode.Visible = true;
            this.colCode.VisibleIndex = 1;
            this.colCode.Width = 70;
            // 
            // colName
            // 
            this.colName.AppearanceHeader.FontStyleDelta = System.Drawing.FontStyle.Bold;
            this.colName.AppearanceHeader.Options.UseFont = true;
            this.colName.AppearanceHeader.Options.UseTextOptions = true;
            this.colName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colName.Caption = "名称";
            this.colName.FieldName = "名称";
            this.colName.MinWidth = 23;
            this.colName.Name = "colName";
            this.colName.OptionsColumn.AllowEdit = false;
            this.colName.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;
            this.colName.OptionsColumn.AllowSize = false;
            this.colName.Visible = true;
            this.colName.VisibleIndex = 2;
            this.colName.Width = 225;
            // 
            // colSCode
            // 
            this.colSCode.AppearanceHeader.FontStyleDelta = System.Drawing.FontStyle.Bold;
            this.colSCode.AppearanceHeader.Options.UseFont = true;
            this.colSCode.AppearanceHeader.Options.UseTextOptions = true;
            this.colSCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.colSCode.Caption = "简码";
            this.colSCode.FieldName = "简码";
            this.colSCode.MinWidth = 23;
            this.colSCode.Name = "colSCode";
            this.colSCode.OptionsColumn.AllowEdit = false;
            this.colSCode.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;
            this.colSCode.OptionsColumn.AllowSize = false;
            this.colSCode.Visible = true;
            this.colSCode.VisibleIndex = 3;
            this.colSCode.Width = 213;
            // 
            // frmcc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(696, 469);
            this.Controls.Add(this.gctMain);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.txtSearch);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmcc";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "主诉选择";
            this.Load += new System.EventHandler(this.frmcc_Load);
            ((System.ComponentModel.ISupportInitialize)(this.txtSearch.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gctMain)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvcc)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.TextEdit txtSearch;
        private DevExpress.XtraEditors.SimpleButton btnSearch;
        private DevExpress.XtraEditors.SimpleButton btnOK;
        private DevExpress.XtraEditors.SimpleButton btnCancel;
        private DevExpress.XtraGrid.GridControl gctMain;
        private DevExpress.XtraGrid.Views.Grid.GridView gvcc;
        private DevExpress.XtraGrid.Columns.GridColumn colChoose;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colType;
        private DevExpress.XtraGrid.Columns.GridColumn colCode;
        private DevExpress.XtraGrid.Columns.GridColumn colName;
        private DevExpress.XtraGrid.Columns.GridColumn colSCode;
    }
}