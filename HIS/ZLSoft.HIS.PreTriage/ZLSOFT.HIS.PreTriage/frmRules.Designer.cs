namespace ZLSOFT.HIS.PreTriage
{
    partial class frmRules
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmRules));
            this.btnSearch = new DevExpress.XtraEditors.SimpleButton();
            this.btnOK = new DevExpress.XtraEditors.SimpleButton();
            this.btnCancel = new DevExpress.XtraEditors.SimpleButton();
            this.txtSearch = new DevExpress.XtraEditors.TextEdit();
            this.gctRules = new DevExpress.XtraGrid.GridControl();
            this.gvRules = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.col_Id = new DevExpress.XtraGrid.Columns.GridColumn();
            this.col_Type = new DevExpress.XtraGrid.Columns.GridColumn();
            this.col_Name = new DevExpress.XtraGrid.Columns.GridColumn();
            this.col_Person = new DevExpress.XtraGrid.Columns.GridColumn();
            this.col_Level = new DevExpress.XtraGrid.Columns.GridColumn();
            this.col_Choose = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemCheckEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSearch.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gctRules)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvRules)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).BeginInit();
            this.SuspendLayout();
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
            // txtSearch
            // 
            this.txtSearch.Location = new System.Drawing.Point(14, 14);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(374, 20);
            this.txtSearch.TabIndex = 0;
            this.txtSearch.EditValueChanged += new System.EventHandler(this.txtSearch_EditValueChanged);
            this.txtSearch.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSearch_KeyPress);
            // 
            // gctRules
            // 
            this.gctRules.Location = new System.Drawing.Point(14, 44);
            this.gctRules.MainView = this.gvRules;
            this.gctRules.Name = "gctRules";
            this.gctRules.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemCheckEdit1});
            this.gctRules.Size = new System.Drawing.Size(670, 373);
            this.gctRules.TabIndex = 2;
            this.gctRules.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvRules});
            this.gctRules.DoubleClick += new System.EventHandler(this.gctRules_DoubleClick);
            this.gctRules.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.gctRules_KeyPress);
            // 
            // gvRules
            // 
            this.gvRules.Appearance.FocusedRow.BackColor = System.Drawing.Color.LightSteelBlue;
            this.gvRules.Appearance.FocusedRow.Options.UseBackColor = true;
            this.gvRules.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.col_Id,
            this.col_Type,
            this.col_Name,
            this.col_Person,
            this.col_Level,
            this.col_Choose});
            this.gvRules.DetailHeight = 408;
            this.gvRules.GridControl = this.gctRules;
            this.gvRules.GroupCount = 1;
            this.gvRules.Name = "gvRules";
            this.gvRules.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.col_Type, DevExpress.Data.ColumnSortOrder.Ascending)});
            // 
            // col_Id
            // 
            this.col_Id.Caption = "ID";
            this.col_Id.FieldName = "ID";
            this.col_Id.MinWidth = 23;
            this.col_Id.Name = "col_Id";
            this.col_Id.Width = 23;
            // 
            // col_Type
            // 
            this.col_Type.Caption = "分类";
            this.col_Type.FieldName = "分类";
            this.col_Type.MinWidth = 23;
            this.col_Type.Name = "col_Type";
            this.col_Type.OptionsColumn.AllowEdit = false;
            this.col_Type.OptionsColumn.AllowSize = false;
            this.col_Type.Visible = true;
            this.col_Type.VisibleIndex = 1;
            this.col_Type.Width = 70;
            // 
            // col_Name
            // 
            this.col_Name.AppearanceHeader.FontStyleDelta = System.Drawing.FontStyle.Bold;
            this.col_Name.AppearanceHeader.Options.UseFont = true;
            this.col_Name.AppearanceHeader.Options.UseTextOptions = true;
            this.col_Name.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.col_Name.Caption = "指标名称";
            this.col_Name.FieldName = "指标名称";
            this.col_Name.MinWidth = 23;
            this.col_Name.Name = "col_Name";
            this.col_Name.OptionsColumn.AllowEdit = false;
            this.col_Name.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;
            this.col_Name.OptionsColumn.AllowSize = false;
            this.col_Name.Visible = true;
            this.col_Name.VisibleIndex = 1;
            this.col_Name.Width = 400;
            // 
            // col_Person
            // 
            this.col_Person.AppearanceHeader.FontStyleDelta = System.Drawing.FontStyle.Bold;
            this.col_Person.AppearanceHeader.Options.UseFont = true;
            this.col_Person.AppearanceHeader.Options.UseTextOptions = true;
            this.col_Person.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.col_Person.Caption = "适用人群";
            this.col_Person.FieldName = "适用人群";
            this.col_Person.MinWidth = 23;
            this.col_Person.Name = "col_Person";
            this.col_Person.OptionsColumn.AllowEdit = false;
            this.col_Person.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;
            this.col_Person.OptionsColumn.AllowSize = false;
            this.col_Person.Visible = true;
            this.col_Person.VisibleIndex = 2;
            this.col_Person.Width = 93;
            // 
            // col_Level
            // 
            this.col_Level.AppearanceHeader.FontStyleDelta = System.Drawing.FontStyle.Bold;
            this.col_Level.AppearanceHeader.Options.UseFont = true;
            this.col_Level.AppearanceHeader.Options.UseTextOptions = true;
            this.col_Level.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.col_Level.Caption = "病情级别";
            this.col_Level.FieldName = "病情级别";
            this.col_Level.MinWidth = 23;
            this.col_Level.Name = "col_Level";
            this.col_Level.OptionsColumn.AllowEdit = false;
            this.col_Level.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;
            this.col_Level.OptionsColumn.AllowSize = false;
            this.col_Level.Visible = true;
            this.col_Level.VisibleIndex = 3;
            this.col_Level.Width = 93;
            // 
            // col_Choose
            // 
            this.col_Choose.ColumnEdit = this.repositoryItemCheckEdit1;
            this.col_Choose.FieldName = "选择";
            this.col_Choose.MinWidth = 23;
            this.col_Choose.Name = "col_Choose";
            this.col_Choose.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;
            this.col_Choose.OptionsColumn.AllowSize = false;
            this.col_Choose.Visible = true;
            this.col_Choose.VisibleIndex = 0;
            this.col_Choose.Width = 86;
            // 
            // repositoryItemCheckEdit1
            // 
            this.repositoryItemCheckEdit1.Appearance.FontStyleDelta = System.Drawing.FontStyle.Bold;
            this.repositoryItemCheckEdit1.Appearance.Options.UseFont = true;
            this.repositoryItemCheckEdit1.Appearance.Options.UseTextOptions = true;
            this.repositoryItemCheckEdit1.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.repositoryItemCheckEdit1.AutoHeight = false;
            this.repositoryItemCheckEdit1.Name = "repositoryItemCheckEdit1";
            this.repositoryItemCheckEdit1.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked;
            // 
            // frmRules
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(696, 471);
            this.Controls.Add(this.gctRules);
            this.Controls.Add(this.txtSearch);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnSearch);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmRules";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "人工评定规则选择";
            this.Load += new System.EventHandler(this.frmRules_Load);
            ((System.ComponentModel.ISupportInitialize)(this.txtSearch.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gctRules)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvRules)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemCheckEdit1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.SimpleButton btnSearch;
        private DevExpress.XtraEditors.SimpleButton btnOK;
        private DevExpress.XtraEditors.SimpleButton btnCancel;
        private DevExpress.XtraEditors.TextEdit txtSearch;
        private DevExpress.XtraGrid.GridControl gctRules;
        private DevExpress.XtraGrid.Views.Grid.GridView gvRules;
        private DevExpress.XtraGrid.Columns.GridColumn col_Id;
        private DevExpress.XtraGrid.Columns.GridColumn col_Type;
        private DevExpress.XtraGrid.Columns.GridColumn col_Name;
        private DevExpress.XtraGrid.Columns.GridColumn col_Person;
        private DevExpress.XtraGrid.Columns.GridColumn col_Level;
        private DevExpress.XtraGrid.Columns.GridColumn col_Choose;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit repositoryItemCheckEdit1;
    }
}