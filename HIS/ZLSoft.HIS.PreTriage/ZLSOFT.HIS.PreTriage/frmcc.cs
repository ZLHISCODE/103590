using System;
using System.Data;
using System.Windows.Forms;

namespace ZLSOFT.HIS.PreTriage
{
    public partial class frmcc : DevExpress.XtraEditors.XtraForm
    {
        public string gstrOut = "";
        public DataTable gdtdate;
        public string gstrTxt = "";
        public bool gbolSearch = true;
        private int mintRow = 0;
        public bool gbolOK = false;

        public frmcc(DataTable dtDate, string strTxt, bool bolSearch)
        {
            gdtdate = dtDate;
            gstrTxt = strTxt;
            gbolSearch = bolSearch;
            InitializeComponent();
        }

        private void txtSearch_EditValueChanged(object sender, EventArgs e)
        {
            mintRow = 0;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                string strtxt = txtSearch.Text.Trim().ToUpper();
                if (strtxt != "")
                {
                    for (int i = mintRow; i < gvcc.DataRowCount + 1; i++)
                    {
                        if (i == gvcc.DataRowCount)
                        {
                            if (mintRow == 0)
                            {
                                DevExpress.XtraEditors.XtraMessageBox.Show("未找到符合条件的数据！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                break;
                            }
                            else
                            {
                                DevExpress.XtraEditors.XtraMessageBox.Show("已查找至最后一行，请重新开始查找！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                mintRow = 0;
                                break;
                            }
                        }
                        var strN = gvcc.GetRowCellValue(i, colName).ToString();
                        var strS = gvcc.GetRowCellValue(i, colSCode).ToString();
                        var strC = gvcc.GetRowCellValue(i, colCode).ToString();
                        if (strN.IndexOf(strtxt) > -1 || strS.IndexOf(strtxt) > -1 || strC.IndexOf(strtxt) > -1)
                        {
                            gvcc.FocusedRowHandle = i;
                            mintRow = i + 1;
                            break;
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            gbolOK = false;
            this.Close();
        }

        private void txtSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnSearch_Click(this, new EventArgs());
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {

                gstrOut = "";
                for (int i = 0; i < gvcc.DataRowCount; i++)
                {
                    if (gvcc.GetRowCellValue(i, colChoose).ToString() == "True")
                    {
                        if (gstrOut == "")
                            gstrOut = gstrOut + gvcc.GetRowCellValue(i, colName).ToString();
                        else
                            gstrOut = gstrOut + "," + gvcc.GetRowCellValue(i, colName).ToString();
                    }
                }
                gbolOK = true;
                this.Close();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void frmcc_Load(object sender, EventArgs e)
        {
            try
            {
                var sql = string.Empty;
                if (gdtdate != null)
                {
                    gdtdate.Columns.Add(new DataColumn("选择", typeof(bool)));
                    gctMain.DataSource = gdtdate;
                    gvcc.ExpandAllGroups();
                    if (gstrTxt != "")
                    {
                        string strtxt = "," + gstrTxt + ",";
                        if (gbolSearch == false)
                        {
                            for (int i = 0; i < gvcc.DataRowCount; i++)
                            {
                                if (gvcc.IsGroupRow(i) == false)
                                {
                                    var str = "," + gvcc.GetRowCellValue(i, colName).ToString() + ",";
                                    if (strtxt.IndexOf(str) > -1)
                                    {
                                        gvcc.SetRowCellValue(i, colChoose, true);
                                    }
                                }
                            }
                        }
                        else
                        {
                            txtSearch.Text = gstrTxt;
                            btnSearch_Click(this, new EventArgs());
                        }
                    }
                }

                this.ActiveControl = gctMain;
                gctMain.Focus();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void gctMain_DoubleClick(object sender, EventArgs e)
        {
            btnOK_Click(sender, e);
        }

        private void gctMain_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                btnOK_Click(sender, e);
            }
        }
    }
}
