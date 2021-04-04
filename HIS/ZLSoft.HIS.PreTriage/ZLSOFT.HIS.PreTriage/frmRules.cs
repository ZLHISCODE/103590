using System;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace ZLSOFT.HIS.PreTriage
{
    public partial class frmRules : DevExpress.XtraEditors.XtraForm
    {
        public string gstrOut = "";
        public DataTable gdtdate;
        public string gstrTxt = "";
        public bool gbolSearch = true;
        private int mintRow = 0;
        public bool gbolOK = false;

        public frmRules(DataTable dtDate, string strTxt, bool bolSearch)
        {
            gdtdate = dtDate;
            gstrTxt = strTxt;
            gbolSearch = bolSearch;
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            gbolOK = false;
            this.Close();
        }

        private void frmRules_Load(object sender, EventArgs e)
        {
            try
            {
                if (gdtdate != null)
                {
                    if (gdtdate.Columns["选择"] == null)
                    {
                        gdtdate.Columns.Add(new DataColumn("选择", typeof(bool)));
                    }
                    gctRules.DataSource = gdtdate;
                    gvRules.ExpandAllGroups();
                    if (gstrTxt != "")
                    {
                        string strtxt = ";" + gstrTxt + ";";
                        //不为查找且文本不为空时，自动勾选文本内容
                        if (gbolSearch == false)
                        {
                            for (int i = 0; i < gdtdate.Rows.Count; i++)
                            {
                                if (gvRules.IsGroupRow(i) == false)
                                {
                                    var str = ";" + gvRules.GetRowCellValue(i, col_Name).ToString() + ";";
                                    if (strtxt.IndexOf(str) > -1)
                                    {
                                        gvRules.SetRowCellValue(i, col_Choose, true);
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

                this.ActiveControl = gctRules;
                gctRules.Focus();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

            try
            {

                string strtxt = txtSearch.Text.Trim().ToUpper();
                string str简码 = strtxt == "" ? "" : ComLib.Pinyin.GetInitials(strtxt, Encoding.GetEncoding("GB2312")).ToUpper();

                if (strtxt != "")
                {
                    for (int i = mintRow; i < gvRules.DataRowCount + 1; i++)
                    {
                        if (i == gvRules.DataRowCount)
                        {
                            if (mintRow == 0)
                            {
                                DevExpress.XtraEditors.XtraMessageBox.Show("未找到符合条件的数据！", "人工评定规则选择", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                break;
                            }
                            else
                            {
                                DevExpress.XtraEditors.XtraMessageBox.Show("已查找至最后一行，请重新开始查找！", "人工评定规则选择", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                mintRow = 0;
                                break;
                            }
                        }
                        var strN = gvRules.GetRowCellValue(i, col_Name).ToString();
                        var strS = gvRules.GetRowCellValue(i, col_Level).ToString();
                        string strN简码 = strN == "" ? "" : ComLib.Pinyin.GetInitials(strN, Encoding.GetEncoding("GB2312")).ToUpper();
                        if (strN.IndexOf(strtxt) > -1 || strS.IndexOf(strtxt) > -1 || strN简码.IndexOf(str简码) > -1)
                        {
                            gvRules.FocusedRowHandle = i;
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

        private void txtSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnSearch_Click(this, new EventArgs());
            }
        }

        private void txtSearch_EditValueChanged(object sender, EventArgs e)
        {
            mintRow = 0;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {

            try
            {

                gstrOut = "";
                for (int i = 0; i < gvRules.DataRowCount; i++)
                {
                    if (gvRules.GetRowCellValue(i, col_Choose).ToString() == "True")
                    {
                        if (gstrOut == "")
                            gstrOut = gstrOut + gvRules.GetRowCellValue(i, col_Name).ToString();
                        else
                            gstrOut = gstrOut + ";" + gvRules.GetRowCellValue(i, col_Name).ToString();
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

        private void gctRules_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                btnOK_Click(sender, e);
            }
        }

        private void gctRules_DoubleClick(object sender, EventArgs e)
        {
            btnOK_Click(sender, e);
        }
    }
}
