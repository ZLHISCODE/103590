using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using ZLSOFT.HIS.PreTriage.DAL;
using ZLSOFT.HIS.PreTriage.Models;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using ADODB;
using DevExpress.XtraEditors;
using ZLSOFT.HIS.ZLVitalSignsCapture;

namespace ZLSOFT.HIS.PreTriage
{
    public partial class frmList : DevExpress.XtraEditors.XtraForm
    {

        ADODB.Connection gcnOracle = new ADODB.Connection();
        public frmList()
        {
            InitializeComponent();
        }

        public bool ISUnload = false;
        DataTable dtLevel = null;
        frmDockEdit frmDock = new frmDockEdit();
        EditService objEditService = new EditService();
        ReadVitalSigns objReadVital = null;
        bool blnNoLoad = false;

        private void frmList_Load(object sender, EventArgs e)
        {


            try
            {

                //【2】打开新的窗体
                frmDock.TopLevel = false;//将子窗体设置成非顶级控件                 //newFrm.FormBorderStyle = FormBorderStyle.None;//去掉子窗体的边框
                frmDock.Parent = palEdit;
                frmDock.Dock = DockStyle.Fill;//这句话保证子窗体会随着容器大小而变化

                frmDock.ChangeEdit(0);//新增窗体
                frmDock.Show();



                clb状态.Items.Clear();
                clb状态.Items.Add("未接诊", true);
                clb状态.Items.Add("已接诊", false);

                DateTime dt = objEditService.GetNowTime();
                dt开始时间.Text = dt.ToString("yyyy-MM-dd 00:00");
                dt结束时间.Text = dt.ToString("yyyy-MM-dd 23:59");

                RefreshList();
                ChangeEdit(false);
                toolEdit.BackColor = palEdit.BackColor;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }


        public void Refreshfrm()
        {
            DateTime dt = objEditService.GetNowTime();
            dt结束时间.Text = dt.ToString("yyyy-MM-dd 23:59");
            RefreshList();
            return;
        }

        public void ShowTimeOut()
        {
            blnNoLoad = true;
            ClearCheckList(clb状态, "未接诊");
            chk已超时.Checked = true;
            blnNoLoad = false;
            Refreshfrm();
            return;
        }

        public DateTime  GetBeginTime()
        {
            return dt开始时间.DateTime;
        }

        /// <summary>
        /// 清空选项并选择缺省选项
        /// </summary>
        /// <param name="chkControl"></param>
        /// <param name="dataTmp"></param>
        private void ClearCheckList(CheckedListBoxControl chkControl, string str缺省)
        {

            try
            {
                for (int i = 0; i < chkControl.Items.Count; i++)
                {
                    if (str缺省 != "" && chkControl.Items[i].Value.ToString() != "" && (",," + str缺省 + ",").IndexOf(("," + chkControl.Items[i].Value + ",").ToString()) > -1)
                    {
                        chkControl.SetItemCheckState(i, System.Windows.Forms.CheckState.Checked);
                    }
                    else
                    {
                        if (str缺省 != "" && chkControl.Items[i].Description != "" && (",," + str缺省 + ",").IndexOf(("," + chkControl.Items[i].Description + ",").ToString()) > -1)
                        {
                            chkControl.SetItemCheckState(i, System.Windows.Forms.CheckState.Checked);
                        }
                        else
                        {
                            chkControl.SetItemCheckState(i, System.Windows.Forms.CheckState.Unchecked);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }



        public void RefreshList()
        {
            if (blnNoLoad) return;
            if (dt开始时间.DateTime == null || dt开始时间.DateTime <= DateTime.MinValue || dt结束时间.DateTime == null || dt结束时间.DateTime <= DateTime.MinValue)
            {
                return;
            }
            try
            {
                string strState = "";
                string strbegin;
                string strend;
                strbegin = dt开始时间.DateTime.ToString("yyyy-MM-dd HH:mm");
                strend = dt结束时间.DateTime.ToString("yyyy-MM-dd HH:mm");
                for (int i = 0; i < clb状态.ItemCount; i++)
                {
                    if (clb状态.GetItemCheckState(i) == CheckState.Checked)
                    {
                        if (strState == "")
                            strState = i.ToString();
                        else
                            strState = strState + "," + i.ToString();
                    }
                }

                DataTable dtpatilist = objEditService.Get病人列表(strbegin, strend, strState, chk已超时.Visible ? (chk已超时.Checked ? "1" : "") : "");
                gridList.DataSource = dtpatilist;
                gvList.ExpandAllGroups();
                cbocount.Properties.Items.Clear();
                cbocount_EditValueChanged(null, null);
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }
        private void btnFind_Click(object sender, EventArgs e)
        {
            RefreshList();
        }

        private void Addcbo(string strid)
        {

            try
            {
                DataTable dtDetail = objEditService.Get病人分诊信息(strid);
                cbocount.Properties.Items.Clear();
                for (int i = 0; i < dtDetail.Rows.Count; i++)
                {
                    string stritem = "";
                    stritem = dtDetail.Rows[i]["病情情况"].ToString();

                    cbocount.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.ImageComboBoxItem[] {
                    new DevExpress.XtraEditors.Controls.ImageComboBoxItem(stritem, dtDetail.Rows[i]["分诊id"].ToString(),1)});
                    if (i == 0)
                        cbocount.SelectItemByDescription(stritem);
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void gvList_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {

            try
            {
                if (dtLevel == null)
                {

                    dtLevel = objEditService.Get急诊等级();
                }

                if (e.Column.Caption == col_Color.Caption)
                {

                    string strColor = "";
                    switch (gvList.GetRowCellValue(e.RowHandle, col_level).ToString())
                    {
                        case "1级":
                            strColor = dtLevel.Rows[0]["患者标识颜色"].ToString();
                            e.Appearance.BackColor = ColorTranslator.FromHtml("#" + strColor);

                            break;
                        case "2级":
                            strColor = dtLevel.Rows[1]["患者标识颜色"].ToString();
                            e.Appearance.BackColor = ColorTranslator.FromHtml("#" + strColor);
                            break;
                        case "3级":
                            strColor = dtLevel.Rows[2]["患者标识颜色"].ToString();
                            e.Appearance.BackColor = ColorTranslator.FromHtml("#" + strColor);
                            break;
                        case "4级":
                            strColor = dtLevel.Rows[3]["患者标识颜色"].ToString();
                            e.Appearance.BackColor = ColorTranslator.FromHtml("#" + strColor);
                            break;
                    }

                }
                else if (e.Column.Caption == col_绿色通道.Caption)
                {
                    e.DisplayText = "";
                    DataRow row = gvList.GetDataRow(e.RowHandle);

                    if (gvList.GetRowCellValue(e.RowHandle, col_绿色通道) == null) return;
                    if (gvList.GetRowCellValue(e.RowHandle, col_绿色通道).ToString() != "1") return;

                    Image image = imgCol16.Images[2];

                    if (row != null && row.RowError == String.Empty)
                    {
                        Rectangle r = e.Bounds;
                        int x = r.X + (r.Width - imgCol16.ImageSize.Width) / 2;
                        int y = r.Y + (r.Height - imgCol16.ImageSize.Height) / 2;
                        e.Graphics.DrawImageUnscaled(imgCol16.Images[2], x, y);
                    }


                }


            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

        private void cbocount_EditValueChanged(object sender, EventArgs e)
        {

            try
            {
                long lng分诊 = 0;
                lng分诊 = frmDock.lng分诊id;
                if (cbocount.EditValue == null || cbocount.Properties.Items.Count == 0)
                {
                    if (lng分诊 != 0) frmDock.IntCtlSet();
                    return;
                }
                if (cbocount.EditValue.ToString() == "" || cbocount.EditValue.ToString() == "0")
                {
                    if (lng分诊 != 0) frmDock.IntCtlSet();
                    return;
                }

                if (lng分诊.ToString() != cbocount.EditValue.ToString())
                {
                    frmDock.LoadEdit(ComLib.ComLib.StrToLng(cbocount.EditValue.ToString()));
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void frmList_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (btn取消.Visible)
            {
                if (DevExpress.XtraEditors.XtraMessageBox.Show("当前病人的再次分诊信息尚未保存,确认是否继续退出？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    btn取消_Click(null, null);
                }
            }
            if (ISUnload == false)
            {
                this.Hide();
                e.Cancel = true;
            }
        }

        private void clb状态_ItemCheck(object sender, DevExpress.XtraEditors.Controls.ItemCheckEventArgs e)
        {
            if (e.State != CheckState.Checked) return;
            chk已超时.Checked = false;
            layItem已超时.ContentVisible = e.State == CheckState.Checked && e.Index == 0;
            RefreshList();
        }

        private void btn再次分诊_Click(object sender, EventArgs e)
        {
            try
            {
                if (gvList.GetFocusedRowCellValue(col_就诊状态) == null) return;
                if (gvList.GetFocusedRowCellValue(col_就诊状态).ToString() != "0") return;
                if (cbocount.EditValue == null) return;
                if (cbocount.EditValue.ToString() == "" || cbocount.EditValue.ToString() == "0") return;
                if (!objEditService.Check就诊状态(cbocount.EditValue.ToString()))
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("当前病人已经就诊,不能再次分诊!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshList();
                    return;
                }
                frmDock.ChangeEdit(2);//修改
                frmDock.Get评分New();//修改
                ChangeEdit(true);
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void ChangeEdit(bool btnEdit)
        {
            try
            {
                //读取就诊状态

                string str就诊状态 = "";
                string str就诊ID = "";
                string str绿色通道 = "";
                str就诊状态 = (gvList.GetFocusedRowCellValue(col_就诊状态) == null) ? "" : gvList.GetFocusedRowCellValue(col_就诊状态).ToString();
                str就诊ID = (gvList.GetFocusedRowCellValue(col_id) == null) ? "" : gvList.GetFocusedRowCellValue(col_id).ToString();
                str绿色通道 = (gvList.GetFocusedRowCellValue(col_绿色通道) == null) ? "" : gvList.GetFocusedRowCellValue(col_绿色通道).ToString();


                //设置按钮状态

                btnSend.Visible = btnEdit;
                btn取消.Visible = btnEdit;
                btn生命体征.Visible = btnEdit;
                sp生命体征.Visible = btnEdit;


                btn再次分诊.Visible = (!(btnSend.Visible)) && (str就诊状态 == "0") && (str就诊ID != "");
                btn删除就诊.Visible = (!(btnSend.Visible)) && (str就诊状态 == "0") && (str就诊ID != "");
                btn绿色通道.Visible = (!(btnSend.Visible)) && (str就诊状态 == "0") && (str就诊ID != "");
                btn绿色通道.Text = str绿色通道 == "1" ? "取消绿色通道" : "标记绿色通道";
                btn打印指引单.Visible = !(btnSend.Visible) && (str绿色通道 != "1") && (str就诊ID != ""); ;
                btn补打腕带.Visible = !(btnSend.Visible) && (str绿色通道 == "1") && (str就诊ID != ""); ;
                btn刷新.Visible = !(btnSend.Visible);
                btnExit.Visible = !(btnSend.Visible);
                toolsplit1.Visible = !(btnSend.Visible) && (str就诊状态 != "0") && (str就诊ID != "");
                toolsplit2.Visible = !(btnSend.Visible);

                Menu再次分诊.Visible = btn再次分诊.Visible;
                Menu删除就诊.Visible = btn删除就诊.Visible;
                Menu补打指引单.Visible = btn打印指引单.Visible;
                Menu补打腕带.Visible = btn补打腕带.Visible;
                Menu绿色通道.Visible = btn绿色通道.Visible;
                Menu绿色通道.Text = btn绿色通道.Text + "(&L)";


                ComLib.ComLib.SetEnabled(dt开始时间, !btnSend.Visible);
                ComLib.ComLib.SetEnabled(dt结束时间, !btnSend.Visible);
                ComLib.ComLib.SetEnabled(btnFind, !btnSend.Visible);
                ComLib.ComLib.SetEnabled(gridList, !btnSend.Visible);
                ComLib.ComLib.SetEnabled(clb状态, !btnSend.Visible);
                ComLib.ComLib.SetEnabled(cbocount, !btnSend.Visible);
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void btn取消_Click(object sender, EventArgs e)
        {
            try
            {
                ChangeEdit(false);
                frmDock.ChangeEdit(0);
                if (frmDock.ISChange)
                {
                    frmDock.LoadEdit(ComLib.ComLib.StrToLng(cbocount.EditValue.ToString()));
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            try
            {
                string str分诊id = "";
                int int绿色通道 = 0;
                int int病情级别 = 0;
                string str病人姓名 = "";
                if (frmDock.SaveData(out str分诊id, out int绿色通道,out int病情级别,out str病人姓名))
                {

                    if (int病情级别 == ComLib.ComLib.StrToInt(gvList.GetFocusedRowCellValue(col_level).ToString()))
                    {
                        Addcbo(gvList.GetFocusedRowCellValue(col_id).ToString());
                        cbocount_EditValueChanged(null, null);
                    }
                    else
                    {
                        string strID = gvList.GetFocusedRowCellValue(col_id).ToString();//记录定位行

                        RefreshList();//重新加载控件

                        for (int i = 0; i < gvList.RowCount; i++)
                        {
                            if (gvList.GetRowCellValue(i, col_id) != null)
                            {

                                if (gvList.GetRowCellValue(i, col_id).ToString() == strID)
                                {
                                    this.gvList.FocusedRowHandle = i;
                                    this.gvList.FocusedColumn = this.gvList.Columns["姓名"];
                                }
                            }
                        }
                    }
                    ChangeEdit(false);
                    frmDock.ChangeEdit(0);


                    //分诊打印
                    if (DevExpress.XtraEditors.XtraMessageBox.Show("病人【" + str病人姓名 + "】的分诊信息发送成功,是否打印" + (int绿色通道 == 0 ? "分诊指引单" : "病人腕带") + "？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                    {
                        try
                        {
                            if (BaseData.gcnOracle == null)
                            {
                                DevExpress.XtraEditors.XtraMessageBox.Show("打印失败：数据库连接异常", "急诊预检分诊工作站", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                ComLib.ComLib.CReportOpen(int绿色通道 == 0 ? "ZL1_REPORT_1244_1" : "ZL1_REPORT_1244_2", "分诊id=" + str分诊id, 2);
                            }
                        }
                        catch (Exception ex)
                        {
                            DevExpress.XtraEditors.XtraMessageBox.Show("打印失败：" + ex.Message, "急诊预检分诊工作站", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void gvList_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {

            try
            {

                string strid = "";
                if (gvList.IsGroupRow(gvList.FocusedRowHandle) == false)
                {
                    if (gvList.GetFocusedRowCellValue(col_id) != null)
                    {
                        strid = gvList.GetFocusedRowCellValue(col_id).ToString();
                        Addcbo(strid);

                        //为触发刷新 单独执行
                        if (frmDock.lng就诊id != ComLib.ComLib.StrToLng(strid))
                        {
                            cbocount_EditValueChanged(null, null);
                        }

                    }
                    else
                    {
                        cbocount.Properties.Items.Clear();
                        cbocount_EditValueChanged(null, null);
                    }
                }
                else
                {
                    cbocount.Properties.Items.Clear();
                    cbocount_EditValueChanged(null, null);
                }

                ChangeEdit(false);
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void btn刷新_Click(object sender, EventArgs e)
        {

            try
            {
                long lng就诊 = 0;
                lng就诊 = frmDock.lng就诊id;
                if (cbocount.EditValue == null || cbocount.Properties.Items.Count == 0)
                {
                    if (lng就诊 != 0) frmDock.IntCtlSet();
                    return;
                }
                if (cbocount.EditValue.ToString() == "" || cbocount.EditValue.ToString() == "0")
                {
                    if (lng就诊 != 0) frmDock.IntCtlSet();
                    return;
                }
                frmDock.LoadEdit(ComLib.ComLib.StrToLng(cbocount.EditValue.ToString()));
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn删除就诊_Click(object sender, EventArgs e)
        {

            try
            {

                if (gvList.GetFocusedRowCellValue(col_就诊状态) == null) return;
                if (gvList.GetFocusedRowCellValue(col_就诊状态).ToString() != "0") return;
                if (gvList.GetFocusedRowCellValue(col_id) == null) return;
                if (gvList.GetFocusedRowCellValue(col_id).ToString() == "") return;
                if (cbocount.EditValue == null) return;
                if (cbocount.EditValue.ToString() == "" || cbocount.EditValue.ToString() == "0") return;
                if (!objEditService.Check就诊状态(cbocount.EditValue.ToString()))
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("当前病人已经就诊,不能删除就诊记录！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshList();
                    return;
                }
                if (DevExpress.XtraEditors.XtraMessageBox.Show("确认要删除病人" + gvList.GetFocusedRowCellValue(col_name).ToString() + "的就诊记录吗？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    try
                    {
                        objEditService.DelPretriage(gvList.GetFocusedRowCellValue(col_id).ToString());
                        ComLib.ComLib.ShowWaitform("病人" + gvList.GetFocusedRowCellValue(col_name).ToString() + "的就诊记录已成功删除", "记录删除成功", 2);

                    }
                    catch (Exception ex)
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show(ComLib.ComLib.GetMessageFromOracleError(ex.Message), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    //更新界面状态

                    foreach (DataRow dtRow in (gridList.DataSource as DataTable).Rows)
                    {
                        if (dtRow["就诊序号"].ToString() == gvList.GetFocusedRowCellValue(col_id).ToString())
                        {
                            (gridList.DataSource as DataTable).Rows.Remove(dtRow);
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

        private void dt开始时间_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                SendKeys.Send("{tab}");
            }
        }

        private void dt结束时间_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                SendKeys.Send("{tab}");
            }
        }

        private void clb状态_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                SendKeys.Send("{tab}");
            }
        }

        [HandleProcessCorruptedStateExceptions] //.net4支持对非托管异常的捕获

        private void btn打印指引单_Click(object sender, EventArgs e)
        {

            if (gvList.GetFocusedRowCellValue(col_level) == null) return;
            if (gvList.GetFocusedRowCellValue(col_level).ToString() == "") return;
            if (gvList.GetFocusedRowCellValue(col_id) == null) return;
            if (gvList.GetFocusedRowCellValue(col_id).ToString() == "") return;
            if (frmDock.lng分诊id == 0) return;

            //分诊打印
            try
            {
                if (BaseData.gcnOracle == null)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("打印失败：数据库连接异常", "急诊预检分诊工作站", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string strReportNO = "ZL1_REPORT_1244_1";
                string strPar = "分诊id=" + Convert.ToString(frmDock.lng分诊id);
                ComLib.ComLib.CReportOpen(strReportNO, strPar,2);

            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("打印失败：" + ex.Message, "急诊预检分诊工作站", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        [HandleProcessCorruptedStateExceptions] //.net4支持对非托管异常的捕获

        private void btn补打腕带_Click(object sender, EventArgs e)
        {
            if (gvList.GetFocusedRowCellValue(col_level) == null) return;
            if (gvList.GetFocusedRowCellValue(col_level).ToString() == "") return;

            if (gvList.GetFocusedRowCellValue(col_id) == null) return;
            if (gvList.GetFocusedRowCellValue(col_id).ToString() == "") return;


            if (frmDock.lng分诊id==0) return;

            //分诊打印
            try
            {
                if (BaseData.gcnOracle == null)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("打印失败：数据库连接异常", "急诊预检分诊工作站", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string strReportNO = "ZL1_REPORT_1244_2";
                string strPar = "分诊id=" + Convert.ToString(frmDock.lng分诊id);
                ComLib.ComLib.CReportOpen(strReportNO, strPar,2);
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("打印失败：" + ex.Message, "急诊预检分诊工作站", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Menu再次分诊_Click(object sender, EventArgs e)
        {
            btn再次分诊_Click(null, null);
        }

        private void Menu删除就诊_Click(object sender, EventArgs e)
        {
            btn删除就诊_Click(null, null);
        }

        private void Menu补打腕带_Click(object sender, EventArgs e)
        {
            btn补打腕带_Click(null, null);
        }

        private void Menu补打指引单_Click(object sender, EventArgs e)
        {
            btn打印指引单_Click(null, null);
        }

        private void btn绿色通道_Click(object sender, EventArgs e)
        {

            try
            {
                if (gvList.GetFocusedRowCellValue(col_就诊状态) == null) return;
                if (gvList.GetFocusedRowCellValue(col_就诊状态).ToString()!= "0") return;
                if (gvList.GetFocusedRowCellValue(col_id) == null) return;
                if (gvList.GetFocusedRowCellValue(col_id).ToString() == "") return;
                if (cbocount.EditValue == null) return;
                if (cbocount.EditValue.ToString() == "" || cbocount.EditValue.ToString() == "0") return;
                if (!objEditService.Check就诊状态(cbocount.EditValue.ToString()))
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("当前病人已经就诊,不能变更绿色通道标记！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshList();
                    return;
                }
                string strMsg = "";
                if (btn绿色通道.Text == "标记绿色通道")
                {
                    strMsg = "你确定要对【" + gvList.GetFocusedRowCellValue(col_name).ToString() + "】标记绿色通道吗？" + System.Environment.NewLine + "(注：绿色通道病人将实行先诊疗后付费)";
                }
                else
                {
                    strMsg = "你确定要对【" + gvList.GetFocusedRowCellValue(col_name).ToString() + "】取消绿色通道标记吗？";
                }

                if (DevExpress.XtraEditors.XtraMessageBox.Show(strMsg, this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    try
                    {
                        objEditService.ChangeGreenchannel(gvList.GetFocusedRowCellValue(col_id).ToString(), (btn绿色通道.Text == "标记绿色通道") ? "1" : "0");
                        ComLib.ComLib.ShowWaitform("病人" + gvList.GetFocusedRowCellValue(col_name).ToString() + "的绿色通道标识已成功变更", "绿色通道标识变更成功", 2);
                        //RefreshList();
                    }
                    catch
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show("变更绿色通道标记失败，服务器返回数据为空！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }


                    //更新界面状态

                    foreach (DataRow dtRow in (gridList.DataSource as DataTable).Rows)
                    {
                        if (dtRow["就诊序号"].ToString() == gvList.GetFocusedRowCellValue(col_id).ToString())
                        {
                            dtRow["是否绿色通道"] = (btn绿色通道.Text == "标记绿色通道") ? "1" : "0";
                        }
                    }

                    //同步处理界面
                    frmDock.Change绿色通道((btn绿色通道.Text == "标记绿色通道"));

                    //更新按钮状态

                    ChangeEdit(false);

                }

            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void Menu绿色通道_Click(object sender, EventArgs e)
        {
            btn绿色通道_Click(null, null);
        }

        bool blnCheckNo = false;
        /// <summary>
        /// 把CheckedListBoxControl设置为单选框
        /// </summary>
        /// <param name="chkControl">CheckedListBoxControl</param>
        /// <param name="index">index当前选中的索引</param>
        public void SingleSelectCheckedListBoxControls(CheckedListBoxControl chkControl, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            try
            {

                if (chkControl.CheckedItems.Count == 1 && chkControl.Items[e.Index].CheckState == CheckState.Checked && blnCheckNo == false)
                {
                    e.Cancel = true;
                    return;
                }
                if (chkControl.CheckedItems.Count > 0)
                {
                    for (int i = 0; i < chkControl.Items.Count; i++)
                    {
                        if (i != e.Index)
                        {
                            blnCheckNo = true;
                            chkControl.SetItemCheckState(i, System.Windows.Forms.CheckState.Unchecked);
                            blnCheckNo = false;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void clb状态_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb状态, e);
        }

        private void clb状态_DrawItem(object sender, ListBoxDrawItemEventArgs e)
        {
            try
            {
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    e.AllowDrawSkinBackground = false;
                    e.Appearance.BackColor2 = clb状态.BackColor;
                    e.Appearance.BorderColor = clb状态.BackColor;
                    e.Appearance.BackColor = clb状态.BackColor;
                }
            }
            catch
            {

            }
        }

        private void chk已超时_CheckedChanged(object sender, EventArgs e)
        {
            RefreshList();
        }

        private void chk已超时_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                SendKeys.Send("{tab}");
            }
        }

        private void btn生命体征_Click(object sender, EventArgs e)
        {
            ComLib.ComLib.ShowWaitform("正在" + btn生命体征.Text, "请稍等", 2);
            try
            {
                VitalSignInfo VitalInfo = Read生命体征();
                if (VitalInfo != null)
                {
                    frmDock.LoadVitalSign(VitalInfo);
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }


        /// <summary>
        /// 读取生命体征
        /// </summary>
        private VitalSignInfo Read生命体征()
        {
            //初始化读生命体征接口
            try
            {
                if (objReadVital == null)
                {
                    objReadVital = new ReadVitalSigns();
                    if (objReadVital != null)
                    {
                        if (objReadVital.IntMain() == false)
                        {
                            if (objReadVital.ErrMeassge != "") DevExpress.XtraEditors.XtraMessageBox.Show(objReadVital.ErrMeassge, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            objReadVital = null;
                            return null;
                        }
                        if (objReadVital.ErrMeassge != "")
                        {
                            DevExpress.XtraEditors.XtraMessageBox.Show(objReadVital.ErrMeassge, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            objReadVital = null;
                            return null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            //读取生命体征
            try
            {
                if (objReadVital != null)
                {
                    VitalSignInfo VitalInfo = objReadVital.ReadInfo();
                    if (objReadVital.ErrMeassge != "")
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show(objReadVital.ErrMeassge, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return null;
                    }
                    return VitalInfo;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }
    }
}