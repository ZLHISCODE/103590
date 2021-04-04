using System;
using System.Windows.Forms;
using ZLSOFT.HIS.PreTriage.ComLib;
using ZLSOFT.HIS.PreTriage.Models;
using ZLSOFT.HIS.PreTriage.DAL;
using DevExpress.XtraBars;
using System.Data;
using ZLSOFT.HIS.ZLIDCard;
using ZLSOFT.HIS.ZLMediCare;
using ZLSOFT.HIS.ZLVitalSignsCapture;

namespace ZLSOFT.HIS.PreTriage
{
    public partial class frmMain : DevExpress.XtraEditors.XtraForm
    {
        public frmMain()
        {
            DevExpress.XtraEditors.Controls.Localizer.Active = new MessboxClass();
            InitializeComponent();
        }

        frmDockEdit frmDock = new frmDockEdit();
        IDCard idCard = null;
        ReadVitalSigns objReadVital = null;
        ReadMediCare objReadCare = null;
        bool blnErr = false;
        EditService objEditSvr = new EditService();

        private void frmMain_Load(object sender, EventArgs e)
        {

            try
            {
                btnTimeOut.Visible = false;
                timeIcon.Enabled = true;
                //【2】打开新的窗体
                frmDock.TopLevel = false;//将子窗体设置成非顶级控件                 //newFrm.FormBorderStyle = FormBorderStyle.None;//去掉子窗体的边框
                frmDock.Parent = grpEdit;
                frmDock.Dock = DockStyle.Fill;//这句话保证子窗体会随着容器大小而变化

                frmDock.ChangeEdit(1);//新增窗体
                frmDock.Show();

                //更新挂号安排
                objEditSvr.Update挂号安排();

                lblUserName.Caption = BaseData.OperatorName;


                //初始化读生命体征接口
                //try
                //{
                //    if (objReadVital == null)
                //    {
                //        objReadVital = new ReadVitalSigns();
                //        if (objReadVital != null)
                //        {
                //            if (objReadVital.IntMain() == false)
                //            {
                //                if (objReadVital.ErrMeassge != "") ComLib.ComLib.SaveErrorlog(objReadVital.ErrMeassge);
                //                objReadVital = null;
                //            }
                //            if (objReadVital.ErrMeassge != "")
                //            {
                //                ComLib.ComLib.SaveErrorlog(objReadVital.ErrMeassge);
                //                objReadVital = null;
                //            }
                //        }
                //    }
                //}
                //catch 
                //{
                //    objReadVital = null;
                //}

                //if (objReadVital == null) {
                //    btn体征采集.Visibility = BarItemVisibility.Never;
                //}



                //初始化报表查看列表

                AddReport();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                blnErr = true;
                this.Close();
            }
        }

        private void AddReport()
        {

            try
            {
                DataTable dtList = objEditSvr.Get报表列表();

                if (dtList == null) return;
                if (dtList.Rows.Count == 0) return;
                BarButtonItem blnReport;

                foreach (DataRow dtRow in dtList.Rows)
                {
                    blnReport = new BarButtonItem() { Caption = ComLib.ComLib.Nvl(dtRow, "名称"), Tag = ComLib.ComLib.Nvl(dtRow, "编号") };
                    blnReport.ImageOptions.Image = this.btn报表查看.ImageOptions.Image;
                    blnReport.ItemClick += ReportClick;
                    popupReport.AddItem(blnReport);
                }

            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 预览报表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReportClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string strReportNO = "";
            string strPar = "";

            //分诊打印
            try
            {
                strReportNO = e.Item.Tag == null ? "" : e.Item.Tag.ToString();
                if (strReportNO == "") return;
                ComLib.ComLib.CReportOpen(strReportNO, strPar, 0);
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("打印失败：" + ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnAdd_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (frmDock.ISChange)
                {
                    if (DevExpress.XtraEditors.XtraMessageBox.Show("新增登记将清空当前录入,确认是否继续？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                    {
                        frmDock.IntCtlSet();
                    }
                }
                else
                {
                    frmDock.IntCtlSet();
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void btnExit_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.Close();
        }

        private void btnHelp_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            DevExpress.XtraEditors.XtraMessageBox.Show("功能模块：" + this.Text + System.Environment.NewLine + "当前用户：" + BaseData.OperatorName + System.Environment.NewLine + "版权所有：中联信息产业有限责任公司", "关于", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnReadCard_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

            ComLib.ComLib.ShowWaitform("正在" + btnReadCard.Caption, "请稍等", 2);
            try
            {
                CareInfor careinfo = Read医保卡();
                if (careinfo != null)
                {
                    if (careinfo.cardinfoName != null)
                    {
                        if (frmDock.ISChange)
                        {
                            if (DevExpress.XtraEditors.XtraMessageBox.Show("读取到病人【" + careinfo.cardinfoName + "】医保信息,请确认是否加载病人信息到当前界面？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                            {
                                return;
                            }
                        }
                        frmDock.IntCtlSet();
                        frmDock.LoadCarePati(careinfo);
                    }
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }


        /// <summary>
        /// 读取医保卡数据
        /// </summary>
        private CareInfor Read医保卡()
        {
            //初始化读医保卡接口
            try
            {
                if (objReadCare == null)
                {
                    objReadCare = new ReadMediCare();
                    if (objReadCare != null)
                    {
                        if (objReadCare.IntMain() == false)
                        {
                            if (objReadCare.ErrMessage != "") DevExpress.XtraEditors.XtraMessageBox.Show(objReadCare.ErrMessage, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            objReadCare = null;
                            return null;
                        }
                        if (objReadCare.ErrMessage != "")
                        {
                            DevExpress.XtraEditors.XtraMessageBox.Show(objReadCare.ErrMessage, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            //读取医保卡
            try
            {
                if (objReadCare != null)
                {
                    CareInfor careinfo = objReadCare.GetCardInfo();
                    if (objReadCare.ErrMessage != "")
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show(objReadCare.ErrMessage, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return null;
                    }
                    return careinfo;
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


        /// <summary>
        /// 读取身份证数据
        /// </summary>
        private PersonInfor Read身份证()
        {
            //初始化读身份证接口

            try
            {
                if (idCard == null)
                {
                    idCard = new IDCard();
                    if (idCard != null)
                    {
                        if (idCard.ErrMeassge != "")
                        {
                            DevExpress.XtraEditors.XtraMessageBox.Show(idCard.ErrMeassge, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            idCard = null;
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

            //读取身份证

            try
            {
                if (idCard != null)
                {
                    PersonInfor patiInfo = idCard.ReadIDCard();
                    if (idCard.ErrMeassge != "")
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show(idCard.ErrMeassge, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return null;
                    }
                    return patiInfo;
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

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (blnErr == false)
            {
                if (DevExpress.XtraEditors.XtraMessageBox.Show("确认要退出当前系统吗？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
            }
            frmlist.ISUnload = true;
            frmlist.Close();
            e.Cancel = false;
        }

        private void btnSend_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string strReportNO = "";
            string strPar = "";

            string str分诊id = "";
            int int绿色通道 = 0;

            string str姓名 = "";
            int int病情等级 = 0;
            if (frmDock.SaveData(out str分诊id, out int绿色通道, out int病情等级, out str姓名))
            {
                //分诊打印
                if (DevExpress.XtraEditors.XtraMessageBox.Show("病人【" + str姓名 + "】的分诊信息发送成功,是否打印" + (int绿色通道 == 0 ? "分诊指引单" : "病人腕带") + "？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    try
                    {
                        if (BaseData.gcnOracle == null)
                        {
                            DevExpress.XtraEditors.XtraMessageBox.Show("打印失败：数据库连接异常", "急诊预检分诊工作站", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        strReportNO = int绿色通道 == 0 ? "ZL1_REPORT_1244_1" : "ZL1_REPORT_1244_2";
                        strPar = "分诊id=" + str分诊id;
                        ComLib.ComLib.CReportOpen(strReportNO, strPar, 2);
                    }
                    catch (Exception ex)
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show("打印失败：" + ex.Message, "急诊预检分诊工作站", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        frmList frmlist = new frmList();
        private void btnView_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (frmlist.Visible == false)
            {
                frmlist.Show(this);
                frmlist.Refreshfrm();
            }
        }

        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                frmlist.Close();
                if (objReadVital != null)//如果创建了生命体征接口  最后卸载接口
                {
                    objReadVital.UnloadMain();
                }
                if (objReadCare != null)//如果创建了医保读取接口  最后卸载接口
                {
                    objReadCare.UnloadMain();
                }
            }
            catch
            {

            }
        }

        private void btnReadIDCard_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

            ComLib.ComLib.ShowWaitform("正在" + btnReadIDCard.Caption, "请稍等", 2);
            try
            {
                PersonInfor IdPatiInfo = Read身份证();
                if (IdPatiInfo != null)
                {
                    if (IdPatiInfo.Name != null)
                    {
                        if (frmDock.ISChange)
                        {
                            if (DevExpress.XtraEditors.XtraMessageBox.Show("读取到病人【" + IdPatiInfo.Name + "】身份信息,请确认是否加载病人信息到当前界面？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                            {
                                return;
                            }
                        }
                        frmDock.IntCtlSet();
                        frmDock.LoadIDPati(IdPatiInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void btn体征采集_ItemClick(object sender, ItemClickEventArgs e)
        {
            ComLib.ComLib.ShowWaitform("正在" + btn体征采集.Caption, "请稍等", 2);
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
                if (objReadVital != null) { 

                    VitalPatiInfo vitaiPatiInfo = new VitalPatiInfo();
  
                     VitalSignInfo VitalInfo = objReadVital.ReadInfo(frmDock.GetVitalPatiInfo());
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

        private void btnReportSet_ItemClick(object sender, ItemClickEventArgs e)
        {
            ComLib.ComLib.CReportSet();
        }

        private void frmMain_Resize(object sender, EventArgs e)
        {
            try
            {
                btnTimeOut.Left = this.Width - btnTimeOut.Width - 20;
                btnTimeOut.Top = 5;
            }
            catch { }
        }

        private void btnTimeOut_Click(object sender, EventArgs e)
        {
            if (frmlist.Visible == false)
            {
                frmlist.Show(this);
                frmlist.ShowTimeOut();
            }
        }

        private void timeIcon_Tick(object sender, EventArgs e)
        {
            try
            {
                DateTime dtNow = objEditSvr.GetNowTime();

                frmDock.IntCtlDate(dtNow);

                if (frmlist.Visible == false)
                {
                    string strbegin = "";
                    string strend = "";
                    DateTime dt开始时间 = frmlist.GetBeginTime();
                    if (dt开始时间 == null)
                    {
                        strbegin = dt开始时间.ToString("yyyy-MM-dd HH:mm");
                    }
                    else
                    {
                        strbegin = dtNow.ToString("yyyy-MM-dd 00:00");
                    }
                    strend = dtNow.ToString("yyyy-MM-dd 23:59");
                    DataTable dtpatilist = objEditSvr.Get病人列表(strbegin, strend, "0", "1");

                    if (dtpatilist != null)
                    {
                        if (dtpatilist.Rows.Count > 0)
                        {
                            btnTimeOut.Text = "当前有" + dtpatilist.Rows.Count + "个病人候诊超时";
                            btnTimeOut.Visible = true;

                        }
                        else
                        {
                            btnTimeOut.Visible = false;
                        }
                    }
                    else
                    {

                        btnTimeOut.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                timeIcon.Enabled = false;
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, "急诊预检分诊工作站", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }
    }
}
