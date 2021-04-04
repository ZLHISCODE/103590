using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using ZLSOFT.HIS.PreTriage.DAL;
using DevExpress.Data.Filtering;
using System.Reflection;
using DevExpress.XtraGrid.Views.Grid;
using Newtonsoft.Json;
using ZLSOFT.HIS.PreTriage.Models;
using System.Xml;
using ZLSOFT.HIS.ZLIDCard;
using ZLSOFT.HIS.ZLVitalSignsCapture;
using ZLSOFT.HIS.ZLMediCare;
using System.Text.RegularExpressions;

namespace ZLSOFT.HIS.PreTriage
{
    public partial class frmDockEdit
        : DevExpress.XtraEditors.XtraForm
    {
        DateTime dtNow = DateTime.Now;
        DataTable data等级 = new DataTable();
        DataTable data主诉 = null;
        DataTable data人工评估 = null;

        int int最大儿童年龄 = 0;
        private bool blnCheckNo;

        public bool ISChange = false;
        DataTable data疼痛评分 = null;
        DataTable data昏迷评分 = null;
        DataTable data客观等级 = null;

        private int EditType = 0;  //编辑状态 0-查看 1-新增 2-修改
        public long lng就诊id = 0;
        public long lng分诊id = 0;
        private bool blnLoad = false;

        EditService objEditService = new EditService();//获取EditService对象
        DataTable data分诊信息 = null;
        DataTable data评分信息 = null;
        DataTable data指标信息 = null;

        private int intMaxLevel = 0;
        private string strMaxLevel = "";
        private string strMaxLevelColor = "";
        private string str病人id = "";
        private string str门诊号 = "";

        //缓存基础数据
        public frmDockEdit()
        {
            InitializeComponent();

            //实现文本控件的回车换行和获取焦点全选

            foreach (System.Windows.Forms.Control item in layoutControl1.Controls)
            {
                item.TextChanged += from_TextChanged;
                if (item is DevExpress.XtraEditors.TextEdit)
                {
                    item.KeyPress += SetNextControl;
                    item.Enter += SetSelectAll;
                    if (("," + item.Name).IndexOf("txt") > -1 && item.Name != "txt科室" && item.Name != "txt民族" && item.Name != "txt姓名")
                    {
                        item.Validated += txt_Validated;
                    }
                }
                else if (item is DevExpress.XtraEditors.CheckedListBoxControl)
                {
                    item.KeyPress += SetNextControl;
                    (item as CheckedListBoxControl).DrawItem += draw;
                }
            }
        }



        //作用 加快界面加载 
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x02000000;
                return cp;
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0014) // 禁掉清除背景消息
                return;
            base.WndProc(ref m);
        }

        private void frmDockEdit_Load(object sender, EventArgs e)
        {

            try
            {
                this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
                this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);

                //加载选择列表
                loadCheckList(clb病人来源, objEditService.Get病人来源());//初始化病人来源

                loadCheckList(clb意识形态, objEditService.Get意识状态());//初始化意识状态

                loadCheckList(clb陪同人员, objEditService.Get陪同人员());//初始化陪同人员

                loadCheckList(clb既往史, objEditService.Get常见既往史());//初始化常见既往史

                load评分指标(objEditService.Get评分指标());//初始化常见既往史


                //缓存急诊等级
                data等级 = objEditService.Get急诊等级();
                loadCheckList(clb修改分级, data等级);//初始化修改分级


                //加载性别选项
                cbo性别.Properties.Items.Clear();
                foreach (DataRow dr in objEditService.Get性别().Rows)
                {
                    if (dr["缺省"].ToString() == "1")
                    {
                        cbo性别.Tag = dr["名称"].ToString();
                    }
                    cbo性别.Properties.Items.Add(new DevExpress.XtraEditors.Controls.ComboBoxItem() { Value = dr["名称"].ToString() });
                }

                //加载国籍选项
                cbo国籍.Tag = "中国";
                cbo国籍.Properties.Items.Clear();
                foreach (DataRow dr in objEditService.Get国籍().Rows)
                {
                    if (dr["缺省"].ToString() == "1")
                    {
                        cbo国籍.Tag = dr["名称"].ToString();
                    }
                    cbo国籍.Properties.Items.Add(new DevExpress.XtraEditors.Controls.ComboBoxItem() { Value = dr["名称"].ToString() });
                }

                //加载年龄单位
                cbo年龄单位.Properties.Items.Clear();
                cbo年龄单位.Properties.Items.Add(new DevExpress.XtraEditors.Controls.ComboBoxItem() { Value = "岁" });
                cbo年龄单位.Properties.Items.Add(new DevExpress.XtraEditors.Controls.ComboBoxItem() { Value = "月" });
                cbo年龄单位.Properties.Items.Add(new DevExpress.XtraEditors.Controls.ComboBoxItem() { Value = "天" });
                cbo年龄单位.Tag = "岁";


                //加载民族
                txt民族.Properties.DataSource = objEditService.Get民族();
                if ((txt民族.Properties.DataSource as DataTable).Rows.Count > 0)
                {
                    txt民族.Tag = (txt民族.Properties.DataSource as DataTable).Rows[0]["名称"].ToString();
                }


                //加载科室
                txt科室.Properties.DataSource = objEditService.Get急诊科室();
                if ((txt科室.Properties.DataSource as DataTable).Rows.Count > 0)
                {
                    txt科室.Tag = (txt科室.Properties.DataSource as DataTable).Rows[0]["名称"].ToString();
                }

                //获取最大儿童年龄

                int最大儿童年龄 = ComLib.ComLib.StrToInt(objEditService.Get儿童年龄上限());

                //窗体控件初始值设置

                IntCtlSet();

                //控件初始化

                this.txt民族.Properties.PopupFilterMode = DevExpress.XtraEditors.PopupFilterMode.Contains;
                this.txt民族.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
                this.txt民族.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.True;
                this.txt民族.Properties.AutoComplete = false;
                this.txt民族.Properties.ImmediatePopup = true;
                this.txt科室.Properties.PopupFilterMode = DevExpress.XtraEditors.PopupFilterMode.Contains;
                this.txt科室.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
                this.txt科室.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.True;
                this.txt科室.Properties.AutoComplete = false;
                this.txt科室.Properties.ImmediatePopup = true;

                //加载控件
                this.KeyPreview = true;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        /// <summary>
        /// 刷新界面日期
        /// </summary>
        public void IntCtlDate(DateTime dtNow) {
            //时间初始化

            if (txt姓名.Text == "") { 
                dt到院时间.DateTime = dtNow;
                dt测量时间.Text = "";
                dt测量时间.EditValue = null;
                dt测量时间.DateTime = dtNow;
            }
        }


        /// <summary>
        /// 窗体控件初始值设置

        /// </summary>
        public void IntCtlSet()
        {
            try
            {
                blnLoad = true;
                //选择框初始化
                ClearCheckList(clb病人来源, clb病人来源.Tag == null ? "" : clb病人来源.Tag.ToString());
                ClearCheckList(clb意识形态, clb意识形态.Tag == null ? "" : clb意识形态.Tag.ToString());
                ClearCheckList(clb陪同人员, clb陪同人员.Tag == null ? "" : clb陪同人员.Tag.ToString());
                ClearCheckList(clb既往史, clb既往史.Tag == null ? "" : clb既往史.Tag.ToString());
                ClearCheckList(clb三无患者, "否");
                ClearCheckList(clb复合伤, "否");
                ClearCheckList(clb成批就诊, "否");
                ClearCheckList(clb绿色通道, "否");
                ClearCheckList(clb睁眼反应, "");
                ClearCheckList(clb语言反应, "");
                ClearCheckList(clb活动反应, "");
                ClearCheckList(clb疼痛描述, "");
                ClearCheckList(clb修改分级, "0");

                //遍历清空文本框数据

                foreach (System.Windows.Forms.Control item in layoutControl1.Controls)
                {
                    if (item is DevExpress.XtraEditors.TextEdit)
                    {
                        if (("," + item.Name).IndexOf("txt") > -1)
                        {
                            item.Text = "";
                            if (item.Name != "txt民族")
                            {
                                item.Tag = "";
                            }
                        }
                    }
                }


                //下拉框初始化
                if (cbo性别.Tag != null)
                {
                    cbo性别.Text = cbo性别.Tag.ToString();
                }
                else
                {
                    cbo性别.Text = "男";
                }
                if (cbo国籍.Tag != null)
                {
                    cbo国籍.Text = cbo国籍.Tag.ToString();
                }
                else
                {
                    cbo国籍.Text = "中国";
                }


                if (cbo年龄单位.Tag != null)
                {
                    cbo年龄单位.Text = cbo年龄单位.Tag.ToString();
                }
                else
                {
                    cbo年龄单位.Text = "岁";
                }
                layItem年龄单位.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;//显示文本框

                if (txt民族.Tag != null)
                {
                    txt民族.Text = txt民族.Tag.ToString();
                }
                else
                {
                    txt民族.Text = "汉族";
                }


                //时间初始化

                DateTime dtNow = objEditService.GetNowTime();


                dt到院时间.DateTime = dtNow;
                dt测量时间.Text = "";
                dt测量时间.EditValue = null;
                dt测量时间.DateTime = dtNow;

                dt出生日期.Text = "";
                dt出生日期.EditValue = null;


                //文本框初始化
                txt人数登记.Enabled = false;
                txt人数登记.Text = "";
                txt修改说明.Enabled = false;
                txt修改说明.Text = "";


                //初始化为最大等级

                txt自动分级.Text = strMaxLevel;
                txt自动分级.Tag = intMaxLevel.ToString();

                txt自动分级.BackColor = ColorTranslator.FromHtml("#" + strMaxLevelColor);

                txt科室.EditValue = null;

                data疼痛评分 = null;
                data昏迷评分 = null;
                data客观等级 = null;

                data分诊信息 = null;
                data评分信息 = null;
                data指标信息 = null;

                layoutGroup疼痛数字评分.Text = "疼痛数字评分";
                layoutGroup疼痛数字评分.Tag = 0;
                layoutGroupGCS.Text = "格拉斯哥昏迷评分";
                layoutGroupGCS.Tag = 0;

                if (EditType == 1)
                {
                    //更新功能状态

                    foreach (System.Windows.Forms.Control item in layoutControl1.Controls)
                    {
                        if (item is DevExpress.XtraEditors.TextEdit)
                        {
                            if (",txt姓名,cbo性别,cbo国籍,dt出生日期,txt身份证号,txt民族,txt家庭地址,txt年龄,cbo年龄单位,".IndexOf(item.Name) > -1)
                            {
                                (item as DevExpress.XtraEditors.TextEdit).ReadOnly = false;
                                (item as DevExpress.XtraEditors.TextEdit).TabStop = true;
                                item.BackColor = Color.White;
                            }
                        }
                    }
                }

                lng就诊id = 0;
                lng分诊id = 0;
                str病人id = "";
                str门诊号 = "";
                txt医保卡号.ToolTipTitle = "";
                txt保险类别.ToolTipTitle = "";

                txt姓名.Focus();
                ISChange = false;
                blnLoad = false;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        /// <summary>
        /// 修改当前编辑器状态

        /// </summary>
        /// <param name="intType">编辑状态 0-查看 1-新增 2-修改</param>
        public void ChangeEdit(int intType)
        {
            try
            {
                //更新功能状态

                foreach (System.Windows.Forms.Control item in layoutControl1.Controls)
                {
                    if (item is DevExpress.XtraEditors.TextEdit || item is DevExpress.XtraEditors.CheckedListBoxControl)
                    {
                        if (item.Name != "txt自动分级" && item.Name != "txt保险类别" && item.Name != "txt医保卡号")
                        {
                            item.BackColor = intType == 0 ? txt保险类别.BackColor : Color.White;
                        }
                        if (intType == 0)
                        {
                            ComLib.ComLib.SetEnabled(item, true);
                            ComLib.ComLib.SetEnabled(item, false);
                        }
                        else if (intType == 1)
                        {
                            ComLib.ComLib.SetEnabled(item, true);
                        }
                        else
                        {
                            if (",txt姓名,cbo性别,cbo国籍,dt出生日期,txt身份证号,txt联系电话,txt民族,txt医保卡号,txt保险类别,txt家庭地址,txt年龄,cbo年龄单位,".IndexOf(item.Name) > -1)
                            {
                                item.BackColor = txt保险类别.BackColor;
                                ComLib.ComLib.SetEnabled(item, true);
                                ComLib.ComLib.SetEnabled(item, false);
                            }
                            else
                            {
                                ComLib.ComLib.SetEnabled(item, true);
                            }
                        }

                    }
                }
                this.Refresh();
                EditType = intType;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }


        /// <summary>
        /// 修改当前绿色通道状态

        /// </summary>
        /// <param name="btn绿色通道">标记绿色通道</param>
        public void Change绿色通道(bool btn绿色通道)
        {
            try
            {
                blnLoad = true;
                ClearCheckList(clb绿色通道, btn绿色通道 ? "是" : "否");
                blnLoad = false;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        /// <summary>
        /// 针对修改界面重新获取评分信息
        /// </summary>
        public void Get评分New()
        {

            try
            {
                string str睁眼反应 = "";
                string str语言反应 = "";
                string str活动反应 = "";

                if (clb语言反应.CheckedItems.Count > 0 && clb活动反应.CheckedItems.Count > 0 && clb睁眼反应.CheckedItems.Count > 0)
                {
                    str睁眼反应 = clb睁眼反应.Tag.ToString() + ":" + clb睁眼反应.CheckedItems[0].ToString();
                    str语言反应 = clb语言反应.Tag.ToString() + ":" + clb语言反应.CheckedItems[0].ToString();
                    str活动反应 = clb活动反应.Tag.ToString() + ":" + clb活动反应.CheckedItems[0].ToString();
                    data昏迷评分 = objEditService.Get昏迷评分(str睁眼反应, str语言反应, str活动反应);
                }
                else
                {
                    data昏迷评分 = null;
                }
                string str年龄 = "";
                string str年龄单位 = "";
                string str指标信息 = "";//指标id：指标结果描述（可多个）

                //检查是否满足获取评分的要求
                if (txt年龄.Text == "" || (txt体温.Text == "" && txt脉搏.Text == "" && txt收缩压.Text == "" && txt舒张压.Text == "" && txt呼吸.Text == "" && txt血氧.Text == "" && txt血糖.Text == "" && txt血钾.Text == ""))
                {
                    data客观等级 = null;
                    Get自动评分();
                    return;
                }

                //解析年龄
                Get年龄解析(txt年龄.Text.Trim() + (cbo年龄单位.SelectedIndex == -1 ? "" : cbo年龄单位.Text), out str年龄, out str年龄单位);

                //组合客观条件
                if (txt体温.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lbl体温单位.Tag.ToString() + ":" + txt体温.Text;
                }
                if (txt脉搏.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lbl脉搏单位.Tag.ToString() + ":" + txt脉搏.Text;
                }
                if (txt收缩压.Text != "")
                {
                    str指标信息 = str指标信息 + "," + layItem收缩压.Tag.ToString() + ":" + txt收缩压.Text;
                }
                if (txt舒张压.Text != "")
                {
                    str指标信息 = str指标信息 + "," + layItem舒张压.Tag.ToString() + ":" + txt舒张压.Text;
                }
                if (txt呼吸.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lbl呼吸单位.Tag.ToString() + ":" + txt呼吸.Text;
                }
                if (txt血氧.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lblSPO2单位.Tag.ToString() + ":" + txt血氧.Text;
                }
                if (txt血糖.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lbl血糖单位.Tag.ToString() + ":" + txt血糖.Text;
                }
                if (txt血钾.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lbl血钾单位.Tag.ToString() + ":" + txt血钾.Text;
                }
                if (str指标信息 != "")
                {
                    str指标信息 = str指标信息.Substring(1);
                }
                data客观等级 = objEditService.Get客观评分(str年龄, str年龄单位, str指标信息);

                string str疼痛描述 = "";

                if (clb疼痛描述.CheckedItems.Count > 0)
                {
                    str疼痛描述 = clb疼痛描述.CheckedItems[0].ToString();
                    data疼痛评分 = objEditService.Get疼痛评分(str疼痛描述);
                }
                else
                {
                    data疼痛评分 = null;
                }

                if (dt测量时间.Text == "")
                {
                    dt测量时间.EditValue = null;
                    dt测量时间.DateTime = dtNow;
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        /// <summary>
        /// 通过分诊ID加载分诊数据
        /// </summary>
        /// <param name="intType">编辑状态 0-查看 1-新增 2-修改</param>
        public void LoadEdit(long lng分诊ID)
        {

            try
            {
                if (lng分诊ID == 0) return;

                data分诊信息 = objEditService.LoadPretriage(lng分诊ID.ToString());
                data评分信息 = objEditService.Load评分信息(lng分诊ID.ToString());
                data指标信息 = objEditService.Load指标信息(lng分诊ID.ToString());

                data疼痛评分 = null;
                data昏迷评分 = null;
                data客观等级 = null;

                if (data分诊信息 == null || data分诊信息.Rows.Count == 0) return;

                //遍历清空文本框数据

                foreach (System.Windows.Forms.Control item in layoutControl1.Controls)
                {
                    if (item is DevExpress.XtraEditors.TextEdit)
                    {
                        if (("," + item.Name).IndexOf("txt") > -1)
                        {
                            if (item.Name != "txt民族")
                            {
                                item.Tag = "";
                            }
                        }
                    }
                }

                lng就诊id = ComLib.ComLib.StrToLng(ComLib.ComLib.Nvl(data分诊信息.Rows[0], "就诊id"));
                lng分诊id = lng分诊ID;
                blnLoad = true;
                txt年龄.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "病人年龄");
                if (txt年龄.Text.Trim() != "")
                {
                    ReCalcOld(txt年龄.Text);
                }

                //选择框读取

                ClearCheckList(clb病人来源, ComLib.ComLib.Nvl(data分诊信息.Rows[0], "病人来源"));
                ClearCheckList(clb意识形态, ComLib.ComLib.Nvl(data分诊信息.Rows[0], "意识状态"));
                ClearCheckList(clb陪同人员, ComLib.ComLib.Nvl(data分诊信息.Rows[0], "陪同人员"));

                txt其他病史.Text = "";
                string str既往史 = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "既往病史");
                if (("," + str既往史).IndexOf("[其他]") == -1)
                {
                    ClearCheckList(clb既往史, str既往史);
                }
                else
                {
                        ClearCheckList(clb既往史, str既往史.Substring(0, str既往史.IndexOf("[其他]")));
                        txt其他病史.Text = str既往史.Substring(str既往史.IndexOf("[其他]")+4, str既往史.Length -(str既往史.IndexOf("[其他]") + 4));
                }


                ClearCheckList(clb三无患者, ComLib.ComLib.Nvl(data分诊信息.Rows[0], "是否三无人员") == "1" ? "是" : "否");
                ClearCheckList(clb成批就诊, ComLib.ComLib.Nvl(data分诊信息.Rows[0], "是否成批就诊") == "1" ? "是" : "否");
                ClearCheckList(clb复合伤, ComLib.ComLib.Nvl(data分诊信息.Rows[0], "是否复合伤") == "1" ? "是" : "否");
                ClearCheckList(clb绿色通道, ComLib.ComLib.Nvl(data分诊信息.Rows[0], "是否绿色通道") == "1" ? "是" : "否");

                //分级信息
                ClearCheckList(clb修改分级, ComLib.ComLib.Nvl(data分诊信息.Rows[0], "人工病情级别") == "" ? "0" : ComLib.ComLib.Nvl(data分诊信息.Rows[0], "人工病情级别"));

                //获取控件指标值

                if (data指标信息 != null && data指标信息.Rows.Count > 0)
                {
                    data指标信息.DefaultView.RowFilter = "指标id=" + clb疼痛描述.Tag.ToString();
                    if (data指标信息.DefaultView.ToTable().Rows.Count > 0)
                    {
                        ClearCheckList(clb疼痛描述, ComLib.ComLib.Nvl(data指标信息.DefaultView.ToTable().Rows[0], "指标结果文本"));
                    }
                    else
                    {
                        ClearCheckList(clb疼痛描述, "");
                    }

                    data指标信息.DefaultView.RowFilter = "指标id=" + clb睁眼反应.Tag.ToString();
                    if (data指标信息.DefaultView.ToTable().Rows.Count > 0)
                    {
                        ClearCheckList(clb睁眼反应, ComLib.ComLib.Nvl(data指标信息.DefaultView.ToTable().Rows[0], "指标结果文本"));
                    }
                    else
                    {
                        ClearCheckList(clb睁眼反应, "");
                    }
                    data指标信息.DefaultView.RowFilter = "指标id=" + clb语言反应.Tag.ToString();
                    if (data指标信息.DefaultView.ToTable().Rows.Count > 0)
                    {
                        ClearCheckList(clb语言反应, ComLib.ComLib.Nvl(data指标信息.DefaultView.ToTable().Rows[0], "指标结果文本"));
                    }
                    else
                    {
                        ClearCheckList(clb语言反应, "");
                    }
                    data指标信息.DefaultView.RowFilter = "指标id=" + clb活动反应.Tag.ToString();
                    if (data指标信息.DefaultView.ToTable().Rows.Count > 0)
                    {
                        ClearCheckList(clb活动反应, ComLib.ComLib.Nvl(data指标信息.DefaultView.ToTable().Rows[0], "指标结果文本"));
                    }
                    else
                    {
                        ClearCheckList(clb活动反应, "");
                    }
                }
                else
                {
                    ClearCheckList(clb疼痛描述, "");
                    ClearCheckList(clb睁眼反应, "");
                    ClearCheckList(clb语言反应, "");
                    ClearCheckList(clb活动反应, "");
                }

                data指标信息.DefaultView.RowFilter = "";



                if (data评分信息 != null && data评分信息.Rows.Count > 0)
                {
                    data评分信息.DefaultView.RowFilter = "方法id=4";
                    if (data评分信息.DefaultView.ToTable().Rows.Count > 0)
                    {
                        layoutGroup疼痛数字评分.Text = "疼痛数字评分             " + ComLib.ComLib.Nvl(data评分信息.DefaultView.ToTable().Rows[0], "评分结果描述") + "：" + ComLib.ComLib.Nvl(data评分信息.DefaultView.ToTable().Rows[0], "评分方法分值");
                        layoutGroup疼痛数字评分.Tag = ComLib.ComLib.StrToInt(ComLib.ComLib.Nvl(data评分信息.DefaultView.ToTable().Rows[0], "评分方法分值"));
                    }
                    else
                    {
                        layoutGroup疼痛数字评分.Text = "疼痛数字评分";
                        layoutGroup疼痛数字评分.Tag = 0;
                    }

                    data评分信息.DefaultView.RowFilter = "方法id=3";
                    if (data评分信息.DefaultView.ToTable().Rows.Count > 0)
                    {
                        layoutGroupGCS.Text = "格拉斯哥昏迷评分         " + ComLib.ComLib.Nvl(data评分信息.DefaultView.ToTable().Rows[0], "评分结果描述") + "：" + ComLib.ComLib.Nvl(data评分信息.DefaultView.ToTable().Rows[0], "评分方法分值");
                        layoutGroupGCS.Tag = ComLib.ComLib.StrToInt(ComLib.ComLib.Nvl(data评分信息.DefaultView.ToTable().Rows[0], "评分方法分值"));
                    }
                    else
                    {
                        layoutGroupGCS.Text = "格拉斯哥昏迷评分";
                        layoutGroupGCS.Tag = 0;
                    }
                }
                else
                {
                    layoutGroup疼痛数字评分.Text = "疼痛数字评分";
                    layoutGroup疼痛数字评分.Tag = 0;
                    layoutGroupGCS.Text = "格拉斯哥昏迷评分";
                    layoutGroupGCS.Tag = 0;
                }
                data评分信息.DefaultView.RowFilter = "";

                //文本框读取

                txt姓名.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "姓名");
                cbo性别.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "性别");
                cbo国籍.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "国籍") == "" ? cbo国籍.Tag.ToString() : ComLib.ComLib.Nvl(data分诊信息.Rows[0], "国籍");
                dt出生日期.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "出生日期");


                txt身份证号.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "身份证号");
                txt联系电话.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "手机号");
                txt民族.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "民族");
                txt医保卡号.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "医保号");
                txt保险类别.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "保险类别");
                txt家庭地址.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "家庭地址");
                dt到院时间.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "到院时间");
                txt人数登记.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "成批就诊人数");
                txt主诉.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "主诉");
                txt备注.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "备注");
                txt人工评定.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "人工评级说明");

                //设置病人科室
                txt科室.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "分诊科室名称");
                txt科室.Tag = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "分诊科室名称");

                txt科室.EditValue = null;
                foreach (DataRow dtRow in (txt科室.Properties.DataSource as DataTable).Rows)
                {
                    if (dtRow["名称"].ToString() == ComLib.ComLib.Nvl(data分诊信息.Rows[0], "分诊科室名称"))
                    {
                        txt科室.EditValue = dtRow["id"];
                        break;
                    }
                }

                //病人体征

                dt测量时间.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "体征测量时间");
                txt收缩压.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "收缩压");
                txt舒张压.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "舒张压");
                txt血糖.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "血糖");
                txt血氧.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "指氧饱和度");
                txt脉搏.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "心率");
                txt血钾.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "血钾");
                txt体温.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "体温");
                txt呼吸.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "呼吸频率");
                txt修改说明.Text = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "修改说明");

                //读取自动分级
                txt自动分级.Tag = ComLib.ComLib.Nvl(data分诊信息.Rows[0], "自动病情级别");
                for (int i = 0; i < data等级.Rows.Count; i++)
                {
                    if (data等级.Rows[i]["序号"].ToString() == ComLib.ComLib.Nvl(data分诊信息.Rows[0], "自动病情级别"))
                    {
                        txt自动分级.Text = data等级.Rows[i]["名称"].ToString();
                        txt自动分级.BackColor = ColorTranslator.FromHtml("#" + data等级.Rows[i]["患者标识颜色"].ToString());
                    }
                }

                blnLoad = false;
                ISChange = false;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        /// <summary>
        /// 通过身份证信息加载病人信息

        /// </summary>
        /// <param name="PatiInfo">病人信息</param>
        public void LoadIDPati(PersonInfor PatiInfo)
        {

            try
            {
                if (PatiInfo == null) return;
                if (PatiInfo.Name == null) return;
                //文本框读取

                txt姓名.Text = PatiInfo.Name;
                txt姓名.Tag = PatiInfo.Name;
                cbo性别.Text = PatiInfo.Sex;
                cbo国籍.Text = cbo国籍.Tag == null ? "中国" : cbo国籍.Tag.ToString();
                dt出生日期.Text = DateTime.ParseExact(PatiInfo.Birthday, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                txt身份证号.Text = PatiInfo.Identity;
                txt民族.Text = PatiInfo.Nation + "族";
                txt家庭地址.Text = PatiInfo.Address;



                //检查是否存在病人信息


                DataTable dt身份证List = objEditService.GetPatlistbyIdCard(txt身份证号.Text.Trim());

                if (dt身份证List != null)
                {
                    if (dt身份证List.Rows.Count > 0)
                    {
                        str病人id = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "病人id");
                        str门诊号 = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "门诊号");
                        txt联系电话.Text = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "手机号");
                        txt医保卡号.Text = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "医保号");
                        txt保险类别.Text = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "保险类别");
                        
                    }
                }
                //更新功能状态

                foreach (System.Windows.Forms.Control item in layoutControl1.Controls)
                {
                    if (item is DevExpress.XtraEditors.TextEdit)
                    {
                        if (",txt姓名,cbo性别,cbo国籍,dt出生日期,txt身份证号,txt民族,txt家庭地址,txt年龄,cbo年龄单位,".IndexOf(item.Name) > -1)
                        {
                            (item as DevExpress.XtraEditors.TextEdit).ReadOnly = true;
                            (item as DevExpress.XtraEditors.TextEdit).TabStop = false;
                            item.BackColor = txt保险类别.BackColor;
                        }
                    }
                }

                ISChange = true;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }



        /// <summary>
        /// 获取病人信息
        /// </summary>
        public VitalPatiInfo GetVitalPatiInfo()
        {
            VitalPatiInfo vitalPatiInfo = new VitalPatiInfo();
            vitalPatiInfo.姓名 = txt姓名.Text;
            vitalPatiInfo.性别 = cbo性别.Text;
            vitalPatiInfo.病人ID = ComLib.ComLib.StrToInt(str病人id) ;
            vitalPatiInfo.身份证号 = txt身份证号.Text;
            vitalPatiInfo.医保卡号 = txt医保卡号.Text;
            return vitalPatiInfo;
        }




        /// <summary>
        /// 通过医保卡信息加载信息

        /// </summary>
        /// <param name="PatiInfo">病人信息</param>
        public void LoadCarePati(CareInfor careinfo)
        {

            try
            {
                if (careinfo == null) return;
                if (careinfo.cardinfoName == null) return;
                //文本框读取
                txt姓名.Text = careinfo.cardinfoName;
                txt姓名.Tag = careinfo.cardinfoName;
                cbo性别.Text = careinfo.Sex;
                cbo国籍.Text = cbo国籍.Tag == null ? "中国" : cbo国籍.Tag.ToString();
                dt出生日期.Text = DateTime.ParseExact(careinfo.Birthday, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                txt身份证号.Text = careinfo.Id_No;
                txt医保卡号.Text = careinfo.Card_No;
                txt保险类别.Text = careinfo.Card_Type;

                if (txt身份证号.Text.Trim() != "")
                {
                    //检查是否存在病人信息
                    DataTable dt身份证List = objEditService.GetPatlistbyIdCard(txt身份证号.Text.Trim());

                    if (dt身份证List != null)
                    {
                        if (dt身份证List.Rows.Count > 0)
                        {
                            str病人id = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "病人id");
                            str门诊号 = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "门诊号");
                            txt联系电话.Text = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "手机号");
                            txt民族.Text = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "民族");
                            txt家庭地址.Text = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "家庭地址");
                            txt医保卡号.ToolTipTitle = str病人id;
                            txt保险类别.ToolTipTitle = str门诊号;
                        }
                    }
                }

                //通过医保卡号来查询
                if (str病人id == "")
                {
                    DataTable dt医保卡List = objEditService.Get_Patlistbymedical(careinfo.Card_No, careinfo.Card_Type);
                    if (dt医保卡List != null)
                    {
                        if (dt医保卡List.Rows.Count > 0)
                        {
                            str病人id = ComLib.ComLib.Nvl(dt医保卡List.Rows[0], "病人id");
                            str门诊号 = ComLib.ComLib.Nvl(dt医保卡List.Rows[0], "门诊号");
                            txt联系电话.Text = ComLib.ComLib.Nvl(dt医保卡List.Rows[0], "手机号");
                            txt民族.Text = ComLib.ComLib.Nvl(dt医保卡List.Rows[0], "民族");
                            txt家庭地址.Text = ComLib.ComLib.Nvl(dt医保卡List.Rows[0], "家庭地址");

                            txt医保卡号.ToolTipTitle = str病人id;
                            txt保险类别.ToolTipTitle = str门诊号;

                            if (txt身份证号.Text == "") txt身份证号.Text = ComLib.ComLib.Nvl(dt医保卡List.Rows[0], "身份证号");
                            if (cbo性别.Text == "") cbo性别.Text = ComLib.ComLib.Nvl(dt医保卡List.Rows[0], "性别");
                            if (dt出生日期.Text == "") dt出生日期.Text = ComLib.ComLib.Nvl(dt医保卡List.Rows[0], "出生日期");
                        }
                    }
                }

                //更新功能状态
                foreach (System.Windows.Forms.Control item in layoutControl1.Controls)
                {
                    if (item is DevExpress.XtraEditors.TextEdit)
                    {
                        if (",txt姓名,cbo性别,cbo国籍,dt出生日期,txt身份证号,txt民族,txt年龄,cbo年龄单位,".IndexOf(item.Name) > -1 && item.Text != "")
                        {
                            (item as DevExpress.XtraEditors.TextEdit).ReadOnly = true;
                            (item as DevExpress.XtraEditors.TextEdit).TabStop = false;
                            item.BackColor = txt保险类别.BackColor;
                        }
                    }
                }

                ISChange = true;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }



        /// <summary>
        /// 加载接口读取病人生命体征
        /// </summary>
        /// <param name="SignInfo">生命体征信息</param>
        public void LoadVitalSign(VitalSignInfo SignInfo)
        {

            try
            {

                if (SignInfo == null) return;
                //读取生命体征
                txt收缩压.Text = SignInfo.收缩压;
                txt舒张压.Text = SignInfo.舒张压;
                txt血糖.Text = SignInfo.血糖;
                txt血氧.Text = SignInfo.指氧饱和度;
                txt脉搏.Text = SignInfo.心率;
                txt血钾.Text = SignInfo.血钾;
                txt体温.Text = SignInfo.体温;
                txt呼吸.Text = SignInfo.呼吸频率;
                Set客观评分();
                ISChange = true;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }




        /// <summary>
        /// 检查分诊录入项
        /// </summary>
        public bool CheckData()
        {

            try
            {
                //检查姓名录入

                if (txt姓名.Text.Trim() == "")
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("病人的姓名不允许为空,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt姓名.Focus();
                    return false;
                }
                //检查到院时间

                if (dt到院时间.DateTime == null || dt到院时间.DateTime <= DateTime.MinValue)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("病人的到院时间不允许为空,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    dt到院时间.Focus();
                    return false;
                }

                //检查主诉

                if (txt主诉.Text.Length > 25)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("病人的主诉不能超过25个字符,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt主诉.Focus();
                    return false;
                }



                //检查测量时间

                if (txt体温.Text.Trim() != "" || txt脉搏.Text.Trim() != "" || txt舒张压.Text.Trim() != "" || txt收缩压.Text.Trim() != "" || txt呼吸.Text.Trim() != "" || txt血氧.Text.Trim() != "" || txt血糖.Text.Trim() != "" || txt血钾.Text.Trim() != "")
                {
                    if (dt测量时间.Text == "")
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show("病人生命体征的测量时间不允许为空,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dt测量时间.Focus();
                        return false;
                    }
                }

                //分诊科室
                if (txt科室.Text.Trim() == "")
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("分诊科室不允许为空,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt科室.Focus();
                    return false;
                }

                //检查昏迷填写

                if (!(clb活动反应.CheckedItems.Count == clb睁眼反应.CheckedItems.Count && clb活动反应.CheckedItems.Count == clb语言反应.CheckedItems.Count))
                {
                    if (clb睁眼反应.CheckedItems.Count == 0)
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show("格拉斯哥昏迷评分未填写完整,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        clb睁眼反应.Focus();
                        return false;
                    }
                    if (clb语言反应.CheckedItems.Count == 0)
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show("格拉斯哥昏迷评分未填写完整,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        clb语言反应.Focus();
                        return false;
                    }
                    if (clb活动反应.CheckedItems.Count == 0)
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show("格拉斯哥昏迷评分未填写完整,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        clb活动反应.Focus();
                        return false;
                    }
                }


                //检查修改依据
                if (txt人工评定.Text.Length > 50)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("病情等级修改依据不能超过50个字符,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt人工评定.Focus();
                    return false;
                }



                return true;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
        }

        /// <summary>
        /// 保存分诊录入项

        /// </summary>
        public bool SaveData(out string str分诊ID, out int int绿色通道, out int int病情等级, out string str姓名)
        {
            str分诊ID = "";
            int绿色通道 = 0;
            int病情等级 = 0;
            str姓名 = "";
            try
            {

                //检查分诊录入项
                if (CheckData() == false)
                {
                    return false;
                }

                //确认是否发送 防止误操作

                if (DevExpress.XtraEditors.XtraMessageBox.Show("确认是否要发送当前病人的分诊信息？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                {
                    return false;
                }

                //组建Json入参

                ComLib.ComLib.ShowWaitform("正在发送病人" + txt姓名.Text.Trim() + "的分诊记录", "分诊发送中", 2);
                //获取病人评分Json
                List<object> List病人评分 = new List<object>();
                //获取评分指标Json
                List<object> List评分指标 = new List<object>();


                if (data疼痛评分 != null)
                {
                    var items = new
                    {
                        方法ID = "4",
                        评分方法分值 = data疼痛评分.Rows[0]["病人分数"].ToString(),
                        评分结果描述 = data疼痛评分.Rows[0]["描述"].ToString(),
                        评分等级 = data疼痛评分.Rows[0]["病人等级"].ToString()
                    };
                    List病人评分.Add(items);

                    var item疼痛指标 = new
                    {
                        方法ID = "4",
                        指标ID = clb疼痛描述.Tag.ToString(),
                        指标结果文本 = clb疼痛描述.CheckedItems[0].ToString(),
                    };
                    List评分指标.Add(item疼痛指标);
                }

                if (data昏迷评分 != null)
                {
                    var items昏迷 = new
                    {
                        方法ID = "3",
                        评分方法分值 = data昏迷评分.Rows[0]["病人分数"].ToString(),
                        评分结果描述 = data昏迷评分.Rows[0]["描述"].ToString(),
                        评分等级 = data昏迷评分.Rows[0]["病人等级"].ToString()
                    };
                    List病人评分.Add(items昏迷);

                    var item睁眼指标 = new
                    {
                        方法ID = "3",
                        指标ID = clb睁眼反应.Tag.ToString(),
                        指标结果文本 = clb睁眼反应.CheckedItems[0].ToString(),
                    };
                    List评分指标.Add(item睁眼指标);


                    var item语言指标 = new
                    {
                        方法ID = "3",
                        指标ID = clb语言反应.Tag.ToString(),
                        指标结果文本 = clb语言反应.CheckedItems[0].ToString(),
                    };
                    List评分指标.Add(item语言指标);


                    var item活动指标 = new
                    {
                        方法ID = "3",
                        指标ID = clb活动反应.Tag.ToString(),
                        指标结果文本 = clb活动反应.CheckedItems[0].ToString(),
                    };
                    List评分指标.Add(item活动指标);

                }

                if (data客观等级 != null)
                {
                    var items客观 = new
                    {
                        方法ID = CheckIS儿童() ? "2" : "1",
                        评分方法分值 = "",
                        评分结果描述 = "",
                        评分等级 = data客观等级.Rows[0]["病人等级"].ToString()
                    };
                    List病人评分.Add(items客观); ;
                }

                string str年龄数值 = "";
                string str年龄单位 = "";

                //解析年龄
                if (txt年龄.Text != "")
                {
                    Get年龄解析(txt年龄.Text.Trim() + (cbo年龄单位.SelectedIndex == -1 ? "" : cbo年龄单位.Text), out str年龄数值, out str年龄单位);
                }

                //获取病人既往病史
                string str既往史 = "";
                for (int i = 0; i <= clb既往史.CheckedItems.Count - 1; i++)
                {
                    str既往史 = str既往史 + "," + clb既往史.CheckedItems[i].ToString();
                }
                if (str既往史 != "")
                {
                    str既往史 = str既往史.Substring(1);
                }
                str既往史 = str既往史 + (txt其他病史.Text.Trim() == "" ? "" : "[其他]" + txt其他病史.Text.Trim());

                //获取病人科室
                string str科室id = "";
                foreach (DataRow dtRow in (txt科室.Properties.DataSource as DataTable).Rows)
                {
                    if (dtRow["名称"].ToString() == txt科室.Text)
                    {
                        str科室id = dtRow["id"].ToString();
                        break;//退出整个循环

                    }
                }

                string message = JsonConvert.SerializeObject(new
                {
                    type = EditType.ToString(),
                    就诊id = lng就诊id != 0 ? lng就诊id.ToString() : "",
                    病人id = str病人id,
                    门诊号 = str门诊号,
                    姓名 = txt姓名.Text.Trim(),
                    性别 = cbo性别.Text.Trim(),
                    国籍 = cbo国籍.Text.Trim(),
                    出生日期 = dt出生日期.Text == "" ? "" : dt出生日期.DateTime.ToString("yyyy-MM-dd"),
                    身份证号 = txt身份证号.Text.Trim(),
                    联系电话 = txt联系电话.Text.Trim(),
                    民族 = txt民族.Text.Trim(),
                    医保卡号 = txt医保卡号.Text.Trim(),
                    保险类别 = txt保险类别.Text.Trim(),
                    家庭地址 = txt家庭地址.Text.Trim(),
                    病人年龄 = txt年龄.Text.Trim() + (cbo年龄单位.SelectedIndex == -1 ? "" : cbo年龄单位.Text),
                    年龄数值 = str年龄数值.Trim(),
                    年龄单位 = str年龄单位.Trim(),
                    到院时间 = dt到院时间.Text == "" ? "" : dt到院时间.DateTime.ToString("yyyy-MM-dd HH:mm"),
                    是否三无人员 = clb三无患者.CheckedItems[0].ToString() == "是" ? "1" : "0",
                    是否复合伤 = clb复合伤.CheckedItems[0].ToString() == "是" ? "1" : "0",
                    是否成批就诊 = clb成批就诊.CheckedItems[0].ToString() == "是" ? "1" : "0",
                    是否绿色通道 = clb绿色通道.CheckedItems[0].ToString() == "是" ? "1" : "0",
                    成批就诊人数 = txt人数登记.Text.Trim(),
                    病人来源 = clb病人来源.CheckedItems[0].ToString(),
                    陪同人员 = clb陪同人员.CheckedItems[0].ToString(),
                    意识状态 = clb意识形态.CheckedItems[0].ToString(),
                    既往病史 = str既往史,
                    主诉 = txt主诉.Text.Trim(),
                    病情级别 = clb修改分级.CheckedItems[0].ToString() == "0" ? txt自动分级.Tag.ToString() : clb修改分级.CheckedItems[0].ToString(),
                    登记人 = BaseData.OperatorName,
                    分诊状态 = "1",
                    备注 = txt备注.Text.Trim(),
                    分诊科室id = str科室id,
                    分诊科室名称 = txt科室.Text.Trim(),
                    体征测量时间 = dt测量时间.Text == "" ? "" : (txt体温.Text == "" && txt脉搏.Text == "" && txt舒张压.Text == "" && txt收缩压.Text == "" && txt呼吸.Text == "" && txt血氧.Text == "" && txt血糖.Text == "" && txt血钾.Text == "") ? "" : dt测量时间.DateTime.ToString("yyyy-MM-dd HH:mm"),
                    舒张压 = txt舒张压.Text.Trim(),
                    收缩压 = txt收缩压.Text.Trim(),
                    血糖 = txt血糖.Text.Trim(),
                    指氧饱和度 = txt血氧.Text.Trim(),
                    心率 = txt脉搏.Text.Trim().ToString(),
                    血钾 = txt血钾.Text.Trim(),
                    体温 = txt体温.Text.Trim(),

                    呼吸频率 = txt呼吸.Text.Trim(),
                    自动病情级别 = txt自动分级.Tag.ToString(),
                    人工病情级别 = clb修改分级.CheckedItems[0].ToString() == "0" ? "" : clb修改分级.CheckedItems[0].ToString(),
                    人工评级说明 = txt人工评定.Text.Trim(),
                    修改说明 = txt修改说明.Text.Trim(),
                    登记人编号 = BaseData.OperatorCode,
                    评分指标 = List评分指标,
                    病人评分 = List病人评分,
                    站点 = BaseData.站点

                });

                DataTable dtOut = new DataTable();
                try
                {
                    dtOut = objEditService.SavePretriage(message);
                    if (dtOut != null)
                    {
                        ComLib.ComLib.ShowWaitform("病人" + txt姓名.Text.Trim() + "的分诊记录已发送", "发送成功", 4);

                        str分诊ID = dtOut.Rows[0]["分诊id"].ToString();
                        int绿色通道 = clb绿色通道.CheckedItems[0].ToString() == "是" ? 1 : 0;
                        int病情等级 = ComLib.ComLib.StrToInt(clb修改分级.CheckedItems[0].ToString() == "0" ? txt自动分级.Tag.ToString() : clb修改分级.CheckedItems[0].ToString());
                        str姓名 = txt姓名.Text.Trim();

                        //编辑界面初始化

                        IntCtlSet();
                    }
                    else
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show("获取保存数据失败，服务器返回数据为空！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return false;
                    }

                }
                catch (Exception ex)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(ComLib.ComLib.GetMessageFromOracleError(ex.Message), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;

            }
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
                        if (str缺省 != "" && chkControl.Items[i].Description!="" && (",," + str缺省 + ",").IndexOf(("," + chkControl.Items[i].Description + ",").ToString()) > -1)
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


        /// <summary>
        /// 通过DataTable加载CheckList
        /// </summary>
        /// <param name="chkControl"></param>
        /// <param name="dataTmp"></param>
        private void loadCheckList(CheckedListBoxControl chkControl, DataTable dataTmp)
        {

            try
            {
                //首先清空CheckList
                chkControl.Items.Clear();

                //

                //修改分级特殊处理
                if (chkControl.Name == "clb修改分级")
                {
                    chkControl.Items.Add("0", "不修改", CheckState.Checked, true);
                    chkControl.Tag = "不修改";
                    //dataTable循环添加
                    for (int i = 0; i < dataTmp.Rows.Count; i++)
                    {
                        chkControl.Items.Add(dataTmp.Rows[i]["序号"].ToString(), dataTmp.Rows[i]["名称"].ToString(), CheckState.Unchecked, true);

                        //记录最大的等级
                        if (ComLib.ComLib.StrToInt(dataTmp.Rows[i]["序号"].ToString()) > intMaxLevel)
                        {
                            intMaxLevel = ComLib.ComLib.StrToInt(dataTmp.Rows[i]["序号"].ToString());
                            strMaxLevel = dataTmp.Rows[i]["名称"].ToString();
                            strMaxLevelColor = dataTmp.Rows[i]["患者标识颜色"].ToString();
                        }

                    }
                }
                else
                {
                    //缓存编码
                    if (chkControl.Name == "clb陪同人员")
                    {
                        //dataTable循环添加
                        for (int i = 0; i < dataTmp.Rows.Count; i++)
                        {
                            if (dataTmp.Rows[i]["缺省"].ToString() == "1")
                            {
                                chkControl.Items.Add(dataTmp.Rows[i]["编码"].ToString(), dataTmp.Rows[i]["名称"].ToString(), CheckState.Checked, true);
                                chkControl.Tag = dataTmp.Rows[i]["名称"].ToString();
                            }
                            else
                            {
                                chkControl.Items.Add(dataTmp.Rows[i]["编码"].ToString(), dataTmp.Rows[i]["名称"].ToString(), CheckState.Unchecked, true);
                            }
                        }

                    }
                    else
                    {

                        //dataTable循环添加
                        for (int i = 0; i < dataTmp.Rows.Count; i++)
                        {
                            if (dataTmp.Rows[i]["缺省"].ToString() == "1")
                            {
                                chkControl.Items.Add(dataTmp.Rows[i]["名称"].ToString(), true);
                                chkControl.Tag = dataTmp.Rows[i]["名称"].ToString();
                            }
                            else
                            {
                                chkControl.Items.Add(dataTmp.Rows[i]["名称"].ToString(), false);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        /// <summary>
        /// 通过DataTable加载评分指标
        /// </summary>
        /// <param name="dataTmp"></param>
        private void load评分指标(DataTable dataTmp)
        {
            try
            {
                string[] strArray = null;
                //首先清空CheckList
                clb睁眼反应.Items.Clear();
                clb语言反应.Items.Clear();
                clb活动反应.Items.Clear();
                clb疼痛描述.Items.Clear();

                //dataTable循环添加
                for (int i = 0; i < dataTmp.Rows.Count; i++)
                {
                    switch (dataTmp.Rows[i]["指标名称"].ToString())
                    {
                        case "睁眼反应":
                            strArray = dataTmp.Rows[i]["值域范围"].ToString().Split('|');
                            foreach (string strValue in strArray) clb睁眼反应.Items.Add(strValue, false);
                            clb睁眼反应.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "语言反应":
                            strArray = dataTmp.Rows[i]["值域范围"].ToString().Split('|');
                            foreach (string strValue in strArray) clb语言反应.Items.Add(strValue, false);
                            clb语言反应.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "活动反应":
                            strArray = dataTmp.Rows[i]["值域范围"].ToString().Split('|');
                            foreach (string strValue in strArray) clb活动反应.Items.Add(strValue, false);
                            clb活动反应.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "疼痛描述":
                            strArray = dataTmp.Rows[i]["值域范围"].ToString().Split('|');
                            foreach (string strValue in strArray) clb疼痛描述.Items.Add(strValue, false);
                            clb疼痛描述.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "舒张压":
                            layItem舒张压.Text = dataTmp.Rows[i]["值域单位"].ToString();
                            layItem舒张压.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "心率":
                            lbl脉搏单位.Text = dataTmp.Rows[i]["值域单位"].ToString();
                            lbl脉搏单位.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "指氧饱和度":
                            lblSPO2单位.Text = dataTmp.Rows[i]["值域单位"].ToString();
                            lblSPO2单位.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "体温":
                            lbl体温单位.Text = dataTmp.Rows[i]["值域单位"].ToString();
                            lbl体温单位.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "血糖":
                            lbl血糖单位.Text = dataTmp.Rows[i]["值域单位"].ToString();
                            lbl血糖单位.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "血钾":
                            lbl血钾单位.Text = dataTmp.Rows[i]["值域单位"].ToString();
                            lbl血钾单位.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "收缩压":
                            layItem收缩压.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                        case "呼吸频率":
                            lbl呼吸单位.Text = dataTmp.Rows[i]["值域单位"].ToString();
                            lbl呼吸单位.Tag = dataTmp.Rows[i]["ID"].ToString();
                            break;
                    }

                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void Get自动评分()
        {
            try
            {
                if (blnLoad) return;
                int int疼痛评分 = 0;
                int int昏迷评分 = 0;
                int int客观评分 = 0;
                int intMin = intMaxLevel;

                layoutGroup疼痛数字评分.Text = "疼痛数字评分";
                layoutGroup疼痛数字评分.Tag = 0;
                layoutGroupGCS.Text = "格拉斯哥昏迷评分";
                layoutGroupGCS.Tag = 0;
                txt自动分级.Text = strMaxLevel;
                txt自动分级.Tag = intMaxLevel.ToString();
                txt自动分级.BackColor = ColorTranslator.FromHtml("#" + strMaxLevelColor);

                if (data疼痛评分 != null)
                {
                    int疼痛评分 = ComLib.ComLib.StrToInt(data疼痛评分.Rows[0]["病人等级"].ToString());
                    layoutGroup疼痛数字评分.Text = "疼痛数字评分             " + data疼痛评分.Rows[0]["描述"].ToString() + "：" + data疼痛评分.Rows[0]["病人分数"].ToString();
                    layoutGroup疼痛数字评分.Tag = ComLib.ComLib.StrToInt(data疼痛评分.Rows[0]["病人分数"].ToString());
                }

                if (data昏迷评分 != null)
                {
                    int昏迷评分 = ComLib.ComLib.StrToInt(data昏迷评分.Rows[0]["病人等级"].ToString());
                    layoutGroupGCS.Text = "格拉斯哥昏迷评分         " + data昏迷评分.Rows[0]["描述"].ToString() + "：" + data昏迷评分.Rows[0]["病人分数"].ToString();
                    layoutGroupGCS.Tag = ComLib.ComLib.StrToInt(data昏迷评分.Rows[0]["病人分数"].ToString());
                }

                if (data客观等级 != null)
                {
                    int客观评分 = ComLib.ComLib.StrToInt(data客观等级.Rows[0]["病人等级"].ToString());
                }

                if (int疼痛评分 != 0)
                {
                    intMin = intMin > int疼痛评分 ? int疼痛评分 : intMin;
                }
                if (int昏迷评分 != 0)
                {
                    intMin = intMin > int昏迷评分 ? int昏迷评分 : intMin;
                }
                if (int客观评分 != 0)
                {
                    intMin = intMin > int客观评分 ? int客观评分 : intMin;
                }

                if (intMin != 0 && intMin != intMaxLevel)
                {
                    foreach (DataRow dtRow in data等级.Rows)
                    {
                        if (ComLib.ComLib.StrToInt(dtRow["序号"].ToString()) == intMin)
                        {
                            txt自动分级.Text = dtRow["名称"].ToString();
                            txt自动分级.Tag = intMin;
                            txt自动分级.BackColor = ColorTranslator.FromHtml("#" + dtRow["患者标识颜色"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void Set昏迷评分()
        {
            try
            {
                if (blnLoad) return;
                string str睁眼反应 = "";
                string str语言反应 = "";
                string str活动反应 = "";

                if (clb语言反应.CheckedItems.Count > 0 && clb活动反应.CheckedItems.Count > 0 && clb睁眼反应.CheckedItems.Count > 0 && blnCheckNo == false)
                {
                    str睁眼反应 = clb睁眼反应.Tag.ToString() + ":" + clb睁眼反应.CheckedItems[0].ToString();
                    str语言反应 = clb语言反应.Tag.ToString() + ":" + clb语言反应.CheckedItems[0].ToString();
                    str活动反应 = clb活动反应.Tag.ToString() + ":" + clb活动反应.CheckedItems[0].ToString();
                    data昏迷评分 = objEditService.Get昏迷评分(str睁眼反应, str语言反应, str活动反应);
                }
                else
                {
                    data昏迷评分 = null;
                }
                if (blnCheckNo == false)
                {
                    Get自动评分();
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void Set客观评分()
        {
            try
            {
                if (blnLoad) return;
                string str年龄 = "";
                string str年龄单位 = "";
                string str指标信息 = "";//指标id：指标结果描述（可多个）

                //检查是否满足获取评分的要求
                if (txt年龄.Text.Trim() == "" || txt年龄.Text == "0" || txt年龄.Text == "0小时" || txt年龄.Text == "0天" || txt年龄.Text == "0月" || (txt体温.Text == "" && txt脉搏.Text == "" && txt收缩压.Text == "" && txt舒张压.Text == "" && txt呼吸.Text == "" && txt血氧.Text == "" && txt血糖.Text == "" && txt血钾.Text == ""))
                {
                    data客观等级 = null;
                    Get自动评分();
                    return;
                }

                //解析年龄
                Get年龄解析(txt年龄.Text.Trim() + (cbo年龄单位.SelectedIndex == -1 ? "" : cbo年龄单位.Text), out str年龄, out str年龄单位);

                //组合客观条件
                if (txt体温.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lbl体温单位.Tag.ToString() + ":" + txt体温.Text;
                }
                if (txt脉搏.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lbl脉搏单位.Tag.ToString() + ":" + txt脉搏.Text;
                }
                if (txt收缩压.Text != "")
                {
                    str指标信息 = str指标信息 + "," + layItem收缩压.Tag.ToString() + ":" + txt收缩压.Text;
                }
                if (txt舒张压.Text != "")
                {
                    str指标信息 = str指标信息 + "," + layItem舒张压.Tag.ToString() + ":" + txt舒张压.Text;
                }
                if (txt呼吸.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lbl呼吸单位.Tag.ToString() + ":" + txt呼吸.Text;
                }
                if (txt血氧.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lblSPO2单位.Tag.ToString() + ":" + txt血氧.Text;
                }
                if (txt血糖.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lbl血糖单位.Tag.ToString() + ":" + txt血糖.Text;
                }
                if (txt血钾.Text != "")
                {
                    str指标信息 = str指标信息 + "," + lbl血钾单位.Tag.ToString() + ":" + txt血钾.Text;
                }
                if (str指标信息 != "")
                {
                    str指标信息 = str指标信息.Substring(1);
                }
                data客观等级 = objEditService.Get客观评分(str年龄, str年龄单位, str指标信息);

                Get自动评分();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void Get年龄解析(string str年龄, out string str年龄值, out string str年龄单位)
        {
            str年龄值 = "";
            str年龄单位 = "";
            try
            {
                str年龄值 = "";
                str年龄单位 = "";
                //解析年龄
                if (str年龄值 == "" && str年龄.IndexOf("岁") > -1)
                {
                    str年龄值 = str年龄.Substring(0, str年龄.IndexOf("岁"));
                    str年龄单位 = "岁";
                }
                if (str年龄值 == "" && str年龄.IndexOf("月") > -1)
                {
                    str年龄值 = str年龄.Substring(0, str年龄.IndexOf("月"));
                    str年龄单位 = "月";
                }
                if (str年龄值 == "" && str年龄.IndexOf("天") > -1)
                {
                    str年龄值 = str年龄.Substring(0, str年龄.IndexOf("天"));
                    str年龄单位 = "天";
                }
                if (str年龄值 == "" && str年龄.IndexOf("小时") > -1)
                {
                    str年龄值 = str年龄.Substring(0, str年龄.IndexOf("小时"));
                    str年龄单位 = "小时";
                }
                if (str年龄值 == "" && str年龄.IndexOf("分钟") > -1)
                {
                    str年龄值 = str年龄.Substring(0, str年龄.IndexOf("分钟"));
                    str年龄单位 = "分钟";
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            return;
        }


        private void GetSelectValue(bool closePopup = true)
        {
            try
            {
                DataRow dtInfo = gv病人信息.GetFocusedDataRow();
                this.txt姓名.ClosePopup();
                setPatiInfo(dtInfo);
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        /// <summary>
        /// 设置病人信息
        /// </summary>
        /// <param name="dtRow"></param>
        private void setPatiInfo(DataRow dtRow)
        {
            try
            {
                bool bln自由录入 = false;
                if (dtRow == null) bln自由录入 = true;
                if (bln自由录入 == false) bln自由录入 = ComLib.ComLib.Nvl(dtRow, "排序id") == "0";

                if (bln自由录入)//自由录入
                {

                    cbo性别.ReadOnly = false;
                    cbo性别.BackColor = Color.White;
                    cbo性别.TabStop = true;

                    dt出生日期.ReadOnly = false;
                    dt出生日期.BackColor = Color.White;
                    dt出生日期.TabStop = true;

                    this.txt姓名.Tag = txt姓名.Text.Trim();

                    if (str病人id != "")
                    {
                        //下拉框初始化
                        if (cbo性别.Tag != null)
                        {
                            cbo性别.Text = cbo性别.Tag.ToString();
                        }
                        else
                        {
                            cbo性别.Text = "男";
                        }
                        if (cbo国籍.Tag != null)
                        {
                            cbo国籍.Text = cbo国籍.Tag.ToString();
                        }
                        else
                        {
                            cbo国籍.Text = "中国";
                        }




                        if (txt民族.Tag != null)
                        {
                            txt民族.Text = txt民族.Tag.ToString();
                        }
                        else
                        {
                            txt民族.Text = "汉族";
                        }

                        dt出生日期.Text = "";
                        txt身份证号.Text = "";

                        txt联系电话.Text = "";
                        txt医保卡号.Text = "";
                        txt保险类别.Text = "";
                        txt家庭地址.Text = "";


                        if (cbo年龄单位.Tag != null)
                        {
                            cbo年龄单位.Text = cbo年龄单位.Tag.ToString();
                        }
                        else
                        {
                            cbo年龄单位.Text = "岁";
                        }
                    }
                    str病人id = "";
                    str门诊号 = "";

                }
                else//选择病人
                {
                    cbo性别.ReadOnly = true;
                    dt出生日期.ReadOnly = ComLib.ComLib.Nvl(dtRow, "出生日期") != "";
                    cbo性别.BackColor = txt保险类别.BackColor;
                    dt出生日期.BackColor = !dt出生日期.ReadOnly ? Color.White : txt保险类别.BackColor;
                    cbo性别.TabStop = false;
                    dt出生日期.TabStop = !dt出生日期.ReadOnly;
                    str病人id = ComLib.ComLib.Nvl(dtRow, "病人id");
                    str门诊号 = ComLib.ComLib.Nvl(dtRow, "门诊号");
                    this.txt姓名.Text = ComLib.ComLib.Nvl(dtRow, "姓名");
                    this.txt姓名.Tag = ComLib.ComLib.Nvl(dtRow, "姓名");
                    cbo性别.Text = ComLib.ComLib.Nvl(dtRow, "性别");
                    cbo国籍.Text = ComLib.ComLib.Nvl(dtRow, "国籍") == "" ? cbo国籍.Tag.ToString() : ComLib.ComLib.Nvl(dtRow, "国籍");
                    dt出生日期.Text = ComLib.ComLib.Nvl(dtRow, "出生日期");
                    txt身份证号.Text = ComLib.ComLib.Nvl(dtRow, "身份证号");
                    txt联系电话.Text = ComLib.ComLib.Nvl(dtRow, "手机号");
                    txt民族.Text = ComLib.ComLib.Nvl(dtRow, "民族");
                    txt医保卡号.Text = ComLib.ComLib.Nvl(dtRow, "医保号");
                    txt保险类别.Text = ComLib.ComLib.Nvl(dtRow, "保险类别");
                    txt家庭地址.Text = ComLib.ComLib.Nvl(dtRow, "家庭地址");
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        /// <summary>
        /// 年龄相关处理
        /// </summary>
        private void ReCalcOld(string strAge)
        {
            try
            {
                string strTmp = "";
                if (strAge != "")
                {
                    string str单位 = strAge.Substring(strAge.Length - 1, 1);
                    if (str单位 == "岁" || str单位 == "月" || str单位 == "天")
                    {
                        strTmp = strAge.Substring(0, strAge.Length - 1);
                        if (ComLib.ComLib.StrToInt(strTmp) != 0)
                        {
                            cbo年龄单位.Text = str单位;
                        }
                        else
                        {
                            strTmp = strAge;
                            cbo年龄单位.SelectedIndex = -1;
                        }

                    }
                    else
                    {
                        strTmp = strAge;
                        if (ComLib.ComLib.StrToInt(strAge) != 0)
                        {
                            cbo年龄单位.SelectedIndex = 0;
                        }
                        else
                        {
                            cbo年龄单位.SelectedIndex = -1;
                        }
                    }
                }

                if (cbo年龄单位.SelectedIndex == -1)
                {

                    layItem年龄单位.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;//隐藏文本框

                }
                else
                {
                    layItem年龄单位.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;//显示文本框

                }
                txt年龄.Text = strTmp;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }



        /// <summary>
        /// 病人年龄检查

        /// </summary>
        /// <returns> 返回错误提示</returns>
        private string CheckAge()
        {
            try
            {
                if (txt年龄.Text.Trim() == "") return "";
                string strOut = "";

                //初步检查年龄输入值的有效性

                if (cbo年龄单位.SelectedIndex != -1)
                {
                    switch (cbo年龄单位.Text)
                    {
                        case "岁":
                            if (ComLib.ComLib.StrToInt(txt年龄.Text) > 200) strOut = "年龄不能大于200岁!";
                            break;
                        case "月":
                            if (ComLib.ComLib.StrToInt(txt年龄.Text) > 2400) strOut = "年龄不能大于2400月!";
                            break;
                        case "天":
                            if (ComLib.ComLib.StrToInt(txt年龄.Text) > 73000) strOut = "年龄不能大于73000天!";
                            break;
                    }
                    if (strOut != "") return strOut;
                }

                string strAge = txt年龄.Text.Trim() + (cbo年龄单位.SelectedIndex == -1 ? "" : cbo年龄单位.Text);

                //调用过程判断年龄
                string strCheck = objEditService.CheckAge(strAge);
                if (strCheck.Trim() != "")
                {
                    strCheck = strCheck.Trim().Split('|')[1];
                    return strCheck;
                }
                return "";
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return "";
            }
        }



        /// <summary>
        /// 通过年龄反算出生日期
        /// </summary>
        /// <returns> 返回错误提示</returns>
        private void AgeToDate()
        {
            try
            {
                DateTime dtTime = dtNow;
                if (txt年龄.Text.Trim() == "" || cbo年龄单位.SelectedIndex == -1 || dt出生日期.Text != "" || dt出生日期.ReadOnly == true) return;
                string strAge = txt年龄.Text.Trim() + (cbo年龄单位.SelectedIndex == -1 ? "" : cbo年龄单位.Text);
                blnDateChange = true;
                if (cbo年龄单位.Text == "岁")
                {
                    dt出生日期.Text = dtTime.AddYears(-(ComLib.ComLib.StrToInt(txt年龄.Text.Trim()))).ToString("yyyy-MM-dd HH:mm");
                }
                else if (cbo年龄单位.Text == "月")
                {
                    dt出生日期.Text = dtTime.AddMonths(-(ComLib.ComLib.StrToInt(txt年龄.Text.Trim()))).ToString("yyyy-MM-dd HH:mm");
                }
                else if (cbo年龄单位.Text == "天")
                {
                    dt出生日期.Text = dtTime.AddDays(-(ComLib.ComLib.StrToInt(txt年龄.Text.Trim()))).ToString("yyyy-MM-dd HH:mm");
                }
                blnDateChange = false;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }


        #region 控件处理


        /// <summary>
        /// 焦点跳转
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetNextControl(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                SendKeys.Send("{tab}");
            }
        }

        /// <summary>
        /// 文本框全选

        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetSelectAll(object sender, EventArgs e)
        {
            DevExpress.XtraEditors.TextEdit txt = sender as DevExpress.XtraEditors.TextEdit;
            txt.SelectAll();
        }

        /// <summary>
        /// 记录窗体控件是否发生改变
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void from_TextChanged(object sender, EventArgs e)
        {
            ISChange = true;
        }

        #region 单选处理


        /// <summary>
        /// 把CheckedListBoxControl设置为单选框
        /// </summary>
        /// <param name="chkControl">CheckedListBoxControl</param>
        /// <param name="index">index当前选中的索引</param>
        public void SingleSelectCheckedListBoxControls(CheckedListBoxControl chkControl, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            if (chkControl.ItemCount > 1 && ",clb睁眼反应,clb语言反应,clb活动反应,clb疼痛描述,".IndexOf(chkControl.Name) == -1)
            {
                if (chkControl.CheckedItems.Count == 1 && chkControl.Items[e.Index].CheckState == CheckState.Checked && blnCheckNo == false)
                {
                    e.Cancel = true;
                    return;
                }
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

        private void clb三无患者_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb三无患者, e);
        }

        private void clb复合伤_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb复合伤, e);
        }

        private void clb成批就诊_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb成批就诊, e);
        }

        private void clb病人来源_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb病人来源, e);
        }

        private void clb陪同人员_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb陪同人员, e);
        }

        private void clb意识形态_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb意识形态, e);
        }

        private void clb睁眼反应_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb睁眼反应, e);
        }

        private void clb语言反应_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb语言反应, e);
        }

        private void clb活动反应_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb活动反应, e);
        }

        private void clb疼痛描述_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb疼痛描述, e);
        }

        private void clb修改分级_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            SingleSelectCheckedListBoxControls(clb修改分级, e);
        }
        private void clb成批就诊_ItemCheck(object sender, DevExpress.XtraEditors.Controls.ItemCheckEventArgs e)
        {
            if (e.State != CheckState.Checked) return;
            txt人数登记.Enabled = e.State == CheckState.Checked && e.Index == 0;
            txt人数登记.Text = "";
        }


        /// <summary>
        /// 重绘CheckListBox颜色
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void draw(object sender, DevExpress.XtraEditors.ListBoxDrawItemEventArgs e)
        {
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                e.AllowDrawSkinBackground = false;
                e.Appearance.BackColor2 = (sender as CheckedListBoxControl).BackColor;
                e.Appearance.BorderColor = (sender as CheckedListBoxControl).BackColor;
                e.Appearance.BackColor = (sender as CheckedListBoxControl).BackColor;
            }
        }



        private void clb疼痛描述_ItemCheck(object sender, DevExpress.XtraEditors.Controls.ItemCheckEventArgs e)
        {
            if (e.State != CheckState.Checked && (sender as CheckedListBoxControl).CheckedItems.Count != 0) return;
            try
            {
                if (blnLoad) return;
                string str疼痛描述 = "";

                if (clb疼痛描述.CheckedItems.Count > 0 && blnCheckNo == false)
                {
                    str疼痛描述 = clb疼痛描述.CheckedItems[0].ToString();
                    data疼痛评分 = objEditService.Get疼痛评分(str疼痛描述);
                }
                else
                {
                    data疼痛评分 = null;
                }
                if (blnCheckNo == false)
                {
                    Get自动评分();
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }


        private void clb睁眼反应_ItemCheck(object sender, DevExpress.XtraEditors.Controls.ItemCheckEventArgs e)
        {
            if (e.State != CheckState.Checked && (sender as CheckedListBoxControl).CheckedItems.Count != 0) return;
            Set昏迷评分();
        }


        private void clb语言反应_ItemCheck(object sender, DevExpress.XtraEditors.Controls.ItemCheckEventArgs e)
        {
            if (e.State != CheckState.Checked && (sender as CheckedListBoxControl).CheckedItems.Count != 0) return;
            Set昏迷评分();
        }

        private void clb活动反应_ItemCheck(object sender, DevExpress.XtraEditors.Controls.ItemCheckEventArgs e)
        {
            if (e.State != CheckState.Checked && (sender as CheckedListBoxControl).CheckedItems.Count != 0) return;
            Set昏迷评分();
        }


        private void clb绿色通道_ItemChecking(object sender, DevExpress.XtraEditors.Controls.ItemCheckingEventArgs e)
        {
            if (blnCheckNo || blnLoad) return;
            string strMsg = "";
            if (e.Index == 0 && e.NewValue == CheckState.Checked)
            {
                strMsg = "你确定要对" + (txt姓名.Text == "" ? "当前病人" : "【" + txt姓名.Text + "】") + "标记绿色通道吗？" + System.Environment.NewLine + "(注：绿色通道病人将实行先诊疗后付费)";

                if (DevExpress.XtraEditors.XtraMessageBox.Show(strMsg, this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
            }
            SingleSelectCheckedListBoxControls(clb绿色通道, e);
        }


        #endregion

        #region 文本框处理


        private void txt血钾_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == 8)
            {
                e.Handled = false;
                return;
            }
            e.Handled = true;
        }


        private void txt主诉_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {

            try
            {
                string strOut = "";
                if (data主诉 == null)
                {
                    data主诉 = objEditService.Get常用主诉();
                }
                if (data主诉 != null)
                {
                    if (ShowForm.ChooseData(data主诉, txt主诉.Text, false, out strOut))
                    {
                        txt主诉.Text = strOut;
                    }
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

        private void txt人工评定_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {

            try
            {
                string strOut = "";
                if (data人工评估 == null)
                {
                    data人工评估 = objEditService.Get人工评估规则();
                }
                if (data人工评估 != null)
                {
                    data人工评估.DefaultView.RowFilter = "适用人群='" + (CheckIS儿童() ? "儿童" : "成人") + "'";
                    if (ShowForm.ChooseRuleData(data人工评估.DefaultView.ToTable(), txt人工评定.Text, false, out strOut))
                    {
                        txt人工评定.Text = strOut;
                        txt人工评定.Tag = txt人工评定.Text;
                        txt人工评定.ToolTipTitle = CheckIS儿童() ? "儿童" : "成人";
                        SetAuto人工评级(strOut, data人工评估.DefaultView.ToTable());
                    }
                    else
                    {
                        txt人工评定.Text = txt人工评定.Tag == null ? "" : txt人工评定.Tag.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private bool CheckIS儿童()
        {
            string strAge = txt年龄.Text.Trim() + (cbo年龄单位.SelectedIndex == -1 ? "" : cbo年龄单位.Text);
            if (strAge == "" || int最大儿童年龄 == 0)
            {
                return false;
            }
            if (strAge.IndexOf("岁") == -1)
            {
                return true;
            }
            if (ComLib.ComLib.StrToInt(strAge.Split('岁')[0].ToString()) > int最大儿童年龄)
            {
                return false;
            }
            else
            {
                return true;
            }
        }


        private void SetAuto人工评级(string str人工指标, DataTable data指标)
        {

            try
            {
                if (str人工指标 == "") return;
                if (data指标 == null) return;

                int intMinLevel = 0;
                foreach (DataRow dtRow in data指标.Rows)
                {
                    if ((";" + str人工指标 + ";").IndexOf(dtRow["指标名称"].ToString()) > -1)
                    {
                        intMinLevel = (ComLib.ComLib.StrToInt(dtRow["病情级别"].ToString()) < intMinLevel || intMinLevel == 0) ? ComLib.ComLib.StrToInt(dtRow["病情级别"].ToString()) : intMinLevel;
                    }
                }

                if (intMinLevel > 0)
                {
                    for (int i = 0; i < clb修改分级.Items.Count; i++)
                    {
                        if (clb修改分级.Items[i].Value.ToString() == intMinLevel.ToString())
                        {
                            clb修改分级.SetItemCheckState(i, System.Windows.Forms.CheckState.Checked);
                        }
                        else
                        {
                            clb修改分级.SetItemCheckState(i, System.Windows.Forms.CheckState.Unchecked);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }



        private void txt人工评定_KeyPress(object sender, KeyPressEventArgs e)
        {

            try
            {
                string strTag = txt人工评定.Tag == null ? "" : txt人工评定.Tag.ToString();
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (strTag != txt人工评定.Text && txt人工评定.Text != "")
                    {
                        string strOut = "";
                        if (data人工评估 == null)
                        {
                            data人工评估 = objEditService.Get人工评估规则();
                        }
                        if (data人工评估 != null)
                        {
                            data人工评估.DefaultView.RowFilter = "适用人群='" + (CheckIS儿童() ? "儿童" : "成人") + "'";

                            if (ShowForm.ChooseRuleData(data人工评估.DefaultView.ToTable(), txt人工评定.Text, true, out strOut))
                            {
                                txt人工评定.Text = strOut;
                                txt人工评定.Tag = txt人工评定.Text;
                                txt人工评定.ToolTipTitle = CheckIS儿童() ? "儿童" : "成人";
                                SetAuto人工评级(strOut, data人工评估.DefaultView.ToTable());
                            }
                            else
                            {
                                txt人工评定.Text = txt人工评定.Tag == null ? "" : txt人工评定.Tag.ToString();
                            }
                            e.Handled = false;
                            e.KeyChar = (char)Keys.Cancel;
                        }
                    }

                }


            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

        private void txt人工评定_Validating(object sender, CancelEventArgs e)
        {

            try
            {
                string strTag = txt人工评定.Tag == null ? "" : txt人工评定.Tag.ToString();
                if (strTag != txt人工评定.Text && txt人工评定.Text != "")
                {
                    txt人工评定.Text = strTag;
                }
                if (txt人工评定.Text == "")
                {
                    txt人工评定.Tag = "";
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }
        private void dt出生日期_EditValueChanged(object sender, EventArgs e)
        {
            if (blnLoad) return;
            if (blnDateChange) return;
            try
            {
                ReCalcOld(objEditService.GetDateToAge(dt出生日期.DateTime).Trim());
                if (txt人工评定.Text != "")
                {
                    if (txt人工评定.ToolTipTitle != (CheckIS儿童() ? "儿童" : "成人"))
                    {
                        txt人工评定.Text = "";
                        txt人工评定.Tag = "";
                        txt人工评定.ToolTipTitle = "";
                    }

                }
                Set客观评分();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txt联系电话_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == 8 || e.KeyChar == '(' || e.KeyChar == ')' || e.KeyChar == '-')
            {
                e.Handled = false;
                return;
            }
            e.Handled = true;
        }



        private void txt医保卡号_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Convert.ToChar(e.KeyChar.ToString().ToUpper());
            if (@"0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ~!@#$%^&*()_+|-=\[]{}<>,./".IndexOf(e.KeyChar) == -1 && e.KeyChar != 8 && e.KeyChar != 13)
            {
                e.Handled = true;
            }
        }

        private void txt人数登记_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == 8)
            {
                e.Handled = false;
                return;
            }
            e.Handled = true;
        }

        //GridLookUpEdit模糊查询
        private void FilterLookup(object sender, string[] arrFilterField)
        {
            try
            {
                //Text += " ! ";
                GridLookUpEdit edit = sender as GridLookUpEdit;
                GridView gridView = edit.Properties.View as GridView;
                FieldInfo fi = gridView.GetType().GetField("extraFilter", BindingFlags.NonPublic | BindingFlags.Instance);
                //Text = edit.AutoSearchText;
                CriteriaOperator[] arrCriteriaOperator = new CriteriaOperator[arrFilterField.Length];
                for (int i = 0; i < arrFilterField.Length; i++)
                {
                    arrCriteriaOperator[i] = new BinaryOperator(arrFilterField[i], "%" + edit.AutoSearchText + "%", BinaryOperatorType.Like);
                }
                string filterCondition = new GroupOperator(GroupOperatorType.Or, arrCriteriaOperator).ToString();
                fi.SetValue(gridView, filterCondition);
                MethodInfo mi = gridView.GetType().GetMethod("ApplyColumnsFilterEx", BindingFlags.NonPublic | BindingFlags.Instance);
                mi.Invoke(gridView, null);
            }
            catch { return; };
        }


        private void txt民族_TextChanged(object sender, EventArgs e)
        {
            GridLookUpEdit edit = sender as GridLookUpEdit;
            edit.SelectionStart = edit.Text.Length;
        }
        private void txt民族_EditValueChanging(object sender, DevExpress.XtraEditors.Controls.ChangingEventArgs e)
        {
            this.BeginInvoke(new System.Windows.Forms.MethodInvoker(delegate
            {
                FilterLookup(sender, new string[] { "名称", "简码" });
            }));
        }

        private void txt民族_Popup(object sender, EventArgs e)
        {
            FilterLookup(sender, new string[] { "名称", "简码" });//中文，编码，拼音
        }

        /// <summary>
        /// 简单支持double截取小数位数  不四舍五入

        /// </summary>
        /// <param name="dlbvalue"></param>
        /// <param name="num"></param>
        /// <returns></returns>
        private double dblSubstring(double dlbvalue, int num)
        {
            try
            {
                if (dlbvalue == 0) return 0;
                double tmp = 0;
                switch (num)
                {
                    case 1:
                        tmp = (int)(dlbvalue * 10) / 10.0;
                        break;
                    case 2:
                        tmp = (int)(dlbvalue * 100) / 100.00;
                        break;
                }
                return tmp;
            }
            catch
            {
                return 0;
            }
        }
        private void txt体温_Validating(object sender, CancelEventArgs e)
        {
            try
            {

                if ((ComLib.ComLib.StrToDlb(txt体温.Text.Trim()) < 33 || ComLib.ComLib.StrToDlb(txt体温.Text) > 43) && txt体温.Text != "")
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("输入病人体温未在指定范围内(33-43),请重新录入。", "急诊预检分诊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true;
                    return;
                }
                if (txt体温.Text != "0" && txt体温.Text != "")
                {
                    txt体温.Text = dblSubstring(ComLib.ComLib.StrToDlb(txt体温.Text), 1).ToString();
                }
                Set客观评分();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void txt脉搏_Validating(object sender, CancelEventArgs e)
        {

            try
            {
                if ((ComLib.ComLib.StrToDlb(txt脉搏.Text) < 1) && txt脉搏.Text != "")
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("输入病人心率未在指定范围内(大于1),请重新录入。", "急诊预检分诊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true;
                    return;
                }
                if (txt脉搏.Text != "0" && txt脉搏.Text != "")
                {
                    txt脉搏.Text = ComLib.ComLib.StrToDlb(txt脉搏.Text).ToString();
                }
                Set客观评分();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void txt呼吸_Validating(object sender, CancelEventArgs e)
        {

            try
            {
                if ((ComLib.ComLib.StrToDlb(txt呼吸.Text) < 1 || ComLib.ComLib.StrToDlb(txt呼吸.Text) > 100) && txt呼吸.Text != "")
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("输入病人呼吸未在指定范围内(1-100),请重新录入。", "急诊预检分诊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true;
                    return;
                }
                if (txt呼吸.Text != "0" && txt呼吸.Text != "")
                {
                    txt呼吸.Text = ComLib.ComLib.StrToDlb(txt呼吸.Text).ToString();
                }
                Set客观评分();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void txt收缩压_Validating(object sender, CancelEventArgs e)
        {

            try
            {
                if ((ComLib.ComLib.StrToDlb(txt收缩压.Text) < 0 || ComLib.ComLib.StrToDlb(txt收缩压.Text) > 300) && txt收缩压.Text != "")
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("输入病人收缩压未在指定范围内(0-300),请重新录入。", "急诊预检分诊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true;
                    return;
                }
                if (txt收缩压.Text != "0" && txt收缩压.Text != "")
                {
                    txt收缩压.Text = ComLib.ComLib.StrToDlb(txt收缩压.Text).ToString();
                }
                Set客观评分();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void txt舒张压_Validating(object sender, CancelEventArgs e)
        {

            try
            {
                if ((ComLib.ComLib.StrToDlb(txt舒张压.Text) < 0 || ComLib.ComLib.StrToDlb(txt舒张压.Text) > 300) && txt舒张压.Text != "")
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("输入病人舒张压未在指定范围内(0-300),请重新录入。", "急诊预检分诊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true;
                    return;
                }
                if (txt舒张压.Text != "0" && txt舒张压.Text != "")
                {
                    txt舒张压.Text = ComLib.ComLib.StrToDlb(txt舒张压.Text).ToString();
                }
                Set客观评分();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void txt血氧_Validating(object sender, CancelEventArgs e)
        {

            try
            {
                if ((ComLib.ComLib.StrToDlb(txt血氧.Text.Trim()) < 0 || ComLib.ComLib.StrToDlb(txt血氧.Text.Trim()) > 100) && txt血氧.Text != "")
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("输入病人指氧饱和度未在指定范围内(0-100),请重新录入。", "急诊预检分诊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true;
                    return;
                }
                if (txt血氧.Text.Trim() != "0" && txt血氧.Text.Trim() != "")
                {
                    txt血氧.Text = dblSubstring(ComLib.ComLib.StrToDlb(txt血氧.Text.Trim()), 1).ToString();
                }
                Set客观评分();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }


        private void txt体温_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == 8)
            {
                e.Handled = false;
                return;
            }
            e.Handled = true;
        }

        private void txt脉搏_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == 8)
            {
                e.Handled = false;
                return;
            }
            e.Handled = true;
        }

        private void txt呼吸_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == 8)
            {
                e.Handled = false;
                return;
            }
            e.Handled = true;
        }

        private void txt收缩压_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == 8)
            {
                e.Handled = false;
                return;
            }
            e.Handled = true;
        }

        private void txt舒张压_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == 8)
            {
                e.Handled = false;
                return;
            }
            e.Handled = true;
        }

        private void txt血氧_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == 8)
            {
                e.Handled = false;
                return;
            }
            e.Handled = true;
        }

        private void txt血糖_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == 8)
            {
                e.Handled = false;
                return;
            }
            e.Handled = true;
        }



        private void clb修改分级_ItemCheck(object sender, DevExpress.XtraEditors.Controls.ItemCheckEventArgs e)
        {
            if ((e.Index == 0 && txt修改说明.Enabled == true) || (e.Index != 0 && txt修改说明.Enabled == false))
            {
                txt修改说明.Enabled = e.State == CheckState.Checked && e.Index != 0;
                txt修改说明.Text = "";
            }
        }

        private void txt科室_EditValueChanging(object sender, DevExpress.XtraEditors.Controls.ChangingEventArgs e)
        {
            this.BeginInvoke(new System.Windows.Forms.MethodInvoker(delegate
            {
                FilterLookup(sender, new string[] { "名称", "简码" });
            }));
        }

        private void txt科室_Popup(object sender, EventArgs e)
        {
            FilterLookup(sender, new string[] { "名称", "简码" });//中文，编码，拼音
        }

        private void txt科室_TextChanged(object sender, EventArgs e)
        {
            GridLookUpEdit edit = sender as GridLookUpEdit;
            edit.SelectionStart = edit.Text.Length;
        }


        private void txt血糖_Validating(object sender, CancelEventArgs e)
        {
            if ((ComLib.ComLib.StrToDlb(txt血糖.Text.Trim()) < 0 || ComLib.ComLib.StrToDlb(txt血糖.Text.Trim()) > 999) && txt血糖.Text != "")
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("输入病人血糖未在指定范围内(0-999),请重新录入。", "急诊预检分诊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                e.Cancel = true;
                return;
            }
            if (txt血糖.Text.Trim() != "0" && txt血糖.Text.Trim() != "")
            {
                txt血糖.Text = dblSubstring(ComLib.ComLib.StrToDlb(txt血糖.Text.Trim()), 2).ToString();
            }
            Set客观评分();
        }

        private void txt血钾_Validating(object sender, CancelEventArgs e)
        {
            if ((ComLib.ComLib.StrToDlb(txt血钾.Text.Trim()) < 0 || ComLib.ComLib.StrToDlb(txt血钾.Text.Trim()) > 999) && txt血钾.Text != "")
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("输入病人血钾未在指定范围内(0-999),请重新录入。", "急诊预检分诊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                e.Cancel = true;
                return;
            }
            if (txt血钾.Text.Trim() != "0" && txt血钾.Text.Trim() != "")
            {
                txt血钾.Text = dblSubstring(ComLib.ComLib.StrToDlb(txt血钾.Text.Trim()), 2).ToString();
            }
            Set客观评分();
        }


        private void txt姓名_Properties_KeyPress(object sender, KeyPressEventArgs e)
        {
            string strTag = txt姓名.Tag == null ? "" : txt姓名.Tag.ToString();
            DataTable dtPati = null;
            if (e.KeyChar == 13 && txt姓名.Text.Trim() != "" && strTag != txt姓名.Text.Trim())
            {
                try
                {
                    dtPati = objEditService.GetPatlistbyName(txt姓名.Text);
                    if (dtPati == null)
                    {
                        if (CheckFindText(txt姓名.Text) || txt姓名.Text.Trim().Substring(0, 1) == "-")
                        {
                            DevExpress.XtraEditors.XtraMessageBox.Show("当前录入项未找到病人,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txt姓名.Text = "";
                            txt姓名.Tag = null;
                            e.Handled = false;
                            e.KeyChar = (char)Keys.Cancel;
                        }
                        return;
                    }
                    if (dtPati.Rows.Count == 1)
                    {
                        if (CheckFindText(txt姓名.Text) || txt姓名.Text.Trim().Substring(0, 1) == "-")
                        {
                            DevExpress.XtraEditors.XtraMessageBox.Show("当前录入项未找到病人,请重新录入！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txt姓名.Text = "";
                            txt姓名.Tag = null;
                            e.Handled = false;
                            e.KeyChar = (char)Keys.Cancel;
                        }
                        return;
                    }

                    grid病人信息.DataSource = dtPati;
                    e.Handled = false;
                    e.KeyChar = (char)Keys.Cancel;

                    txt姓名.ShowPopup();
                    this.grid病人信息.Focus();
                    canAcceptReturn = false;

                }
                catch (Exception ex)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
        }


        /// <summary>
        /// 设置一个标识，是否在GridView中可以接受回车键
        /// </summary>
        bool canAcceptReturn = false;
        private void grid病人信息_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (canAcceptReturn)
                    {
                        GetSelectValue();
                    }

                    canAcceptReturn = true;
                }

            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }


        private void grid病人信息_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            GetSelectValue();
        }


        private void txt姓名_Validated(object sender, EventArgs e)
        {
            try
            {
                if (txt姓名.Text.Trim() != "")
                {
                    if (CheckFindText(txt姓名.Text) || txt姓名.Text.Trim().Substring(0, 1) == "-")
                    {
                        txt姓名.Text = "";
                        txt姓名.Tag = null;
                    }
                }

                if (txt姓名.Tag == null || (txt姓名.Tag == null ? "" : txt姓名.Tag.ToString()) != txt姓名.Text.Trim())
                {
                    setPatiInfo(null);
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static bool CheckFindText(string Id)
        {
            try
            {
                Id = Id.Trim();
                if (Id == "" || Id.Length < 18)
                {
                    return false;
                }

                if (Id.Substring(0, 17) == ComLib.ComLib.StrToLng(Id.Substring(0, 17)).ToString())
                {
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }



        private void txt年龄_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (ComLib.ComLib.StrToInt(txt年龄.Text.Trim()) != 0 || txt年龄.Text.Trim() == "" || txt年龄.Text.Trim() == "0")
                {
                    if (cbo年龄单位.SelectedIndex == -1)
                    {
                        cbo年龄单位.SelectedIndex = 0;
                        layItem年龄单位.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;//显示文本框

                    }
                }
                else
                {
                    cbo年龄单位.SelectedIndex = -1;
                    layItem年龄单位.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;//隐藏文本框

                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt年龄_Validating(object sender, CancelEventArgs e)
        {
            if (blnLoad) return;
            try
            {
                if (txt年龄.Text.Trim() != "")
                {
                    string strCheck = CheckAge();
                    if (strCheck != "")
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show(strCheck, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        e.Cancel = true;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }


        private void cbo年龄单位_Validating(object sender, CancelEventArgs e)
        {
            if (blnLoad) return;
            try
            {
                if (txt年龄.Text != "")
                {
                    string strCheck = CheckAge();
                    if (strCheck != "")
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show(strCheck, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        e.Cancel = true;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        bool blnDateChange = false;

        private void txt年龄_Validated(object sender, EventArgs e)
        {
            if (blnLoad) return;
            try
            {
                AgeToDate();
                if (txt人工评定.Text != "")
                {
                    if (txt人工评定.ToolTipTitle != (CheckIS儿童() ? "儿童" : "成人"))
                    {
                        txt人工评定.Text = "";
                        txt人工评定.Tag = "";
                        txt人工评定.ToolTipTitle = "";
                    }
                }
                Set客观评分();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void cbo年龄单位_Validated(object sender, EventArgs e)
        {
            if (blnLoad) return;
            try
            {
                AgeToDate();
                if (txt人工评定.Text != "")
                {
                    if (txt人工评定.ToolTipTitle != (CheckIS儿童() ? "儿童" : "成人"))
                    {
                        txt人工评定.Text = "";
                        txt人工评定.Tag = "";
                        txt人工评定.ToolTipTitle = "";
                    }

                }
                Set客观评分();
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }



        #endregion

        #endregion

        private void txt身份证号_Validating(object sender, CancelEventArgs e)
        {
            if (txt身份证号.Text.Trim() == "") return;
            if (blnLoad) return;
            if (txt身份证号.Tag == null) return;
            if (cbo国籍.Text != "中国") return;



            bool blnFind = false;
            try
            {
                if (txt姓名.ReadOnly && EditType == 1)//读取医保卡没有返回身份证情况
                {
                    str病人id = txt医保卡号.ToolTipTitle;
                    str门诊号 = txt保险类别.ToolTipTitle;
                }


                string strOut = objEditService.CheckIDCard(txt身份证号.Text.Trim());

                if (strOut != "")
                {
                    var doc = new XmlDocument();
                    doc.LoadXml(strOut);

                    //首先读取是否报错
                    if (doc.SelectSingleNode("/OUTPUT/MSG").InnerText.Trim() != "")
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show(doc.SelectSingleNode("/OUTPUT/MSG").InnerText.Trim(), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        e.Cancel = true;
                        return;
                    }

                    //检查是否存在病人信息

                    DataTable dt身份证List = objEditService.GetPatlistbyIdCard(txt身份证号.Text.Trim());

                    if (dt身份证List != null)
                    {
                        if (dt身份证List.Rows.Count > 0)
                        {
                            if (txt姓名.ReadOnly && EditType == 1)//读取医保卡没有返回身份证情况
                            {
                                str病人id = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "病人id");
                                str门诊号 = ComLib.ComLib.Nvl(dt身份证List.Rows[0], "门诊号");
                            }
                            else
                            {
                                if (str病人id == "")
                                {
                                    if (DevExpress.XtraEditors.XtraMessageBox.Show("根据身份证号找到【" + ComLib.ComLib.Nvl(dt身份证List.Rows[0], "姓名") + "】的病人信息,是否填充病人信息？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                                    {
                                        e.Cancel = true;
                                        return;
                                    }
                                    blnFind = true;
                                    setPatiInfo(dt身份证List.Rows[0]);

                                    //更新功能状态

                                    foreach (System.Windows.Forms.Control item in layoutControl1.Controls)
                                    {
                                        if (item is DevExpress.XtraEditors.TextEdit)
                                        {
                                            if (",txt姓名,cbo性别,cbo国籍,dt出生日期,txt身份证号,txt民族,txt家庭地址,txt年龄,cbo年龄单位,".IndexOf(item.Name) > -1)
                                            {
                                                (item as DevExpress.XtraEditors.TextEdit).ReadOnly = true;
                                                (item as DevExpress.XtraEditors.TextEdit).TabStop = false;
                                                item.BackColor = txt保险类别.BackColor;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = 0; i <= dt身份证List.Rows.Count - 1; i++)
                                    {
                                        if (str病人id == ComLib.ComLib.Nvl(dt身份证List.Rows[i], "病人id").Trim() && cbo国籍.Text == (ComLib.ComLib.Nvl(dt身份证List.Rows[i], "国籍").Trim() == "" ? cbo国籍.Tag.ToString() : ComLib.ComLib.Nvl(dt身份证List.Rows[i], "国籍").Trim()))
                                        {
                                            blnFind = true;
                                        }
                                    }

                                    if (!blnFind)
                                    {
                                        DevExpress.XtraEditors.XtraMessageBox.Show("已存在身份证号为【" + txt身份证号.Text + "】的建档病人信息,请重新录入！" + System.Environment.NewLine + "(注：可以在姓名录入框录入身份证号查找病人信息)", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        e.Cancel = true;
                                        return;
                                    }
                                }
                            }
                        }
                    }

                    string str出生日期 = doc.SelectSingleNode("/OUTPUT/BIRTHDAY").InnerText.Trim();
                    string str性别 = doc.SelectSingleNode("/OUTPUT/SEX").InnerText.Trim();
                    string str年龄 = doc.SelectSingleNode("/OUTPUT/AGE").InnerText.Trim();

                    if (!blnFind)
                    {
                        if (str性别 != cbo性别.Text || str出生日期 != dt出生日期.Text || str年龄 != txt年龄.Text + (cbo年龄单位.SelectedIndex == -1 ? "" : cbo年龄单位.Text))
                        {
                            if (DevExpress.XtraEditors.XtraMessageBox.Show("身份证号码获取的病人信息和当前录入的病人信息不一致,是否继续？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                            {
                                e.Cancel = true;
                                return;
                            }
                        }
                    }

                    blnLoad = true;
                    //读取病人出生日期

                    if (dt出生日期.ReadOnly == false)
                    {

                        dt出生日期.Text = str出生日期;
                    }
                    //读取病人年龄
                    if (txt年龄.Text != str年龄 && txt年龄.ReadOnly == false)
                    {
                        txt年龄.Text = str年龄;
                        if (txt年龄.Text.Trim() != "")
                        {
                            ReCalcOld(txt年龄.Text);
                        }
                    }

                    //读取病人性别
                    if (cbo性别.ReadOnly == false)
                    {
                        cbo性别.Text = str性别;
                    }

                    blnLoad = false;
                }

            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void txt身份证号_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Convert.ToChar(e.KeyChar.ToString().ToUpper());
            if (@"1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ*".IndexOf(e.KeyChar) == -1 && e.KeyChar != 8 && e.KeyChar != 3 && e.KeyChar != 24 && e.KeyChar != 22)
            {
                e.Handled = true;
            }
        }

        private void txt身份证号_TextChanged(object sender, EventArgs e)
        {
            txt身份证号.Tag = "1";
        }

        private void txt身份证号_Validated(object sender, EventArgs e)
        {
            txt身份证号.Tag = null;
        }

        private void txt_Validated(object sender, EventArgs e)
        {
            (sender as TextEdit).Text = (sender as TextEdit).Text.Trim();
        }

        private void cbo国籍_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (blnLoad) return;

            txt身份证号.Tag = "1";
            txt身份证号.Focus();

        }
    }




}