using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ZLSOFT.HIS.ZyEdit;
using ZLSOFT.HIS.ZyEdit.Model;

namespace ZLSOFT.HIS.ZyEdit.From
{
    partial class frmBase : Form
    {

        DataTable mdtDrug = null;//草药目录记录集

        DataTable mdt方剂 = null;//方剂目录记录集

        DataTable mdtHISDrug = null;
        DataTable mdt疾病 = null;//疾病目录记录集

        DataTable mdt证型 = null;//证型目录记录集

        DataTable mdt方剂对照 = null;//方剂对照记录集

        DataTable mdt加症 = null;//临症加症记录集

        DataTable mdt治法 = null;//临症治法记录集

        DataTable mdt用药 = null;//临症用药记录集


        private BaseData m_baseData = new BaseData(); //基础数据
        public frmBase(string message)
        {
            try
            {
                JObject jObject = (JObject)JsonConvert.DeserializeObject(message);
                if (jObject == null || message.Equals("[]"))
                {
                    throw new Exception("传入参数格式错误，应传入JSON格式字符串。" + message);
                }
                m_baseData.System = jObject["来源系统"].ToString();
                DataCenter.TranPasswd = m_baseData.System.Equals("0");
                m_baseData.站点 = jObject["站点"].ToString();
                m_baseData.OperatorID = jObject["操作员ID"].ToString();
                m_baseData.OperatorName = jObject["操作员姓名"].ToString();
                m_baseData.UserName = jObject["用户名"].ToString();
                m_baseData.UserPassword = jObject["用户密码"].ToString();
                if(m_baseData.System.Equals("0")){ 
                    m_baseData.TNSNAME = jObject["TNSNAME"].ToString();
                    if (jObject["TNSNAME"].ToString() != "")
                    {
                        ComLib.strTnsName = jObject["TNSNAME"].ToString();
                    }
                }

                SQLHelper_Oracle.IntData(jObject["用户名"].ToString(), jObject["用户密码"].ToString(), jObject["TNSNAME"].ToString());
                //初始化Oracle连接
                SQLHelper_Oracle.OracleCnn = SQLHelper_Oracle.GetOdpConnection();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            InitializeComponent();
        }

        private void IntdgvList(DataGridView dgv, string[] CH = null)
        {
            try
            {

                ComLib.DoubleBuffered(dgv, true);

                //设置dgvList设置表头
                if (CH != null)
                {
                    //初始化行数

                    dgv.RowCount = 1;

                    if (CH[0] == "状态")
                    {
                        DataGridViewImageColumn dgvImageColumn = new DataGridViewImageColumn();
                        dgvImageColumn.HeaderText = "状态";
                        dgvImageColumn.ImageLayout = DataGridViewImageCellLayout.Normal;
                        dgvImageColumn.Name = "状态";
                        dgvImageColumn.Image = picNone.Image;

                        dgv.Columns.Add(dgvImageColumn);
                        dgv.Columns[dgv.ColumnCount - 1].Width = 30;
                    }

                    dgv.ColumnCount = CH.Count() + 1;
                    for (int i = 1; i < CH.Length; i++)
                    {
                        dgv.Columns[i + 1].Name = CH[i];
                    }
                }

                //DataGridView控件基础设置
                // 表格上下左右自适应
                dgv.AllowUserToAddRows = false;
                dgv.AllowUserToDeleteRows = false;
                dgv.ReadOnly = true;
                // 列手工排序

                dgv.AllowUserToOrderColumns = false;
                // 列头系统样式，设置为false，自定义才生效

                dgv.EnableHeadersVisualStyles = false;
                // 列头高度大小模式
                dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                // 列头高度大小
                dgv.ColumnHeadersHeight = 25;
                // 列头居中
                dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                dgv.ColumnHeadersDefaultCellStyle.Font = new Font("宋体", 9);

                dgv.RowsDefaultCellStyle.Font = new Font("宋体", 9);

                // 列头边框样式
                dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
                // 列头背景色

                dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 236, 239);
                // 列头前景色

                dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
                // 网格线颜色

                dgv.GridColor = Color.FromArgb(245, 245, 247);
                // 背景色

                dgv.BackgroundColor = ColorTranslator.FromHtml("#FFFFFF");
                // 行头边框样式
                dgv.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
                // 行头背景色

                dgv.RowHeadersDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#FFFFFF");
                // 行高（要在窗体初始化的地方InitializeComponent调用才生效）
                dgv.RowTemplate.Height = 25;
                // 单元格内容居中

                dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                // 单元格背景色
                dgv.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#FFFFFF");
                // 隔行背景色

                dgv.AlternatingRowsDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#FFFFFF");

                dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//设置为整行


                dgv.DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#66CCFF");//设置背景色

                dgv.DefaultCellStyle.SelectionForeColor = Color.Black;//设置前景色


                dgv.MultiSelect = false;

                dgv.RowHeadersVisible = false;
                //选中第一行

                if (dgv.Rows.Count > 0)
                {
                    dgv.Rows[0].Selected = true;
                    //dgv.CurrentCell = dgv.Rows[0].Cells[2];
                }

                //设置隐藏列

                if (dgv.Columns[""] != null)
                {
                    dgv.Columns[""].Visible = false;
                }
                if (dgv.Columns["ID"] != null)
                {
                    dgv.Columns["ID"].Visible = false;
                }
                if (dgv.Columns["简码"] != null)
                {
                    dgv.Columns["简码"].Visible = false;
                }
                if (dgv.Columns["别名简码"] != null)
                {
                    dgv.Columns["别名简码"].Visible = false;
                }

                //列宽适应
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void frmBase_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = "中医基础数据维护(当前用户：" + m_baseData.OperatorName + ")";

                //初始化用药详情

                string[] CHDINFO = { "草药详情" };
                IntdgvList(dgvDrugInfo, CHDINFO);
                dgvDrugInfo.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvDrugInfo.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dgvDrugInfo.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                dgvDrugInfo.ColumnHeadersVisible = false;
                dgvDrugInfo.GridColor = Color.White;
                dgvDrugInfo.DefaultCellStyle.SelectionBackColor = Color.White;
                dgvDrugInfo.RowCount = 27;
                dgvDrugInfo.Rows[0].Cells[0].Value = "草药描述：";
                DataGridViewCellStyle style = dgvDrugInfo.Rows[0].Cells[0].Style;
                style.Font = new Font(dgvDrugInfo.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgvDrugInfo.Rows[3].Cells[0].Value = "草药性状：";
                style = dgvDrugInfo.Rows[3].Cells[0].Style;
                style.Font = new Font(dgvDrugInfo.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgvDrugInfo.Rows[6].Cells[0].Value = "草药药性：";
                style = dgvDrugInfo.Rows[6].Cells[0].Style;
                style.Font = new Font(dgvDrugInfo.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgvDrugInfo.Rows[9].Cells[0].Value = "适应证：";
                style = dgvDrugInfo.Rows[9].Cells[0].Style;
                style.Font = new Font(dgvDrugInfo.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgvDrugInfo.Rows[12].Cells[0].Value = "草药用法：";
                style = dgvDrugInfo.Rows[12].Cells[0].Style;
                style.Font = new Font(dgvDrugInfo.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgvDrugInfo.Rows[15].Cells[0].Value = "草药服法：";
                style = dgvDrugInfo.Rows[15].Cells[0].Style;
                style.Font = new Font(dgvDrugInfo.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgvDrugInfo.Rows[18].Cells[0].Value = "草药禁忌：";
                style = dgvDrugInfo.Rows[18].Cells[0].Style;
                style.Font = new Font(dgvDrugInfo.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgvDrugInfo.Rows[21].Cells[0].Value = "草药成分：";
                style = dgvDrugInfo.Rows[21].Cells[0].Style;
                style.Font = new Font(dgvDrugInfo.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgvDrugInfo.Rows[24].Cells[0].Value = "药理作用：";
                style = dgvDrugInfo.Rows[24].Cells[0].Style;
                style.Font = new Font(dgvDrugInfo.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                //初始化方剂详情

                string[] FJINFO = { "方剂详情" };
                IntdgvList(dgv方剂Info, FJINFO);
                dgv方剂Info.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv方剂Info.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dgv方剂Info.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                dgv方剂Info.ColumnHeadersVisible = false;
                dgv方剂Info.GridColor = Color.White;
                dgv方剂Info.DefaultCellStyle.SelectionBackColor = Color.White;
                dgv方剂Info.RowCount = 19;
                dgv方剂Info.Rows[0].Cells[0].Value = "组成摘要：";
                style = dgv方剂Info.Rows[0].Cells[0].Style;
                style.Font = new Font(dgv方剂Info.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgv方剂Info.Rows[3].Cells[0].Value = "服法描述：";
                style = dgv方剂Info.Rows[3].Cells[0].Style;
                style.Font = new Font(dgv方剂Info.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgv方剂Info.Rows[6].Cells[0].Value = "作用描述：";
                style = dgv方剂Info.Rows[6].Cells[0].Style;
                style.Font = new Font(dgv方剂Info.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgv方剂Info.Rows[9].Cells[0].Value = "制法描述：";
                style = dgv方剂Info.Rows[9].Cells[0].Style;
                style.Font = new Font(dgv方剂Info.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgv方剂Info.Rows[12].Cells[0].Value = "适应证描述：";
                style = dgv方剂Info.Rows[12].Cells[0].Style;
                style.Font = new Font(dgv方剂Info.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;

                dgv方剂Info.Rows[15].Cells[0].Value = "方剂组成作用描述：";
                style = dgv方剂Info.Rows[15].Cells[0].Style;
                style.Font = new Font(dgv方剂Info.Font, FontStyle.Bold);
                style.ForeColor = SystemColors.ActiveCaption;


                LoadDrug();

                //增加表格下拉项

                dgvDrug.Controls.Add(txt品种);

                //初始化方剂对应列表

                IntdgvList(dgv方剂对应);
                dgv方剂对应.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv方剂对应.Columns["方剂名称"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgv方剂对应.Columns["组成摘要"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgv方剂对应.Columns["状态"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                //初始化临症加症列表

                string[] CH = { "状态", "加症名称", "简码", "ID" };
                IntdgvList(dgv加症, CH);
                dgv加症.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv加症.Columns["简码"].Visible = true;
                dgv加症.Columns["状态"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgv加症.Columns["加症名称"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                //初始化临症治法列表

                string[] CHC = { "状态", "治法名称", "简码", "ID", "加症ID" };
                IntdgvList(dgv加症治法, CHC);
                dgv加症治法.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv加症治法.Columns["简码"].Visible = true;
                dgv加症治法.Columns["加症ID"].Visible = false;
                dgv加症治法.Columns["状态"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgv加症治法.Columns["治法名称"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;


                //初始化用药列表

                string[] CHY = { "状态", "草药名称", "简码", "用量", "单位", "草药ID", "ID", "治法ID" };
                IntdgvList(dgv加症用药, CHY);
                dgv加症用药.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv加症用药.Columns["简码"].Visible = true;
                dgv加症用药.Columns["治法ID"].Visible = false;
                dgv加症用药.Columns["草药ID"].Visible = false;
                dgv加症用药.Columns["状态"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgv加症用药.Columns["草药名称"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgv加症用药.Columns["单位"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgv加症用药.Columns["用量"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            EditDrug(0);
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            EditDrug(1);
        }
        private void EditZx(int intType)
        {
            try
            {

                if (intType == 1)
                {
                    DataRow dataRow = null;//返回dataRow
                    if (dgv证型.Rows.Count > 0 && dgv证型.SelectedRows.Count > 0)
                    {
                        if (dgv证型.SelectedRows[0].Cells["ID"].Value != null)
                        {
                            dataRow = (dgv证型.Rows[dgv证型.SelectedRows[0].Index < 0 ? 0 : dgv证型.SelectedRows[0].Index].DataBoundItem as DataRowView).Row;
                            frmZxEdit frm = new frmZxEdit(1, dataRow, m_baseData, dgv疾病.SelectedRows[0].Cells["ID"].Value.ToString());
                            frm.Text = "修改中医证型";
                            frm.ShowDialog();
                            if (frm.gblnOK == true)
                            {
                                Load证型();
                                FindItem(dgv证型, "", frm.gstrID);
                            }
                        }
                    }
                }
                else
                {
                    if (dgv疾病.SelectedRows.Count > 0)
                    {
                        frmZxEdit frm = new frmZxEdit(0, null, m_baseData, dgv疾病.SelectedRows[0].Cells["ID"].Value.ToString());
                        frm.Text = "新增中医证型";
                        frm.ShowDialog();
                        if (frm.gblnOK == true)
                        {
                            Load证型();
                            FindItem(dgv证型, "", frm.gstrID);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void EditDiag(int intType)
        {
            try
            {

                if (mdt疾病 == null)
                {
                    mdt疾病 = DataCenter.GetTable(m_baseData, "Get_Distype", "");
                }
                if (intType == 1)
                {
                    DataRow dataRow = null;//返回dataRow
                    if (dgv疾病.Rows.Count > 0 && dgv疾病.SelectedRows.Count > 0)
                    {
                        if (dgv疾病.SelectedRows[0].Cells["ID"].Value != null)
                        {
                            dataRow = (dgv疾病.Rows[dgv疾病.SelectedRows[0].Index < 0 ? 0 : dgv疾病.SelectedRows[0].Index].DataBoundItem as DataRowView).Row;
                            frmDiagEdit frm = new frmDiagEdit(1, dataRow, mdt疾病, m_baseData);
                            frm.Text = "修改中医疾病";
                            frm.ShowDialog();
                            if (frm.gblnOK == true)
                            {
                                Update疾病(intType, frm.gstrID, frm.gstrOut);
                                FindItem(dgv疾病, "", frm.gstrID);
                            }
                        }
                    }
                }
                else
                {
                    frmDiagEdit frm = new frmDiagEdit(0, null, mdt疾病, m_baseData);
                    frm.Text = "新增中医疾病";
                    frm.ShowDialog();
                    if (frm.gblnOK == true)
                    {
                        Update疾病(intType, frm.gstrID, frm.gstrOut);
                        FindItem(dgv疾病, "", frm.gstrID);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Edit加症(int intType)
        {
            try
            {
                if (intType == 1)
                {
                    DataGridViewRow dataRow = null;//返回dataRow
                    if (dgv加症.Rows.Count > 0 && dgv加症.SelectedRows.Count > 0)
                    {
                        if (dgv加症.SelectedRows[0].Cells["ID"].Value != null)
                        {
                            dataRow = dgv加症.SelectedRows[0];
                            frmAddDisEdit frm = new frmAddDisEdit(0, 1, dataRow, m_baseData);
                            frm.Text = "修改临证加症";
                            frm.ShowDialog();
                            if (frm.gblnOK == true)
                            {
                                Update加症(intType, frm.gstrID, frm.gstrOut);
                                FindItem(dgv加症, "", frm.gstrID);
                            }
                        }
                    }
                }
                else
                {
                    frmAddDisEdit frm = new frmAddDisEdit(0, 0, null, m_baseData);
                    frm.Text = "新增临证加症";
                    frm.ShowDialog();
                    if (frm.gblnOK == true)
                    {
                        Update加症(intType, frm.gstrID, frm.gstrOut);
                        FindItem(dgv加症, "", frm.gstrID);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Edit治法(int intType)
        {
            try
            {

                if (intType == 1)
                {
                    DataGridViewRow dataRow = null;//返回dataRow
                    if (dgv加症治法.Rows.Count > 0 && dgv加症治法.SelectedRows.Count > 0)
                    {
                        if (dgv加症治法.SelectedRows[0].Cells["ID"].Value != null)
                        {
                            dataRow = dgv加症治法.SelectedRows[0];
                            frmAddDisEdit frm = new frmAddDisEdit(1, 1, dataRow, m_baseData);
                            frm.Text = "修改加症治法";
                            frm.ShowDialog();
                            if (frm.gblnOK == true)
                            {
                                Load治法();
                                FindItem(dgv加症治法, "", frm.gstrID);
                            }
                        }
                    }
                }
                else
                {
                    if (dgv加症.SelectedRows.Count > 0 && dgv加症.SelectedRows[0].Cells["ID"].Value != null)
                    {
                        frmAddDisEdit frm = new frmAddDisEdit(1, 0, null, m_baseData, dgv加症.SelectedRows[0].Cells["ID"].Value.ToString());
                        frm.Text = "新增加症治法";
                        frm.ShowDialog();
                        if (frm.gblnOK == true)
                        {
                            Load治法();
                            FindItem(dgv加症治法, "", frm.gstrID);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Change状态(DataGridView dgv)
        {
            string strIDName = string.Empty;
            string strExName = string.Empty;
            string message = string.Empty;

            try
            {
                if (dgv.SelectedRows.Count > 0 && dgv.SelectedRows[0].Cells["ID"].Value != null)
                {
                    switch (dgv.Name)
                    {
                        case "dgv加症":
                            strIDName = "加症ID";
                            strExName = "Set_Adddis";
                            //保存
                            message = JsonConvert.SerializeObject(new
                            {
                                USETYPE = "2",////使用类型(0-新增/1-修改/3-更换状态)
                                加症ID = dgv.SelectedRows[0].Cells["ID"].Value.ToString(),
                                状态 = dgv.SelectedRows[0].Cells["状态"].Tag.ToString() == "1" ? "0" : "1",
                                操作员名称 = m_baseData.OperatorName,
                                操作员ID = m_baseData.OperatorID
                            });
                            break;
                        case "dgv加症治法":
                            strIDName = "治法ID";
                            strExName = "Set_Addzf";
                            //保存
                            message = JsonConvert.SerializeObject(new
                            {
                                USETYPE = "2",////使用类型(0-新增/1-修改/3-更换状态)
                                治法ID = dgv.SelectedRows[0].Cells["ID"].Value.ToString(),
                                状态 = dgv.SelectedRows[0].Cells["状态"].Tag.ToString() == "1" ? "0" : "1",
                                操作员名称 = m_baseData.OperatorName,
                                操作员ID = m_baseData.OperatorID
                            });
                            break;
                        case "dgv加症用药":
                            strIDName = "用药ID";
                            strExName = "Set_Zftozy";
                            //保存
                            message = JsonConvert.SerializeObject(new
                            {
                                USETYPE = "2",////使用类型(0-新增/1-修改/2-更换状态)
                                用药ID = dgv.SelectedRows[0].Cells["ID"].Value.ToString(),
                                状态 = dgv.SelectedRows[0].Cells["状态"].Tag.ToString() == "1" ? "0" : "1",
                                操作员名称 = m_baseData.OperatorName,
                                操作员ID = m_baseData.OperatorID
                            });
                            break;
                    }


                    DataTable dtOut = new DataTable();

                    dtOut = DataCenter.ExecuteFunc(strExName, message, m_baseData.UserName, m_baseData.UserPassword);
                    if (dtOut != null)
                    {

                        switch (dgv.Name)
                        {
                            case "dgv加症":
                                Load加症();
                                break;
                            case "dgv加症治法":
                                Load治法();
                                break;
                            case "dgv加症用药":
                                Load用药();
                                break;
                        }
                        FindItem(dgv, "", dtOut.Rows[0][strIDName].ToString());
                    }
                    else
                    {
                        throw new Exception("获取保存数据失败，服务器返回数据为空！");
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void DelItem(DataGridView dgv)
        {
            int intType = 0;
            string strID = string.Empty;
            try
            {
                if (dgv.SelectedRows.Count > 0 && dgv.SelectedRows[0].Cells["ID"].Value != null)
                {
                    strID = dgv.SelectedRows[0].Cells["ID"].Value.ToString();
                    if (strID == string.Empty)
                    {
                        return;
                    }
                    if (MessageBox.Show("请确认是否删除当前选择项目？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                    {
                        return;
                    }

                    switch (dgv.Name)
                    {
                        case "dgvDrug":
                            intType = 0;
                            break;
                        case "dgv方剂List":
                            intType = 4;
                            break;
                        case "dgv疾病":
                            intType = 1;
                            break;
                        case "dgv证型":
                            intType = 2;
                            break;
                        case "dgv方剂对应":
                            intType = 3;
                            break;
                        case "dgv加症":
                            intType = 5;
                            break;
                        case "dgv加症治法":
                            intType = 6;
                            break;
                        case "dgv加症用药":
                            intType = 7;
                            break;
                    }

                    string message = JsonConvert.SerializeObject(new
                    {
                        USETYPE = intType.ToString(),////使用类型(0, '草药目录', 1, '中医疾病', 2, '中医证型', 3, '证型方剂对照', 4, '治法方剂', 5, '临证加症', 6, '加症治法', 7, '加症用药')
                        ID = strID
                    });
                    DataTable dtOut = new DataTable();

                    dtOut = DataCenter.ExecuteFunc("Del_Zydata", message, m_baseData.UserName, m_baseData.UserPassword);
                    if (dtOut != null)
                    {
                        //本地同步
                        dgv.Rows.RemoveAt(dgv.SelectedRows[0].Index);
                        switch (dgv.Name)
                        {
                            case "dgvDrug":
                                mdtDrug.AcceptChanges();
                                break;
                            case "dgv方剂List":
                                mdt方剂.AcceptChanges();
                                break;
                            case "dgv疾病":
                                mdt疾病.AcceptChanges();
                                break;
                            case "dgv证型":
                                mdt证型.AcceptChanges();
                                break;
                            case "dgv方剂对应":
                                mdt方剂对照.AcceptChanges();
                                break;
                            case "dgv加症":
                                mdt加症.AcceptChanges();
                                break;
                            case "dgv加症治法":
                                mdt治法.AcceptChanges();
                                break;
                            case "dgv加症用药":
                                mdt用药.AcceptChanges();
                                break;
                        }
                    }
                    else
                    {
                        throw new Exception("获取保存数据失败，服务器返回数据为空！");
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ComLib.GetMessageFromOracleError(ex.Message), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Edit方剂对应(int intType)
        {
            DataRow drTmp = null;
            try
            {
                //新增对应
                if (intType == 0)
                {
                    if (dgv证型.SelectedRows.Count > 0 && txt方剂.Tag != null)
                    {
                        drTmp = (DataRow)txt方剂.Tag;
                        //检查是否有重复的方剂

                        foreach (DataGridViewRow row in this.dgv方剂对应.Rows)
                        {
                            if (row.Cells["ID"].Value != null)
                            {
                                if (drTmp["ID"].ToString() == row.Cells["方剂名称"].Tag.ToString())
                                {
                                    MessageBox.Show(@"方剂 """ + drTmp["方剂名称"].ToString() + @""" 已经录入，不能添加相同的方剂。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    return;
                                }
                            }
                        }

                        //保存
                        string message = JsonConvert.SerializeObject(new
                        {
                            USETYPE = intType.ToString(),////使用类型(0-新增/1-修改)
                            证型ID = dgv证型.SelectedRows[0].Cells["ID"].Value.ToString(),
                            方剂ID = drTmp["ID"].ToString(),
                            操作员名称 = m_baseData.OperatorName,
                            操作员ID = m_baseData.OperatorID
                        });

                        DataTable dtOut = new DataTable();

                        dtOut = DataCenter.ExecuteFunc("Set_Zxtofj", message, m_baseData.UserName, m_baseData.UserPassword);
                        if (dtOut != null)
                        {
                            Load方剂对应();
                            FindItem(dgv方剂对应, "", dtOut.Rows[0]["对照ID"].ToString());
                        }
                        else
                        {
                            throw new Exception("获取保存数据失败，服务器返回数据为空！");
                        }
                    }
                }
                else if (intType == 1)//更改对应状态

                {
                    if (dgv方剂对应.SelectedRows.Count > 0 && dgv方剂对应.SelectedRows[0].Cells["ID"].Value != null)
                    {

                        //保存
                        string message = JsonConvert.SerializeObject(new
                        {
                            USETYPE = intType.ToString(),////使用类型(0-新增/1-修改)
                            对照ID = dgv方剂对应.SelectedRows[0].Cells["ID"].Value.ToString(),
                            状态 = btn方剂状态.Text == "启用" ? "1" : "0",
                            操作员名称 = m_baseData.OperatorName,
                            操作员ID = m_baseData.OperatorID
                        });

                        DataTable dtOut = new DataTable();

                        dtOut = DataCenter.ExecuteFunc("Set_Zxtofj", message, m_baseData.UserName, m_baseData.UserPassword);
                        if (dtOut != null)
                        {
                            Load方剂对应();
                            FindItem(dgv方剂对应, "", dtOut.Rows[0]["对照ID"].ToString());
                        }
                        else
                        {
                            throw new Exception("获取保存数据失败，服务器返回数据为空！");
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ComLib.GetMessageFromOracleError(ex.Message), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void EditDrug(int intType)
        {
            try
            {
                if (mdtHISDrug == null)
                {
                    mdtHISDrug = DataCenter.GetTable(m_baseData, "Get_Hisdrug", "");
                }
                if (intType == 1)
                {
                    DataRow dataRow = null;//返回dataRow
                    if (dgvDrug.Rows.Count > 0 && dgvDrug.SelectedRows.Count > 0)
                    {
                        if (dgvDrug.SelectedRows[0].Cells["ID"].Value != null)
                        {
                            dataRow = (dgvDrug.Rows[dgvDrug.SelectedRows[0].Index < 0 ? 0 : dgvDrug.SelectedRows[0].Index].DataBoundItem as DataRowView).Row ;
                            frmDrugEdit frm = new frmDrugEdit(1, dataRow, mdtHISDrug, m_baseData);
                            frm.Text = "修改草药信息";
                            frm.ShowDialog();
                            if (frm.gblnOK == true)
                            {
                                UpdateDrug(intType, frm.gstrID, frm.gstrOut);
                                FindItem(dgvDrug, "", frm.gstrID);
                                LoadDrugInfo();
                            }
                        }
                    }
                }
                else
                {
                    frmDrugEdit frm = new frmDrugEdit(0, null, mdtHISDrug, m_baseData);
                    frm.Text = "新增草药信息";
                    frm.ShowDialog();
                    if (frm.gblnOK == true)
                    {
                        UpdateDrug(intType, frm.gstrID, frm.gstrOut);
                        FindItem(dgvDrug, "", frm.gstrID);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //药品界面更新
        //intType 0=新增 1=修改
        private void UpdateDrug(int intType, string strID, string strOut)
        {
            string strNow = string.Empty;
            DataGridViewRow dgvRow = null;
            DataRow dtRow = mdtDrug.NewRow();
            try
            {
                if (strOut == string.Empty || strID == string.Empty)
                {
                    throw new Exception("草药数据本地同步失败，请刷新！");
                }
                if (intType == 1 && dgvDrug.Rows.Count > 0 && dgvDrug.SelectedRows.Count > 0)
                {
                    dgvRow = dgvDrug.SelectedRows[0];
                }
                else
                {
                    dtRow["id"] = strID;
                    ((DataTable)dgvDrug.DataSource).Rows.Add(dtRow);
                    //获取对应dataRow
                    foreach (DataGridViewRow row in this.dgvDrug.Rows)
                    {
                        if (row.Cells["ID"].Value.ToString() == strID)
                        {
                            dgvRow = row;
                            break;
                        }
                    }
                }
                if (dgvRow != null)
                {

                    strNow = DataCenter.Get_Now_Time(m_baseData).ToString("yyyy-MM-dd HH:mm");
                    JObject jObject = (JObject)JsonConvert.DeserializeObject(strOut);
                    if (jObject == null || strOut.Equals("[]"))
                    {
                        throw new Exception("传入参数格式错误，应传入JSON格式字符串。" + strOut);
                    }
                    dgvRow.Cells["ID"].Value = strID;
                    dgvRow.Cells["草药名称"].Value = jObject["草药名称"].ToString();
                    dgvRow.Cells["简码"].Value = jObject["简码"].ToString();
                    dgvRow.Cells["别名"].Value = jObject["别名"].ToString();
                    dgvRow.Cells["别名简码"].Value = jObject["别名简码"].ToString();
                    dgvRow.Cells["单位"].Value = jObject["单位"].ToString();
                    dgvRow.Cells["来源"].Value = jObject["来源"].ToString();
                    dgvRow.Cells["草药描述"].Value = jObject["草药描述"].ToString();
                    dgvRow.Cells["性状"].Value = jObject["性状"].ToString();
                    dgvRow.Cells["药性"].Value = jObject["药性"].ToString();
                    dgvRow.Cells["适应证"].Value = jObject["适应证"].ToString();
                    dgvRow.Cells["用法"].Value = jObject["用法"].ToString();
                    dgvRow.Cells["服法"].Value = jObject["服法"].ToString();
                    dgvRow.Cells["禁忌"].Value = jObject["禁忌"].ToString();
                    dgvRow.Cells["成分"].Value = jObject["成分"].ToString();
                    dgvRow.Cells["药理作用"].Value = jObject["药理作用"].ToString();
                    if (jObject["HIS品种ID"].ToString() != "")
                    {
                        dgvRow.Cells["HIS品种ID"].Value = ComLib.StrToInt(jObject["HIS品种ID"].ToString());
                        if (mdtHISDrug != null)
                        {
                            mdtHISDrug.DefaultView.RowFilter = "His品种id =" + jObject["HIS品种ID"].ToString();
                            if (mdtHISDrug.DefaultView.ToTable().Rows.Count != 0)
                            {
                                dgvRow.Cells["His对码"].Value = "[" + jObject["HIS品种ID"].ToString() + "]" + mdtHISDrug.DefaultView.ToTable().Rows[0]["名称"].ToString();
                            }
                            mdtHISDrug.DefaultView.RowFilter = null;
                        }
                    }
                    dgvRow.Cells["最后修改时间"].Value = strNow;
                    dgvRow.Cells["最后修改人"].Value = m_baseData.OperatorName;
                    if (intType == 0)
                    {
                        dgvRow.Cells["创建时间"].Value = strNow;
                        dgvRow.Cells["创建人"].Value = m_baseData.OperatorName;
                    }
                    dgvRow.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //方剂界面更新
        //intType 0=新增 1=修改
        private void Update方剂(int intType, string strID, string strOut)
        {
            string strNow = string.Empty;
            DataGridViewRow dgvRow = null;
            DataRow dtRow = mdt方剂.NewRow();
            try
            {
                if (strOut == string.Empty || strID == string.Empty)
                {
                    throw new Exception("方剂数据本地同步失败，请刷新！");
                }
                if (intType == 1 && dgv方剂List.Rows.Count > 0 && dgv方剂List.SelectedRows.Count > 0)
                {
                    dgvRow = dgv方剂List.SelectedRows[0];
                }
                else
                {
                    dtRow["id"] = strID;
                    ((DataTable)dgv方剂List.DataSource).Rows.Add(dtRow);
                    //获取对应dataRow
                    foreach (DataGridViewRow row in this.dgv方剂List.Rows)
                    {
                        if (row.Cells["ID"].Value.ToString() == strID)
                        {
                            dgvRow = row;
                            break;
                        }
                    }
                }
                if (dgvRow != null)
                {

                    strNow = DataCenter.Get_Now_Time(m_baseData).ToString("yyyy-MM-dd HH:mm");
                    JObject jObject = (JObject)JsonConvert.DeserializeObject(strOut);
                    if (jObject == null || strOut.Equals("[]"))
                    {
                        throw new Exception("传入参数格式错误，应传入JSON格式字符串。" + strOut);
                    }
                    dgvRow.Cells["ID"].Value = strID;
                    dgvRow.Cells["方剂名称"].Value = jObject["方剂名称"].ToString();
                    dgvRow.Cells["简码"].Value = jObject["简码"].ToString();
                    dgvRow.Cells["别名"].Value = jObject["别名"].ToString();
                    dgvRow.Cells["别名简码"].Value = jObject["别名简码"].ToString();
                    dgvRow.Cells["来源"].Value = jObject["来源"].ToString();
                    dgvRow.Cells["组成摘要"].Value = jObject["组成摘要"].ToString();
                    dgvRow.Cells["服法描述"].Value = jObject["服法描述"].ToString();
                    dgvRow.Cells["作用描述"].Value = jObject["作用描述"].ToString();
                    dgvRow.Cells["制法描述"].Value = jObject["制法描述"].ToString();
                    dgvRow.Cells["适应证描述"].Value = jObject["适应证描述"].ToString();
                    dgvRow.Cells["方剂组成作用描述"].Value = jObject["方剂组成作用描述"].ToString();
                    dgvRow.Cells["是否保密"].Value = jObject["是否保密"].ToString();
                    dgvRow.Cells["密"].Value = jObject["是否保密"].ToString()=="1"? "√":"";
                    dgvRow.Cells["最后修改时间"].Value = strNow;
                    dgvRow.Cells["最后修改人"].Value = m_baseData.OperatorName;
                    if (intType == 0)
                    {
                        dgvRow.Cells["创建时间"].Value = strNow;
                        dgvRow.Cells["创建人"].Value = m_baseData.OperatorName;
                    }

                    dgvRow.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //疾病列表更新
        //intType 0=新增 1=修改
        private void Update疾病(int intType, string strID, string strOut)
        {
            string strNow = string.Empty;
            DataGridViewRow dgvRow = null;
            DataRow dtRow = mdt疾病.NewRow();
            try
            {
                if (strOut == string.Empty || strID == string.Empty)
                {
                    throw new Exception("中医疾病数据本地同步失败，请刷新！");
                }
                if (intType == 1 && dgv疾病.Rows.Count > 0 && dgv疾病.SelectedRows.Count > 0)
                {
                    dgvRow = dgv疾病.SelectedRows[0];
                }
                else
                {
                    dtRow["id"] = strID;
                    ((DataTable)dgv疾病.DataSource).Rows.Add(dtRow);
                    //获取对应dataRow
                    foreach (DataGridViewRow row in this.dgv疾病.Rows)
                    {
                        if (row.Cells["ID"].Value.ToString() == strID)
                        {
                            dgvRow = row;
                            break;
                        }
                    }
                }
                if (dgvRow != null)
                {
                    JObject jObject = (JObject)JsonConvert.DeserializeObject(strOut);
                    if (jObject == null || strOut.Equals("[]"))
                    {
                        throw new Exception("传入参数格式错误，应传入JSON格式字符串。" + strOut);
                    }

                    dgvRow.Cells["ID"].Value = strID;
                    dgvRow.Cells["科别"].Value = jObject["科别"].ToString();
                    dgvRow.Cells["疾病名称"].Value = jObject["疾病名称"].ToString();
                    dgvRow.Cells["简码"].Value = jObject["简码"].ToString();
                    dgvRow.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //加症列表更新
        //intType 0=新增 1=修改
        private void Update加症(int intType, string strID, string strOut)
        {
            string strNow = string.Empty;
            DataGridViewRow dgvRow = null;
            try
            {
                if (strOut == string.Empty || strID == string.Empty)
                {
                    throw new Exception("加症数据本地同步失败，请刷新！");
                }
                if (intType == 1 && dgv加症.Rows.Count > 0 && dgv加症.SelectedRows.Count > 0)
                {
                    dgvRow = dgv加症.SelectedRows[0];
                }
                else
                {
                    dgvRow = this.dgv加症.Rows[this.dgv加症.Rows.Add()];
                }
                if (dgvRow != null)
                {
                    JObject jObject = (JObject)JsonConvert.DeserializeObject(strOut);
                    if (jObject == null || strOut.Equals("[]"))
                    {
                        throw new Exception("传入参数格式错误，应传入JSON格式字符串。" + strOut);
                    }
                    if (intType == 0)
                    {
                        dgvRow.Cells["状态"].Tag = "1";
                        dgvRow.Cells["状态"].Value = picOpen.Image;
                    }
                    dgvRow.Cells["加症名称"].Value = jObject["加症名称"].ToString();
                    dgvRow.Cells["简码"].Value = jObject["简码"].ToString();
                    dgvRow.Cells["ID"].Value = strID;
                    dgvRow.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Edit方剂(int intType)
        {
            try
            {
                //缓存草药用于
                if (mdtDrug == null)
                {
                    mdtDrug = DataCenter.GetTable(m_baseData, "Get_Drugitems", "");
                }
                if (intType == 1)
                {
                    DataRow dataRow = null;//返回dataRow
                    if (dgv方剂List.Rows.Count > 0 && dgv方剂List.SelectedRows.Count > 0)
                    {
                        if (dgv方剂List.SelectedRows[0].Cells["ID"].Value != null)
                        {
                            dataRow = (dgv方剂List.Rows[dgv方剂List.SelectedRows[0].Index < 0 ? 0 : dgv方剂List.SelectedRows[0].Index].DataBoundItem as DataRowView).Row;
                            frmFjEdit frm = new frmFjEdit(1, dataRow, mdtDrug.DefaultView.ToTable(false, "ID", "草药名称", "简码", "单位"), m_baseData);
                            frm.Text = "修改方剂信息";
                            frm.ShowDialog();
                            if (frm.gblnOK == true)
                            {
                                Update方剂(intType, frm.gstrID, frm.gstrOut);
                                FindItem(dgv方剂List, "", frm.gstrID);
                                Load方剂Info();
                            }
                        }
                    }
                }
                else
                {
                    frmFjEdit frm = new frmFjEdit(0, null, mdtDrug.DefaultView.ToTable(false, "ID", "草药名称", "简码", "单位"), m_baseData);
                    frm.Text = "新增方剂信息";
                    frm.ShowDialog();
                    if (frm.gblnOK == true)
                    {
                        Update方剂(intType, frm.gstrID, frm.gstrOut);
                        FindItem(dgv方剂List, "", frm.gstrID);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Edit用药(int intType)
        {
            try
            {
                //缓存草药用于
                if (mdtDrug == null)
                {
                    mdtDrug = DataCenter.GetTable(m_baseData, "Get_Drugitems", "");
                }
                if (intType == 1)
                {
                    DataGridViewRow dataRow = null;//返回dataRow
                    if (dgv加症用药.Rows.Count > 0 && dgv加症用药.SelectedRows.Count > 0)
                    {
                        if (dgv加症用药.SelectedRows[0].Cells["ID"].Value != null)
                        {
                            dataRow = dgv加症用药.SelectedRows[0];
                            frmDrugUseEdit frm = new frmDrugUseEdit(1, dataRow, mdtDrug.DefaultView.ToTable(false, "ID", "草药名称", "简码", "单位"), m_baseData, dgv加症治法.SelectedRows[0].Cells["ID"].Value.ToString());
                            frm.Text = "修改加症用药";
                            frm.ShowDialog();
                            if (frm.gblnOK == true)
                            {
                                Load用药();
                                FindItem(dgv加症用药, "", frm.gstrID);
                            }
                        }
                    }
                }
                else
                {
                    if (dgv加症治法.SelectedRows.Count > 0 && dgv加症治法.SelectedRows[0].Cells["ID"].Value != null)
                    {
                        frmDrugUseEdit frm = new frmDrugUseEdit(0, null, mdtDrug.DefaultView.ToTable(false, "ID", "草药名称", "简码", "单位"), m_baseData, dgv加症治法.SelectedRows[0].Cells["ID"].Value.ToString());
                        frm.Text = "新增加症用药";
                        frm.ShowDialog();
                        if (frm.gblnOK == true)
                        {
                            Load用药();
                            FindItem(dgv加症用药, "", frm.gstrID);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void LoadDrug()
        {
            try
            {
                //初始化草药目录

                mdtDrug = DataCenter.GetTable(m_baseData, "Get_Drugitems", "");
                if (mdtDrug != null)
                {
                    //生成对码数据
                    if (mdtHISDrug == null)
                    {
                        mdtHISDrug = DataCenter.GetTable(m_baseData, "Get_Hisdrug", "");
                    }

                    if (mdtHISDrug != null)
                    {
                        for (int i = 0; i < mdtDrug.Rows.Count; i++)
                        {
                            if (mdtDrug.Rows[i]["His品种id"] != null)
                            {
                                if (mdtDrug.Rows[i]["His品种id"].ToString() != string.Empty)
                                {
                                    mdtHISDrug.DefaultView.RowFilter = "His品种id =" + mdtDrug.Rows[i]["His品种id"].ToString();
                                    if (mdtHISDrug.DefaultView.ToTable().Rows.Count != 0)
                                    {
                                        mdtDrug.Rows[i]["His对码"] = "[" + mdtDrug.Rows[i]["His品种id"].ToString() + "]" + mdtHISDrug.DefaultView.ToTable().Rows[0]["名称"].ToString();
                                    }
                                }
                            }
                        }
                        mdtHISDrug.DefaultView.RowFilter = null;
                    }
                }

                dgvDrug.DataSource = null;
                dgvDrug.DataSource = mdtDrug;
                //隐藏列

                dgvDrug.Columns["别名"].Visible = false;
                dgvDrug.Columns["用法"].Visible = false;
                dgvDrug.Columns["服法"].Visible = false;
                dgvDrug.Columns["禁忌"].Visible = false;
                dgvDrug.Columns["适应证"].Visible = false;
                dgvDrug.Columns["药理作用"].Visible = false;
                dgvDrug.Columns["最后修改人"].Visible = false;
                dgvDrug.Columns["最后修改时间"].Visible = false;
                dgvDrug.Columns["his品种id"].Visible = false;
                dgvDrug.Columns["草药描述"].Visible = false;
                dgvDrug.Columns["性状"].Visible = false;
                dgvDrug.Columns["成分"].Visible = false;
                dgvDrug.Columns["药性"].Visible = false;

                IntdgvList(dgvDrug);

                dgvDrug.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dgvDrug.Columns["草药名称"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgvDrug.Columns["单位"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgvDrug.Columns["HIS对码"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgvDrug.Columns["来源"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (ex.Message == "用户名或密码错误！") //ORA-01017: invalid username/password; logon denied
                {
                    this.Close();
                }
            }
        }
        private void Load方剂()
        {
            try
            {
                //初始化方剂目录

                mdt方剂 = DataCenter.GetTable(m_baseData, "Get_Fjall", "");
                dgv方剂List.DataSource = null;
                dgv方剂List.DataSource = mdt方剂;

                //隐藏列

                dgv方剂List.Columns["组成摘要"].Visible = false;
                dgv方剂List.Columns["服法描述"].Visible = false;
                dgv方剂List.Columns["作用描述"].Visible = false;
                dgv方剂List.Columns["制法描述"].Visible = false;
                dgv方剂List.Columns["适应证描述"].Visible = false;
                dgv方剂List.Columns["方剂组成作用描述"].Visible = false;
                dgv方剂List.Columns["最后修改人"].Visible = false;
                dgv方剂List.Columns["最后修改时间"].Visible = false;
                dgv方剂List.Columns["是否保密"].Visible = false;

                IntdgvList(dgv方剂List);

                dgv方剂List.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dgv方剂List.Columns["密"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;

                dgv方剂List.Columns["来源"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgv方剂List.Columns["别名"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgv方剂List.Columns["方剂名称"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        private void LoadDrugInfo()
        {
            try
            {
                if (dgvDrug.Rows.Count > 0 && dgvDrug.SelectedRows.Count > 0 && dgvDrugInfo.Rows.Count >= 27)
                {
                    if (dgvDrug.SelectedRows[0].Cells["id"].Value != null)
                    {
                        dgvDrugInfo.Rows[1].Cells[0].Value = dgvDrug.SelectedRows[0].Cells["草药描述"].Value.ToString();
                        dgvDrugInfo.Rows[4].Cells[0].Value = dgvDrug.SelectedRows[0].Cells["性状"].Value.ToString();
                        dgvDrugInfo.Rows[7].Cells[0].Value = dgvDrug.SelectedRows[0].Cells["药性"].Value.ToString();
                        dgvDrugInfo.Rows[10].Cells[0].Value = dgvDrug.SelectedRows[0].Cells["适应证"].Value.ToString();
                        dgvDrugInfo.Rows[13].Cells[0].Value = dgvDrug.SelectedRows[0].Cells["用法"].Value.ToString();
                        dgvDrugInfo.Rows[16].Cells[0].Value = dgvDrug.SelectedRows[0].Cells["服法"].Value.ToString();
                        dgvDrugInfo.Rows[19].Cells[0].Value = dgvDrug.SelectedRows[0].Cells["禁忌"].Value.ToString();
                        dgvDrugInfo.Rows[22].Cells[0].Value = dgvDrug.SelectedRows[0].Cells["成分"].Value.ToString();
                        dgvDrugInfo.Rows[25].Cells[0].Value = dgvDrug.SelectedRows[0].Cells["药理作用"].Value.ToString();
                        dgvDrugInfo.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        dgvDrugInfo.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        private void Load方剂Info()
        {
            try
            {
                if (dgv方剂List.Rows.Count > 0 && dgv方剂List.SelectedRows.Count > 0 && dgv方剂Info.Rows.Count >= 19)
                {
                    if (dgv方剂List.SelectedRows[0].Cells["id"].Value != null)
                    {
                        dgv方剂Info.Rows[1].Cells[0].Value = dgv方剂List.SelectedRows[0].Cells["组成摘要"].Value.ToString();
                        dgv方剂Info.Rows[4].Cells[0].Value = dgv方剂List.SelectedRows[0].Cells["服法描述"].Value.ToString();
                        dgv方剂Info.Rows[7].Cells[0].Value = dgv方剂List.SelectedRows[0].Cells["作用描述"].Value.ToString();
                        dgv方剂Info.Rows[10].Cells[0].Value = dgv方剂List.SelectedRows[0].Cells["制法描述"].Value.ToString();
                        dgv方剂Info.Rows[13].Cells[0].Value = dgv方剂List.SelectedRows[0].Cells["适应证描述"].Value.ToString();
                        dgv方剂Info.Rows[16].Cells[0].Value = dgv方剂List.SelectedRows[0].Cells["方剂组成作用描述"].Value.ToString();
                        dgv方剂Info.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        dgv方剂Info.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        private void Load疾病()
        {
            try
            {
                //初始化疾病目录

                mdt疾病 = DataCenter.GetTable(m_baseData, "Get_Distype", "");
                dgv疾病.DataSource = null;
                dgv疾病.DataSource = mdt疾病;
                IntdgvList(dgv疾病);
                dgv疾病.Columns["科别"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgv疾病.Columns["疾病名称"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgv疾病.Refresh();
                Load证型();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Load证型()
        {
            try
            {
                //初始化证型目录

                mdt证型 = null;
                string message = JsonConvert.SerializeObject(new { 疾病ID = dgv疾病.SelectedRows.Count == 0 ? "0" : dgv疾病.SelectedRows[0].Cells["ID"].Value.ToString() });
                mdt证型 = DataCenter.GetTable(m_baseData, "Get_Zxtype", message);
                dgv证型.DataSource = null;
                dgv证型.DataSource = mdt证型;
                IntdgvList(dgv证型);
                dgv证型.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                dgv证型.Columns["证型名称"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgv证型.Columns["证型治法"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                Load方剂对应();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Load方剂对应()
        {
            try
            {
                dgv方剂对应.Rows.Clear();
                mdt方剂对照 = null;
                if (dgv证型.SelectedRows.Count != 0)
                {
                    string message = JsonConvert.SerializeObject(new { USETYPE = "1", 证型ID = dgv证型.SelectedRows.Count == 0 ? "0" : dgv证型.SelectedRows[0].Cells["ID"].Value.ToString() });
                    mdt方剂对照 = DataCenter.GetTable(m_baseData, "Get_Fjlist", message);
                    if (mdt方剂对照 != null)
                    {
                        if (mdt方剂对照.Rows.Count > 0)
                        {
                            dgv方剂对应.Rows.Clear();

                            for (int i = 0; i < mdt方剂对照.Rows.Count; i++)
                            {
                                DataGridViewRow Dgvr = this.dgv方剂对应.Rows[this.dgv方剂对应.Rows.Add()];

                                Dgvr.Cells["状态"].Tag = mdt方剂对照.Rows[i]["状态"].ToString();
                                if (mdt方剂对照.Rows[i]["状态"].ToString() == "1")
                                {
                                    Dgvr.Cells["状态"].Value = picOpen.Image;
                                }
                                else
                                {
                                    Dgvr.Cells["状态"].Value = picStop.Image;
                                }
                                Dgvr.Cells["方剂名称"].Value = mdt方剂对照.Rows[i]["方剂名称"].ToString();
                                Dgvr.Cells["方剂名称"].Tag = mdt方剂对照.Rows[i]["ID"].ToString();
                                Dgvr.Cells["组成摘要"].Value = mdt方剂对照.Rows[i]["组成摘要"].ToString();
                                Dgvr.Cells["ID"].Value = mdt方剂对照.Rows[i]["对照ID"].ToString();
                                Dgvr.Cells["作用描述"].Value = mdt方剂对照.Rows[i]["作用描述"].ToString();
                                Dgvr.Dispose();
                            }
                        }
                    }

                }

                if (dgv方剂对应.RowCount == 0)
                {
                    dgv方剂对应.RowCount = 1;
                }
                //选中第一行

                if (dgv方剂对应.Rows.Count > 0)
                {
                    dgv方剂对应.Rows[0].Selected = true;
                    dgv方剂对应.CurrentCell = dgv方剂对应.Rows[0].Cells[0];
                }

                dgv方剂对应_SelectionChanged(null, null);

                txt方剂.TextBoxText = "";
                txt方剂.TextBoxToolTipText = "";
                txt方剂.Tag = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Load加症()
        {
            try
            {
                mdt加症 = null;
                string message = JsonConvert.SerializeObject(new { USETYPE = "1" });
                mdt加症 = DataCenter.GetTable(m_baseData, "Get_Adddis", message);
                dgv加症.Rows.Clear();
                if (mdt加症 != null)
                {
                    if (mdt加症.Rows.Count > 0)
                    {
                        dgv加症.Rows.Clear();

                        for (int i = 0; i < mdt加症.Rows.Count; i++)
                        {
                            DataGridViewRow Dgvr = this.dgv加症.Rows[this.dgv加症.Rows.Add()];

                            Dgvr.Cells["状态"].Tag = mdt加症.Rows[i]["状态"].ToString();
                            if (mdt加症.Rows[i]["状态"].ToString() == "1")
                            {
                                Dgvr.Cells["状态"].Value = picOpen.Image;
                            }
                            else
                            {
                                Dgvr.Cells["状态"].Value = picStop.Image;
                            }
                            Dgvr.Cells["加症名称"].Value = mdt加症.Rows[i]["加症名称"].ToString();
                            Dgvr.Cells["简码"].Value = mdt加症.Rows[i]["简码"].ToString();
                            Dgvr.Cells["ID"].Value = mdt加症.Rows[i]["ID"].ToString();
                            Dgvr.Dispose();
                        }
                    }
                }

                if (dgv加症.RowCount == 0)
                {
                    dgv加症.RowCount = 1;
                }
                dgv加症.Refresh();
                //选中第一行

                if (dgv加症.Rows.Count > 0)
                {
                    dgv加症.Rows[0].Selected = true;
                    dgv加症.CurrentCell = dgv加症.Rows[0].Cells["加症名称"];
                }


                dgv加症_SelectionChanged(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Load治法()
        {
            try
            {
                dgv加症治法.Rows.Clear();
                mdt治法 = null;
                if (dgv加症.SelectedRows.Count != 0 && dgv加症.SelectedRows[0].Cells["ID"].Value != null)
                {
                    string message = JsonConvert.SerializeObject(new { USETYPE = "1", 加症ID = dgv加症.SelectedRows.Count == 0 ? "0" : dgv加症.SelectedRows[0].Cells["ID"].Value.ToString() });
                    mdt治法 = DataCenter.GetTable(m_baseData, "Get_Addzf", message);
                    if (mdt治法 != null)
                    {
                        if (mdt治法.Rows.Count > 0)
                        {
                            for (int i = 0; i < mdt治法.Rows.Count; i++)
                            {
                                DataGridViewRow Dgvr = this.dgv加症治法.Rows[this.dgv加症治法.Rows.Add()];

                                Dgvr.Cells["状态"].Tag = mdt治法.Rows[i]["状态"].ToString();
                                if (mdt治法.Rows[i]["状态"].ToString() == "1")
                                {
                                    Dgvr.Cells["状态"].Value = picOpen.Image;
                                }
                                else
                                {
                                    Dgvr.Cells["状态"].Value = picStop.Image;
                                }
                                Dgvr.Cells["治法名称"].Value = mdt治法.Rows[i]["治法名称"].ToString();
                                Dgvr.Cells["简码"].Value = mdt治法.Rows[i]["简码"].ToString();
                                Dgvr.Cells["ID"].Value = mdt治法.Rows[i]["ID"].ToString();
                                Dgvr.Cells["加症ID"].Value = mdt治法.Rows[i]["加症ID"].ToString();
                                Dgvr.Dispose();
                            }
                        }
                    }

                }

                if (dgv加症治法.RowCount == 0)
                {
                    dgv加症治法.RowCount = 1;
                }
                //选中第一行

                if (dgv加症治法.Rows.Count > 0)
                {
                    dgv加症治法.Rows[0].Selected = true;
                    dgv加症治法.CurrentCell = dgv加症治法.Rows[0].Cells["治法名称"];
                }
                dgv加症治法_SelectionChanged(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void Load用药()
        {
            try
            {
                dgv加症用药.Rows.Clear();
                mdt用药 = null;
                if (dgv加症治法.SelectedRows.Count != 0 && dgv加症治法.SelectedRows[0].Cells["ID"].Value != null)
                {
                    string message = JsonConvert.SerializeObject(new { USETYPE = "1", 治法ID = dgv加症治法.SelectedRows.Count == 0 ? "0" : dgv加症治法.SelectedRows[0].Cells["ID"].Value.ToString() });
                    mdt用药 = DataCenter.GetTable(m_baseData, "Get_Additems", message);
                    if (mdt用药 != null)
                    {
                        if (mdt用药.Rows.Count > 0)
                        {
                            for (int i = 0; i < mdt用药.Rows.Count; i++)
                            {
                                DataGridViewRow Dgvr = this.dgv加症用药.Rows[this.dgv加症用药.Rows.Add()];

                                Dgvr.Cells["状态"].Tag = mdt用药.Rows[i]["状态"].ToString();
                                if (mdt用药.Rows[i]["状态"].ToString() == "1")
                                {
                                    Dgvr.Cells["状态"].Value = picOpen.Image;
                                }
                                else
                                {
                                    Dgvr.Cells["状态"].Value = picStop.Image;
                                }
                                Dgvr.Cells["草药名称"].Value = mdt用药.Rows[i]["草药名称"].ToString();
                                Dgvr.Cells["简码"].Value = mdt用药.Rows[i]["简码"].ToString();
                                Dgvr.Cells["用量"].Value = mdt用药.Rows[i]["用量"].ToString();
                                Dgvr.Cells["单位"].Value = mdt用药.Rows[i]["单位"].ToString();
                                Dgvr.Cells["ID"].Value = mdt用药.Rows[i]["ID"].ToString();
                                Dgvr.Cells["草药ID"].Value = mdt用药.Rows[i]["草药id"].ToString();
                                Dgvr.Cells["治法ID"].Value = mdt用药.Rows[i]["治法id"].ToString();
                                Dgvr.Dispose();
                            }
                        }
                    }

                }

                if (dgv加症用药.RowCount == 0)
                {
                    dgv加症用药.RowCount = 1;
                }
                //选中第一行

                if (dgv加症用药.Rows.Count > 0)
                {
                    dgv加症用药.Rows[0].Selected = true;
                    dgv加症用药.CurrentCell = dgv加症用药.Rows[0].Cells["草药名称"];
                }
                dgv加症用药_SelectionChanged(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //查找Dgv项目
        private void FindItem(DataGridView dgv, string strFind, string strID = "")
        {
            try
            {
                int intFind = 0;
                string strName = string.Empty;

                if (strID == "")
                {
                    if (strFind == string.Empty)
                    {
                        return;
                    }
                    if (dgv.SelectedRows.Count > 0)
                    {
                        if ((dgv.Tag == null ? "" : dgv.Tag.ToString()) != strFind)
                        {
                            dgv.Rows[0].Selected = true;
                            dgv.CurrentCell = dgv.Rows[0].Cells[1];
                            intFind = 0;
                        }
                        else
                        {

                            intFind = dgv.SelectedRows[0].Index + 1 < dgv.Rows.Count ? dgv.SelectedRows[0].Index + 1 : dgv.SelectedRows[0].Index;
                        }

                        dgv.Tag = strFind;
                    }
                    else
                    {
                        intFind = 0;
                    }
                    if (dgv.Name == "dgvDrug")
                    {
                        strName = "草药名称";
                    }
                    else if (dgv.Name == "dgv方剂List")
                    {
                        strName = "方剂名称";
                    }
                    else if (dgv.Name == "dgv疾病")
                    {
                        strName = "疾病名称";
                    }
                    else if (dgv.Name == "dgv加症")
                    {
                        strName = "加症名称";
                    }
                }
                for (int j = intFind; j < dgv.Rows.Count; j++)
                {
                    if (strID == "")
                    {
                        if (dgv.Rows[j].Cells[strName].Value.ToString().IndexOf(strFind) > -1)
                        {
                            dgv.Rows[j].Selected = true;
                            dgv.CurrentCell = dgv.Rows[j].Cells[1];
                            break;
                        }
                        if (dgv.Rows[j].Cells["简码"].Value != null)
                        {
                            if (dgv.Rows[j].Cells["简码"].Value.ToString().ToUpper().IndexOf(strFind.ToUpper()) > -1)
                            {
                                dgv.Rows[j].Selected = true;
                                dgv.CurrentCell = dgv.Rows[j].Cells[1];
                                break;
                            }
                        }
                    }
                    else
                    {
                        if (dgv.Rows[j].Cells["ID"].Value.ToString() == strID)
                        {
                            dgv.Rows[j].Selected = true;
                            dgv.CurrentCell = dgv.Rows[j].Cells[1];
                            break;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgvDrug_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            EditDrug(1);
        }

        private void txtFind_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {

                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txtFind.Text != "" && dgvDrug.Rows.Count > 0)
                    {
                        FindItem(dgvDrug, txtFind.Text);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadDrug();
        }

        private void btnAuto_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("请确认是否将未对码的草药通过草药名称和HIS系统中的草药品种进行自动对码？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.OK)
            {

                try
                {
                    DataTable dtOut = new DataTable();
                    //保存

                    string message = JsonConvert.SerializeObject(new
                    {
                        操作员名称 = m_baseData.OperatorName,
                        操作员ID = m_baseData.OperatorID
                    });
                    dtOut = DataCenter.ExecuteFunc("Set_Autodrug", message, m_baseData.UserName, m_baseData.UserPassword);
                    if (dtOut != null)
                    {
                        if (ComLib.StrToInt(dtOut.Rows[0]["数量"].ToString()) > 0)
                        {
                            MessageBox.Show("本次成功对码了" + dtOut.Rows[0]["数量"].ToString() + "条草药数据。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadDrug();
                        }
                        else
                        {
                            MessageBox.Show("目前尚未匹配的可以自动对码的草药。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        throw new Exception("获取保存数据失败，服务器返回数据为空！");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ComLib.GetMessageFromOracleError(ex.Message), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            DelItem(dgvDrug);
        }

        private void mytab_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (mytab.SelectedIndex)
            {
                case 0://草药目录
                    if (mdtDrug == null)
                    {
                        LoadDrug();
                    }
                    break;
                case 1://方剂目录
                    if (mdt方剂 == null)
                    {
                        Load方剂();
                    }
                    break;
                case 2://疾病目录
                    if (mdt疾病 == null)
                    {
                        Load疾病();
                    }
                    break;
                case 3://加症目录
                    if (mdt加症 == null)
                    {
                        Load加症();
                    }
                    break;
                default:
                    break;
            }
        }

        private void btn方剂刷新_Click(object sender, EventArgs e)
        {
            Load方剂();
        }

        private void txt方剂Find_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {

                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt方剂Find.Text != "" && dgv方剂List.Rows.Count > 0)
                    {
                        FindItem(dgv方剂List, txt方剂Find.Text);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn新增方剂_Click(object sender, EventArgs e)
        {
            Edit方剂(0);
        }

        private void btn修改方剂_Click(object sender, EventArgs e)
        {
            Edit方剂(1);
        }

        private void dgv方剂List_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Edit方剂(1);
        }

        private void btn疾病刷新_Click(object sender, EventArgs e)
        {
            Load疾病();
        }

        private void dgv疾病_SelectionChanged(object sender, EventArgs e)
        {
            if (dgv疾病.RowCount > 0)
            {
                Load证型();
            }
        }

        private void frmBase_Resize(object sender, EventArgs e)
        {
            mytab_Resize(null, null);
        }

        private void mytab_Resize(object sender, EventArgs e)
        {
            foreach (TabPage p in mytab.TabPages) p.Refresh();
        }

        private void txtFind疾病_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {

                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txtFind疾病.Text != "" && dgv疾病.Rows.Count > 0)
                    {
                        FindItem(dgv疾病, txtFind疾病.Text);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgv证型_SelectionChanged(object sender, EventArgs e)
        {
            if (dgv证型.RowCount > 0)
            {
                Load方剂对应();
            }
        }

        private void dgv方剂对应_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (dgv方剂对应.SelectedRows.Count > 0 && dgv方剂对应.RowCount > 0)
                {
                    if (dgv方剂对应.SelectedRows[0].Cells["状态"].Tag == null)
                    {
                        btn方剂状态.Image = picOpen.Image;
                        btn方剂状态.Text = "启用";
                    }
                    else
                    {
                        if (dgv方剂对应.SelectedRows[0].Cells["状态"].Tag.ToString() == "1")
                        {
                            btn方剂状态.Image = picStop.Image;
                            btn方剂状态.Text = "停用";
                        }
                        else
                        {
                            btn方剂状态.Image = picOpen.Image;
                            btn方剂状态.Text = "启用";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgv疾病_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {

                if (e.Button == MouseButtons.Right && e.ColumnIndex > -1 && e.RowIndex > -1)  //点击的是鼠标右键，并且不是表头

                {
                    //右键选中单元格

                    this.dgv疾病.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
                    MenuOpen.Visible = false;
                    MenuStop.Visible = false;
                    this.menu疾病.Show(MousePosition.X, MousePosition.Y); //MousePosition.X, MousePosition.Y 是为了让菜单在所选行的位置显示

                    this.menu疾病.Tag = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgv证型_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right && e.ColumnIndex > -1 && e.RowIndex > -1)  //点击的是鼠标右键，并且不是表头

                {
                    //右键选中单元格

                    this.dgv证型.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
                    MenuOpen.Visible = false;
                    MenuStop.Visible = false;
                    this.menu疾病.Show(MousePosition.X, MousePosition.Y); //MousePosition.X, MousePosition.Y 是为了让菜单在所选行的位置显示

                    this.menu疾病.Tag = "1";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn新增疾病_Click(object sender, EventArgs e)
        {
            EditDiag(0);
        }

        private void btn修改疾病_Click(object sender, EventArgs e)
        {
            EditDiag(1);
        }

        private void dgv疾病_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            EditDiag(1);
        }

        private void MenuAdd_Click(object sender, EventArgs e)
        {
            if (menu疾病.Tag.ToString() == "0")
            {
                EditDiag(0);
            }
            else if (menu疾病.Tag.ToString() == "1")
            {
                EditZx(0);
            }
            else if (menu疾病.Tag.ToString() == "2")
            {
                Edit加症(0);
            }
            else if (menu疾病.Tag.ToString() == "3")
            {
                Edit治法(0);
            }
            else if (menu疾病.Tag.ToString() == "4")
            {
                Edit用药(0);
            }
            else if (menu疾病.Tag.ToString() == "5")
            {
                EditDrug(0);
            }
            else if (menu疾病.Tag.ToString() == "6")
            {
                Edit方剂(0);
            }
        }

        private void MenuEdit_Click(object sender, EventArgs e)
        {
            if (menu疾病.Tag.ToString() == "0")
            {
                EditDiag(1);
            }
            else if (menu疾病.Tag.ToString() == "1")
            {
                EditZx(1);
            }
            else if (menu疾病.Tag.ToString() == "2")
            {
                Edit加症(1);
            }
            else if (menu疾病.Tag.ToString() == "3")
            {
                Edit治法(1);
            }
            else if (menu疾病.Tag.ToString() == "4")
            {
                Edit用药(1);
            }
            else if (menu疾病.Tag.ToString() == "5")
            {
                EditDrug(1);
            }
            else if (menu疾病.Tag.ToString() == "6")
            {
                Edit方剂(1);
            }
        }

        private void menuDel_Click(object sender, EventArgs e)
        {
            if (menu疾病.Tag.ToString() == "0")
            {
                DelItem(dgv疾病);
            }
            else if (menu疾病.Tag.ToString() == "1")
            {
                DelItem(dgv证型);
            }
            else if (menu疾病.Tag.ToString() == "2")
            {
                DelItem(dgv加症);
            }
            else if (menu疾病.Tag.ToString() == "3")
            {
                DelItem(dgv加症治法);
            }
            else if (menu疾病.Tag.ToString() == "4")
            {
                DelItem(dgv加症用药);
            }
            else if (menu疾病.Tag.ToString() == "5")
            {
                DelItem(dgvDrug);
            }
            else if (menu疾病.Tag.ToString() == "6")
            {
                DelItem(dgv方剂List);
            }
        }

        private void btn新增证型_Click(object sender, EventArgs e)
        {
            EditZx(0);
        }

        private void btn修改证型_Click(object sender, EventArgs e)
        {
            EditZx(1);
        }

        private void dgv证型_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            EditZx(1);
        }

        private void txt方剂_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (mdt方剂 == null) { Load方剂(); }

                if (mdt方剂 != null)
                {
                    mdt方剂.DefaultView.RowFilter = "";
                    if (mdt方剂.Rows.Count > 0)
                    {
                        frmZySelect frmSelect = new frmZySelect("方剂选择器", mdt方剂.DefaultView.ToTable(false, "ID", "方剂名称", "作用描述", "组成摘要"));
                        frmSelect.ShowDialog(this);
                        if (frmSelect.dataRowOut != null)
                        {
                            txt方剂.TextBoxText = frmSelect.dataRowOut["方剂名称"].ToString();
                            txt方剂.TextBoxToolTipText = frmSelect.dataRowOut["方剂名称"].ToString();
                            txt方剂.Tag = frmSelect.dataRowOut;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("未获取到有效的方剂。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt方剂_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {

                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt方剂.TextBoxText != txt方剂.TextBoxToolTipText || txt方剂.TextBoxToolTipText == "")
                    {
                        if (mdt方剂 == null) { Load方剂(); }

                        if (mdt方剂 != null)
                        {
                            mdt方剂.DefaultView.RowFilter = txt方剂.TextBoxText == "" ? "" : "方剂名称 like '%" + txt方剂.TextBoxText + "%' Or 简码 like '%" + txt方剂.TextBoxText + "%'";
                            if (mdt方剂.Rows.Count > 0)
                            {
                                frmZySelect frmSelect = new frmZySelect("方剂选择器", mdt方剂.DefaultView.ToTable(false, "ID", "方剂名称", "作用描述", "组成摘要"));
                                frmSelect.ShowDialog(this);
                                if (frmSelect.dataRowOut != null)
                                {
                                    txt方剂.TextBoxText = frmSelect.dataRowOut["方剂名称"].ToString();
                                    txt方剂.TextBoxToolTipText = frmSelect.dataRowOut["方剂名称"].ToString();
                                    txt方剂.Tag = frmSelect.dataRowOut;
                                }
                            }
                            mdt方剂.DefaultView.RowFilter = "";
                        }
                        else
                        {
                            MessageBox.Show("未获取到有效的方剂。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        btn方剂对应Add.Focus();
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt方剂_Validating(object sender, CancelEventArgs e)
        {
            if (txt方剂.TextBoxToolTipText != txt方剂.TextBoxText)
            {
                txt方剂.TextBoxText = txt方剂.TextBoxToolTipText;
            }
        }

        private void btn方剂对应Add_Click(object sender, EventArgs e)
        {
            Edit方剂对应(0);
        }


        private void btn方剂状态_Click(object sender, EventArgs e)
        {
            Edit方剂对应(1);
        }

        private void dgv加症_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                bool blnShow = false;
                //更新图标
                if (dgv加症.SelectedRows.Count > 0 && dgv加症.RowCount > 0)
                {
                    if (dgv加症.SelectedRows[0].Cells["状态"].Tag == null)
                    {
                        blnShow = false;
                    }
                    else
                    {
                        if (dgv加症.SelectedRows[0].Cells["状态"].Tag.ToString() == "1")
                        {
                            blnShow = true;
                        }
                        else
                        {
                            blnShow = false;
                        }
                    }
                }
                btn停用加症.Visible = blnShow;
                btn启用加症.Visible = !blnShow;

                Load治法();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgv加症用药_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                bool blnShow = false;
                //更新图标
                if (dgv加症用药.SelectedRows.Count > 0 && dgv加症用药.RowCount > 0)
                {
                    if (dgv加症用药.SelectedRows[0].Cells["状态"].Tag == null)
                    {
                        blnShow = false;
                    }
                    else
                    {
                        if (dgv加症用药.SelectedRows[0].Cells["状态"].Tag.ToString() == "1")
                        {
                            blnShow = true;
                        }
                        else
                        {
                            blnShow = false;
                        }
                    }
                }
                btn停用用药.Visible = blnShow;
                btn启用用药.Visible = !blnShow;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgv加症治法_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                bool blnShow = false;
                //更新图标
                if (dgv加症治法.SelectedRows.Count > 0 && dgv加症治法.RowCount > 0)
                {
                    if (dgv加症治法.SelectedRows[0].Cells["状态"].Tag == null)
                    {
                        blnShow = false;
                    }
                    else
                    {
                        if (dgv加症治法.SelectedRows[0].Cells["状态"].Tag.ToString() == "1")
                        {
                            blnShow = true;
                        }
                        else
                        {
                            blnShow = false;
                        }
                    }
                }
                btn停用治法.Visible = blnShow;
                btn启用治法.Visible = !blnShow;
                Load用药();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn加症刷新_Click(object sender, EventArgs e)
        {
            Load加症();
        }

        private void txt加症Find_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {

                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt加症Find.Text != "" && dgv加症.Rows.Count > 0)
                    {
                        FindItem(dgv加症, txt加症Find.Text);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn加症Add_Click(object sender, EventArgs e)
        {
            Edit加症(0);
        }

        private void btn加症Edit_Click(object sender, EventArgs e)
        {
            Edit加症(1);
        }

        private void dgv加症_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Edit加症(1);
        }

        private void btn治法Add_Click(object sender, EventArgs e)
        {
            Edit治法(0);
        }

        private void btn治法Edit_Click(object sender, EventArgs e)
        {
            Edit治法(1);
        }

        private void dgv加症治法_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {

                if (e.Button == MouseButtons.Right && e.ColumnIndex > -1 && e.RowIndex > -1)  //点击的是鼠标右键，并且不是表头

                {
                    //右键选中单元格

                    this.dgv加症治法.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
                    bool blnShow = false;

                    //更新图标
                    if (dgv加症治法.SelectedRows.Count > 0 && dgv加症治法.RowCount > 0)
                    {
                        if (dgv加症治法.SelectedRows[0].Cells["状态"].Tag == null)
                        {
                            blnShow = false;
                        }
                        else
                        {
                            if (dgv加症治法.SelectedRows[0].Cells["状态"].Tag.ToString() == "1")
                            {
                                blnShow = true;
                            }
                            else
                            {
                                blnShow = false;
                            }
                        }
                    }

                    MenuStop.Visible = blnShow;
                    MenuOpen.Visible = !blnShow;
                    this.menu疾病.Show(MousePosition.X, MousePosition.Y); //MousePosition.X, MousePosition.Y 是为了让菜单在所选行的位置显示

                    this.menu疾病.Tag = "3";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgv加症治法_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Edit治法(1);
        }

        private void dgv加症_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {

                if (e.Button == MouseButtons.Right && e.ColumnIndex > -1 && e.RowIndex > -1)  //点击的是鼠标右键，并且不是表头

                {
                    //右键选中单元格

                    this.dgv加症.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;

                    bool blnShow = false;
                    //更新图标
                    if (dgv加症.SelectedRows.Count > 0 && dgv加症.RowCount > 0)
                    {
                        if (dgv加症.SelectedRows[0].Cells["状态"].Tag == null)
                        {
                            blnShow = false;
                        }
                        else
                        {
                            if (dgv加症.SelectedRows[0].Cells["状态"].Tag.ToString() == "1")
                            {
                                blnShow = true;
                            }
                            else
                            {
                                blnShow = false;
                            }
                        }
                    }
                    MenuStop.Visible = blnShow;
                    MenuOpen.Visible = !blnShow;

                    this.menu疾病.Show(MousePosition.X, MousePosition.Y); //MousePosition.X, MousePosition.Y 是为了让菜单在所选行的位置显示

                    this.menu疾病.Tag = "2";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn新增用药_Click(object sender, EventArgs e)
        {
            Edit用药(0);
        }

        private void btn修改用药_Click(object sender, EventArgs e)
        {
            Edit用药(1);
        }

        private void dgv加症用药_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Edit用药(1);
        }

        private void btn停用加症_Click(object sender, EventArgs e)
        {
            Change状态(dgv加症);
        }

        private void btn启用加症_Click(object sender, EventArgs e)
        {
            Change状态(dgv加症);
        }

        private void MenuOpen_Click(object sender, EventArgs e)
        {
            if (menu疾病.Tag.ToString() == "2")
            {
                Change状态(dgv加症);
            }
            else if (menu疾病.Tag.ToString() == "3")
            {
                Change状态(dgv加症治法);
            }
            else if (menu疾病.Tag.ToString() == "4")
            {
                Change状态(dgv加症用药);
            }
        }

        private void MenuStop_Click(object sender, EventArgs e)
        {
            if (menu疾病.Tag.ToString() == "2")
            {
                Change状态(dgv加症);
            }
            else if (menu疾病.Tag.ToString() == "3")
            {
                Change状态(dgv加症治法);
            }
            else if (menu疾病.Tag.ToString() == "4")
            {
                Change状态(dgv加症用药);
            }
        }

        private void dgv加症用药_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right && e.ColumnIndex > -1 && e.RowIndex > -1)  //点击的是鼠标右键，并且不是表头

                {
                    //右键选中单元格

                    this.dgv加症用药.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;

                    bool blnShow = false;
                    //更新图标
                    if (dgv加症用药.SelectedRows.Count > 0 && dgv加症用药.RowCount > 0)
                    {
                        if (dgv加症用药.SelectedRows[0].Cells["状态"].Tag == null)
                        {
                            blnShow = false;
                        }
                        else
                        {
                            if (dgv加症用药.SelectedRows[0].Cells["状态"].Tag.ToString() == "1")
                            {
                                blnShow = true;
                            }
                            else
                            {
                                blnShow = false;
                            }
                        }
                    }
                    MenuStop.Visible = blnShow;
                    MenuOpen.Visible = !blnShow;

                    this.menu疾病.Show(MousePosition.X, MousePosition.Y); //MousePosition.X, MousePosition.Y 是为了让菜单在所选行的位置显示

                    this.menu疾病.Tag = "4";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn停用治法_Click(object sender, EventArgs e)
        {
            Change状态(dgv加症治法);
        }

        private void btn启用治法_Click(object sender, EventArgs e)
        {
            Change状态(dgv加症治法);
        }

        private void btn停用用药_Click(object sender, EventArgs e)
        {
            Change状态(dgv加症用药);
        }

        private void btn启用用药_Click(object sender, EventArgs e)
        {
            Change状态(dgv加症用药);
        }

        private void btn删除方剂_Click(object sender, EventArgs e)
        {
            DelItem(dgv方剂List);
        }

        private void btn疾病Del_Click(object sender, EventArgs e)
        {
            DelItem(dgv疾病);
        }

        private void btn删除证型_Click(object sender, EventArgs e)
        {
            DelItem(dgv证型);
        }

        private void btn方剂对应Del_Click(object sender, EventArgs e)
        {
            DelItem(dgv方剂对应);
        }

        private void btn加症Del_Click(object sender, EventArgs e)
        {
            DelItem(dgv加症);
        }

        private void btn治法Del_Click(object sender, EventArgs e)
        {
            DelItem(dgv加症治法);
        }

        private void btn删除用药_Click(object sender, EventArgs e)
        {
            DelItem(dgv加症用药);
        }

        private void dgvDrug_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right && e.ColumnIndex > -1 && e.RowIndex > -1)  //点击的是鼠标右键，并且不是表头

                {
                    //右键选中单元格

                    this.dgvDrug.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
                    MenuOpen.Visible = false;
                    MenuStop.Visible = false;
                    this.menu疾病.Show(MousePosition.X, MousePosition.Y); //MousePosition.X, MousePosition.Y 是为了让菜单在所选行的位置显示

                    this.menu疾病.Tag = "5";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgv方剂List_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right && e.ColumnIndex > -1 && e.RowIndex > -1)  //点击的是鼠标右键，并且不是表头

                {
                    //右键选中单元格

                    this.dgv方剂List.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
                    MenuOpen.Visible = false;
                    MenuStop.Visible = false;
                    this.menu疾病.Show(MousePosition.X, MousePosition.Y); //MousePosition.X, MousePosition.Y 是为了让菜单在所选行的位置显示

                    this.menu疾病.Tag = "6";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgvDrug_SelectionChanged(object sender, EventArgs e)
        {
            LoadDrugInfo();
        }

        private void dgvDrug_CurrentCellChanged(object sender, EventArgs e)
        {
            try
            {
                DataGridViewCell CurrnetCell = this.dgvDrug.CurrentCell;

                this.txt品种.Visible = false;

                if (CurrnetCell != null)
                {
                    Rectangle Rect = this.dgvDrug.GetCellDisplayRectangle(dgvDrug.Columns["His对码"].Index, CurrnetCell.RowIndex, true);
                    txt品种.Visible = true;
                    txt品种.Top = Rect.Top;
                    txt品种.Left = Rect.Left;
                    txt品种.Height = Rect.Height;
                    txt品种.Width = Rect.Width;
                    //缓存上次注脚
                    txt品种.Tag = CurrnetCell.RowIndex;

                    if (dgvDrug.Rows[CurrnetCell.RowIndex].Cells[dgvDrug.Columns["His对码"].Index].Value != null)
                    {
                        txt品种.TextBoxText = dgvDrug.Rows[CurrnetCell.RowIndex].Cells[dgvDrug.Columns["His对码"].Index].Value.ToString();
                        txt品种.TextBoxToolTipText = dgvDrug.Rows[CurrnetCell.RowIndex].Cells[dgvDrug.Columns["His对码"].Index].Value.ToString();
                    }
                    else
                    {
                        txt品种.TextBoxText = "";
                        txt品种.TextBoxToolTipText = "";
                    }
                    if (CurrnetCell.ColumnIndex == dgvDrug.Columns["His对码"].Index)
                    {
                        txt品种.Focus();
                    }

                    txt品种.Refresh();
                }
                else
                {
                    this.txt品种.Visible = false;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgvDrug_Scroll(object sender, ScrollEventArgs e)
        {
            txt品种.Visible = false;
            dgvDrug.Focus();
        }


        //表格光标纵向跳转
        private bool m_txt品种Edit = false;

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            string strTmp = string.Empty;

            strTmp = txt品种.TextBoxToolTipText.Replace("'", "").Replace("[", "[[ ")
                            .Replace("]", " ]]")
                            .Replace("*", "[*]")
                            .Replace("%", "[%]")
                            .Replace("[[ ", "[[]")
                            .Replace(" ]]", "[]]")
                            .Replace("\'", "''");
            if (keyData == Keys.Enter && (m_txt品种Edit == true) && (txt品种.TextBoxText != strTmp || txt品种.TextBoxText == "") && txt品种.Visible)
            {
                try
                {

                    mdtHISDrug.DefaultView.RowFilter = "名称 like '%" + txt品种.TextBoxText + "%' Or 简码 like '%" + txt品种.TextBoxText + "%'";
                    frmZySelect frmSelect = new frmZySelect("HIS品种选择", mdtHISDrug.DefaultView.ToTable());
                    frmSelect.ShowDialog(this);
                    if (frmSelect.dataRowOut != null)
                    {
                        txt品种.TextBoxText = "[" + frmSelect.dataRowOut["His品种id"].ToString() + "]" + frmSelect.dataRowOut["名称"].ToString();
                        txt品种.TextBoxToolTipText = "[" + frmSelect.dataRowOut["His品种id"].ToString() + "]" + frmSelect.dataRowOut["名称"].ToString();
                        dgvDrug.SelectedRows[0].Cells["his对码"].Value = "[" + frmSelect.dataRowOut["His品种id"].ToString() + "]" + frmSelect.dataRowOut["名称"].ToString();
                        dgvDrug.SelectedRows[0].Cells["his品种ID"].Value = frmSelect.dataRowOut["His品种id"].ToString();
                        SaveHIS对码();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void txt品种_Enter(object sender, EventArgs e)
        {
            m_txt品种Edit = true;
        }

        private void txt品种_Leave(object sender, EventArgs e)
        {
            m_txt品种Edit = false;
            dgvDrug.Focus();
            txt品种.Visible = false;
        }

        private void txt品种_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                mdtHISDrug.DefaultView.RowFilter = "";
                frmZySelect frmSelect = new frmZySelect("HIS品种选择", mdtHISDrug.DefaultView.ToTable());
                frmSelect.ShowDialog(this);
                if (frmSelect.dataRowOut != null)
                {
                    txt品种.TextBoxText = "[" + frmSelect.dataRowOut["His品种id"].ToString() + "]" + frmSelect.dataRowOut["名称"].ToString();
                    txt品种.TextBoxToolTipText = "[" + frmSelect.dataRowOut["His品种id"].ToString() + "]" + frmSelect.dataRowOut["名称"].ToString();

                    dgvDrug.SelectedRows[0].Cells["his对码"].Value = "[" + frmSelect.dataRowOut["His品种id"].ToString() + "]" + frmSelect.dataRowOut["名称"].ToString();
                    dgvDrug.SelectedRows[0].Cells["his品种ID"].Value = frmSelect.dataRowOut["His品种id"].ToString();
                    SaveHIS对码();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void SaveHIS对码()
        {
            try
            {

                if (dgvDrug.SelectedRows.Count > 0 && dgvDrug.SelectedRows[0].Cells["ID"].Value != null && dgvDrug.SelectedRows[0].Cells["His品种ID"].Value != null)
                {
                    //保存
                    string message = JsonConvert.SerializeObject(new
                    {
                        USETYPE = "2",////使用类型(0-新增/1-修改/2-修改对码)
                        草药ID = dgvDrug.SelectedRows[0].Cells["ID"].Value.ToString(),
                        HIS品种ID = dgvDrug.SelectedRows[0].Cells["His品种ID"].Value.ToString(),
                        操作员名称 = m_baseData.OperatorName,
                        操作员ID = m_baseData.OperatorID
                    });

                    DataTable dtOut = new DataTable();

                    dtOut = DataCenter.ExecuteFunc("Save_Drugitem", message, m_baseData.UserName, m_baseData.UserPassword);
                    if (dtOut == null)
                    {
                        throw new Exception("获取保存数据失败，服务器返回数据为空！");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ComLib.GetMessageFromOracleError(ex.Message), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void dgvDrugInfo_SelectionChanged(object sender, EventArgs e)
        {
            dgvDrugInfo.Rows[dgvDrugInfo.RowCount - 1].Selected = true;
        }

        private void dgv方剂List_SelectionChanged(object sender, EventArgs e)
        {
            Load方剂Info();
        }

        private void dgv方剂Info_SelectionChanged(object sender, EventArgs e)
        {
            dgv方剂Info.Rows[dgv方剂Info.RowCount - 1].Selected = true;
        }
    }
}
