using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ZLSOFT.HIS.ZyEdit.Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ZLSOFT.HIS.ZyEdit
{
    partial class frmZyEdit : Form
    {

        #region 模块变量
        //基础数据
        private BaseData m_baseData = new BaseData();
        //基础输入项缓存数据

        DataTable mdtDisType = new DataTable();
        DataTable mdt证型 = new DataTable();
        DataTable mdt方剂 = new DataTable();
        DataTable mdt方剂组成 = new DataTable();
        DataTable mdt临症加症 = new DataTable();
        DataTable mdt加症治法 = new DataTable();
        DataTable mdt加症草药 = new DataTable();
        DataTable mdt煎法 = new DataTable();
        DataTable mdt用法 = new DataTable();
        DataTable mdt频率 = new DataTable();
        DataTable mdt药房 = new DataTable();
        private TextBoxEx mtxtTmp; //当前操作文本框对象

        private DataGridViewTextBoxEditingControl CellEdit = null;//用于绑定表格输入事件
        private string mstr方剂名 = "";
        private Boolean mbln保密 = false;
        public string gstrOut = "";

        //让Combobox支持Tag
        struct itemEx
        {
            public object Tag;
            public string Text;
            public string ItemData;
            public itemEx(object tag, string text, string ItemData = "")
            {
                this.Tag = tag;
                this.Text = text;
                this.ItemData = ItemData;
            }
            public override string ToString()
            {
                return this.Text;
            }
        }

        #endregion

        #region 窗体初始化

        public frmZyEdit(string message)
        {
            try
            {
                JObject jObject = (JObject)JsonConvert.DeserializeObject(message);
                if (jObject == null || message.Equals("[]"))
                {
                    throw new Exception("传入参数格式错误，应传入JSON格式字符串。" + message);
                }
                m_baseData.System = jObject["来源系统"].ToString();
                m_baseData.站点 = jObject["站点"].ToString();
                m_baseData.UseType = jObject["使用类型"].ToString();
                m_baseData.病人ID = jObject["病人ID"].ToString();
                m_baseData.挂号单 = jObject["挂号单"].ToString();
                m_baseData.门诊号 = jObject["门诊号"].ToString();

                m_baseData.Name = jObject["病人姓名"].ToString();
                m_baseData.Sex = jObject["病人性别"].ToString();
                m_baseData.Age = jObject["病人年龄"].ToString();
                m_baseData.民族 = jObject["病人民族"].ToString();
                m_baseData.Date = jObject["出生日期"].ToString();

                m_baseData.诊断ID = jObject["诊断ID"].ToString();
                m_baseData.DeptID = jObject["当前科室ID"].ToString();
                m_baseData.DeptName = jObject["当前科室名"].ToString();
                m_baseData.OperatorID = jObject["操作员ID"].ToString();
                m_baseData.OperatorName = jObject["操作员姓名"].ToString();
                m_baseData.UserName = jObject["用户名"].ToString();
                m_baseData.UserPassword = jObject["用户密码"].ToString();
                if (m_baseData.System.Equals("0"))
                {
                    m_baseData.TNSNAME = jObject["TNSNAME"].ToString();
                    if (jObject["TNSNAME"].ToString() != "")
                    {
                        ComLib.strTnsName = jObject["TNSNAME"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            InitializeComponent();


        }


        private void IntdgvList(DataGridView dgv, string[] CH, int[] CHWidth)
        {
            ComLib.DoubleBuffered(dgv, true);
            //DataGridView控件基础设置
            // 表格上下左右自适应
            dgv.Anchor = (AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom | AnchorStyles.Left);
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;

            // 列手工排序

            dgv.AllowUserToOrderColumns = false;
            // 列头系统样式，设置为false，自定义才生效

            dgv.EnableHeadersVisualStyles = false;
            // 列头高度大小模式
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            // 列头高度大小
            dgv.ColumnHeadersHeight = 25;
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

            // 单元格背景色
            dgv.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#FFFFFF");

            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;//设置前景色


            dgv.MultiSelect = false;

            // 单元格内容居中

            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            //初始化行数

            dgv.RowCount = 1;

            //设置dgvList设置表头
            if (CH != null)
            {
                dgv.ColumnCount = CH.Count() + 1;
                dgv.Columns[0].Width = 20;

                for (int i = 0; i < CH.Length; i++)
                {
                    dgv.Columns[i + 1].Name = CH[i];
                    dgv.Columns[i + 1].Width = CHWidth[i];
                }

                DataGridViewImageColumn dgvImageColumn = new DataGridViewImageColumn();
                dgvImageColumn.HeaderText = "";
                dgvImageColumn.ImageLayout = DataGridViewImageCellLayout.Normal;
                dgvImageColumn.Name = "btnDel";
                dgvImageColumn.Image = picNone.Image;
                dgv.Columns.Add(dgvImageColumn);
                dgv.Columns[dgv.ColumnCount - 1].Width = 30;

            }

            //dgvList设置用量和脚注可以编辑

            if (dgv.Name == "dgvList")
            {

                dgv.ReadOnly = false;
                foreach (DataGridViewColumn c in dgv.Columns)
                {
                    if (c.Name == "用量")
                    {
                        c.ReadOnly = false;
                        c.DefaultCellStyle.BackColor = Color.AliceBlue;
                        c.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if (c.Name == "脚注")
                    {
                        c.ReadOnly = true;
                        c.DefaultCellStyle.BackColor = Color.AliceBlue;
                        c.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else if (c.Name == "规格")
                    {
                        c.ReadOnly = true;
                        c.DefaultCellStyle.BackColor = Color.AliceBlue;
                        c.DefaultCellStyle.ForeColor = Color.Maroon;
                    }
                    else
                    {
                        c.ReadOnly = true;
                    }
                }

                dgv.DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#87CEFA");//设置背景色

                dgv.SelectionMode = DataGridViewSelectionMode.CellSelect;//设置为单个选择项


                //设置不允许排序

                for (int i = 0; i < dgv.Columns.Count; i++)
                    dgv.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;

                // 设置列头
                dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dgv.Columns["用量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgv.Columns["总量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgv.Columns["库存情况"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv.ForeColor = Color.Black;
            }
            else
            {
                // 列头居中
                dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                dgv.ReadOnly = true;
                dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//设置为整行

                dgv.DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#66CCFF");//设置背景色


            }
            dgv.RowHeadersVisible = false;

        }

        private void frmZyEdit_Load(object sender, EventArgs e)
        {
            try
            {

                DataCenter.TranPasswd = m_baseData.System.Equals("0");
                mbln保密 = false;

                this.Text = "中医辩证论治(当前用户：" + m_baseData.OperatorName + ")";
                //控件信息
                string[] CH = { "来源", "草药名", "用量", "单位", "脚注", "总量", "规格", "库存情况" };//它的数组的长度必须和datagridview的列数保持一致

                int[] CHWidth = { 130, 100, 60, 50, 100, 60, 180, 80 };//列宽
                IntdgvList(dgvList, CH, CHWidth);
                IntdgvList(dgvSelect, null, null);



                //增加表格下拉项

                dgvList.Controls.Add(cbo脚注);
                dgvList.Controls.Add(cbo规格);

                cboSelect.SetDropDown(dgvSelect);

                //基础信息
                LoadBaseInfo(m_baseData);


                //病人信息
                LoadPatientInfo(m_baseData);

                //处方信息
                if (m_baseData.UseType != "0")
                {
                    LoadData();
                }

                //查看处方设置控件状态不可用
                if (m_baseData.UseType == "2")
                {
                    btnAdd.Enabled = false;
                    btnSave.Enabled = false;
                    ComLib.SetEnabled(palPatiInfo, false);
                    ComLib.SetEnabled(palEditInfo, false);
                    ComLib.SetEnabled(tableLayoutPanel3, false);
                    ComLib.SetEnabled(tableLayoutPanel4, false);
                    ComLib.SetEnabled(tableLayoutPanel5, false);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }

        }

        //加载病人信息
        private void LoadPatientInfo(BaseData baseData)
        {
            txtName.Text = baseData.Name;
            txtSex.Text = baseData.Sex;
            txtAge.Text = baseData.Age;
            txtDate.Text = baseData.Date;
            txtMzh.Text = baseData.门诊号;
            txt民族.Text = baseData.民族;
        }

        //加载基础数据
        private void LoadBaseInfo(BaseData baseData)
        {
            //初始化DataTable
            mdt临症加症 = null;
            mdtDisType = null;
            mdt证型 = null;
            mdt方剂 = null;
            mdt方剂组成 = null;
            mdt加症治法 = null;
            mdt加症草药 = null;
            mdt煎法 = null;
            mdt用法 = null;
            mdt频率 = null;
            mdt药房 = null;

            string strTmp = ",";
            //缓存疾病种类
            mdtDisType = DataCenter.GetTable(baseData, "Get_Distype", "");
            if (mdtDisType.Rows.Count == 0)
            {
                throw new Exception("获取中药疾病数据失败，服务器返回数据为空！");
            }

            //加载科别
            cbo科别.Items.Clear();
            for (int i = 0; i < mdtDisType.Rows.Count; i++)
            {
                if (strTmp.IndexOf(mdtDisType.Rows[i]["科别"].ToString()) < 0)
                {
                    cbo科别.Items.Add(mdtDisType.Rows[i]["科别"].ToString());
                    strTmp = strTmp + mdtDisType.Rows[i]["科别"].ToString() + ",";
                }
            }

            //加载药房
            cbo药房.Items.Clear();

            string message = JsonConvert.SerializeObject(new { 药房ID = 0, 科室ID = m_baseData.DeptID, 站点 = m_baseData.站点 });
            mdt药房 = DataCenter.GetTable(baseData, "Get_Drugdept", message);
            if (mdt药房.Rows.Count == 0)
            {
                throw new Exception("获取可用药房数据失败，服务器返回数据为空！");
            }
            for (int i = 0; i < mdt药房.Rows.Count; i++)
            {
                itemEx item = new itemEx(mdt药房.Rows[i]["ID"].ToString(), mdt药房.Rows[i]["编码"].ToString() + "-" + mdt药房.Rows[i]["名称"].ToString());
                cbo药房.Items.Add(item);
                if (mdt药房.Rows[i]["缺省"].ToString() == "1")
                {
                    cbo药房.SelectedIndex = cbo药房.Items.Count - 1;
                }
            }

            if (cbo药房.Items.Count > 0 && cbo药房.SelectedIndex == -1)
            {
                cbo药房.SelectedIndex = 0;
            }

            //加载脚注
            cbo脚注.Items.Clear();

            DataTable dtTmp = DataCenter.GetTable(baseData, "Get_Jzitems", message);
            if (dtTmp != null)
            {
                for (int i = 0; i < dtTmp.Rows.Count; i++)
                {
                    cbo脚注.Items.Add(dtTmp.Rows[i]["名称"].ToString());
                }
            }

        }

        private void frmZyEdit_Shown(object sender, EventArgs e)
        {
            if (myTab.SelectedIndex == 0)
            {
                cbo科别.Focus();
            }
            else
            {

                txtDiag.Focus();
            }
        }

        #endregion

        #region 控件交互

        private void txt方剂_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (txt方剂.TextBoxText != "")
                {
                    DataTable dtTmp = new DataTable();
                    string message = JsonConvert.SerializeObject(new { 证型ID = 0, 匹配项 = txt方剂.TextBoxText });
                    dtTmp = DataCenter.GetTable(m_baseData, "Get_Fjlist", message);
                    if (dtTmp != null)
                    {
                        if (dtTmp.Rows.Count > 0)
                        {
                            frmZySelect frmSelect = new frmZySelect("方剂选择器", dtTmp);
                            frmSelect.ShowDialog(this);
                            //设置处方
                            if (frmSelect.dataRowOut != null)
                            {
                                SetDrugInfo(frmSelect.dataRowOut);
                                txt方剂.TextBoxText = frmSelect.dataRowOut["方剂名称"].ToString();
                                txt方剂.TextBoxToolTipText = frmSelect.dataRowOut["方剂名称"].ToString();
                                txt方剂.Tag = frmSelect.dataRowOut["ID"].ToString();
                            }
                        }
                        else
                        {
                            MessageBox.Show("当前查询未获取有效的方剂，请重新查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txt方剂.Focus();
                        }
                    }
                    else
                    {
                        MessageBox.Show("当前查询未获取有效的方剂，请重新查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txt方剂.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("请录入方剂的名称或者简码进行匹配查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt方剂.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
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
                        if (txt方剂.TextBoxText != "")
                        {
                            DataTable dtTmp = new DataTable();
                            string message = JsonConvert.SerializeObject(new { 证型ID = 0, 匹配项 = txt方剂.TextBoxText });
                            dtTmp = DataCenter.GetTable(m_baseData, "Get_Fjlist", message);
                            if (dtTmp != null)
                            {
                                if (dtTmp.Rows.Count > 0)
                                {
                                    frmZySelect frmSelect = new frmZySelect("方剂选择器", dtTmp);
                                    frmSelect.ShowDialog(this);
                                    //设置处方
                                    if (frmSelect.dataRowOut != null)
                                    {
                                        SetDrugInfo(frmSelect.dataRowOut);
                                        txt方剂.TextBoxText = frmSelect.dataRowOut["方剂名称"].ToString();
                                        txt方剂.TextBoxToolTipText = frmSelect.dataRowOut["方剂名称"].ToString();
                                        txt方剂.Tag = frmSelect.dataRowOut["ID"].ToString();
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("当前查询未获取有效的方剂，请重新查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    txt方剂.Focus();
                                }
                            }
                            else
                            {
                                MessageBox.Show("当前查询未获取有效的方剂，请重新查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                txt方剂.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("请录入方剂的名称或者简码进行匹配查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txt方剂.Focus();
                        }
                    }
                    else
                    {
                        dgvList.Focus();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void txt方剂_Validating(object sender, CancelEventArgs e)
        {
            if (txt方剂.TextBoxToolTipText != txt方剂.TextBoxText)
            {
                txt方剂.TextBoxText = txt方剂.TextBoxToolTipText;
            }
        }

        private void myTab_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (myTab.SelectedIndex == 1)
            {
                txtDiag.Focus();
            }
            else
            {
                cbo科别.Focus();
            }

        }

        private void txtDiag_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (txtDiag.TextBoxText != "")
                {
                    DataTable dtTmp = new DataTable();
                    string message = JsonConvert.SerializeObject(new { 匹配项 = txtDiag.TextBoxText });
                    dtTmp = DataCenter.GetTable(m_baseData, "Get_Diaglist", message);
                    if (dtTmp != null)
                    {
                        if (dtTmp.Rows.Count > 0)
                        {
                            frmZySelect frmSelect = new frmZySelect("诊断选择器", dtTmp);
                            frmSelect.ShowDialog(this);
                            //设置处方
                            if (frmSelect.dataRowOut != null)
                            {
                                txtDiag.TextBoxText = frmSelect.dataRowOut["诊断名称"].ToString();
                                txtDiag.TextBoxToolTipText = frmSelect.dataRowOut["诊断名称"].ToString();
                                txtDiag.Tag = frmSelect.dataRowOut["ID"].ToString();
                            }
                        }
                        else
                        {
                            MessageBox.Show("当前查询未获取有效的诊断，请重新查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txtDiag.Focus();
                        }
                    }
                    else
                    {
                        MessageBox.Show("当前查询未获取有效的诊断，请重新查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtDiag.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("请录入诊断的名称或者简码进行匹配查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtDiag.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void txtDiag_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {

                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txtDiag.TextBoxText != txtDiag.TextBoxToolTipText || txtDiag.TextBoxToolTipText == "")
                    {
                        if (txtDiag.TextBoxText != "")
                        {
                            DataTable dtTmp = new DataTable();
                            string message = JsonConvert.SerializeObject(new { 匹配项 = txtDiag.TextBoxText });
                            dtTmp = DataCenter.GetTable(m_baseData, "Get_Diaglist", message);
                            if (dtTmp != null)
                            {
                                if (dtTmp.Rows.Count > 0)
                                {
                                    frmZySelect frmSelect = new frmZySelect("诊断选择器", dtTmp);
                                    frmSelect.ShowDialog(this);
                                    //设置处方
                                    if (frmSelect.dataRowOut != null)
                                    {
                                        txtDiag.TextBoxText = frmSelect.dataRowOut["诊断名称"].ToString();
                                        txtDiag.TextBoxToolTipText = frmSelect.dataRowOut["诊断名称"].ToString();
                                        txtDiag.Tag = frmSelect.dataRowOut["ID"].ToString();
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("当前查询未获取有效的诊断，请重新查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    txtDiag.Focus();
                                }
                            }
                            else
                            {
                                MessageBox.Show("当前查询未获取有效的诊断，请重新查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                txtDiag.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("请录入诊断的名称或者简码进行匹配查询。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txtDiag.Focus();
                        }
                    }
                    else
                    {
                        txt方剂.Focus();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void txtDiag_Validating(object sender, CancelEventArgs e)
        {
            {
                if (txtDiag.TextBoxToolTipText != txtDiag.TextBoxText)
                {
                    txtDiag.TextBoxText = txtDiag.TextBoxToolTipText;
                }
            }
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void myTab_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (dgvList.Rows[0].Cells["草药名"].Value != null)
            {
                if (MessageBox.Show("切换中药就诊方式将清空当前录入的处方,请确认是否继续？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.OK)
                {
                    dgvList.Rows.Clear();
                    dgvList.RowCount = 1;
                }
                else
                {
                    e.Cancel = true;
                }
            }

        }

        //设置科别
        private void cbo科别_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbo科别.Tag == null)
            {
                //清空疾病
                txt疾病.Tag = null;
                txt疾病.TextBoxText = "";
                txt疾病.TextBoxToolTipText = "";
                //清空证型
                txt证型.TextBoxText = "";
                txt证型.Tag = null;
                txt证型.TextBoxToolTipText = "";
            }
        }
        private void cbo科别_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (txt疾病.Enabled && cbo科别.Text != "")
                {
                    txt疾病.Focus();
                }
                else
                {
                    cbo科别.DroppedDown = true;
                }
            }
        }

        //设置疾病
        private void txt疾病_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt疾病.TextBoxText != txt疾病.TextBoxToolTipText || txt疾病.TextBoxToolTipText == "")
                    {
                        mdtDisType.DefaultView.RowFilter = txt疾病.TextBoxText == "" ? (cbo科别.Text == "" ? "" : "科别 = '" + cbo科别.Text + "'") : (cbo科别.Text == "" ? "" : "科别 = '" + cbo科别.Text + "' and") + " (疾病名称 like '%" + txt疾病.TextBoxText + "%' Or 简码 like '%" + txt疾病.TextBoxText + "%')";
                        ShowSelect(txt疾病, mdtDisType, 350);
                    }
                    else
                    {
                        txt证型.Focus();
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        private void txt疾病_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                mdtDisType.DefaultView.RowFilter = cbo科别.Text == "" ? "" : "科别 = '" + cbo科别.Text + "'";
                ShowSelect(txt疾病, mdtDisType, 350);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void txt疾病_Validating(object sender, CancelEventArgs e)
        {
            if (txt疾病.TextBoxToolTipText != txt疾病.TextBoxText)
            {
                txt疾病.TextBoxText = txt疾病.TextBoxToolTipText;
            }
        }

        //设置证型
        private void txt证型_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (txt疾病.Tag != null)
                {
                    if (mdt证型 != null)
                    {
                        mdt证型.DefaultView.RowFilter = "";
                        ShowSelect(txt证型, mdt证型, 350);
                    }
                    else
                    {
                        MessageBox.Show("未获取有效的中医证型。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("请先选择中医疾病。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (txt疾病.Enabled)
                    {
                        txt疾病.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt证型_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {

                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt证型.TextBoxText != txt证型.TextBoxToolTipText || txt证型.TextBoxToolTipText == "")
                    {

                        if (txt疾病.Tag != null)
                        {
                            if (mdt证型 != null)
                            {
                                mdt证型.DefaultView.RowFilter = txt证型.TextBoxText == "" ? "" : "证型名称 like '%" + txt证型.TextBoxText + "%' Or 简码 like '%" + txt证型.TextBoxText + "%'";
                                ShowSelect(txt证型, mdt证型, 350);
                            }
                            else
                            {
                                MessageBox.Show("未获取有效的中医证型。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            MessageBox.Show("请先选择中医疾病。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txt疾病.Focus();
                        }
                    }
                    else
                    {
                        dgvList.Focus();
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt证型_Validating(object sender, CancelEventArgs e)
        {
            if (txt证型.TextBoxToolTipText != txt证型.TextBoxText)
            {
                txt证型.TextBoxText = txt证型.TextBoxToolTipText;
            }
        }

        private void txt付数_Validating(object sender, CancelEventArgs e)
        {
            try
            {
                if (ComLib.StrToInt(txt付数.Text) < 1)
                {
                    MessageBox.Show("当前中药付数录入不正确，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true;
                }
                else
                {
                    //刷新总量
                    foreach (DataGridViewRow row in this.dgvList.Rows)
                    {
                        //计算总量
                        if (row.Cells["草药名"].Value != null)
                        {
                            row.Cells["总量"].Tag = ComLib.StrToDlb(row.Cells["用量"].Value.ToString()) * ComLib.StrToInt(txt付数.Text);
                            row.Cells["总量"].Value = row.Cells["总量"].Tag.ToString() + row.Cells["单位"].Value.ToString();
                            Check库存(row.Index);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt付数_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                txt嘱托.Focus();
            }
            else
            {
                if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == 8)
                {
                    e.Handled = false;
                    return;
                }
                e.Handled = true;
            }
        }

        private void txt症状_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (mdt临症加症 == null)
                {
                    string message = JsonConvert.SerializeObject(new { USETYPE = "0" });
                    mdt临症加症 = DataCenter.GetTable(m_baseData, "Get_Adddis", message);
                }
                if (mdt临症加症 != null)
                {
                    mdt临症加症.DefaultView.RowFilter = "";
                    if (mdt临症加症.Rows.Count > 0)
                    {
                        ShowSelect(txt症状, mdt临症加症, 200);
                    }
                    else
                    {
                        MessageBox.Show("获取临症加症数据失败，服务器返回数据为空。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("获取临症加症数据失败，服务器返回数据为空。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt症状_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt症状.TextBoxText == "")
                    {
                        txt症状.TextBoxText = "";
                        txt症状.TextBoxToolTipText = "";
                        txt症状.Tag = null;

                        //清空治法
                        txt治法.TextBoxText = "";
                        txt治法.Tag = null;
                        txt治法.TextBoxToolTipText = "";

                        //清空草药
                        txt草药.TextBoxText = "";
                        txt草药.Tag = null;
                        txt草药.TextBoxToolTipText = "";
                        btnAdd.Tag = null;

                        txt煎法.Focus();
                        return;
                    }

                    if (txt症状.TextBoxText != txt症状.TextBoxToolTipText || txt症状.TextBoxToolTipText == "")
                    {
                        if (mdt临症加症 == null)
                        {
                            string message = JsonConvert.SerializeObject(new { USETYPE = "0" });
                            mdt临症加症 = DataCenter.GetTable(m_baseData, "Get_Adddis", message);
                        }
                        if (mdt临症加症 != null)
                        {
                            mdt临症加症.DefaultView.RowFilter = txt症状.TextBoxText == "" ? "" : "加症名称 like '%" + txt症状.TextBoxText + "%' Or 简码 like '%" + txt症状.TextBoxText + "%'";
                            if (mdt临症加症.Rows.Count > 0)
                            {
                                ShowSelect(txt症状, mdt临症加症, 200);
                            }
                            else
                            {
                                MessageBox.Show("当前过滤未找到临症加症数据。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                txt症状.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("获取临症加症数据失败，服务器返回数据为空。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txt症状.Focus();
                        }

                    }
                    else
                    {
                        txt治法.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt症状_Validating(object sender, CancelEventArgs e)
        {
            if (txt症状.TextBoxToolTipText != txt症状.TextBoxText)
            {
                txt症状.TextBoxText = txt症状.TextBoxToolTipText;
            }
        }

        private void txt治法_ButtonClick(object sender, EventArgs e)
        {
            try
            {

                if (txt症状.Tag != null)
                {
                    if (mdt加症治法 != null)
                    {
                        mdt加症治法.DefaultView.RowFilter = "";
                        if (mdt加症治法.Rows.Count > 0)
                        {
                            ShowSelect(txt治法, mdt加症治法, 200);
                        }
                        else
                        {
                            MessageBox.Show("未获取有效的加症治法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("未获取有效的加症治法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("请先选择加症症状。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (txt症状.Enabled)
                    {
                        txt症状.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt治法_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {

                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt治法.TextBoxText != txt治法.TextBoxToolTipText || txt治法.TextBoxToolTipText == "")
                    {

                        if (txt症状.Tag != null)
                        {
                            if (mdt加症治法 != null)
                            {
                                mdt加症治法.DefaultView.RowFilter = txt治法.TextBoxText == "" ? "" : "治法名称 like '%" + txt治法.TextBoxText + "%' Or 简码 like '%" + txt治法.TextBoxText + "%'";
                                ShowSelect(txt治法, mdt加症治法, 200);
                            }
                            else
                            {
                                MessageBox.Show("未获取有效的加症治法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            MessageBox.Show("请先选择加症症状。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (txt症状.Enabled)
                            {
                                txt症状.Focus();
                            }
                        }
                    }
                    else
                    {
                        txt草药.Focus();
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt治法_Validating(object sender, CancelEventArgs e)
        {
            if (txt治法.TextBoxToolTipText != txt治法.TextBoxText)
            {
                txt治法.TextBoxText = txt治法.TextBoxToolTipText;
            }
        }

        private void txt草药_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (txt治法.Tag != null)
                {
                    if (mdt加症草药 != null)
                    {
                        mdt加症草药.DefaultView.RowFilter = "";
                        if (mdt加症草药.Rows.Count > 0)
                        {
                            ShowSelect(txt草药, mdt加症草药, 200);
                        }
                        else
                        {
                            MessageBox.Show("未获取有效的加症草药。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("未获取有效的加症草药。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("请先选择加症治法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (txt治法.Enabled)
                    {
                        txt治法.Focus();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt草药_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {

                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt草药.TextBoxText != txt草药.TextBoxToolTipText || txt草药.TextBoxToolTipText == "")
                    {

                        if (txt治法.Tag != null)
                        {
                            if (mdt加症草药 != null)
                            {
                                mdt加症草药.DefaultView.RowFilter = txt草药.TextBoxText == "" ? "" : "草药名称 like '%" + txt草药.TextBoxText + "%' Or 简码 like '%" + txt草药.TextBoxText + "%'";
                                ShowSelect(txt草药, mdt加症草药, 200);
                            }
                            else
                            {
                                MessageBox.Show("未获取有效的加症治法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            MessageBox.Show("请先选择加症治法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (txt治法.Enabled)
                            {
                                txt治法.Focus();
                            }
                        }
                    }
                    else
                    {
                        btnAdd.Focus();
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt草药_Validating(object sender, CancelEventArgs e)
        {
            if (txt草药.TextBoxToolTipText != txt草药.TextBoxText)
            {
                txt草药.TextBoxText = txt草药.TextBoxToolTipText;
            }
        }
        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                if (txt草药.Tag != null)
                {
                    //检查是否有重复的中药

                    foreach (DataGridViewRow row in this.dgvList.Rows)
                    {
                        if (row.Cells["草药名"].Tag != null)
                        {
                            if (txt草药.Tag.ToString() == row.Cells["草药名"].Tag.ToString())
                            {
                                MessageBox.Show(@"中草药 """ + txt草药.TextBoxText + @""" 已经录入，不能添加相同的草药。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                        }
                    }

                    if (dgvList.Rows[0].Cells["草药名"].Value == null)
                    {
                        dgvList.Rows.Remove(dgvList.Rows[0]);
                    }
                    DataGridViewRow Dgvr = this.dgvList.Rows[this.dgvList.Rows.Add()];
                    Dgvr.Cells["来源"].Value = txt症状.TextBoxText + "-" + txt治法.TextBoxText;
                    Dgvr.Cells["草药名"].Value = txt草药.TextBoxText;
                    Dgvr.Cells["草药名"].Tag = txt草药.Tag;
                    Dgvr.Cells[0].Tag = txt草药.TextBoxText;

                    Dgvr.Cells["用量"].Value = btnAdd.Tag.ToString().Split('|')[0].ToString();
                    Dgvr.Cells["单位"].Value = btnAdd.Tag.ToString().Split('|')[1].ToString();
                    Dgvr.Cells["单位"].Tag = btnAdd.Tag.ToString().Split('|')[2].ToString();

                    //计算总量
                    Dgvr.Cells["总量"].Tag = ComLib.StrToDlb(Dgvr.Cells["用量"].Value.ToString()) * ComLib.StrToInt(txt付数.Text);
                    Dgvr.Cells["总量"].Value = (Dgvr.Cells["总量"].Tag).ToString() + Dgvr.Cells["单位"].Value.ToString();

                    //把窗体上的picturebox控件赋值给单元格

                    dgvList.Rows[dgvList.RowCount - 1].Cells[0].Value = this.picAdd.BackgroundImage;

                    //选中最后一行

                    dgvList.Focus();
                    dgvList.CurrentCell = dgvList.Rows[dgvList.RowCount - 1].Cells[9];
                    dgvList.CurrentCell = dgvList.Rows[dgvList.RowCount - 1].Cells[3];
                    dgvList.Rows[dgvList.RowCount - 1].Selected = true;

                    Load药品库存(1, dgvList.RowCount - 1);

                    Dgvr.Dispose();

                    //清空草药信息
                    txt草药.TextBoxText = "";
                    txt草药.Tag = null;
                    txt草药.TextBoxToolTipText = "";
                    btnAdd.Tag = null;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void txt煎法_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (mdt煎法 == null)
                {
                    string message = JsonConvert.SerializeObject(new { 科室ID = m_baseData.DeptID, 站点 = m_baseData.站点 });
                    mdt煎法 = DataCenter.GetTable(m_baseData, "Get_Jftype", message);
                }
                if (mdt煎法 != null)
                {
                    mdt煎法.DefaultView.RowFilter = "";
                    if (mdt煎法.Rows.Count > 0)
                    {
                        ShowSelect(txt煎法, mdt煎法, 200);
                    }
                    else
                    {
                        MessageBox.Show("未获取到有效的中药煎法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("未获取到有效的中药煎法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt煎法_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt煎法.TextBoxText != txt煎法.TextBoxToolTipText || txt煎法.TextBoxToolTipText == "")
                    {
                        if (mdt煎法 == null)
                        {
                            string message = JsonConvert.SerializeObject(new { 科室ID = m_baseData.DeptID, 站点 = m_baseData.站点 });
                            mdt煎法 = DataCenter.GetTable(m_baseData, "Get_Jftype", message);
                        }
                        if (mdt煎法 != null)
                        {
                            mdt煎法.DefaultView.RowFilter = txt煎法.TextBoxText == "" ? "" : "煎法名称 like '%" + txt煎法.TextBoxText + "%' Or 简码 like '%" + txt煎法.TextBoxText + "%'";
                            if (mdt煎法.Rows.Count > 0)
                            {
                                ShowSelect(txt煎法, mdt煎法, 200);
                            }
                            else
                            {
                                MessageBox.Show("当前过滤未找到有效的中药煎法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                txt煎法.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("未获取到有效的中药煎法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txt煎法.Focus();
                        }

                    }
                    else
                    {
                        txt煎量.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt煎法_Validating(object sender, CancelEventArgs e)
        {
            if (txt煎法.TextBoxToolTipText != txt煎法.TextBoxText)
            {
                txt煎法.TextBoxText = txt煎法.TextBoxToolTipText;
            }
        }

        private void txt用法_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (mdt用法 == null)
                {
                    string message = JsonConvert.SerializeObject(new { 科室ID = m_baseData.DeptID, 站点 = m_baseData.站点 });
                    mdt用法 = DataCenter.GetTable(m_baseData, "Get_Usetype", message);
                }
                if (mdt用法 != null)
                {
                    mdt用法.DefaultView.RowFilter = "";
                    if (mdt用法.Rows.Count > 0)
                    {
                        ShowSelect(txt用法, mdt用法, 200);
                    }
                    else
                    {
                        MessageBox.Show("未获取到有效的中药用法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("未获取到有效的中药用法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt用法_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt用法.TextBoxText != txt用法.TextBoxToolTipText || txt用法.TextBoxToolTipText == "")
                    {
                        if (mdt用法 == null)
                        {
                            string message = JsonConvert.SerializeObject(new { 科室ID = m_baseData.DeptID, 站点 = m_baseData.站点 });
                            mdt用法 = DataCenter.GetTable(m_baseData, "Get_Usetype", message);
                        }
                        if (mdt用法 != null)
                        {
                            mdt用法.DefaultView.RowFilter = txt用法.TextBoxText == "" ? "" : "用法名称 like '%" + txt用法.TextBoxText + "%' Or 简码 like '%" + txt用法.TextBoxText + "%'";
                            if (mdt用法.Rows.Count > 0)
                            {
                                ShowSelect(txt用法, mdt用法, 200);
                            }
                            else
                            {
                                MessageBox.Show("当前过滤未找到有效的中药用法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                txt用法.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("未获取到有效的中药用法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txt用法.Focus();
                        }

                    }
                    else
                    {
                        txt频率.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt用法_Validating(object sender, CancelEventArgs e)
        {
            if (txt用法.TextBoxToolTipText != txt用法.TextBoxText)
            {
                txt用法.TextBoxText = txt用法.TextBoxToolTipText;
            }
        }

        private void txt频率_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (mdt频率 == null)
                {
                    mdt频率 = DataCenter.GetTable(m_baseData, "Get_UseTime", "");
                }
                if (mdt频率 != null)
                {
                    mdt频率.DefaultView.RowFilter = "";
                    if (mdt频率.Rows.Count > 0)
                    {
                        ShowSelect(txt频率, mdt频率, 200);
                    }
                    else
                    {
                        MessageBox.Show("未获取到有效的中药用法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("未获取到有效的中药用法。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt频率_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt频率.TextBoxText != txt频率.TextBoxToolTipText || txt频率.TextBoxToolTipText == "")
                    {
                        if (mdt频率 == null)
                        {
                            mdt频率 = DataCenter.GetTable(m_baseData, "Get_UseTime", "");
                        }
                        if (mdt频率 != null)
                        {
                            mdt频率.DefaultView.RowFilter = txt频率.TextBoxText == "" ? "" : "频率名称 like '%" + txt频率.TextBoxText + "%' Or 简码 like '%" + txt频率.TextBoxText + "%'";
                            if (mdt频率.Rows.Count > 0)
                            {
                                ShowSelect(txt频率, mdt频率, 200);
                            }
                            else
                            {
                                MessageBox.Show("当前过滤未找到有效的中药频率。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                txt频率.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("未获取到有效的中药频率。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txt频率.Focus();
                        }

                    }
                    else
                    {
                        cbo药房.Focus();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt频率_Validating(object sender, CancelEventArgs e)
        {
            if (txt频率.TextBoxToolTipText != txt频率.TextBoxText)
            {
                txt频率.TextBoxText = txt频率.TextBoxToolTipText;
            }
        }

        private void txt煎量_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                txt用法.Focus();
            }
        }

        private void txt嘱托_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                btnSave.Focus();
            }
        }

        private void cbo药房_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (cbo药房.Enabled && cbo药房.Text != "")
                {
                    txt付数.Focus();
                }
                else
                {
                    cbo药房.DroppedDown = true;
                }
            }
            else
            {
                cbo药房.DroppedDown = true;
                e.Handled = true;
            }
        }


        //表格输入框绑定KeyPress事件
        private void dgvList_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            try
            {
                if (dgvList.Columns[dgvList.CurrentCellAddress.X].Name == "用量")//获取当前处于活动状态的单元格索引

                {
                    CellEdit = (DataGridViewTextBoxEditingControl)e.Control;
                    CellEdit.SelectAll();
                    CellEdit.KeyPress += Cells_KeyPress; //绑定事件
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //表格输入项限制

        private void Cells_KeyPress(object sender, KeyPressEventArgs e) //自定义事件

        {

            if (dgvList.Columns[dgvList.CurrentCellAddress.X].Name == "用量")//获取当前处于活动状态的单元格索引

            {
                if (mbln保密 == false)
                {
                    if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == 8)
                    {
                        e.Handled = false;
                        return;
                    }
                }
                e.Handled = true;

            }
        }

        //表格光标纵向跳转
        private bool m_isdataGridViewBegingEdit = false;

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            try
            {
                if (keyData == Keys.Enter && (m_isdataGridViewBegingEdit == true || this.dgvList.Focused == true || this.cbo脚注.Focused == true || this.cbo规格.Focused == true))
                {
                    if (dgvList.CurrentCell.ColumnIndex < dgvList.Columns["用量"].Index)
                    {
                        dgvList.CurrentCell = this.dgvList.Rows[dgvList.CurrentRow.Index].Cells["用量"];
                    }
                    else if (dgvList.CurrentCell.ColumnIndex >= dgvList.Columns["用量"].Index && dgvList.CurrentCell.ColumnIndex < dgvList.Columns["脚注"].Index)
                    {
                        dgvList.CurrentCell = this.dgvList.Rows[dgvList.CurrentRow.Index].Cells["脚注"];
                    }
                    else if (dgvList.CurrentCell.ColumnIndex >= dgvList.Columns["脚注"].Index && dgvList.CurrentCell.ColumnIndex < dgvList.Columns["规格"].Index)
                    {
                        dgvList.CurrentCell = this.dgvList.Rows[dgvList.CurrentRow.Index].Cells["规格"];
                    }
                    else if (dgvList.CurrentCell.ColumnIndex >= dgvList.Columns["规格"].Index)
                    {
                        if (dgvList.CurrentRow.Index + 1 >= dgvList.RowCount)
                        {
                            if (mbln保密) {
                                txt煎法.Focus();
                            }
                            else { 
                                txt症状.Focus();
                            }
                        }
                        else
                        {
                            dgvList.CurrentCell = this.dgvList.Rows[dgvList.CurrentRow.Index + 1].Cells["用量"];
                        }

                    }

                    return true;
                }
            }
            catch { }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void dgvList_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            m_isdataGridViewBegingEdit = false;
        }

        private void dgvList_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            m_isdataGridViewBegingEdit = true;
        }

        private void dgvList_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                if (e.RowIndex > -1 && e.ColumnIndex > -1 && m_isdataGridViewBegingEdit)
                {
                    DataGridView grid = (DataGridView)sender;
                    grid.Rows[e.RowIndex].ErrorText = "";

                    if (grid.Columns[e.ColumnIndex].Name == "用量" && grid.Rows[e.RowIndex].Cells["草药名"].Value != null)
                    {
                        try
                        {
                            double newDB = 0;
                            if (!double.TryParse(e.FormattedValue.ToString(), out newDB))
                            {
                                e.Cancel = true;
                                grid.Rows[e.RowIndex].ErrorText = "请输入有效的用量";
                                MessageBox.Show("单味用量输入错误，请输入有效的用量！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
                            else
                            {
                                if (newDB <= 0 || newDB > 99999)
                                {
                                    e.Cancel = true;
                                    grid.Rows[e.RowIndex].ErrorText = "请输入大于0的整数";
                                    MessageBox.Show("不是大于零的数字或输入数值过大！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    return;
                                }
                                else
                                {
                                    //计算总量
                                    e.Cancel = false;
                                    dgvList.Rows[e.RowIndex].Cells["用量"].Tag = Math.Round(ComLib.StrToDlb(e.FormattedValue.ToString()), 4);
                                    dgvList.Rows[e.RowIndex].Cells["总量"].Tag = Math.Round(ComLib.StrToDlb(e.FormattedValue.ToString()), 4) * ComLib.StrToInt(txt付数.Text);
                                    dgvList.Rows[e.RowIndex].Cells["总量"].Value = dgvList.Rows[e.RowIndex].Cells["总量"].Tag.ToString() + dgvList.Rows[e.RowIndex].Cells["单位"].Value.ToString();
                                    Check库存(e.RowIndex);
                                }
                            }
                        }
                        catch
                        {
                            e.Cancel = true;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgvList_CurrentCellChanged(object sender, EventArgs e)
        {
            try
            {
                DataGridViewCell CurrnetCell = this.dgvList.CurrentCell;

                //同步combox控件数据
                if (cbo脚注.Visible && cbo脚注.Tag != null)
                {
                    if (ComLib.StrToInt(cbo脚注.Tag.ToString()) >= 0)
                    {
                        dgvList.Rows[ComLib.StrToInt(cbo脚注.Tag.ToString())].Cells["脚注"].Value = cbo脚注.Text;
                        cbo脚注.Tag = null;
                    }
                }

                if (cbo规格.Visible && cbo规格.Tag != null)
                {
                    if (ComLib.StrToInt(cbo规格.Tag.ToString()) >= 0)
                    {
                        dgvList.Rows[ComLib.StrToInt(cbo规格.Tag.ToString())].Cells["规格"].Value = cbo规格.Text;
                        cbo规格.Tag = null;
                    }
                }
                this.cbo脚注.Visible = false;
                this.cbo规格.Visible = false;


                if (CurrnetCell != null && CurrnetCell.OwningColumn.Name == "脚注")
                {
                    Rectangle Rect = this.dgvList.GetCellDisplayRectangle(CurrnetCell.ColumnIndex, CurrnetCell.RowIndex, true);
                    cbo脚注.Visible = (mbln保密 == false);
                    cbo脚注.Top = Rect.Top;
                    cbo脚注.Left = Rect.Left;
                    cbo脚注.Height = Rect.Height;
                    cbo脚注.Width = Rect.Width;
                    //缓存上次注脚
                    cbo脚注.Tag = CurrnetCell.RowIndex;

                    if (dgvList.Rows[CurrnetCell.RowIndex].Cells[CurrnetCell.ColumnIndex].Value != null)
                    {
                        cbo脚注.Text = dgvList.Rows[CurrnetCell.RowIndex].Cells[CurrnetCell.ColumnIndex].Value.ToString();
                    }
                    else
                    {
                        cbo脚注.Text = "";
                    }
                    cbo脚注.Focus();
                    cbo脚注.Refresh();
                    this.cbo规格.Visible = false;
                }
                else if (CurrnetCell != null && CurrnetCell.OwningColumn.Name == "规格")
                {
                    Rectangle Rect = this.dgvList.GetCellDisplayRectangle(CurrnetCell.ColumnIndex, CurrnetCell.RowIndex, true);
                    cbo规格.Top = Rect.Top;
                    cbo规格.Left = Rect.Left;
                    cbo规格.Height = Rect.Height;
                    cbo规格.Width = Rect.Width;
                    cbo规格.DropDownWidth = Rect.Width;

                    //重新刷新规格选择项

                    cbo规格.Items.Clear();
                    if (dgvList.Rows[CurrnetCell.RowIndex].Cells["脚注"].Tag != null)
                    {
                        DataTable dtTmp = (DataTable)dgvList.Rows[CurrnetCell.RowIndex].Cells["脚注"].Tag;
                        int dWidth = cbo规格.Width, cTemp = 0;
                        Graphics g = cbo规格.CreateGraphics();
                        for (int i = 0; i < dtTmp.Rows.Count; i++)
                        {
                            itemEx item = new itemEx(dtTmp.Rows[i]["药品id"].ToString(), "[" + dtTmp.Rows[i]["编码"].ToString() + "]" + ((mbln保密 == false ? true : false) ? dtTmp.Rows[i]["名称"].ToString() : "***") + (dtTmp.Rows[i]["规格"].ToString() == string.Empty ? "" : " " + dtTmp.Rows[i]["规格"].ToString()) + (dtTmp.Rows[i]["产地"].ToString() == string.Empty ? "" : " 产地：" + dtTmp.Rows[i]["产地"].ToString()), dtTmp.Rows[i]["可用数量"].ToString());
                            //计算宽度
                            cTemp = (int)g.MeasureString(item.Text, cbo规格.Font).Width;
                            if (cTemp > dWidth)
                                dWidth = cTemp;
                            cbo规格.Items.Add(item);
                        }
                        cbo规格.DropDownWidth = (int)((double)dWidth * 1.1);
                    }

                    cbo规格.Visible = true;
                    //缓存上次注脚
                    cbo规格.Tag = CurrnetCell.RowIndex;

                    if (dgvList.Rows[CurrnetCell.RowIndex].Cells[CurrnetCell.ColumnIndex].Value != null)
                    {
                        cbo规格.Text = dgvList.Rows[CurrnetCell.RowIndex].Cells[CurrnetCell.ColumnIndex].Value.ToString();
                    }
                    else
                    {
                        cbo规格.Text = "";
                    }
                    cbo规格.Focus();
                    cbo规格.Refresh();
                    this.cbo脚注.Visible = false;
                }
                else
                {
                    this.cbo脚注.Visible = false;
                    this.cbo规格.Visible = false;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgvList_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            try
            {
                if (cbo脚注.Visible)
                {
                    this.cbo脚注.Visible = false;
                    if (cbo脚注.Tag != null)
                    {
                        if (ComLib.StrToInt(cbo脚注.Tag.ToString()) >= 0)
                        {
                            dgvList.Rows[ComLib.StrToInt(cbo脚注.Tag.ToString())].Cells["脚注"].Value = cbo脚注.Text;
                            cbo脚注.Tag = null;
                        }
                    }
                }
                if (cbo规格.Visible)
                {
                    this.cbo规格.Visible = false;
                    if (cbo规格.Tag != null)
                    {
                        if (ComLib.StrToInt(cbo规格.Tag.ToString()) >= 0)
                        {
                            dgvList.Rows[ComLib.StrToInt(cbo规格.Tag.ToString())].Cells["规格"].Value = cbo规格.Text;
                            cbo规格.Tag = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void cbo脚注_Leave(object sender, EventArgs e)
        {
            dgvList.Focus();
        }

        private void cbo规格_Leave(object sender, EventArgs e)
        {
            dgvList.Focus();
        }

        //设置删除图标
        private void dgvList_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvList.ColumnCount == 10)
            {
                try
                {
                    //把窗体上的picturebox控件赋值给单元格

                    foreach (DataGridViewRow row in this.dgvList.Rows)
                    {
                        if (row.Index == dgvList.CurrentRow.Index)
                        {
                            row.Cells[9].Value = this.picDel.Image;
                        }
                        else
                        {
                            row.Cells[9].Value = this.picNone.Image;
                        }
                    }
                }
                catch { }
            }
        }

        private void dgvList_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 9 && dgvList.RowCount > 0)
                {
                    if (mbln保密 == false)
                    {
                        if (dgvList.Rows[e.RowIndex].Cells["草药名"].Value == null)
                        {
                            dgvList.Rows.Remove(dgvList.Rows[e.RowIndex]);
                        }
                        else
                        {
                            if (MessageBox.Show(@"确定要删除草药 """ + dgvList.Rows[e.RowIndex].Cells["草药名"].Value.ToString() + @""" 吗？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.OK)
                            {
                                dgvList.Rows.Remove(dgvList.Rows[e.RowIndex]);
                            }
                        }
                    }
                }
                //保持初始行

                if (dgvList.RowCount == 0)
                {
                    dgvList.RowCount = 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void cbo规格_KeyPress(object sender, KeyPressEventArgs e)
        {
            cbo规格.DroppedDown = true;
            e.Handled = true;
        }

        private void cbo规格_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbo规格.Items.Count > 0 && cbo规格.Tag != null)
                {
                    itemEx item = (itemEx)cbo规格.SelectedItem;
                    dgvList.Rows[ComLib.StrToInt(cbo规格.Tag.ToString())].Cells["规格"].Tag = item.Tag.ToString();

                    //设置库存
                    dgvList.Rows[ComLib.StrToInt(cbo规格.Tag.ToString())].Cells["库存情况"].Tag = item.ItemData.ToString();
                    Check库存(ComLib.StrToInt(cbo规格.Tag.ToString()));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Check库存(int intRow)
        {
            try
            {
                if (ComLib.StrToDlb(dgvList.Rows[intRow].Cells["单位"].Tag == null ? "0" : dgvList.Rows[intRow].Cells["单位"].Tag.ToString()) <= 0)
                {
                    dgvList.Rows[intRow].Cells["库存情况"].Value = "尚未对码";
                    dgvList.Rows[intRow].Cells["库存情况"].Style.ForeColor = Color.Red;
                }
                else if (ComLib.StrToDlb(dgvList.Rows[intRow].Cells["库存情况"].Tag == null ? "0" : dgvList.Rows[intRow].Cells["库存情况"].Tag.ToString()) <= 0)
                {
                    dgvList.Rows[intRow].Cells["库存情况"].Value = "没有库存";
                    dgvList.Rows[intRow].Cells["库存情况"].Style.ForeColor = Color.Red;
                }
                else if (ComLib.StrToDlb(dgvList.Rows[intRow].Cells["库存情况"].Tag.ToString()) > ComLib.StrToDlb(dgvList.Rows[intRow].Cells["总量"].Tag == null ? "0" : dgvList.Rows[intRow].Cells["总量"].Tag.ToString()))
                {
                    dgvList.Rows[intRow].Cells["库存情况"].Value = "库存充足";
                    dgvList.Rows[intRow].Cells["库存情况"].Style.ForeColor = Color.Black;
                }
                else
                {
                    dgvList.Rows[intRow].Cells["库存情况"].Value = "库存不足";
                    dgvList.Rows[intRow].Cells["库存情况"].Style.ForeColor = Color.Red;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void cbo药房_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbo药房.Items.Count > 0)
            {
                Load药品库存();
            }
        }

        private void dgvList_Leave(object sender, EventArgs e)
        {
            try
            {
                if (cbo脚注.Visible)
                {
                    this.cbo脚注.Visible = false;
                    if (cbo脚注.Tag != null)
                    {
                        if (ComLib.StrToInt(cbo脚注.Tag.ToString()) >= 0)
                        {
                            dgvList.Rows[ComLib.StrToInt(cbo脚注.Tag.ToString())].Cells["脚注"].Value = cbo脚注.Text;
                            cbo脚注.Tag = null;
                        }
                    }
                }
                if (cbo规格.Visible)
                {
                    this.cbo规格.Visible = false;
                    if (cbo规格.Tag != null)
                    {
                        if (ComLib.StrToInt(cbo规格.Tag.ToString()) >= 0)
                        {
                            dgvList.Rows[ComLib.StrToInt(cbo规格.Tag.ToString())].Cells["规格"].Value = cbo规格.Text;
                            cbo规格.Tag = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void txtSex_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt诊断_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt煎量_Enter(object sender, EventArgs e)
        {
            txt煎量.SelectAll();
        }

        private void txt付数_Enter(object sender, EventArgs e)
        {
            txt付数.SelectAll();
        }

        private void txt嘱托_Enter(object sender, EventArgs e)
        {
            txt嘱托.SelectAll();
        }

        private void cbo脚注_Enter(object sender, EventArgs e)
        {
            cbo脚注.SelectAll();
        }

        private void cbo规格_Enter(object sender, EventArgs e)
        {
            cbo规格.SelectAll();
        }

        private void dgvList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                try
                {

                    if (dgvList.CurrentRow.Index >= 0 && dgvList.RowCount > 0)
                    {
                        if (mbln保密 == false)
                        {
                            if (dgvList.Rows[dgvList.CurrentRow.Index].Cells["草药名"].Value == null)
                            {
                                dgvList.Rows.Remove(dgvList.Rows[dgvList.CurrentRow.Index]);
                            }
                            else
                            {
                                if (MessageBox.Show(@"确定要删除草药 """ + dgvList.Rows[dgvList.CurrentRow.Index].Cells["草药名"].Value.ToString() + @""" 吗？", this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.OK)
                                {
                                    dgvList.Rows.Remove(dgvList.Rows[dgvList.CurrentRow.Index]);
                                }
                            }
                        }
                    }
                    //保持初始行

                    if (dgvList.RowCount == 0)
                    {
                        dgvList.RowCount = 1;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void dgvList_Scroll(object sender, ScrollEventArgs e)
        {
            try
            {
                if (cbo脚注.Visible)
                {
                    this.cbo脚注.Visible = false;
                    if (cbo脚注.Tag != null)
                    {
                        if (ComLib.StrToInt(cbo脚注.Tag.ToString()) >= 0)
                        {
                            dgvList.Rows[ComLib.StrToInt(cbo脚注.Tag.ToString())].Cells["脚注"].Value = cbo脚注.Text;
                            cbo脚注.Tag = null;
                        }
                    }
                }
                if (cbo规格.Visible)
                {
                    this.cbo规格.Visible = false;
                    if (cbo规格.Tag != null)
                    {
                        if (ComLib.StrToInt(cbo规格.Tag.ToString()) >= 0)
                        {
                            dgvList.Rows[ComLib.StrToInt(cbo规格.Tag.ToString())].Cells["规格"].Value = cbo规格.Text;
                            cbo规格.Tag = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }


        #endregion

        #region 表格选择器处理

        private void dgvSelect_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    SetSelectData(1);
                    cboSelect.CloseDiag();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //选中数据进行设置
        private void SetSelectData(int intType = 0)
        {
            try
            {
                if (dgvSelect.Rows.Count > 0 && dgvSelect.CurrentCell != null)
                {
                    if (dgvSelect.CurrentCell.RowIndex >= 0)
                    {
                        int index = 0;
                        if (intType == 1)
                        {
                            index = ComLib.StrToInt(dgvSelect.Tag.ToString()); //获取选中行的行号
                        }
                        else
                        {
                            index = dgvSelect.CurrentRow.Index; //获取选中行的行号
                        }
                        index = index < 0 ? 0 : index;
                        mtxtTmp.Tag = dgvSelect.Rows[index].Cells["ID"].Value.ToString();

                        switch (mtxtTmp.Name)
                        {
                            case "txt疾病":
                                mtxtTmp.TextBoxToolTipText = dgvSelect.Rows[index].Cells["疾病名称"].Value.ToString();
                                mtxtTmp.TextBoxText = mtxtTmp.TextBoxToolTipText;
                                if (cbo科别.Text == "")
                                {
                                    cbo科别.Tag = "1";
                                    cbo科别.Text = dgvSelect.Rows[index].Cells["科别"].Value.ToString();
                                    cbo科别.Tag = null;
                                }
                                //清空证型
                                txt证型.TextBoxText = "";
                                txt证型.Tag = null;
                                txt证型.TextBoxToolTipText = "";
                                //选择疾病时缓存证型种类

                                if (mtxtTmp.Tag != null)
                                {
                                    string message = JsonConvert.SerializeObject(new { 疾病ID = mtxtTmp.Tag });
                                    mdt证型 = DataCenter.GetTable(m_baseData, "Get_Zxtype", message);
                                }
                                txt证型.Focus();
                                break;
                            case "txt证型":
                                mtxtTmp.TextBoxToolTipText = dgvSelect.Rows[index].Cells["证型名称"].Value.ToString();
                                mtxtTmp.TextBoxText = mtxtTmp.TextBoxToolTipText;
                                txt病人治法.Text = dgvSelect.Rows[index].Cells["证型治法"].Value.ToString();
                                //缓存疾病ID-证型ID
                                txt诊断.Text = txt疾病.TextBoxToolTipText.Replace(" ", "") + "-" + mtxtTmp.TextBoxToolTipText;

                                //清空当前处方
                                dgvList.Rows.Clear();
                                dgvList.RowCount = 1;

                                //选择证型后弹出方剂选择器

                                if (mtxtTmp.Tag != null)
                                {
                                    DataTable dtTmp = new DataTable();
                                    string message = JsonConvert.SerializeObject(new { 证型ID = mtxtTmp.Tag.ToString() });
                                    dtTmp = DataCenter.GetTable(m_baseData, "Get_Fjlist", message);
                                    if (dtTmp != null)
                                    {
                                        if (dtTmp.Rows.Count > 0)
                                        {
                                            frmZySelect frmSelect = new frmZySelect("方剂选择器", dtTmp);
                                            frmSelect.ShowDialog(this);
                                            //设置处方
                                            SetDrugInfo(frmSelect.dataRowOut);
                                        }
                                        else
                                        {
                                            MessageBox.Show("当前证型未获取有效的方剂，请重新选择。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            mtxtTmp.Focus();
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("当前证型未获取有效的方剂，请重新选择。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        mtxtTmp.Focus();
                                    }
                                }

                                break;
                            case "txt症状":
                                mtxtTmp.TextBoxToolTipText = dgvSelect.Rows[index].Cells["加症名称"].Value.ToString();
                                mtxtTmp.TextBoxText = mtxtTmp.TextBoxToolTipText;
                                //清空治法
                                txt治法.TextBoxText = "";
                                txt治法.Tag = null;
                                txt治法.TextBoxToolTipText = "";

                                //清空草药
                                txt草药.TextBoxText = "";
                                txt草药.Tag = null;
                                txt草药.TextBoxToolTipText = "";
                                btnAdd.Tag = null;

                                //选择症状时缓存治法种类

                                if (mtxtTmp.Tag != null)
                                {
                                    string message = JsonConvert.SerializeObject(new { 加症ID = mtxtTmp.Tag });
                                    mdt加症治法 = DataCenter.GetTable(m_baseData, "Get_Addzf", message);
                                }
                                txt治法.Focus();
                                break;
                            case "txt治法":
                                mtxtTmp.TextBoxToolTipText = dgvSelect.Rows[index].Cells["治法名称"].Value.ToString();
                                mtxtTmp.TextBoxText = mtxtTmp.TextBoxToolTipText;

                                //清空草药
                                txt草药.TextBoxText = "";
                                txt草药.Tag = null;
                                txt草药.TextBoxToolTipText = "";
                                btnAdd.Tag = null;

                                //选择治法时加症草药

                                if (mtxtTmp.Tag != null)
                                {
                                    string message = JsonConvert.SerializeObject(new { 治法ID = mtxtTmp.Tag });
                                    mdt加症草药 = DataCenter.GetTable(m_baseData, "Get_Additems", message);
                                }
                                txt草药.Focus();
                                break;
                            case "txt草药":
                                mtxtTmp.TextBoxToolTipText = dgvSelect.Rows[index].Cells["草药名称"].Value.ToString();
                                mtxtTmp.TextBoxText = mtxtTmp.TextBoxToolTipText;
                                btnAdd.Tag = dgvSelect.Rows[index].Cells["用量"].Value.ToString() + "|" + dgvSelect.Rows[index].Cells["单位"].Value.ToString() + "|" + dgvSelect.Rows[index].Cells["HIS品种ID"].Value.ToString();
                                break;
                            case "txt煎法":
                                mtxtTmp.TextBoxToolTipText = dgvSelect.Rows[index].Cells["煎法名称"].Value.ToString();
                                mtxtTmp.TextBoxText = mtxtTmp.TextBoxToolTipText;
                                txt煎量.Focus();
                                break;
                            case "txt用法":
                                mtxtTmp.TextBoxToolTipText = dgvSelect.Rows[index].Cells["用法名称"].Value.ToString();
                                mtxtTmp.TextBoxText = mtxtTmp.TextBoxToolTipText;
                                txt频率.Focus();
                                break;
                            case "txt频率":
                                mtxtTmp.TextBoxToolTipText = dgvSelect.Rows[index].Cells["频率名称"].Value.ToString();
                                mtxtTmp.TextBoxText = mtxtTmp.TextBoxToolTipText;
                                mtxtTmp.Tag = dgvSelect.Rows[index].Cells["频率次数"].Value.ToString() + "|" + dgvSelect.Rows[index].Cells["频率间隔"].Value.ToString() + "|" + dgvSelect.Rows[index].Cells["间隔单位"].Value.ToString();
                                cbo药房.Focus();
                                break;
                            default:
                                Console.WriteLine("NO");
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

        //根据方剂选择设置处方
        private void SetDrugInfo(DataRow DrSelect)
        {
            try
            {
                dgvList.Tag = null;
                if (DrSelect != null)
                {
                    string message = JsonConvert.SerializeObject(new { 方剂ID = DrSelect["ID"].ToString() });
                    mdt方剂组成 = DataCenter.GetTable(m_baseData, "Get_Fjitems", message);
                    if (mdt方剂组成 != null)
                    {
                        if (mdt方剂组成.Rows.Count > 0)
                        {
                            //移除方剂处方
                            dgvList.Rows.Clear();

                            dgvList.Tag = DrSelect["ID"].ToString();
                            mbln保密 = (DrSelect["是否保密"].ToString() == "1");

                            if (mbln保密) {
                                txt症状.TextBoxText = "";
                                txt症状.TextBoxToolTipText = "";
                                txt症状.Tag = null;

                                //清空治法
                                txt治法.TextBoxText = "";
                                txt治法.Tag = null;
                                txt治法.TextBoxToolTipText = "";

                                //清空草药
                                txt草药.TextBoxText = "";
                                txt草药.Tag = null;
                                txt草药.TextBoxToolTipText = "";
                                btnAdd.Tag = null;
                            }

                            if (m_baseData.UseType != "2")
                            {
                                tableLayoutPanel3.Enabled = !mbln保密;
                            }

                            for (int i = 0; i < mdt方剂组成.Rows.Count; i++)
                            {
                                DataGridViewRow Dgvr = this.dgvList.Rows[this.dgvList.Rows.Add()];
                                Dgvr.Cells["来源"].Value = DrSelect["方剂名称"].ToString();
                                Dgvr.Cells["来源"].Tag = DrSelect["ID"].ToString();

                                Dgvr.Cells["草药名"].Value = DrSelect["是否保密"].ToString() == "1" ? "***" : mdt方剂组成.Rows[i]["草药名称"].ToString();
                                Dgvr.Cells["草药名"].Tag = mdt方剂组成.Rows[i]["草药id"].ToString();
                                Dgvr.Cells[0].Tag = mdt方剂组成.Rows[i]["草药名称"].ToString();

                                Dgvr.Cells["用量"].Value = mdt方剂组成.Rows[i]["用量"].ToString();
                                Dgvr.Cells["单位"].Value = mdt方剂组成.Rows[i]["单位"].ToString();
                                Dgvr.Cells["单位"].Tag = mdt方剂组成.Rows[i]["His品种id"].ToString();

                                //计算总量
                                Dgvr.Cells["总量"].Tag = ComLib.StrToDlb(mdt方剂组成.Rows[i]["用量"].ToString()) * ComLib.StrToInt(txt付数.Text);
                                Dgvr.Cells["总量"].Value = (Dgvr.Cells["总量"].Tag).ToString() + mdt方剂组成.Rows[i]["单位"].ToString();

                                Dgvr.Cells["脚注"].Value = mdt方剂组成.Rows[i]["用法备注"].ToString();
                                //设置悬浮提示
                                Dgvr.Cells["单位"].ToolTipText = mdt方剂组成.Rows[i]["古法用量"].ToString() == "" ? "" : "古法用量：" + mdt方剂组成.Rows[i]["古法用量"].ToString();

                                //把窗体上的picturebox控件赋值给单元格
                                dgvList.Rows[dgvList.RowCount - 1].Cells[0].Value = this.picSelect.Image;

                                Dgvr.Dispose();
                            }

                            Load药品库存();
                            dgvList.Focus();
                        }
                        else
                        {
                            MessageBox.Show("当前方剂未对应草药，请重新选择。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txt证型.Focus();
                        }
                    }
                    else
                    {
                        MessageBox.Show("当前方剂未对应草药，请重新选择。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txt证型.Focus();
                    }

                    if (dgvList.RowCount == 0)
                    {
                        dgvList.RowCount = 1;
                    }
                    //选中第一行

                    if (dgvList.Rows.Count > 0)
                    {
                        dgvList.Rows[0].Selected = true;
                        dgvList.CurrentCell = dgvList.Rows[0].Cells[3];
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        //根据当前设置的药房检查库存

        //intOne =0 刷新全部的库存    =1 刷新一行草药的库存
        private void Load药品库存(int intOne = 0, int intRow = 0, Boolean blnLoad = false)
        {
            try
            {

                string strIds = "";
                if (intOne == 0)
                {
                    foreach (DataGridViewRow row in this.dgvList.Rows)
                    {
                        if (row.Cells["草药名"].Value != null)
                        {
                            strIds = strIds + "," + row.Cells["单位"].Tag;
                        }
                    }
                    if (strIds != "") { strIds = strIds.Substring(1).ToString(); }
                }
                else
                {
                    if (intRow >= 0)
                    {
                        if (dgvList.Rows[intRow].Cells["单位"].Tag != null)
                        {
                            strIds = dgvList.Rows[intRow].Cells["单位"].Tag.ToString();
                        }
                    }
                }
                //获取药房ID
                string str药房ID = "";
                itemEx item = (itemEx)this.cbo药房.SelectedItem;
                str药房ID = item.Tag.ToString();

                //获取有库存的药名ID
                if (strIds != "")
                {
                    DataTable dtTmp = new DataTable();
                    string message = JsonConvert.SerializeObject(new
                    {
                        药房ID = str药房ID,
                        站点 = m_baseData.站点,
                        药名IDS = strIds
                    });
                    dtTmp = DataCenter.GetTable(m_baseData, "Get_Drugstock", message);

                    foreach (DataGridViewRow row in this.dgvList.Rows)
                    {
                        if (row.Cells["草药名"].Value != null)
                        {
                            if (intOne == 0 || (intOne == 1 && intRow == row.Index))
                            {
                                if (blnLoad == false)
                                {
                                    row.Cells["规格"].Value = null;
                                }
                                row.Cells["规格"].Tag = null;
                                row.Cells["脚注"].Tag = null;

                                if (dtTmp == null)
                                {
                                    row.Cells["库存情况"].Value = (ComLib.StrToDlb(row.Cells["单位"].Tag == null ? "0" : row.Cells["单位"].Tag.ToString()) <= 0) ? "尚未对码" : "没有库存";
                                    row.Cells["库存情况"].Style.ForeColor = Color.Red;
                                }
                                else
                                {
                                    //缓存记录集

                                    dtTmp.DefaultView.RowFilter = "药名ID=" + row.Cells["单位"].Tag.ToString();
                                    dtTmp.DefaultView.Sort = "可用数量 desc";
                                    row.Cells["脚注"].Tag = dtTmp.DefaultView.ToTable();
                                    if (blnLoad || row.Cells["规格"].Value != null)
                                    {
                                        dtTmp.DefaultView.RowFilter = "药名ID=" + row.Cells["单位"].Tag.ToString() + " And 药品id=" + row.Cells["规格"].Value.ToString();
                                    }
                                    //清空缓存规格ID数据
                                    row.Cells["规格"].Value = null;
                                    if (dtTmp.DefaultView.ToTable().Rows.Count != 0)
                                    {
                                        row.Cells["库存情况"].Tag = dtTmp.DefaultView.ToTable().Rows[0]["可用数量"].ToString();

                                        Check库存(row.Index);

                                        row.Cells["规格"].Value = "[" + dtTmp.DefaultView.ToTable().Rows[0]["编码"].ToString() + "]" + (mbln保密 == false ? dtTmp.DefaultView.ToTable().Rows[0]["名称"].ToString() : "***") + (dtTmp.DefaultView.ToTable().Rows[0]["规格"].ToString() == string.Empty ? "" : " " + dtTmp.DefaultView.ToTable().Rows[0]["规格"].ToString());
                                        row.Cells["规格"].Value = row.Cells["规格"].Value + (dtTmp.DefaultView.ToTable().Rows[0]["产地"].ToString() == string.Empty ? "" : " 产地：" + dtTmp.DefaultView.ToTable().Rows[0]["产地"].ToString());
                                        row.Cells["规格"].Tag = dtTmp.DefaultView.ToTable().Rows[0]["药品id"].ToString();
                                        DataGridViewCellStyle style = row.Cells["规格"].Style;
                                        style.Font = new Font(dgvList.Font, FontStyle.Bold);
                                    }
                                    else
                                    {
                                        row.Cells["库存情况"].Value = (ComLib.StrToDlb(row.Cells["单位"].Tag == null ? "0" : row.Cells["单位"].Tag.ToString()) <= 0) ? "尚未对码" : "没有库存";
                                        row.Cells["库存情况"].Style.ForeColor = Color.Red;
                                    }
                                    dtTmp.DefaultView.RowFilter = null;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgvSelect_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                SetSelectData();
                cboSelect.CloseDiag();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //显示选择器

        private Boolean ShowSelect(TextBoxEx objCtl, DataTable dt, int intHeight)
        {
            try
            {
                //调整控件大小和位置

                int intLeft = 0;
                int intTop = 0;

                cboSelect.Parent = this;
                cboSelect.Width = objCtl.Width;
                cboSelect.Height = objCtl.Height;
                ComLib.GetCtlPos(objCtl, ref intTop, ref intLeft);

                //计算屏幕高度
                int iActulaHeight = Screen.PrimaryScreen.WorkingArea.Height;
                if (this.Top + intTop + cboSelect.Height + intHeight + 25 > iActulaHeight)
                {
                    intTop = intTop - intHeight - cboSelect.Height;
                }


                cboSelect.Top = intTop;
                cboSelect.Left = intLeft;

                //绑定数据源

                dgvSelect.DataSource = null;
                dgvSelect.DataSource = dt.DefaultView;


                dgvSelect.Parent = this;
                cboSelect.GridViewTmp = dgvSelect;


                //选中第一行

                if (dgvSelect.Rows.Count > 0)
                {
                    dgvSelect.Rows[0].Selected = true;
                    dgvSelect.CurrentCell = dgvSelect.Rows[0].Cells[2];
                }

                dgvSelect.Columns[""].Visible = false;
                dgvSelect.Columns["ID"].Visible = false;
                if (objCtl.Name == "txt证型")
                {
                    dgvSelect.Columns["简码"].Visible = false;
                }
                else
                {
                    dgvSelect.Columns["简码"].Visible = true;
                }

                if (objCtl.Name == "txt草药")
                {
                    dgvSelect.Columns["HIS品种ID"].Visible = false;
                }

                if (dgvSelect.Columns["状态"] != null)
                {
                    dgvSelect.Columns["状态"].Visible = false;
                }

                //列宽适应
                dgvSelect.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                //计算列宽
                int intWidth = 35;
                for (int j = 1; j < dgvSelect.Columns.Count; j++)
                {
                    if (dgvSelect.Columns[j].Width > 100)
                    {
                        dgvSelect.Columns[j].Width = 100;
                    }
                    if (dgvSelect.Columns[j].Visible)
                    {
                        intWidth = intWidth + dgvSelect.Columns[j].Width;
                    }
                }
                intWidth = intWidth > 500 ? 500 : intWidth;
                intWidth = intWidth < cboSelect.Width ? cboSelect.Width : intWidth;

                cboSelect.ShowDiag(intWidth, intHeight);
                //表格获取焦点
                dgvSelect.Focus();

                //缓存控件对象
                mtxtTmp = objCtl;

                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
        }

        private void dgvSelect_KeyDown(object sender, KeyEventArgs e)
        {
            if (dgvSelect.CurrentRow != null)
            {
                dgvSelect.Tag = dgvSelect.CurrentRow.Index;
            }
        }

        #endregion

        private void btnSave_Click(object sender, EventArgs e)
        {
            //检查处方数据

            if (CheckData() == true)
            {
                if (SaveData() == true)
                {
                    this.Close();
                }
            }
        }

        //保存数据
        private Boolean SaveData()
        {
            itemEx item = (itemEx)this.cbo药房.SelectedItem;

            //获取处方明细
            List<object> details = new List<object>();
            foreach (DataGridViewRow row in this.dgvList.Rows)
            {
                var items = new
                {
                    草药ID = row.Cells["草药名"].Tag.ToString(),
                    是否加药 = (row.Cells["来源"].Tag != null ? 0 : 1).ToString(),
                    来源 = row.Cells["来源"].Value,
                    草药名称 = row.Cells[0].Tag.ToString(),
                    用量 = row.Cells["用量"].Value.ToString(),
                    单位 = row.Cells["单位"].Value.ToString(),
                    脚注 = row.Cells["脚注"].Value != null ? row.Cells["脚注"].Value.ToString() : "",
                    HIS品种ID = row.Cells["单位"].Tag.ToString(),
                    HIS规格ID = row.Cells["规格"].Tag.ToString()
                };
                details.Add(items);
            }


            string message = JsonConvert.SerializeObject(new
            {
                USETYPE = m_baseData.UseType,////使用类型(0-新增/1-修改)
                病人ID = m_baseData.病人ID,
                挂号单 = m_baseData.挂号单,
                门诊号 = m_baseData.门诊号,
                诊断ID = m_baseData.诊断ID,
                科室ID = m_baseData.DeptID,
                科室名称 = m_baseData.DeptName,
                操作员ID = m_baseData.OperatorID,
                操作员姓名 = m_baseData.OperatorName,
                姓名 = m_baseData.Name,
                性别 = m_baseData.Sex,
                年龄 = m_baseData.Age,
                民族 = m_baseData.民族,
                出生日期 = m_baseData.Date,
                就诊方式 = (myTab.SelectedIndex == 0 ? 1 : 2).ToString(),
                科别 = myTab.SelectedIndex == 1 ? "" : cbo科别.Text,
                疾病ID = myTab.SelectedIndex == 1 ? "" : txt疾病.Tag.ToString(),
                疾病名称 = myTab.SelectedIndex == 1 ? "" : txt疾病.TextBoxToolTipText,
                证型ID = myTab.SelectedIndex == 1 ? "" : txt证型.Tag.ToString(),
                证型名称 = myTab.SelectedIndex == 1 ? "" : txt证型.TextBoxToolTipText,
                中医诊断 = myTab.SelectedIndex == 1 ? txtDiag.TextBoxToolTipText.ToString() : txt诊断.Text.ToString(),
                中医治法 = myTab.SelectedIndex == 1 ? "" : txt病人治法.Text.ToString(),
                方剂ID = dgvList.Tag.ToString().ToString(),
                方剂名称 = mstr方剂名,
                付数 = txt付数.Text,
                中药用法 = txt用法.TextBoxToolTipText,
                HIS用法ID = txt用法.Tag.ToString(),
                中药煎法 = txt煎法.TextBoxToolTipText,
                HIS煎法ID = txt煎法.Tag.ToString(),
                煎量 = txt煎量.Text,
                用药频率 = txt频率.TextBoxToolTipText,
                频率次数 = txt频率.Tag.ToString().Split('|')[0].ToString(),
                频率间隔 = txt频率.Tag.ToString().Split('|')[1].ToString(),
                间隔单位 = txt频率.Tag.ToString().Split('|')[2].ToString(),
                医生嘱托 = txt嘱托.Text,
                HIS药房ID = item.Tag.ToString(),
                处方明细 = details
            });

            DataTable dtOut = new DataTable();
            try
            {
                dtOut = DataCenter.ExecuteFunc("Save_Zyedit", message, m_baseData.UserName, m_baseData.UserPassword);
                if (dtOut != null)
                {
                    gstrOut = JsonConvert.SerializeObject(new
                    {
                        His医嘱id = dtOut.Rows[0]["HIS医嘱ID"].ToString(),
                        His诊断id = dtOut.Rows[0]["HIS诊断ID"].ToString(),
                        诊断id = dtOut.Rows[0]["诊断ID"].ToString(),
                        处方id = dtOut.Rows[0]["处方ID"].ToString()
                    });
                }
                else
                {
                    throw new Exception("获取保存数据失败，服务器返回数据为空！");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ComLib.GetMessageFromOracleError(ex.Message), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            return true;
        }

        //检查数据

        private Boolean CheckData()
        {
            try
            {

                mstr方剂名 = "";
                int intTmp = 0;
                string strTmp = string.Empty;
                //检查诊断

                if (myTab.SelectedIndex == 0)
                //辩证施治
                {
                    if (txt疾病.TextBoxToolTipText == "" || txt疾病.Tag == null)
                    {
                        MessageBox.Show("病人中医疾病不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txt疾病.Focus();
                        return false;
                    }
                    if (txt证型.TextBoxToolTipText == "" || txt证型.Tag == null)
                    {
                        MessageBox.Show("病人中医证型不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txt证型.Focus();
                        return false;
                    }
                }
                else
                //直接选方
                {
                    if (txt方剂.TextBoxToolTipText == "" || txt方剂.Tag == null)
                    {
                        MessageBox.Show("处方方剂不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txt方剂.Focus();
                        return false;
                    }
                    if (txtDiag.TextBoxToolTipText == "")
                    {
                        MessageBox.Show("病人诊断不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtDiag.Focus();
                        return false;
                    }
                }

                //必须选择方剂
                if (dgvList.Tag == null || dgvList.RowCount == 0)
                {
                    MessageBox.Show("当前处方尚未选择方剂，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (myTab.SelectedIndex == 0)
                    {
                        txt证型.Focus();
                    }
                    else
                    {
                        txt方剂.Focus();
                    }
                    return false;
                }

                foreach (DataGridViewRow row1 in this.dgvList.Rows)
                {
                    if (row1.Cells["草药名"].Tag != null)
                    {

                        //记录是否有方剂记录

                        if (row1.Cells["来源"].Tag != null)
                        {
                            intTmp++;

                            mstr方剂名 = row1.Cells["来源"].Value.ToString();
                        }

                        //检查用量

                        if (ComLib.StrToDlb(row1.Cells["用量"].Value == null ? "" : row1.Cells["用量"].Value.ToString()) <= 0 || ComLib.StrToDlb(row1.Cells["用量"].Value == null ? "" : row1.Cells["用量"].Value.ToString()) >= 9999)
                        {
                            MessageBox.Show(@"中草药""" + row1.Cells["草药名"].Value.ToString() + @"""的用量录入不正确，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //定位用量行

                            dgvList.Rows[row1.Index].Selected = true;
                            dgvList.CurrentCell = dgvList.Rows[row1.Index].Cells[3];
                            dgvList.Focus();
                            return false;
                        }

                        //检查规格

                        if (row1.Cells["规格"].Tag == null)
                        {
                            MessageBox.Show(@"中草药""" + row1.Cells["草药名"].Value.ToString() + @"""的规格录入不正确，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //定位规格行

                            dgvList.Rows[row1.Index].Selected = true;
                            dgvList.CurrentCell = dgvList.Rows[row1.Index].Cells[7];
                            dgvList.Focus();
                            return false;
                        }
                        else
                        {
                            //检查库存

                            if (row1.Cells["库存情况"].Value.ToString() != "库存充足")
                            {
                                strTmp = strTmp + row1.Cells["草药名"].Value.ToString() + "：所需总量 " + row1.Cells["总量"].Value.ToString() + "\r\n";
                            }
                        }

                        //检查是否有重复的中药

                        foreach (DataGridViewRow row in this.dgvList.Rows)
                        {
                            if (row.Cells["草药名"].Tag != null && row1.Index != row.Index)
                            {
                                if (row1.Cells["草药名"].Tag.ToString() == row.Cells["草药名"].Tag.ToString())
                                {
                                    MessageBox.Show(@"中草药 """ + row.Cells["草药名"].Value.ToString() + @""" 已经录入，不能添加相同的草药。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    //定位行

                                    dgvList.Rows[row.Index].Selected = true;
                                    dgvList.CurrentCell = dgvList.Rows[row.Index].Cells[3];
                                    dgvList.Focus();
                                    return false;
                                }
                            }
                        }
                    }
                }

                //检查是否录入方剂

                if (intTmp == 0)
                {
                    MessageBox.Show("当前处方尚未选择方剂，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (myTab.SelectedIndex == 0)
                    {
                        txt证型.Focus();
                    }
                    else
                    {
                        txt方剂.Focus();
                    }
                    return false;
                }

                //检查药房

                itemEx item = (itemEx)this.cbo药房.SelectedItem;
                if (item.Tag == null || item.ItemData == "停用")
                {
                    MessageBox.Show("当前未选择药房或者选择的药房不可用，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cbo药房.Focus();
                    return false;
                }


                //检查煎法

                if (txt煎法.TextBoxToolTipText == "" || txt煎法.Tag == null)
                {
                    MessageBox.Show("中药煎法不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt煎法.Focus();
                    return false;
                }
                //检查用法

                if (txt用法.TextBoxToolTipText == "" || txt用法.Tag == null)
                {
                    MessageBox.Show("中药用法不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt用法.Focus();
                    return false;
                }
                //检查频率

                if (txt频率.TextBoxToolTipText == "" || txt频率.Tag == null)
                {
                    MessageBox.Show("中药频率不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt频率.Focus();
                    return false;
                }

                //检查库存

                if (strTmp != string.Empty)
                {
                    if (MessageBox.Show("当前处方中以下中草药库存不足,请确认是否继续？\r\n" + strTmp, this.Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
        }

        //加载数据
        private void LoadData()
        {
            bool bln药房 = false;
            try
            {
                string message = JsonConvert.SerializeObject(new { 诊断ID = m_baseData.诊断ID });
                DataTable dtTmp = DataCenter.GetTable(m_baseData, "Load_Zyedit", message);
                if (dtTmp == null)
                {
                    throw new Exception("加载病人中医诊断数据失败，服务器返回数据为空！");
                }
                else
                {
                    myTab.SelectedIndex = dtTmp.Rows[0]["就诊方式"].ToString() == "1" ? 0 : 1;

                    foreach (itemEx item in cbo药房.Items)
                    {
                        if (item.Tag != null)
                        {
                            if (item.Tag.ToString() == dtTmp.Rows[0]["His药房id"].ToString())
                            {
                                cbo药房.SelectedItem = item;
                                bln药房 = true;
                                break;
                            }
                        }
                    }

                    //未获取到保存的药房时，重新读取但不允许保存

                    if (bln药房 == false)
                    {
                        message = JsonConvert.SerializeObject(new { 药房ID = dtTmp.Rows[0]["His药房id"].ToString() });
                        DataTable dt药房 = DataCenter.GetTable(m_baseData, "Get_Drugdept", message);
                        if (dt药房 != null)
                        {
                            itemEx item = new itemEx(dt药房.Rows[0]["ID"].ToString(), dt药房.Rows[0]["编码"].ToString() + "-" + dt药房.Rows[0]["名称"].ToString(), "停用");
                            cbo药房.Items.Add(item);
                            cbo药房.SelectedItem = item;
                        }
                    }

                    if (myTab.SelectedIndex == 0)
                    {
                        //加载科别
                        cbo科别.Text = dtTmp.Rows[0]["科别"].ToString();

                        //加载疾病
                        txt疾病.Tag = dtTmp.Rows[0]["疾病ID"].ToString();
                        txt疾病.TextBoxToolTipText = dtTmp.Rows[0]["疾病名称"].ToString();
                        txt疾病.TextBoxText = txt疾病.TextBoxToolTipText;


                        //选择疾病时缓存证型种类

                        if (txt疾病.Tag != null)
                        {
                            message = JsonConvert.SerializeObject(new { 疾病ID = txt疾病.Tag });
                            mdt证型 = DataCenter.GetTable(m_baseData, "Get_Zxtype", message);
                        }

                        //加载证型
                        txt证型.Tag = dtTmp.Rows[0]["证型ID"].ToString();
                        txt证型.TextBoxToolTipText = dtTmp.Rows[0]["证型名称"].ToString();
                        txt证型.TextBoxText = txt证型.TextBoxToolTipText;

                        //加载治法和诊断

                        txt病人治法.Text = dtTmp.Rows[0]["中医治法"].ToString();
                        txt诊断.Text = dtTmp.Rows[0]["中医诊断"].ToString();
                    }

                    else
                    {

                        txt方剂.TextBoxToolTipText = dtTmp.Rows[0]["方剂名称"].ToString();
                        txt方剂.TextBoxText = txt方剂.TextBoxToolTipText;
                        txt方剂.Tag = dtTmp.Rows[0]["方剂ID"].ToString();

                        txtDiag.TextBoxToolTipText = dtTmp.Rows[0]["中医诊断"].ToString();
                        txtDiag.TextBoxText = txtDiag.TextBoxToolTipText;
                    }

                    dgvList.Tag = dtTmp.Rows[0]["方剂ID"].ToString();

                    txt煎法.TextBoxToolTipText = dtTmp.Rows[0]["中药煎法"].ToString();
                    txt煎法.TextBoxText = txt煎法.TextBoxToolTipText;
                    txt煎法.Tag = dtTmp.Rows[0]["His煎法id"].ToString();

                    txt煎量.Text = dtTmp.Rows[0]["煎量"].ToString();

                    txt用法.TextBoxToolTipText = dtTmp.Rows[0]["中药用法"].ToString();
                    txt用法.TextBoxText = txt用法.TextBoxToolTipText;
                    txt用法.Tag = dtTmp.Rows[0]["His用法id"].ToString();

                    txt频率.TextBoxToolTipText = dtTmp.Rows[0]["用药频率"].ToString();
                    txt频率.TextBoxText = txt频率.TextBoxToolTipText;
                    txt频率.Tag = dtTmp.Rows[0]["频率次数"].ToString() + "|" + dtTmp.Rows[0]["频率间隔"].ToString() + "|" + dtTmp.Rows[0]["间隔单位"].ToString();

                    txt付数.Text = dtTmp.Rows[0]["付数"].ToString();
                    txt嘱托.Text = dtTmp.Rows[0]["医生嘱托"].ToString();
                    mbln保密 = (dtTmp.Rows[0]["是否保密"].ToString() == "1");

                    if (mbln保密)
                    {
                        txt症状.TextBoxText = "";
                        txt症状.TextBoxToolTipText = "";
                        txt症状.Tag = null;

                        //清空治法
                        txt治法.TextBoxText = "";
                        txt治法.Tag = null;
                        txt治法.TextBoxToolTipText = "";

                        //清空草药
                        txt草药.TextBoxText = "";
                        txt草药.Tag = null;
                        txt草药.TextBoxToolTipText = "";
                        btnAdd.Tag = null;
                    }

                    if (m_baseData.UseType != "2")
                    {
                        tableLayoutPanel3.Enabled = !mbln保密;
                    }

                    //加载草药清单
                    //a.处方明细id, a.处方id, a.序号, a.草药id, a.是否加药, a.来源, a.草药名称, a.用量, a.单位, a.脚注, a.His品种id, a.His规格id
                    message = JsonConvert.SerializeObject(new { 处方ID = dtTmp.Rows[0]["处方ID"].ToString() });
                    DataTable dtInfo = DataCenter.GetTable(m_baseData, "Load_Zyinfo", message);
                    if (dtInfo == null)
                    {
                        throw new Exception("加载病人处方明细数据失败，服务器返回数据为空！");
                    }
                    else
                    {
                        //移除方剂处方
                        dgvList.Rows.Clear();
                        for (int i = 0; i < dtInfo.Rows.Count; i++)
                        {
                            DataGridViewRow Dgvr = this.dgvList.Rows[this.dgvList.Rows.Add()];
                            Dgvr.Cells["来源"].Value = dtInfo.Rows[i]["来源"].ToString();
                            if (dtInfo.Rows[i]["是否加药"].ToString() == "0")
                            {
                                Dgvr.Cells["来源"].Tag = dtTmp.Rows[0]["方剂ID"].ToString();
                                dgvList.Rows[dgvList.RowCount - 1].Cells[0].Value = this.picSelect.Image;
                            }
                            else
                            {
                                dgvList.Rows[dgvList.RowCount - 1].Cells[0].Value = this.picAdd.BackgroundImage;
                            }
                            Dgvr.Cells["草药名"].Value = (mbln保密 == false) ? dtInfo.Rows[i]["草药名称"].ToString() : "***";
                            Dgvr.Cells[0].Tag = dtInfo.Rows[i]["草药名称"].ToString();
                            Dgvr.Cells["草药名"].Tag = dtInfo.Rows[i]["草药id"].ToString();
                            Dgvr.Cells["用量"].Value = dtInfo.Rows[i]["用量"].ToString();
                            Dgvr.Cells["单位"].Value = dtInfo.Rows[i]["单位"].ToString();
                            Dgvr.Cells["单位"].Tag = dtInfo.Rows[i]["His品种id"].ToString();

                            Dgvr.Cells["脚注"].Tag = dtInfo.Rows[i]["脚注"].ToString();

                            Dgvr.Cells["规格"].Value = dtInfo.Rows[i]["His规格id"].ToString();
                            //计算总量
                            Dgvr.Cells["总量"].Tag = ComLib.StrToDlb(dtInfo.Rows[i]["用量"].ToString()) * ComLib.StrToInt(txt付数.Text);
                            Dgvr.Cells["总量"].Value = (Dgvr.Cells["总量"].Tag).ToString() + dtInfo.Rows[i]["单位"].ToString();

                            Dgvr.Dispose();
                        }

                        Load药品库存(0, 0, true);

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

    }
}