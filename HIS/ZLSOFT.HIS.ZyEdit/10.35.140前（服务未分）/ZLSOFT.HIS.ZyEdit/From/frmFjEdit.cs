using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ZLSOFT.HIS.ZyEdit.Model;

namespace ZLSOFT.HIS.ZyEdit.From
{
    partial class frmFjEdit : Form
    {

        public BaseData m_baseData = new BaseData(); //基础数据
        private int mintType;//0=新增,1=修改
        private DataRow mdtRow;//修改时数据加载集
        private DataTable mdtDrug;//加载草药目录，用于选择方剂
        private DataTable mdt方剂组成;
        private DataGridViewTextBoxEditingControl CellEdit = null;//用于绑定表格输入事件
        public string gstrOut;
        public string gstrID;
        public bool gblnOK;
        public frmFjEdit(int intType, DataRow dtRow, DataTable dtDrug, BaseData baseData)
        {
            mintType = intType;
            mdtRow = dtRow;
            mdtDrug = dtDrug;
            m_baseData = baseData;
            gstrID = string.Empty;
            InitializeComponent();
        }

        private void IntdgvList(DataGridView dgv, string[] CH, int[] CHWidth)
        {
            try
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

                dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
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

                dgv.ColumnCount = 0;

                //设置dgvList设置表头
                if (CH != null)
                {
                    dgv.ColumnCount = CH.Count();

                    for (int i = 0; i < CH.Length; i++)
                    {
                        dgv.Columns[i].Name = CH[i];
                        dgv.Columns[i].Width = CHWidth[i];
                    }

                    DataGridViewImageColumn dgvImageColumn = new DataGridViewImageColumn();
                    dgvImageColumn.HeaderText = "";
                    dgvImageColumn.ImageLayout = DataGridViewImageCellLayout.Normal;
                    dgvImageColumn.Name = "btnDel";
                    dgvImageColumn.Image = picNone.Image;
                    dgv.Columns.Add(dgvImageColumn);
                    dgv.Columns[dgv.ColumnCount - 1].Width = 22;
                }

                //dgvList设置用量和脚注可以编辑

                dgv.ReadOnly = false;
                foreach (DataGridViewColumn c in dgv.Columns)
                {
                    if (c.Name == "用量" || c.Name == "用法备注" || c.Name == "古法用量")
                    {
                        c.ReadOnly = false;
                        c.DefaultCellStyle.BackColor = Color.AliceBlue;
                        c.DefaultCellStyle.ForeColor = Color.Black;
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
                dgv.ForeColor = Color.Black;

                dgv.RowHeadersVisible = false;

                //设置输入限制
                ((DataGridViewTextBoxColumn)dgv.Columns["古法用量"]).MaxInputLength = 25;
                ((DataGridViewTextBoxColumn)dgv.Columns["用法备注"]).MaxInputLength = 50;

                //初始化行数

                dgv.RowCount = 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string str组成摘要 = "";
            //获取处方明细
            List<object> details = new List<object>();

            try
            {
                //检查录入项
                if (txtName.Text == "")
                {
                    MessageBox.Show("方剂名称不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtName.Focus();
                    return;
                }

                //检查方剂组成录入项
                if (dgvList.Rows[0].Cells["草药名"].Tag == null && dgvList.Rows.Count == 1)
                {
                    MessageBox.Show("方剂组成不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt草药.Focus();
                    return;
                }

                //检查表格录入项
                foreach (DataGridViewRow row1 in this.dgvList.Rows)
                {
                    if (row1.Cells["草药名"].Tag != null)
                    {

                        //生成中药组成摘要
                        str组成摘要 = str组成摘要 + "," + row1.Cells["草药名"].Value.ToString();

                        //检查用量

                        if (ComLib.StrToDlb(row1.Cells["用量"].Value == null ? "" : row1.Cells["用量"].Value.ToString()) <= 0 || ComLib.StrToDlb(row1.Cells["用量"].Value == null ? "" : row1.Cells["用量"].Value.ToString()) >= 9999)
                        {
                            MessageBox.Show(@"中草药""" + row1.Cells["草药名"].Value.ToString() + @"""的用量录入不正确，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //定位用量行

                            dgvList.Rows[row1.Index].Selected = true;
                            dgvList.CurrentCell = dgvList.Rows[row1.Index].Cells[1];
                            dgvList.Focus();
                            return;
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
                                    dgvList.CurrentCell = dgvList.Rows[row.Index].Cells[0];
                                    dgvList.Focus();
                                    return;
                                }
                            }
                        }


                        var items = new
                        {
                            草药ID = row1.Cells["草药名"].Tag.ToString(),
                            用量 = row1.Cells["用量"].Value.ToString(),
                            用法备注 = row1.Cells["用法备注"].Value != null ? row1.Cells["用法备注"].Value.ToString() : "",
                            古法用量 = row1.Cells["古法用量"].Value != null ? row1.Cells["古法用量"].Value.ToString() : ""
                        };
                        details.Add(items);
                    }
                }

                str组成摘要 = str组成摘要.Substring(1);

                //保存
                string message = JsonConvert.SerializeObject(new
                {
                    USETYPE = mintType.ToString(),////使用类型(0-新增/1-修改)
                    方剂ID = gstrID,
                    方剂名称 = txtName.Text.Replace(":",""),
                    简码 = txt简码.Text.Replace(":", ""),
                    别名 = txt别名.Text,
                    别名简码 = txt别名简码.Text,
                    来源 = txt来源.Text,
                    组成摘要 = str组成摘要,
                    服法描述 = txt服法描述.Text,
                    作用描述 = txt作用描述.Text,
                    制法描述 = txt制法描述.Text,
                    适应证描述 = txt适应证.Text,
                    方剂组成作用描述 = txt组成描述.Text,
                    操作员名称 = m_baseData.OperatorName,
                    操作员ID = m_baseData.OperatorID,
                    是否保密= chk保密.Checked == true ?  "1": "0",
                方剂构成 = details
                });

                DataTable dtOut = new DataTable();

                dtOut = DataCenter.ExecuteFunc("Save_Fjitem", message, m_baseData.UserName, m_baseData.UserPassword);
                if (dtOut != null)
                {
                    gstrID = dtOut.Rows[0]["方剂ID"].ToString();
                    gstrOut = message;
                    gblnOK = true;
                    this.Close();
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

        private void frmDrugEdit_Load(object sender, EventArgs e)
        {
            try
            {
                //初始化表格

                //控件信息
                string[] CH = { "草药名", "用量", "单位", "古法用量", "用法备注" };//它的数组的长度必须和datagridview的列数保持一致

                int[] CHWidth = { 100, 60, 50, 105, 105 };//列宽
                IntdgvList(dgvList, CH, CHWidth);

                //增加表格下拉项

                dgvList.Controls.Add(txt草药);

                if (mintType == 1)
                {
                    LoadData();
                    LoadDrugInfo();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void LoadData()
        {
            try
            {
                txtName.Text = mdtRow["方剂名称"].ToString();
                txt简码.Text = mdtRow["简码"].ToString();
                txt别名.Text = mdtRow["别名"].ToString();
                txt别名简码.Text = mdtRow["别名简码"].ToString();
                txt服法描述.Text = mdtRow["服法描述"].ToString();
                txt来源.Text = mdtRow["来源"].ToString();
                txt作用描述.Text = mdtRow["作用描述"].ToString();
                txt制法描述.Text = mdtRow["制法描述"].ToString();
                txt适应证.Text = mdtRow["适应证描述"].ToString();
                txt组成描述.Text = mdtRow["方剂组成作用描述"].ToString();
                chk保密.Checked = mdtRow["是否保密"].ToString() == "1" ? true : false;
                gstrID = mdtRow["id"].ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmDrugEdit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }


        private void txtName_Validating(object sender, CancelEventArgs e)
        {

            txt简码.Text = txtName.Text == "" ? "" : Pinyin.GetInitials(txtName.Text, Encoding.GetEncoding("GB2312"));
        }

        private void txt别名_Validating(object sender, CancelEventArgs e)
        {
            txt别名简码.Text = txt别名.Text == "" ? "" : Pinyin.GetInitials(txt别名.Text, Encoding.GetEncoding("GB2312"));
        }

        private void txtName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
                SendKeys.Send("{tab}");
                return;
            }
        }

        //设置删除图标
        private void dgvList_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvList.ColumnCount >= 6)
            {
                try
                {
                    //把窗体上的picturebox控件赋值给单元格

                    foreach (DataGridViewRow row in this.dgvList.Rows)
                    {
                        if (row.Index == dgvList.CurrentRow.Index)
                        {
                            row.Cells[5].Value = this.picDel.Image;
                        }
                        else
                        {
                            row.Cells[5].Value = this.picNone.Image;
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
                if (e.ColumnIndex == 5 && dgvList.RowCount > 0)
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
        private void txt草药_ButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (mdtDrug != null && dgvList.CurrentRow != null)
                {
                    mdtDrug.DefaultView.RowFilter = "";
                    if (mdtDrug.Rows.Count > 0)
                    {
                        mdtDrug.DefaultView.RowFilter = "";
                        frmZySelect frmSelect = new frmZySelect("草药选择器", mdtDrug.DefaultView.ToTable());
                        frmSelect.ShowDialog(this);



                        if (frmSelect.dataRowOut != null)
                        {
                            //检查是否有重复的中药

                            foreach (DataGridViewRow row in this.dgvList.Rows)
                            {
                                if (row.Cells["草药名"].Tag != null)
                                {
                                    if (frmSelect.dataRowOut["ID"].ToString() == row.Cells["草药名"].Tag.ToString())
                                    {
                                        MessageBox.Show(@"中草药 """ + frmSelect.dataRowOut["草药名称"].ToString() + @""" 已经录入，不能添加相同的草药。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        return;
                                    }
                                }
                            }

                            txt草药.TextBoxText = frmSelect.dataRowOut["草药名称"].ToString();
                            txt草药.TextBoxToolTipText = frmSelect.dataRowOut["草药名称"].ToString();
                            dgvList.CurrentRow.Cells["草药名"].Value = frmSelect.dataRowOut["草药名称"].ToString();
                            dgvList.CurrentRow.Cells["草药名"].Tag = frmSelect.dataRowOut["ID"].ToString();
                            dgvList.CurrentRow.Cells["单位"].Value = frmSelect.dataRowOut["单位"].ToString();
                            if (dgvList.Rows[dgvList.RowCount - 1].Cells["草药名"].Value != null)
                            {
                                dgvList.RowCount = dgvList.RowCount + 1;
                            }
                            dgvList.CurrentCell = this.dgvList.Rows[dgvList.CurrentRow.Index].Cells["用量"];
                            dgvList.Focus();
                        }
                    }
                    else
                    {
                        MessageBox.Show("未获取有效的草药。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("未获取有效的草药。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //表格光标纵向跳转
        private bool m_isdataGridViewBegingEdit = false;

        //表格光标纵向跳转
        private bool m_txtEdit = false;


        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            try
            {
                if (keyData == Keys.Enter && (m_txtEdit == true && this.txt草药.Visible == true))
                {
                    if (txt草药.TextBoxText != txt草药.TextBoxToolTipText || txt草药.TextBoxToolTipText == "")
                    {
                        if (mdtDrug != null)
                        {
                            mdtDrug.DefaultView.RowFilter = txt草药.TextBoxText == "" ? "" : "草药名称 like '%" + txt草药.TextBoxText + "%' Or 简码 like '%" + txt草药.TextBoxText + "%'";
                            if (mdtDrug.Rows.Count > 0)
                            {
                                frmZySelect frmSelect = new frmZySelect("草药选择器", mdtDrug.DefaultView.ToTable());
                                frmSelect.ShowDialog(this);
                                if (frmSelect.dataRowOut != null)
                                {
                                    //检查是否有重复的中药

                                    foreach (DataGridViewRow row in this.dgvList.Rows)
                                    {
                                        if (row.Cells["草药名"].Tag != null)
                                        {
                                            if (frmSelect.dataRowOut["ID"].ToString() == row.Cells["草药名"].Tag.ToString())
                                            {
                                                MessageBox.Show(@"中草药 """ + frmSelect.dataRowOut["草药名称"].ToString() + @""" 已经录入，不能添加相同的草药。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                return true;
                                            }
                                        }
                                    }

                                    txt草药.TextBoxText = frmSelect.dataRowOut["草药名称"].ToString();
                                    txt草药.TextBoxToolTipText = frmSelect.dataRowOut["草药名称"].ToString();
                                    dgvList.CurrentRow.Cells["草药名"].Value = frmSelect.dataRowOut["草药名称"].ToString();
                                    dgvList.CurrentRow.Cells["草药名"].Tag = frmSelect.dataRowOut["ID"].ToString();
                                    dgvList.CurrentRow.Cells["单位"].Value = frmSelect.dataRowOut["单位"].ToString();

                                    dgvList.CurrentCell = this.dgvList.Rows[dgvList.CurrentRow.Index].Cells["用量"];

                                    //添加空行
                                    if (dgvList.Rows[dgvList.RowCount - 1].Cells["草药名"].Value != null)
                                    {
                                        dgvList.RowCount = dgvList.RowCount + 1;
                                    }
                                    dgvList.Focus();
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("未获取有效的草药。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        dgvList.CurrentCell = this.dgvList.Rows[dgvList.CurrentRow.Index].Cells["用量"];
                        dgvList.Focus();
                    }
                    return true;
                }
                else if (keyData == Keys.Enter && (m_isdataGridViewBegingEdit == true || this.dgvList.Focused == true))
                {
                    SendKeys.Send("{tab}");
                    //添加空行
                    if (dgvList.Rows[dgvList.RowCount - 1].Cells["草药名"].Value != null)
                    {
                        dgvList.RowCount = dgvList.RowCount + 1;
                    }
                    return true;
                }
                return base.ProcessCmdKey(ref msg, keyData);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
        }

        private void dgvList_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            m_isdataGridViewBegingEdit = false;
        }

        private void dgvList_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            m_isdataGridViewBegingEdit = true;
        }

        //表格输入项限制

        private void Cells_KeyPress(object sender, KeyPressEventArgs e) //自定义事件

        {

            if (dgvList.Columns[dgvList.CurrentCellAddress.X].Name == "用量")//获取当前处于活动状态的单元格索引

            {
                if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == 8)
                {
                    e.Handled = false;
                    return;
                }
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



        //根据方剂选择设置处方
        private void LoadDrugInfo()
        {
            try
            {
                if (gstrID != "")
                {
                    string message = JsonConvert.SerializeObject(new { 方剂ID = gstrID });
                    mdt方剂组成 = DataCenter.GetTable(m_baseData, "Get_Fjitems", message);
                    if (mdt方剂组成 != null)
                    {
                        if (mdt方剂组成.Rows.Count > 0)
                        {
                            //移除方剂处方
                            dgvList.Rows.Clear();

                            for (int i = 0; i < mdt方剂组成.Rows.Count; i++)
                            {
                                DataGridViewRow Dgvr = this.dgvList.Rows[this.dgvList.Rows.Add()];
                                Dgvr.Cells["草药名"].Value = mdt方剂组成.Rows[i]["草药名称"].ToString();
                                Dgvr.Cells["草药名"].Tag = mdt方剂组成.Rows[i]["草药id"].ToString();

                                Dgvr.Cells["用量"].Value = mdt方剂组成.Rows[i]["用量"].ToString();
                                Dgvr.Cells["单位"].Value = mdt方剂组成.Rows[i]["单位"].ToString();

                                Dgvr.Cells["古法用量"].Value = mdt方剂组成.Rows[i]["古法用量"].ToString();
                                Dgvr.Cells["用法备注"].Value = mdt方剂组成.Rows[i]["用法备注"].ToString();

                                Dgvr.Dispose();
                            }
                        }
                    }
                    dgvList.RowCount = dgvList.RowCount + 1;
                    //选中第一行

                    if (dgvList.Rows.Count > 0)
                    {
                        dgvList.Rows[0].Selected = true;
                        dgvList.CurrentCell = dgvList.Rows[0].Cells[1];
                        dgvList.CurrentCell = dgvList.Rows[0].Cells[0];
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txtName_Enter(object sender, EventArgs e)
        {
            txtName.SelectAll();
        }

        private void txt简码_Enter(object sender, EventArgs e)
        {
            txt简码.SelectAll();
        }

        private void txt别名_Enter(object sender, EventArgs e)
        {
            txt别名.SelectAll();
        }

        private void txt别名简码_Enter(object sender, EventArgs e)
        {
            txt别名简码.SelectAll();
        }

        private void txt来源_Enter(object sender, EventArgs e)
        {
            txt来源.SelectAll();
        }

        private void txt服法描述_Enter(object sender, EventArgs e)
        {
            txt服法描述.SelectAll();
        }

        private void txt作用描述_Enter(object sender, EventArgs e)
        {
            txt作用描述.SelectAll();
        }

        private void txt制法描述_Enter(object sender, EventArgs e)
        {
            txt制法描述.SelectAll();
        }

        private void txt适应证_Enter(object sender, EventArgs e)
        {
            txt适应证.SelectAll();
        }

        private void txt组成描述_Enter(object sender, EventArgs e)
        {
            txt组成描述.SelectAll();
        }

        private void dgvList_CurrentCellChanged(object sender, EventArgs e)
        {
            try
            {
                DataGridViewCell CurrnetCell = this.dgvList.CurrentCell;

                this.txt草药.Visible = false;

                if (CurrnetCell != null)
                {
                    Rectangle Rect = this.dgvList.GetCellDisplayRectangle(dgvList.Columns["草药名"].Index, CurrnetCell.RowIndex, true);
                    txt草药.Visible = true;
                    txt草药.Top = Rect.Top;
                    txt草药.Left = Rect.Left;
                    txt草药.Height = Rect.Height;
                    txt草药.Width = Rect.Width;
                    //缓存上次注脚
                    txt草药.Tag = CurrnetCell.RowIndex;

                    if (dgvList.Rows[CurrnetCell.RowIndex].Cells[dgvList.Columns["草药名"].Index].Value != null)
                    {
                        txt草药.TextBoxText = dgvList.Rows[CurrnetCell.RowIndex].Cells[dgvList.Columns["草药名"].Index].Value.ToString();
                        txt草药.TextBoxToolTipText = dgvList.Rows[CurrnetCell.RowIndex].Cells[dgvList.Columns["草药名"].Index].Value.ToString();
                    }
                    else
                    {
                        txt草药.TextBoxText = "";
                        txt草药.TextBoxToolTipText = "";
                    }
                    if (CurrnetCell.ColumnIndex == dgvList.Columns["草药名"].Index)
                    {
                        txt草药.Focus();
                    }

                    txt草药.Refresh();
                }
                else
                {
                    this.txt草药.Visible = false;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt草药_Enter(object sender, EventArgs e)
        {
            m_txtEdit = true;
        }

        private void txt草药_Leave(object sender, EventArgs e)
        {
            m_txtEdit = false;
            dgvList.Focus();
            txt草药.Visible = false;
        }

        private void dgvList_Scroll(object sender, ScrollEventArgs e)
        {
            txt草药.Visible = false;
            dgvList.Focus();
        }

        private void dgvList_Enter(object sender, EventArgs e)
        {
            txt草药.Visible = true;
            txt草药.Focus();
        }
    }
}
