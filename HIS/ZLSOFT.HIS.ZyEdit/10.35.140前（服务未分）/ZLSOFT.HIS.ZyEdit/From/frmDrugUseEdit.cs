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
    partial class frmDrugUseEdit : Form
    {

        public BaseData m_baseData = new BaseData(); //基础数据

        private int mintType;//0=新增,1=修改
        private DataGridViewRow mdtRow;//修改时数据加载集
        private DataTable mdtDrug;
        public string gstrID;
        private string mstrTmp;
        public bool gblnOK;
        public frmDrugUseEdit(int intType, DataGridViewRow dtRow, DataTable dtDrug, BaseData baseData, string strTmp = "")
        {
            mintType = intType;
            mdtRow = dtRow;
            m_baseData = baseData;
            mdtDrug = dtDrug;
            mstrTmp = strTmp;
            gstrID = string.Empty;
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                //检查录入项
                if (txt草药.TextBoxToolTipText == "" || txt草药.Tag == null)
                {
                    MessageBox.Show("草药不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt草药.Focus();
                    return;
                }

                //检查用量
                if (ComLib.StrToDlb(txt用量.Text) <= 0 || ComLib.StrToDlb(txt用量.Text) >= 99999)
                {
                    MessageBox.Show(@"中草药""" + txt草药.TextBoxText + @"""的用量录入不正确，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //定位用量行
                    txt用量.Focus();
                    return;
                }

                //保存

                string message = JsonConvert.SerializeObject(new
                {
                    USETYPE = mintType.ToString(),////使用类型(0-新增/1-修改)
                    用药ID = gstrID,
                    治法ID = mstrTmp,
                    草药ID = txt草药.Tag.ToString(),
                    用量 = txt用量.Text,
                    操作员名称 = m_baseData.OperatorName,
                    操作员ID = m_baseData.OperatorID
                });

                DataTable dtOut = new DataTable();

                dtOut = DataCenter.ExecuteFunc("Set_Zftozy", message, m_baseData.UserName, m_baseData.UserPassword);
                if (dtOut != null)
                {
                    gstrID = dtOut.Rows[0]["用药ID"].ToString();
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

        private void LoadData()
        {
            try
            {
                txt草药.TextBoxText = mdtRow.Cells["草药名称"].Value.ToString();
                txt草药.TextBoxToolTipText = mdtRow.Cells["草药名称"].Value.ToString();
                txt草药.Tag = mdtRow.Cells["草药ID"].Value.ToString();
                txt用量.Text = mdtRow.Cells["用量"].Value.ToString();
                lbl单位.Text = mdtRow.Cells["单位"].Value.ToString();
                gstrID = mdtRow.Cells["ID"].Value.ToString();
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

        private void frmDiagEdit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void frmDiagEdit_Load(object sender, EventArgs e)
        {
            if (mintType == 1)
            {
                LoadData();
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

                        if (mdtDrug != null)
                        {
                            mdtDrug.DefaultView.RowFilter = txt草药.TextBoxText == "" ? "" : "草药名称 like '%" + txt草药.TextBoxText + "%' Or 简码 like '%" + txt草药.TextBoxText + "%'";
                            if (mdtDrug.Rows.Count > 0)
                            {
                                frmZySelect frmSelect = new frmZySelect("草药选择器", mdtDrug.DefaultView.ToTable());
                                frmSelect.ShowDialog(this);
                                if (frmSelect.dataRowOut != null)
                                {
                                    txt草药.TextBoxText = frmSelect.dataRowOut["草药名称"].ToString();
                                    txt草药.TextBoxToolTipText = frmSelect.dataRowOut["草药名称"].ToString();
                                    txt草药.Tag = frmSelect.dataRowOut["ID"].ToString();
                                    lbl单位.Text = frmSelect.dataRowOut["单位"].ToString();
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
                        txt用量.Focus();
                    }
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
                if (mdtDrug != null)
                {
                    mdtDrug.DefaultView.RowFilter = "";
                    if (mdtDrug.Rows.Count > 0)
                    {
                        mdtDrug.DefaultView.RowFilter = "";
                        frmZySelect frmSelect = new frmZySelect("草药选择器", mdtDrug.DefaultView.ToTable());
                        frmSelect.ShowDialog(this);
                        if (frmSelect.dataRowOut != null)
                        {
                            txt草药.TextBoxText = frmSelect.dataRowOut["草药名称"].ToString();
                            txt草药.TextBoxToolTipText = frmSelect.dataRowOut["草药名称"].ToString();
                            txt草药.Tag = frmSelect.dataRowOut["ID"].ToString();
                            lbl单位.Text = frmSelect.dataRowOut["单位"].ToString();
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

        private void txt草药_Validating(object sender, CancelEventArgs e)
        {
            if (txt草药.TextBoxToolTipText != txt草药.TextBoxText)
            {
                txt草药.TextBoxText = txt草药.TextBoxToolTipText;
            }
        }

        private void txt用量_Validating(object sender, CancelEventArgs e)
        {
            try
            {
                double newDB = 0;
                if (!double.TryParse(txt用量.Text, out newDB))
                {
                    e.Cancel = true;
                    MessageBox.Show("用量输入错误，请输入有效的用量！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else
                {
                    if (newDB <= 0 || newDB > 99999)
                    {
                        e.Cancel = true;
                        MessageBox.Show("不是大于零的数字或输入数值过大！", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                    {
                        //计算总量
                        e.Cancel = false;
                        txt用量.Text = Math.Round(ComLib.StrToDlb(txt用量.Text), 4).ToString();
                    }
                }
            }
            catch
            {
                e.Cancel = true;
            }
        }

        private void txt用量_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                btnSave.Focus();
            }
            else
            {
                if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == 8 || e.KeyChar == '.')
                {
                    e.Handled = false;
                    return;
                }
                e.Handled = true;
            }
        }

        private void frmDrugUseEdit_Shown(object sender, EventArgs e)
        {
            txt草药.Focus();
        }

        private void txt用量_Enter(object sender, EventArgs e)
        {
            txt用量.SelectAll();
        }
    }
}
