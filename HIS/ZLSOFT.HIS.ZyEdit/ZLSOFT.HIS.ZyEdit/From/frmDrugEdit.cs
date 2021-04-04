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
    partial class frmDrugEdit : Form
    {

        public BaseData m_baseData = new BaseData(); //基础数据
        private int mintType;//0=新增,1=修改
        private DataRow mdtRow;//修改时数据加载集
        private DataTable mdtHISDrug;//HIS草药数据集

        public string gstrID;
        public string gstrOut;
        public bool gblnOK;
        public frmDrugEdit(int intType, DataRow dtRow, DataTable dtHISDrug, BaseData baseData)
        {
            mintType = intType;
            mdtRow = dtRow;
            mdtHISDrug = dtHISDrug;
            m_baseData = baseData;
            gstrID = string.Empty;
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {

            try
            {
                //检查录入项
                if (txtName.Text == "")
                {
                    MessageBox.Show("草药名称不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtName.Focus();
                    return;
                }

                if (txtName.Text.IndexOf("*") >0)
                {
                    MessageBox.Show("草药名称不能录入非法字符，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtName.Focus();
                    return;
                }


                //保存

                string message = JsonConvert.SerializeObject(new
                {
                    USETYPE = mintType.ToString(),////使用类型(0-新增/1-修改)
                    草药ID = gstrID,
                    草药名称 = txtName.Text,
                    简码 = txt简码.Text,
                    别名 = txt别名.Text,
                    别名简码 = txt别名简码.Text,
                    单位 = txt单位.Text,
                    来源 = txt来源.Text,
                    草药描述 = txt描述.Text,
                    性状 = txt性状.Text,
                    药性 = txt药性.Text,
                    适应证 = txt适应证.Text,
                    用法 = txt用法.Text,
                    服法 = txt服法.Text,
                    禁忌 = txt禁忌.Text,
                    成分 = txt成分.Text,
                    药理作用 = txt作用.Text,
                    HIS品种ID = txt品种.Tag == null ? "" : txt品种.Tag.ToString(),
                    操作员名称 = m_baseData.OperatorName,
                    操作员ID = m_baseData.OperatorID
                });

                DataTable dtOut = new DataTable();

                dtOut = DataCenter.ExecuteFunc("Save_Drugitem", message, m_baseData.UserName, m_baseData.UserPassword);
                if (dtOut != null)
                {
                    gstrID = dtOut.Rows[0]["草药ID"].ToString();
                    gblnOK = true;
                    gstrOut = message;
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
            if (mintType == 1)
            {
                LoadData();
            }
        }

        private void LoadData()
        {
            try
            {
                txtName.Text = mdtRow["草药名称"].ToString();
                txt简码.Text = mdtRow["简码"].ToString();
                txt别名.Text = mdtRow["别名"].ToString();
                txt别名简码.Text = mdtRow["别名简码"].ToString();
                txt单位.Text = mdtRow["单位"].ToString();
                txt来源.Text = mdtRow["来源"].ToString();
                txt描述.Text = mdtRow["草药描述"].ToString();
                txt性状.Text = mdtRow["性状"].ToString();
                txt药性.Text = mdtRow["药性"].ToString();
                txt适应证.Text = mdtRow["适应证"].ToString();
                txt用法.Text = mdtRow["用法"].ToString();
                txt服法.Text = mdtRow["服法"].ToString();
                txt禁忌.Text = mdtRow["禁忌"].ToString();
                txt成分.Text = mdtRow["成分"].ToString();
                txt作用.Text = mdtRow["药理作用"].ToString();
                gstrID = mdtRow["id"].ToString();
                if (mdtRow["His品种id"].ToString() != "" && mdtRow["His品种id"].ToString() != "0")
                {
                    mdtHISDrug.DefaultView.RowFilter = "His品种id =" + mdtRow["His品种id"].ToString();
                    txt品种.Tag = mdtRow["His品种id"].ToString();
                    txt品种.TextBoxToolTipText = "[" + mdtHISDrug.DefaultView.ToTable().Rows[0]["His品种id"].ToString() + "]" + mdtHISDrug.DefaultView.ToTable().Rows[0]["名称"].ToString();
                    txt品种.TextBoxText = txt品种.TextBoxToolTipText;
                }
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
                    txt品种.Tag = frmSelect.dataRowOut["His品种id"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void txt品种_EditBoxKeyPress(object sender, KeyPressEventArgs e)
        {

            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt品种.TextBoxText != txt品种.TextBoxToolTipText.Replace("'", "").Replace("[", "[[ ")
                            .Replace("]", " ]]")
                            .Replace("*", "[*]")
                            .Replace("%", "[%]")
                            .Replace("[[ ", "[[]")
                            .Replace(" ]]", "[]]")
                            .Replace("\'", "''") || txt品种.TextBoxToolTipText == "")
                    {
                        mdtHISDrug.DefaultView.RowFilter = "名称 like '%" + txt品种.TextBoxText + "%' Or 简码 like '%" + txt品种.TextBoxText + "%'";
                        frmZySelect frmSelect = new frmZySelect("HIS品种选择", mdtHISDrug.DefaultView.ToTable());
                        frmSelect.ShowDialog(this);
                        if (frmSelect.dataRowOut != null)
                        {
                            txt品种.TextBoxText = "[" + frmSelect.dataRowOut["His品种id"].ToString() + "]" + frmSelect.dataRowOut["名称"].ToString();
                            txt品种.TextBoxToolTipText = "[" + frmSelect.dataRowOut["His品种id"].ToString() + "]" + frmSelect.dataRowOut["名称"].ToString();
                            txt品种.Tag = frmSelect.dataRowOut["His品种id"].ToString();
                        }
                    }
                    else
                    {
                        e.Handled = true;
                        SendKeys.Send("{tab}");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txt品种_Validating(object sender, CancelEventArgs e)
        {
            if (txt品种.TextBoxToolTipText != txt品种.TextBoxText)
            {
                txt品种.TextBoxText = txt品种.TextBoxToolTipText;
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

        private void txt单位_Enter(object sender, EventArgs e)
        {
            txt单位.SelectAll();
        }

        private void txt描述_Enter(object sender, EventArgs e)
        {
            txt描述.SelectAll();
        }

        private void txt性状_Enter(object sender, EventArgs e)
        {
            txt性状.SelectAll();
        }

        private void txt药性_Enter(object sender, EventArgs e)
        {
            txt药性.SelectAll();
        }

        private void txt成分_Enter(object sender, EventArgs e)
        {
            txt成分.SelectAll();
        }

        private void txt用法_Enter(object sender, EventArgs e)
        {
            txt用法.SelectAll();
        }

        private void txt服法_Enter(object sender, EventArgs e)
        {
            txt服法.SelectAll();
        }

        private void txt禁忌_Enter(object sender, EventArgs e)
        {
            txt禁忌.SelectAll();
        }

        private void txt适应证_Enter(object sender, EventArgs e)
        {
            txt适应证.SelectAll();
        }

        private void txt作用_Enter(object sender, EventArgs e)
        {
            txt作用.SelectAll();
        }
    }
}
