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
    partial class frmDiagEdit : Form
    {

        public BaseData m_baseData = new BaseData(); //基础数据
        private int mintType;//0=新增,1=修改
        private DataRow mdtRow;//修改时数据加载集
        private DataTable mdt疾病;//疾病数据集
        public string gstrID;
        public string gstrOut;
        public bool gblnOK;
        public frmDiagEdit(int intType, DataRow dtRow, DataTable dt疾病, BaseData baseData)
        {
            mintType = intType;
            mdtRow = dtRow;
            mdt疾病 = dt疾病;
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
                    MessageBox.Show("疾病名称不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtName.Focus();
                    return;
                }

                if (cbo科别.Text == "")
                {
                    MessageBox.Show("科别不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cbo科别.Focus();
                    return;
                }

                //保存

                string message = JsonConvert.SerializeObject(new
                {
                    USETYPE = mintType.ToString(),////使用类型(0-新增/1-修改)
                    疾病ID = gstrID,
                    疾病名称 = txtName.Text,
                    简码 = txt简码.Text,
                    科别 = cbo科别.Text,
                    操作员名称 = m_baseData.OperatorName,
                    操作员ID = m_baseData.OperatorID
                });

                DataTable dtOut = new DataTable();

                dtOut = DataCenter.ExecuteFunc("Set_Distype", message, m_baseData.UserName, m_baseData.UserPassword);
                if (dtOut != null)
                {
                    gstrID = dtOut.Rows[0]["疾病ID"].ToString();
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

        private void LoadData()
        {
            try
            {
                txtName.Text = mdtRow["疾病名称"].ToString();
                txt简码.Text = mdtRow["简码"].ToString();
                cbo科别.Text = mdtRow["科别"].ToString();
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

        private void txtName_Validating(object sender, CancelEventArgs e)
        {

            txt简码.Text = txtName.Text == "" ? "" : Pinyin.GetInitials(txtName.Text, Encoding.GetEncoding("GB2312"));
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

        private void frmDiagEdit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void frmDiagEdit_Load(object sender, EventArgs e)
        {
            try
            {
                string strTmp = "";

                //加载科别
                if (mdt疾病 != null)
                {
                    cbo科别.Items.Clear();
                    for (int i = 0; i < mdt疾病.Rows.Count; i++)
                    {
                        if (strTmp.IndexOf(mdt疾病.Rows[i]["科别"].ToString()) < 0)
                        {
                            cbo科别.Items.Add(mdt疾病.Rows[i]["科别"].ToString());
                            strTmp = strTmp + mdt疾病.Rows[i]["科别"].ToString() + ",";
                        }
                    }
                }

                if (mintType == 1)
                {
                    LoadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void cbo科别_KeyPress(object sender, KeyPressEventArgs e)
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

        private void cbo科别_Enter(object sender, EventArgs e)
        {
            cbo科别.SelectAll();
        }
    }
}
