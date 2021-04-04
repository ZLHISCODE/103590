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
    partial class frmAddDisEdit : Form
    {

        public BaseData m_baseData = new BaseData(); //基础数据

        private int mintUseType;//0=加症疾病 1=加症治法
        private int mintType;//0=新增,1=修改
        private DataGridViewRow mdtRow;//修改时数据加载集
        public string gstrID;
        public string gstrOut;
        private string mstrTmp;
        public bool gblnOK;
        public frmAddDisEdit(int intUseType, int intType, DataGridViewRow dtRow, BaseData baseData, string strTmp = "")
        {
            mintType = intType;
            mintUseType = intUseType;
            mdtRow = dtRow;
            m_baseData = baseData;
            mstrTmp = strTmp;
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
                    MessageBox.Show(mintUseType == 0 ? "加症" : "治法" + "名称不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtName.Focus();
                    return;
                }
                //保存
                string message = "";
                if (mintUseType == 0)
                {
                    message = JsonConvert.SerializeObject(new
                    {
                        USETYPE = mintType.ToString(),////使用类型(0-新增/1-修改)
                        加症ID = gstrID,
                        加症名称 = txtName.Text,
                        简码 = txt简码.Text,
                        操作员名称 = m_baseData.OperatorName,
                        操作员ID = m_baseData.OperatorID
                    });
                }
                else
                {
                    message = JsonConvert.SerializeObject(new
                    {
                        USETYPE = mintType.ToString(),////使用类型(0-新增/1-修改)
                        治法ID = gstrID,
                        治法名称 = txtName.Text,
                        简码 = txt简码.Text,
                        加症ID = mstrTmp,
                        操作员名称 = m_baseData.OperatorName,
                        操作员ID = m_baseData.OperatorID
                    });
                }

                DataTable dtOut = new DataTable();

                dtOut = DataCenter.ExecuteFunc(mintUseType == 0 ? "Set_Adddis" : "Set_Addzf", message, m_baseData.UserName, m_baseData.UserPassword);
                if (dtOut != null)
                {
                    gstrID = dtOut.Rows[0][mintUseType == 0 ? "加症ID" : "治法ID"].ToString();
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
                txtName.Text = mdtRow.Cells[mintUseType == 0 ? "加症名称" : "治法名称"].Value.ToString();
                txt简码.Text = mdtRow.Cells["简码"].Value.ToString();
                gstrID = mdtRow.Cells["id"].Value.ToString();
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
            if (mintUseType == 0)
            {
                label1.Text = "加症名称";
            }
            else
            {
                label1.Text = "治法名称";
            }
            if (mintType == 1)
            {
                LoadData();
            }
        }

        private void frmAddDisEdit_Shown(object sender, EventArgs e)
        {
            txtName.Focus();
        }

        private void txtName_Enter(object sender, EventArgs e)
        {
            txtName.SelectAll();
        }

        private void txt简码_Enter(object sender, EventArgs e)
        {
            txt简码.SelectAll();
        }
    }
}
