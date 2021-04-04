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
    partial class frmZxEdit : Form
    {

        public BaseData m_baseData = new BaseData(); //基础数据
        private int mintType;//0=新增,1=修改
        private DataRow mdtRow;//修改时数据加载集
        public string gstrID;
        private string mstr疾病ID;
        public bool gblnOK;
        public frmZxEdit(int intType, DataRow dtRow, BaseData baseData, string str疾病ID)
        {
            mintType = intType;
            mdtRow = dtRow;
            m_baseData = baseData;
            mstr疾病ID = str疾病ID;
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
                    MessageBox.Show("证型名称不能为空，请重新录入。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtName.Focus();
                    return;
                }

                //保存

                string message = JsonConvert.SerializeObject(new
                {
                    USETYPE = mintType.ToString(),////使用类型(0-新增/1-修改)
                    证型ID = gstrID,
                    证型名称 = txtName.Text,
                    简码 = txt简码.Text,
                    疾病ID = mstr疾病ID,
                    证型描述 = txt证型描述.Text,
                    证型治法 = txt证型治法.Text,
                    症状表现 = txt症状表现.Text,
                    操作员名称 = m_baseData.OperatorName,
                    操作员ID = m_baseData.OperatorID
                });

                DataTable dtOut = new DataTable();

                dtOut = DataCenter.ExecuteFunc("Set_Zxtype", message, m_baseData.UserName, m_baseData.UserPassword);
                if (dtOut != null)
                {
                    gstrID = dtOut.Rows[0]["证型ID"].ToString();
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
                txtName.Text = mdtRow["证型名称"].ToString();
                txt简码.Text = mdtRow["简码"].ToString();
                txt证型描述.Text = mdtRow["证型描述"].ToString();
                txt证型治法.Text = mdtRow["证型治法"].ToString();
                txt症状表现.Text = mdtRow["症状表现"].ToString();
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


        private void frmZxEdit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void frmZxEdit_Load(object sender, EventArgs e)
        {
            if (mintType == 1)
            {
                LoadData();
            }
        }

        private void frmZxEdit_Shown(object sender, EventArgs e)
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

        private void txt证型治法_Enter(object sender, EventArgs e)
        {
            txt证型治法.SelectAll();
        }

        private void txt证型描述_Enter(object sender, EventArgs e)
        {
            txt证型描述.SelectAll();
        }

        private void txt症状表现_Enter(object sender, EventArgs e)
        {
            txt症状表现.SelectAll();
        }
    }
}
