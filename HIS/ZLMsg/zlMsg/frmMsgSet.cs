using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using zlMsgModel;
using zlMsgBLL;
using System.Reflection;

namespace zlShortMsg
{
    public partial class frmMsgSet : Form
    {
        private MsgTemplateBLL templateBLL = new MsgTemplateBLL();

        public frmMsgSet()
        {
            InitializeComponent();
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            frmAddSetting frm = new frmAddSetting();
            MsgTemplate m = frm.ShowDialog();

            if (m != null)
            {
                AddTemplate2Grid(m);
            }
        }

        private void AddTemplate2Grid(MsgTemplate msgTemplate)
        {
            DataGridViewRow  gridViewRow = new DataGridViewRow();
            dgvSetting.Rows.Add(gridViewRow);

            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["Templateid"].Value = msgTemplate.TemplateID.ToString() ?? "";
            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["Providercode"].Value = msgTemplate.Provider.ProviderCode.ToString() ?? "";
            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["Providername"].Value = msgTemplate.Provider.PrividerName ?? "";
            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["Templatecode"].Value = msgTemplate.TemplateCode ?? "";
            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["Signcode"].Value = msgTemplate.SignCode?? "";
            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["SignNumber"].Value = msgTemplate.SignNumber ?? "";
            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["Templatepara"].Value = msgTemplate.TemplatePara ?? "";
            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["Appkey"].Value = msgTemplate.Provider.AppKey?? "";
            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["Appsecret"].Value = msgTemplate.Provider.AppSecret ?? "";
            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["Templatetext"].Value = msgTemplate.TemplateText ?? "";
            dgvSetting.Rows[dgvSetting.RowCount - 1].Cells["Templatekind"].Value = msgTemplate.TemplateKind.ToString() ?? "";
        }


        private void LoadSetting()
        {
            dgvSetting.Rows.Clear();
            List<MsgTemplate> lists = templateBLL.GetTemplates();
           
            foreach (MsgTemplate msgTemplate in lists)
            {
                AddTemplate2Grid(msgTemplate);
            }
        }

        private void frmMsgSet_Load(object sender, EventArgs e)
        {
            //添加首列的选项框
            DataGridViewCheckBoxHeaderCell boxColumn = new DataGridViewCheckBoxHeaderCell();
            boxColumn.OnCheckBoxClicked += BoxColumn_OnCheckBoxClicked; ;

            DataGridViewCheckBoxColumn boxHeader = new DataGridViewCheckBoxColumn();
            boxHeader.HeaderCell = boxColumn;
            boxHeader.HeaderText = "  ";
            dgvSetting.Columns.Add(boxHeader);

            //初始化表格样式
            DgvDrawer.ChangeDgvStyle(ref dgvSetting, "Templateid:Templateid:1;Providercode:Providercode:0;平台:Providername:1;" +
                                                                            "模板Code:Templatecode:1;模版签名:Signcode:1;签名通道:SignNumber:1;" +
                                                                            "参数名:Templatepara:1;分类:Templatekind:1;" +
                                                                            "模版文本:Templatetext:1;Appkey:Appkey:0;Appsecret:Appsecret:0");
            //加载数据
            LoadSetting();
        }

        private void BoxColumn_OnCheckBoxClicked(object sender, datagridviewCheckboxHeaderEventArgs e)
        {
            dgvSetting.EndEdit();   //需要停止编辑状态,否则在编辑的行不会发生值变化

            foreach (DataGridViewRow r in dgvSetting.Rows)
            {
                r.Cells[0].Value = e.CheckedState;
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            List<long> templateIdList = new List<long>();
            List<DataGridViewRow> selectRows = new List<DataGridViewRow>();

            foreach (DataGridViewRow row in dgvSetting.Rows)
            {
                if ((row.Cells[0].Value==null ? "False" : row.Cells[0].Value.ToString()) == "True")
                {
                    templateIdList.Add(long.Parse(row.Cells["Templateid"].Value.ToString()));
                    selectRows.Add(row);
                }
            }

            if (templateIdList.Count > 0)
            {
                if (templateBLL.DeleteTempLateByid(templateIdList))
                {
                    //删除成功后,在界面删除行
                    foreach (DataGridViewRow r in selectRows)
                    {
                        dgvSetting.Rows.Remove(r);
                    }

                    MessageBox.Show("删除成功!", "提示");
                }
                else
                {
                    MessageBox.Show("删除失败!", "提示");
                }
            }
            else
            {
                MessageBox.Show("请勾选后再进行删除操作。", "提示");
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = dgvSetting.CurrentRow;

            MsgProvider provider = new MsgProvider(long.Parse(row.Cells["Providercode"].Value.ToString()),
                                                                            row.Cells["Providername"].Value.ToString(),
                                                                            row.Cells["Appkey"].Value.ToString(),
                                                                            row.Cells["Appsecret"].Value.ToString()
                                                                            );

            MsgTemplate template = new MsgTemplate(long.Parse(row.Cells["Templateid"].Value.ToString()),
                                                                                provider,
                                                                                row.Cells["Templatecode"].Value.ToString(), 
                                                                                row.Cells["Signcode"].Value.ToString(),
                                                                                row.Cells["SignNumber"].Value.ToString(),
                                                                                row.Cells["Templatepara"].Value.ToString(),
                                                                                row.Cells["Templatekind"].Value.ToString(),
                                                                                row.Cells["Templatetext"].Value.ToString());

            template = new frmAddSetting().ShowDialog(template);

            if (template != null)
            {
                //修改当前行的信息
                row.Cells["Templateid"].Value = template.TemplateID.ToString() ?? "";
                row.Cells["Providercode"].Value = template.Provider.ProviderCode.ToString() ?? "";
                row.Cells["Providername"].Value = template.Provider.PrividerName ?? "";
                row.Cells["Templatecode"].Value = template.TemplateCode ?? "";
                row.Cells["Signcode"].Value = template.SignCode?? "";
                row.Cells["Templatepara"].Value = template.TemplatePara ?? "";
                row.Cells["Templatetext"].Value = template.TemplateText ?? "";
                row.Cells["Templatekind"].Value = template.TemplateKind ?? "";

                //如果Provider的Key和Value发生了ji变化 ,需要同步修改其他行的数据
                if (row.Cells["Appkey"].Value.ToString() != template.Provider.AppKey ||row.Cells["Appsecret"].Value.ToString() != template.Provider.AppSecret)
                {
                    row.Cells["Appkey"].Value = template.Provider.AppKey ?? "";
                    row.Cells["Appsecret"].Value = template.Provider.AppSecret ?? "";

                    foreach (DataGridViewRow r in  dgvSetting.Rows)
                    {
                        if (long.Parse(r.Cells["Providercode"].Value.ToString()) == template.Provider.ProviderCode)
                        {
                            r.Cells["Appkey"].Value = template.Provider.AppKey ?? "";
                            r.Cells["Appsecret"].Value = template.Provider.AppSecret ?? "";
                        }
                    }
                }
            }

        }
    }
}
