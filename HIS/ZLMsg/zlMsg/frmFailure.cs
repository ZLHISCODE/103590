using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using zlMsgBLL;
using zlMsgModel;

namespace zlShortMsg
{
    public partial class frmFailure : Form
    {
        MsgDataBll dataBll = new MsgDataBll();
        private long lngTmp = 0;
        private MsgDataBll DataBll = new MsgDataBll();
        private string mstrLogFile = "";


        public frmFailure()
        {
            InitializeComponent();
        }

        private void frmFailure_Load(object sender, EventArgs e)
        {
            //添加首列的选项框
            DataGridViewCheckBoxHeaderCell boxColumn = new DataGridViewCheckBoxHeaderCell();
            boxColumn.OnCheckBoxClicked += BoxColumn_OnCheckBoxClicked; ;

            DataGridViewCheckBoxColumn boxHeader = new DataGridViewCheckBoxColumn();
            boxHeader.HeaderCell = boxColumn;
            boxHeader.HeaderText = "  ";
            dgvList.Columns.Add(boxHeader);

            DgvDrawer.ChangeDgvStyle(ref dgvList, "Id:Id:0;接收人:Receiver:1;参数:Para:1;发送人:Sender:1;" +
                                                    "发送端:Terminal:1;状态:State:1;发送时间:Senddate:1;Extend:Extend:0;" +
                                                     "短信文本:MsgText:1;" +
                                                      "Templateid:Templateid:0;Providercode:Providercode:0;平台:Providername:1" +
                                                     "Templatecode:Templatecode:0;Signcode:Signcode:0;Templatecode:Templatecode:0;" +
                                                     "Templatepara:Templatepara:0;Templatekind:Templatekind:0;" +
                                                     "Templatetext:Templatetext:0;Appkey:Appkey:0;Appsecret:Appsecret:0");
            LoadMsgData();
            mstrLogFile = RegistryHelper.GetValue("LogPath") == "" ? Environment.CurrentDirectory : RegistryHelper.GetValue("LogPath") + @"\MsgLog" + DateTime.Today.ToString("yyyy-M-d") + ".txt";

            chkSend.Checked = RegistryHelper.GetValue("Resend").Val() == 1;
        }

        private void BoxColumn_OnCheckBoxClicked(object sender, datagridviewCheckboxHeaderEventArgs e)
        {
            dgvList.EndEdit();   //需要停止编辑状态,否则在编辑的行不会发生值变化

            foreach (DataGridViewRow r in dgvList.Rows)
            {
                r.Cells[0].Value = e.CheckedState;
            }
        }

        private void LoadMsgData()
        {
            dgvList.Rows.Clear();
            List<MsgData> lists = dataBll.GetMsgDatas();

            if (lists==null) return;

            foreach (MsgData msgData in lists)
            {
                AddMsg2Grid(msgData);
            }
        }

        private void AddMsg2Grid(MsgData msgData)
        {
            DataGridViewRow gridViewRow = new DataGridViewRow();
            dgvList.Rows.Add(gridViewRow);

            dgvList.Rows[dgvList.RowCount - 1].Cells["Id"].Value = msgData.ID.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Receiver"].Value = msgData.Receiver.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Para"].Value = msgData.Para.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Sender"].Value = msgData.Sender.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["State"].Value = msgData.State.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Kind"].Value = msgData.Kind.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Senddate"].Value = msgData.SendDate.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Extend"].Value = msgData.Extend.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["MsgText"].Value = msgData.MsgText.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Terminal"].Value = msgData.Terminal.ToString() ?? "";

            dgvList.Rows[dgvList.RowCount - 1].Cells["Templateid"].Value = msgData.Template.TemplateID.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Providercode"].Value = msgData.Template.Provider.ProviderCode.ToString() ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Providername"].Value = msgData.Template.Provider.PrividerName ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Templatecode"].Value = msgData.Template.TemplateCode ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Signcode"].Value = msgData.Template.SignCode ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Templatepara"].Value = msgData.Template.TemplatePara ?? "";

            dgvList.Rows[dgvList.RowCount - 1].Cells["Appkey"].Value = msgData.Template.Provider.AppKey ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Appsecret"].Value = msgData.Template.Provider.AppSecret ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Templatetext"].Value = msgData.Template.TemplateText ?? "";
            dgvList.Rows[dgvList.RowCount - 1].Cells["Templatekind"].Value = msgData.Template.TemplateKind.ToString() ?? "";
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            lngTmp++;
            //每小时重新获取一次列表
            if (lngTmp >= 360) LoadMsgData();

            if (!chkSend.Checked) return;

            //每隔10S检查一次发送失败的短信,进行重发
            MsgProvider provider = null;
            MsgTemplate template = null;
            MsgData data = null;

            foreach (DataGridViewRow row in dgvList.Rows)
            {
                provider = new MsgProvider((long)row.Cells["Providercode"].Value.ToString().Val(), row.Cells["Providername"].Value.ToString(),
                                                            row.Cells["Appkey"].Value.ToString(), row.Cells["Appsecret"].Value.ToString());

                template = new MsgTemplate((long)row.Cells["Templateid"].Value.ToString().Val(), provider, row.Cells["Templatecode"].Value.ToString(), row.Cells["Signcode"].Value.ToString(), row.Cells["SignNumber"].Value.ToString() , row.Cells["Templatepara"].Value.ToString()
                                                            , row.Cells["Templatekind"].Value.ToString(), row.Cells["Templatetext"].Value.ToString());

                data = new MsgData((long)row.Cells["Id"].Value.ToString().Val(), template, (long)row.Cells["Receiver"].Value.ToString().Val(), row.Cells["Para"].Value.ToString(),
                                                row.Cells["Sender"].Value.ToString(), row.Cells["Terminal"].Value.ToString(), (int)row.Cells["State"].Value.ToString().Val(), (int)row.Cells["Kind"].Value.ToString().Val(),
                                                row.Cells["MsgText"].Value.ToString(), row.Cells["Extend"].Value.ToString(), Convert.ToDateTime(row.Cells["Senddate"].Value.ToString()));
                SendShortMsg(data);
                dgvList.Rows.Remove(row);
            }

        }


        /// <summary>
        /// 传入短信数据,发送短信
        /// </summary>
        /// <param name="msgData"></param>
        private void SendShortMsg(MsgData msgData)
        {
            //初始化对象并赋值
            SendMsg sendMsg = new SendMsg
            {
                AppKey = msgData.Template.Provider.AppKey,
                AppSecret = msgData.Template.Provider.AppSecret,
                SignName = msgData.Template.SignCode,
                Tel = msgData.Receiver.ToString(),
                TemplateCode = msgData.Template.TemplateCode,
                TemplateParas = msgData.Template.TemplatePara,
                Paras = msgData.Para
            };

            string strUrl = "";
            string strResponse = "";
            string strErrCode = "";
            string strErrMessage = "";
            string strTip = "";

            switch (msgData.Template.Provider.PrividerName)
            {
                case ("阿里云"):
                    {
                        SendMsgAli s = new SendMsgAli();
                        s.Clone(sendMsg);       //复制父类属性

                        //使用Lambda表达式开启线程,避免阻塞,同时回写数据
                        ThreadPool.QueueUserWorkItem(state =>
                        {
                            strUrl = s.GetMessageUrl() + ";";
                            strResponse = HttpHelper.HttpGet(strUrl);
                            if (s.ResolveResponse(strResponse, ref strErrMessage, ref strErrCode))
                            {
                                strTip = DateTime.Now.ToString() + "  向手机号" + msgData.Receiver + "重发短信成功。";
                                msgData.State = 1;
                                msgData.MsgText = s.ResolveText(msgData.Template.TemplateText, msgData.Para);
                                msgData.SendDate = DateTime.Now;
                            }
                            else
                            {
                                strTip = DateTime.Now.ToString() + "  向手机号" + msgData.Receiver + "重发短信失败，错误信息： " + strErrMessage + " 错误代码 : " + strErrCode;
                                msgData.State = 2;
                                msgData.MsgText = s.ResolveText(msgData.Template.TemplateText, msgData.Para);
                                msgData.SendDate = DateTime.Now;
                            }
                            
                            LogWriter.WriteLog(mstrLogFile,strTip);
                            DataBll.UpdateMsgdata(msgData);
                        });
                        break;
                    }
                case ("腾讯云"):
                    {
                        break;
                    }
                default:
                    {
                        break;
                    }
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            List<long> msgList = new List<long>();
            List<DataGridViewRow> selectRows = new List<DataGridViewRow>();

            foreach (DataGridViewRow row in dgvList.Rows)
            {
                if ((row.Cells[0].Value == null ? "False" : row.Cells[0].Value.ToString()) == "True")
                {
                    msgList.Add(long.Parse(row.Cells["Id"].Value.ToString()));
                    selectRows.Add(row);
                }
            }

            if (msgList.Count > 0)
            {
                if (dataBll.DeleteMsgData(msgList))
                {
                    //删除成功后,在界面删除行
                    foreach (DataGridViewRow r in selectRows)
                    {
                        dgvList.Rows.Remove(r);
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

        private void btnResend_Click(object sender, EventArgs e)
        {

            DataGridViewRow row = dgvList.SelectedRows[0];

            MsgProvider provider = new MsgProvider((long)row.Cells["Providercode"].Value.ToString().Val(), row.Cells["Providername"].Value.ToString(),
                row.Cells["Appkey"].Value.ToString(), row.Cells["Appsecret"].Value.ToString());

            MsgTemplate template = new MsgTemplate((long)row.Cells["Templateid"].Value.ToString().Val(), provider, row.Cells["Templatecode"].Value.ToString(), row.Cells["Signcode"].Value.ToString(), row.Cells["SignNumber"].Value.ToString(), row.Cells["Templatepara"].Value.ToString()
                , row.Cells["Templatekind"].Value.ToString(), row.Cells["Templatetext"].Value.ToString());

            MsgData data = new MsgData((long)row.Cells["Id"].Value.ToString().Val(), template, (long)row.Cells["Receiver"].Value.ToString().Val(), row.Cells["Para"].Value.ToString(),
                row.Cells["Sender"].Value.ToString(), row.Cells["Terminal"].Value.ToString(), (int)row.Cells["State"].Value.ToString().Val(), (int)row.Cells["Kind"].Value.ToString().Val(),
                row.Cells["MsgText"].Value.ToString(), row.Cells["Extend"].Value.ToString(), Convert.ToDateTime(row.Cells["Senddate"].Value.ToString()));

            SendShortMsg(data);
        }

        private void chkSend_CheckedChanged(object sender, EventArgs e)
        {
            RegistryHelper.SetValue("Resend",chkSend.Checked?"1":"0");
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadMsgData();
        }
    }
}
  