using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using zlMsgBLL;
using Oracle.ManagedDataAccess.Client;
using zlMsgModel;

namespace zlShortMsg
{
    public partial class frmLog : Form
    {
        private delegate void UpdateLogDelegate(string strLog);
        private UpdateLogDelegate UpdateLog;
        
        private MsgDataBll DataBll = new MsgDataBll();
        private string mstrPath; //文件保存路径
        private string mstrLogFile;

        private void WriteLog(string strLog)
        {
            if (txtLog.Text == "")
            {
                txtLog.Text = strLog;
            }
            else
            {
                txtLog.Text = txtLog.Text + Environment.NewLine + strLog;
            }

            if (chkSave.Checked)
            {
                LogWriter.WriteLog(mstrLogFile, strLog);
            }
        }

        public frmLog()
        {
            UpdateLog = WriteLog;
            InitializeComponent();
        }

        private void frmLog_Load(object sender, EventArgs e)
        {
            chkSave.Checked = RegistryHelper.GetValue("SaveLog") != "0";

            mstrPath = RegistryHelper.GetValue("LogPath") == "" ? Environment.CurrentDirectory : RegistryHelper.GetValue("LogPath");
            txtPath.Text = mstrPath;
        }

        private void ChangeEventHandler(object sender, OracleNotificationEventArgs eventargs)
        {
            string strRowid = eventargs.Details.Rows[eventargs.Details.Rows.Count - 1]["Rowid"].ToString();
            int changeType = int.Parse(eventargs.Details.Rows[eventargs.Details.Rows.Count - 1]["Info"].ToString());

            if (strRowid != "" && changeType == (int)OracleNotificationInfo.Insert)
            {
                MsgData msgData = DataBll.GetMsgDataByRowid(strRowid);
                if (msgData == null) return;

                SendShortMsg(msgData);
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
                AppSecret = ZLSM4.Sm4DecryptEcb(msgData.Template.Provider.AppSecret),
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
                            strUrl = s.GetMessageUrl();
                            strResponse = HttpHelper.HttpGet(strUrl);
                            if (s.ResolveResponse(strResponse, ref strErrMessage, ref strErrCode))
                            {
                                strTip = DateTime.Now.ToString() + "  向手机号" + msgData.Receiver + "发送短信成功。";
                                msgData.State = 1;
                                msgData.MsgText = s.ResolveText(msgData.Template.TemplateText, msgData.Para);
                                msgData.SendDate = DateTime.Now;
                            }
                            else
                            {
                                strTip = DateTime.Now.ToString() + "  向手机号" + msgData.Receiver + "发送短信失败，错误信息： " + strErrMessage + " 错误代码 : " + strErrCode;
                                msgData.State = 2;
                                msgData.MsgText = s.ResolveText(msgData.Template.TemplateText, msgData.Para);
                                msgData.SendDate = DateTime.Now;
                            }

                            DataBll.UpdateMsgdata(msgData);
                            this.Invoke(UpdateLog, strTip);
                        });
                        break;
                    }
                case ("腾讯云"):
                    {
                        SendMsgTencent  s = new SendMsgTencent();
                        s.Clone(sendMsg);

                        string Header = "Conetent-Type||application/json";

                        ThreadPool.QueueUserWorkItem(state =>
                        {
                            strUrl = s.GetMessageUrl();
                            JObject jData = s.GetRequestBody();
                            strResponse = HttpHelper.HttpPost(strUrl, jData, Header);

                            if (s.ResolveResponse(strResponse, ref strErrMessage, ref strErrCode))
                            {
                                strTip = DateTime.Now.ToString() + "  向手机号" + msgData.Receiver + "发送短信成功。";
                                msgData.State = 1;
                                msgData.MsgText = s.ResolveText(msgData.Template.TemplateText, msgData.Para);
                                msgData.SendDate = DateTime.Now;
                            }
                            else
                            {
                                strTip = DateTime.Now.ToString() + "  向手机号" + msgData.Receiver + "发送短信失败，错误信息： " + strErrMessage + " 错误代码 : " + strErrCode;
                                msgData.State = 2;
                                msgData.MsgText = s.ResolveText(msgData.Template.TemplateText, msgData.Para);
                                msgData.SendDate = DateTime.Now;
                            }

                            DataBll.UpdateMsgdata(msgData);
                            this.Invoke(UpdateLog, strTip);
                        });

                        break;
                    }
                case ("华为云"):
                    {
                        SendMsgHuawei s = new SendMsgHuawei();
                        s.Clone(sendMsg);

                        ThreadPool.QueueUserWorkItem(state =>
                        {
                            strUrl = s.GetMessageUrl();
                            strResponse = HttpHelper.HttpPost(strUrl, s.GetRequestBody(), s.GetHeader());

                            if (s.ResolveResponse(strResponse, ref strErrMessage, ref strErrCode))
                            {
                                strTip = DateTime.Now.ToString() + "  向手机号" + msgData.Receiver + "发送短信成功。";
                                msgData.State = 1;
                                msgData.MsgText = s.ResolveText(msgData.Template.TemplateText, msgData.Para);
                                msgData.SendDate = DateTime.Now;
                            }
                            else
                            {
                                strTip = DateTime.Now.ToString() + "  向手机号" + msgData.Receiver + "发送短信失败，错误信息： " + strErrMessage + " 错误代码 : " + strErrCode;
                                msgData.State = 2;
                                msgData.MsgText = s.ResolveText(msgData.Template.TemplateText, msgData.Para);
                                msgData.SendDate = DateTime.Now;
                            }

                            DataBll.UpdateMsgdata(msgData);
                            this.Invoke(UpdateLog, strTip);
                        });

                        break;    
                    }
                default:
                    {
                        break;
                    }
            }
        }

        /// <summary>
        /// 创建日志文件
        /// </summary>
        private void CreateFile()
        {
            if (!chkSave.Checked) return;

            if (!Directory.Exists(mstrPath))
            {
                Directory.CreateDirectory(mstrPath);
            }

            mstrLogFile = mstrPath + @"\MsgLog" + DateTime.Today.ToString("yyyy-M-d") + ".txt";
            if (!File.Exists(mstrLogFile))
            {
                File.CreateText(mstrLogFile).Close();
            }
        }

        private void txtPath_TextChanged(object sender, EventArgs e)
        {
            CreateFile();       //当路径发生变化时,就创建新的日志文件
        }

        private void chkSave_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSave.Checked)
            {
                CreateFile();
                RegistryHelper.SetValue("SaveLog", "1");
            }
            else
            {
                RegistryHelper.SetValue("SaveLog", "0");
            }
            
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtLog.Text = "";
        }

        public void StartDcn()
        {
            WriteLog(DateTime.Now.ToString() + "正在开启短信服务...");
            OnChangeEventHandler changeEventHandler = ChangeEventHandler;
            DataBll.RegistDcn(changeEventHandler);
            WriteLog(DateTime.Now.ToString() + "短信服务开启成功");
        }

        public void EndDcn()
        {
            WriteLog(DateTime.Now.ToString() + "正在关闭短信服务...");
            DataBll.UnRegistDcn();
            WriteLog(DateTime.Now.ToString() + "短信服务关闭成功");
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = folderBrowser.ShowDialog(this);
            if (dialogResult == DialogResult.Yes || dialogResult == DialogResult.OK)
            {
                mstrPath = folderBrowser.SelectedPath;
                txtPath.Text = mstrPath;
                RegistryHelper.SetValue("LogPath", mstrPath);
            }
        }
    }
}
