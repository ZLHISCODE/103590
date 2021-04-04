using DevExpress.LookAndFeel;
using DevExpress.UserSkins;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZLSOFT.HIS.PreTriage.ComLib;
using ZLSOFT.HIS.PreTriage.Models;

namespace ZLSOFT.HIS.PreTriage
{
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class ShowFrm
    {
        /// <summary>
        /// 急诊预检分诊工作站

        /// </summary>
        /// <param name="message">JSON字符串:
        /// {
        /// tnsname,用户名, 用户密码,站点
        /// 操作员id,操作员姓名,操作员编码

        /// }
        /// 说明：传参均为字符串类型
        /// </param>
        /// <returns></returns>
        /// 
        [HandleProcessCorruptedStateExceptions] //.net4支持对非托管异常的捕获

        public bool EditPreTriage(string message, ADODB.Connection cn)
        {
            try
            {
                JObject jObject = (JObject)JsonConvert.DeserializeObject(message);
                if (jObject == null || message.Equals("[]"))
                {
                    throw new Exception("传入参数格式错误，应传入JSON格式字符串。" + message);
                }

                BaseData.SYS = jObject["系统"].ToString();
                BaseData.站点 = jObject["站点"].ToString();
                BaseData.OperatorID = jObject["操作员id"].ToString();
                BaseData.OperatorName = jObject["操作员姓名"].ToString();
                BaseData.OperatorCode = jObject["操作员编码"].ToString();
                BaseData.gcnOracle = cn;//用于调用报表打印

                //正式用的连接方法
                SQLHelper_Oracle.IntData(jObject["用户名"].ToString(), jObject["用户密码"].ToString(), jObject["tnsname"].ToString());
                string dataMessage = "程序入口：EditPreTriage";

                ComLib.ComLib.SaveDatalog(dataMessage);

                BonusSkins.Register();

                UserLookAndFeel.Default.SetSkinStyle("Office 2007 Pink");

                //初始化Oracle连接
                BaseData.OracleCnn = SQLHelper_Oracle.GetOdpConnection();

                frmMain frm = new frmMain();
                frm.ShowDialog();
                return true;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, "急诊预检分诊工作站", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
        }
    }
}
