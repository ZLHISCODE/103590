using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ZLSOFT.HIS.ZyEdit.From;
using ZLSOFT.HIS.ZyEdit.Model;

namespace ZLSOFT.HIS.ZyEdit
{

    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class ShowFrm : Ishowfrm
    {

        /// <summary>
        /// 中医辩证论治
        /// </summary>
        /// <param name="message">JSON字符串:
        /// {
        /// 来源系统, 使用类型, 病人ID, 挂号单,
        /// 门诊号,病人姓名,病人性别,病人年龄,病人民族,出生日期, 诊断ID,
        /// 当前科室ID, 当前科室名, 操作员ID,操作员姓名,
        /// 用户名, 用户密码
        /// }
        /// 说明：传参均为字符串类型
        /// 来源系统(0-ZLHIS/1-新门诊)
        /// 使用类型(0-新增/1-修改)
        /// </param>
        /// <returns>是否进行了新增或修改中医处方</returns>
        /// <returns>strOut JSON字符串:HIS医嘱ID,HIS诊断ID,诊断ID,处方ID</returns>
        public bool EditZyInfo(string message, out string strOut)
        {
            try
            {
                CheckDot();//检测dot4.0
                string dataMessage = "程序入口：EditZyInfo";
                ComLib.SaveDatalog(dataMessage);

                frmZyEdit frm = new frmZyEdit(message);
                frm.ShowDialog();
                strOut = frm.gstrOut;

                return frm.gstrOut != string.Empty;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "中医辩证论治", MessageBoxButtons.OK, MessageBoxIcon.Information);
                strOut = "";
                return false;
            }
        }

        /// <summary>
        /// 中医基础数据维护
        /// </summary>
        /// <param name="message">JSON字符串:
        /// {
        /// 来源系统, 站点, 操作员ID,操作员姓名,
        /// 用户名, 用户密码
        /// }
        /// 说明：传参均为字符串类型
        /// 来源系统(0-ZLHIS/1-新门诊)
        /// 使用类型(0-新增/1-修改)
        /// </param>
        /// <returns></returns>
        public bool EditZyBase(string message)
        {
            try
            {
                CheckDot();//检测dot4.0
                string dataMessage = "程序入口：EditZyBase";
                ComLib.SaveDatalog(dataMessage);
                frmBase frm = new frmBase(message);
                frm.ShowDialog();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "中医基础数据维护", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
        }


        /// <summary>
        /// 检查是否为中医辩证论治下达的医嘱或者诊断
        /// </summary>
        /// <param name="message">JSON字符串:
        /// {
        /// 来源系统, 病人ID, 挂号单,
        /// 门诊号,HIS医嘱ID,HIS诊断ID,用户名, 用户密码
        /// }
        /// 说明：传参均为字符串类型
        /// 来源系统(0-ZLHIS/1-新门诊)
        /// </param>
        /// <returns>CheckDiag:是否为中医辩证论治下达的医嘱或者诊断</returns>
        /// <returns>strOut:诊断ID|处方ID</returns>
        public bool CheckDiag(string message, out string strOut)
        {
            //基础数据
            BaseData baseData = new BaseData();
            string str医嘱id = string.Empty;
            string str诊断id = string.Empty;
            try
            {
                JObject jObject = (JObject)JsonConvert.DeserializeObject(message);
                if (jObject == null || message.Equals("[]"))
                {
                    throw new Exception("传入参数格式错误，应传入JSON格式字符串。" + message);

                }
                baseData.System = jObject["来源系统"].ToString();
                baseData.病人ID = jObject["病人ID"].ToString();
                baseData.挂号单 = jObject["挂号单"].ToString();
                baseData.门诊号 = jObject["门诊号"].ToString();
                baseData.UserName = jObject["用户名"].ToString();
                baseData.UserPassword = jObject["用户密码"].ToString();
                str医嘱id = jObject["HIS医嘱ID"].ToString();
                str诊断id = jObject["HIS诊断ID"].ToString();

                if (str医嘱id == "")
                {
                    throw new Exception("HIS医嘱ID不能为空,请检查传入参数格式。" + message);
                }

                DataTable dtTmp = new DataTable();
                message = JsonConvert.SerializeObject(new { HIS医嘱ID = str医嘱id, HIS诊断ID = str诊断id });
                dtTmp = DataCenter.GetTable(baseData, "Get_Diagid", message);
                if (dtTmp == null)
                {
                    strOut = "";
                    return false;
                }
                else
                {
                    if (dtTmp.Rows.Count == 0)
                    {
                        strOut = "";
                        return false;
                    }
                    else
                    {
                        strOut = dtTmp.Rows[0]["诊断ID"].ToString() + "|" + dtTmp.Rows[0]["处方ID"].ToString();
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "中医辩证论治", MessageBoxButtons.OK, MessageBoxIcon.Information);
                strOut = "";
                return false;
            }
        }

        public static bool CheckDot()
        {
            DirectoryInfo[] directories = new DirectoryInfo(
                Environment.SystemDirectory + @"\..\Microsoft.NET\Framework").GetDirectories("v?.?.*");
            foreach (DirectoryInfo info2 in directories)
            {
                if (ComLib.StrToInt(info2.Name.Substring(1, 1)) >= 4)
                {
                    return true;
                }
            }
            throw new Exception("检测到当前系统未安装Net4.0以上版本，无法执行当前功能。");
        }


    }
}
