using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using ZLSOFT.HIS.ZyEdit.Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;

namespace ZLSOFT.HIS.ZyEdit
{
    /// <summary>
    /// 数据处理中心
    /// </summary>
    static class DataCenter
    {
        static SQLHelper_Oracle m_SQLHelper_Oracle;

        /// <summary>
        /// 是否需要密码转换

        /// </summary>
        public static bool TranPasswd { get; set; }

        //从配置文件获取Oracle过程名

        private static string GetMethodName(string methodName)
        {
            //HIS调用
            if (TranPasswd == true) {
                return "Pkg_ZyEdit." + methodName;
            }
            else
            {
                object newName = ComLib.GetAppSettingValue(methodName, methodName);
                if (string.IsNullOrWhiteSpace(newName.ToString()))
                {
                    return methodName;
                }
                else
                {
                    return newName.ToString();
                }
            }
        }


        /// <summary>
        /// 执行功能
        /// </summary>
        /// <param name="methodName">执行方法名</param>
        /// <param name="message">传入消息，DataTable</param>
        /// <param name="userName">用户名</param>
        /// <param name="password">用户密码</param>
        /// <returns>返回结果，DataTable</returns>
        public static DataTable ExecuteFunc(string methodName, string message, string userName, string password)
        {
            string procedureName = GetMethodName(methodName);
            //记录数据日志
            ComLib.SaveDatalog("过程名：" + procedureName + "【" + methodName + "】");
            ComLib.SaveDatalog("入参：Input_In=" + message, false);

            if (m_SQLHelper_Oracle == null) { m_SQLHelper_Oracle = new SQLHelper_Oracle(); }
            DataTable result = m_SQLHelper_Oracle.ExcuteProcedureQueryData(userName, password, procedureName, message, TranPasswd);

            //记录数据日志
            ComLib.SaveDatalog("输出：" + result.Rows.Count, false);
            return result;
          
        }

        /// <summary>
        /// 执行功能
        /// </summary>
        /// <param name="methodName">执行方法名</param>
        /// <param name="message">传入消息，DataTable</param>
        /// <param name="userName">用户名</param>
        /// <param name="password">用户密码</param>
        /// <returns>返回结果，Json出参</returns>
        public static string DrugFunc(string methodName, string message, string userName, string password)
        {

            string procedureName = methodName;
            //记录数据日志
            ComLib.SaveDatalog("药品服务名：" + procedureName + "【" + methodName + "】");
            ComLib.SaveDatalog("药品服务入参：Input_In=" + message, false);

            if (m_SQLHelper_Oracle == null) { m_SQLHelper_Oracle = new SQLHelper_Oracle(); }
            string strOut = m_SQLHelper_Oracle.ExcuteProcedureJson(userName, password, procedureName, message, TranPasswd);

            //记录数据日志
            ComLib.SaveDatalog("药品服务调用成功", false);
            return strOut;
        }


        /// <summary>
        /// 获取DataTable数据
        /// </summary>
        /// <param name="data">调用方传入数据</param>
        /// <param name="pkName">包名</param>
        /// <param name="message">传入JSON参数</param>
        /// <returns></returns>
        public static DataTable GetTable(BaseData data,string pkName,string message)
        {
            try {
                DataTable dtTmp = ExecuteFunc(pkName, message, data.UserName, data.UserPassword);
                return dtTmp;
            }
            catch(Exception ex) {
                if (ex.Message == "用户名或密码错误！") //ORA-01017: invalid username/password; logon denied
                {
                    throw new Exception("用户名或密码错误！");
                }
                else {
                    return null;
                }
            }
        }

        /// <summary>
        /// 获取药品服务数据
        /// </summary>
        /// <param name="data">调用方传入数据</param>
        /// <param name="pkName">包名</param>
        /// <param name="message">传入JSON参数</param>
        /// <returns></returns>
        public static string GetDrugService(BaseData data, string pkName, string message)
        {
            try
            {
                string strOut = DrugFunc(pkName, message, data.UserName, data.UserPassword);
                return strOut;
            }
            catch (Exception ex)
            {
                if (ex.Message == "用户名或密码错误！") //ORA-01017: invalid username/password; logon denied
                {
                    throw new Exception("用户名或密码错误！");
                }
                else
                {
                    return null;
                }
            }
        }



        /// <summary>
        /// 获取数据库系统时间

        /// </summary>
        /// <param name="data">调用方传入数据</param>
        /// <returns></returns>
        public static DateTime Get_Now_Time(BaseData data)
        {
            DataTable result = ExecuteFunc("Get_Now_Time", "", data.UserName, data.UserPassword);
            if (result == null)
            {
                throw new Exception("获取服务器系统时间失败，服务器返回数据格式错误！");
            }
            return Convert.ToDateTime(result.Rows[0]["当前时间"]);
        }
    }
}
