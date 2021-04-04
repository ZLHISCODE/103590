using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;

namespace ZLSOFT.HIS.ZyEdit
{
    /// <summary>
    /// 公共方法
    /// </summary>
    internal static class ComLib
    {

        public static bool blnConfig=true;
        public static Dictionary<string, string> dtcTns = null;
        public static string strErrLog = string.Empty;
        public static string strDataLog = string.Empty;
        public static string strTnsName = string.Empty;

        /// <summary>
        /// 从Oracle错误信息中获取提示信息
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public static string GetMessageFromOracleError(string message)
        {
            //处理Oracle过程抛出的错误信息，取出提示信息
            if (message.Contains("[ZLSOFT]"))
            {
                string[] messages = message.Split(new string[] { "[ZLSOFT]" }, StringSplitOptions.None);
                if (messages.Length > 2) { message = messages[1]; }
            }
            return message;
        }

        /// <summary>
        /// 读取配置文件值
        /// </summary>
        /// <param name="name">Key值</param>
        /// <returns></returns>
        public static object GetAppSettingValue(string name, object defaultValue)
        {
            object saveErrorlog = null;
            if (blnConfig == true) { 
                try
                {
                    Configuration fig = GetDllConfig();//读取Dll自己的配置文件
                    saveErrorlog = fig.AppSettings.Settings[name].Value;
                }
                catch
                {
                    //忽略错误
                }
            }
            return saveErrorlog == null ? defaultValue : saveErrorlog;
        }

        /// <summary>
        /// 获取当前应用程序集文件的位置(完全路径)
        /// </summary>
        /// <returns></returns>
        public static string GetAssemblyLocation()
        {
            //获取调用当前正在执行的方法的方法的 Assembly  
            Assembly assembly = Assembly.GetCallingAssembly();
            return assembly.Location;
        }

        /// <summary>
        /// 类库调用自身所带的配置文件中的配置信息，而不是读取应用程序所带的配置信息
        /// </summary>
        /// <returns></returns>
        public static Configuration GetDllConfig()
        {
            string path = string.Format("{0}.config", GetAssemblyLocation());

            if (File.Exists(path) == false)
            {
                blnConfig = false;
                return null;
            }

            ExeConfigurationFileMap configFile = new ExeConfigurationFileMap();
            configFile.ExeConfigFilename = path;
            return ConfigurationManager.OpenMappedExeConfiguration(configFile, ConfigurationUserLevel.None);
        }

        /// <summary>
        /// 保存错误日志
        /// </summary>
        /// <param name="errorMessage"></param>
        public static void SaveErrorlog(string errorMessage, bool addDateTime = true)
        {
            if (strErrLog == string.Empty) {
                strErrLog = GetAppSettingValue("SaveErrorlog", "0").ToString();
            }
            if (!strErrLog.Equals("1")) { return; }
            try
            {
                //记录错误日志       
                string message = string.Empty;
                if (addDateTime)
                {
                    message += DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + ":" + Environment.NewLine;
                }
                message += errorMessage.Replace("\\r\\n   在", Environment.NewLine + "在");
                message += Environment.NewLine;

                string filePath = Path.GetDirectoryName(GetAssemblyLocation());
                string fileName = "ZlSoft_HIS_ZyEdit_ErrorLog_" + DateTime.Now.ToString("yyyy_MM_dd") + ".log";

                Directory.CreateDirectory(filePath);//目录存在时不会创建
                File.AppendAllText(Path.Combine(filePath, fileName), message);
            }
            catch
            {
                //忽略错误
            }
        }




        /// <summary>
        /// 保存数据日志
        /// </summary>
        /// <param name="dataMessage"></param>
        public static void SaveDatalog(string dataMessage, bool addDateTime = true)
        {
            if (strDataLog == string.Empty)
            {
                strDataLog = GetAppSettingValue("SaveDatalog", "0").ToString();
            }
            if (!strDataLog.Equals("1")) { return; }
            try
            {
                //记录日志    
                string message = string.Empty;
                if (addDateTime)
                {
                    message += DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + ":" + Environment.NewLine;
                }
                message += dataMessage;
                message += Environment.NewLine;

                string filePath = Path.GetDirectoryName(GetAssemblyLocation());
                string fileName = "ZlSoft_HIS_ZyEdit_DataLog_" + DateTime.Now.ToString("yyyy_MM_dd") + ".log";

                Directory.CreateDirectory(filePath);//目录存在时不会创建
                File.AppendAllText(Path.Combine(filePath, fileName), message);
            }
            catch
            {
                //忽略错误
            }
        }

        /// <summary>
        /// DataRow扩展方法，转换为字符
        /// </summary>
        /// <param name="dataRow"></param>
        /// <param name="propertyName"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static string Nvl(this DataRow dataRow, string propertyName, string defaultValue = "")
        {
            if (dataRow.IsNull(propertyName))
            {
                return defaultValue;
            }
            else
            {
                return dataRow[propertyName].ToString();
            }
        }


        /// <summary>
        /// 获取控件在父窗体的位置
        /// </summary>
        /// <param name="objCtl"></param>
        /// <param name="intTop"></param>
        /// <param name="intLeft"></param>
        /// <returns></returns>
        public static Boolean GetCtlPos(Control objCtl, ref int intTop, ref int intLeft)
        {
            try {
                Control objTmp = new Control();
                intTop = intTop + objCtl.Top;
                intLeft = intLeft + objCtl.Left;
                objTmp = objCtl;
                do
                {
                    objTmp = objTmp.Parent;
                    if(objTmp != null && !(objTmp is  Form))
                    {
                        intTop = intTop + objTmp.Top;
                        intLeft = intLeft + objTmp.Left;
                    }
                 }
                 while (objTmp != null);
            }
            catch
            {
                //忽略错误
            }
            return true;
        }


        /// <summary>
        /// DataRow的扩展方法，克隆
        /// </summary>
        /// <param name="drItem">DataRow对象</param>
        /// <param name="drNewItem">要复制到的DataTable对象</param>
        /// <returns></returns>
        public static void CloneTo(this DataRow drItem, DataRow drNewItem)
        {
            DataTable dtblTable = drItem.Table;
            if (dtblTable == null) { return; }
            if (drNewItem == null) { return; }

            foreach (DataColumn column in dtblTable.Columns)
            {
                DataRow drColumn = drItem[column.ColumnName] as DataRow;
                if (drColumn == null)
                {
                    drNewItem[column.ColumnName] = drItem[column.ColumnName];
                }
                else
                {
                    drColumn.CloneTo((DataRow)drNewItem[column.ColumnName]);
                }
            }
        }

        public static int StrToInt(string intStr)
        {
            try
            {
                int parseInt;
                if (int.TryParse(intStr, out parseInt))
                    return parseInt;
                return 0;
            }
            catch
            {
                return 0; //忽略错误
            }

        }

        public static double StrToDlb(string intStr)
        {
            try
            {
                double parseDlb;
                if (double.TryParse(intStr, out parseDlb))
                    return parseDlb;
                return 0;
            }
            catch
            {
                return 0; //忽略错误
            }

        }

        public static void DoubleBuffered(DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }



        /// <summary>
        /// 从注册表中读取oracle主目录含有tnsnames.ora文件 最新


        /// </summary>
        /// <param name="key">Oracle Key</param>
        /// <param name="dbhome">主目录</param>
        public static void GetOracleHomePath(RegistryKey key, ref string dbhome)
        {
            string installLocation = key.GetValue("inst_loc") + @"\ContentsXML\inventory.xml";
            if (!File.Exists(installLocation))
            {
                //不存在继续判断当前位置是否有oracle_home变量
                string home = (string)key.GetValue("oracle_home");
                string tnsfile = home + @"\network\ADMIN\tnsnames.ora";
                if (File.Exists(tnsfile))
                {
                    dbhome = home;
                    return;
                }
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(installLocation);
            XmlNodeList nodeList = xmlDoc.SelectNodes("/INVENTORY/HOME_LIST/HOME");
            int installIndex = -1;
            foreach (XmlNode node in nodeList)
            {
                if (node.Attributes["TYPE"] != null && node.Attributes["IDX"] != null)
                {
                    if (!node.Attributes["TYPE"].Value.Equals("O")) continue; //指示是否为dbhome
                    int index = -1;
                    if (int.TryParse(node.Attributes["IDX"].Value, out index))
                    {
                        if (installIndex < index)
                        {
                            string home = node.Attributes["LOC"].Value;
                            string tnsfile = home + @"\network\ADMIN\tnsnames.ora";
                            if (File.Exists(tnsfile))
                            {
                                dbhome = home;
                                installIndex = index;
                            }
                        }
                    }


                }

            }
        }


        /// <summary>
        /// 解析tnsname.ora文件
        /// </summary>
        /// <param name="path">oraclehome路径</param>
        /// <returns>TNS Alias(key)-TNS(value)</returns>
        public static Dictionary<string, string> GetDictionaryByOracleHomePath()
        {
            string strFile = "";
            string file = "";
            string dbPath = Environment.GetEnvironmentVariable("TNS_ADMIN");   //首先从环境变量中获取文件
            if (dbPath != "")
            {
                strFile = dbPath + "\\tnsnames.ora";
                if (!File.Exists(strFile))
                {
                    strFile = dbPath + "\\network\\ADMIN\\tnsnames.ora";
                    if (!File.Exists(strFile))
                    {
                        strFile = string.Empty;
                    }
                }
            }

            if (strFile == string.Empty)
            {
                dbPath = Environment.GetEnvironmentVariable("ORACLE_HOME");   //首先从环境变量中获取文件
                if (dbPath != "")
                {
                    strFile = dbPath + "\\tnsnames.ora";
                    if (!File.Exists(strFile))
                    {
                        strFile = dbPath + "\\network\\ADMIN\\tnsnames.ora";
                        if (!File.Exists(strFile))
                        {
                            strFile = string.Empty;
                        }
                    }
                }
            }

            if (strFile != string.Empty)
            {
                file = strFile;
            }
            else
            {
                RegistryKey key = Registry.LocalMachine.OpenSubKey((@"SOFTWARE\Oracle"));
                RegistryKey keyWow = Registry.LocalMachine.OpenSubKey((@"SOFTWARE\WOW6432Node\Oracle"));
                if (key == null && keyWow == null) return null;
                string path = "";
                if (keyWow != null)
                    GetOracleHomePath(keyWow, ref path);
                if (key != null)
                    GetOracleHomePath(key, ref path);
                if (string.IsNullOrEmpty(path)) return null;
                file = path + @"\network\ADMIN\tnsnames.ora";
            }
            if (!File.Exists(file)) return null;
            string output = "";
            string outputDataSource = "";
            string fileLine;
            string tempLine = "";
            Stack parens = new Stack();
            StreamReader sr;
            try
            {
                sr = new StreamReader(file);
            }
            catch (System.IO.FileNotFoundException ex)
            {
                throw ex;
            }
            #region 解析文件
            fileLine = sr.ReadLine();
            while (fileLine != null)
            {
                if (fileLine.Length > 0 && fileLine.Trim().Substring(0, 1) != "#")
                {
                    char lineChar;
                    for (int i = 0; i < fileLine.Length; i++)
                    {
                        lineChar = fileLine[i];


                        if (lineChar == '(')
                        {
                            parens.Push(lineChar);
                            outputDataSource += lineChar;
                        }
                        else if (lineChar == ')')
                        {
                            parens.Pop();
                            outputDataSource += lineChar;
                            if (parens.Count == 0)  //为0的时候代表一个完整的tns alias 读取完成
                            {
                                tempLine += output + outputDataSource + "|";
                                outputDataSource = "";
                                output = "";
                            }
                        }
                        else
                        {
                            if (parens.Count == 0)
                            {
                                output += lineChar;
                            }
                            else
                            {
                                outputDataSource += lineChar;
                            }
                        }
                    }
                }
                fileLine = sr.ReadLine();
            }
            sr.Close();
            tempLine = tempLine.Replace("  ", "");
            #endregion
            #region 解析tns alias成字典


            Dictionary<string, string> dict = new Dictionary<string, string>();
            string[] arrStr = tempLine.Split('|');
            if (arrStr.Length < 1) return null;
            foreach (string str in arrStr)
            {
                if (str.Equals("")) continue;
                string tns = str.Substring(str.IndexOf('('), str.LastIndexOf(')') - str.IndexOf('(') + 1).Trim();
                string tnsAlias = str.Substring(0, str.IndexOf('=')).ToLower().Trim();
                if (dict.ContainsKey(tnsAlias)) continue;
                dict.Add(tnsAlias, tns);
            }
            return dict;
            #endregion
        }

        /// <summary>
        /// 获取tnsname.ora文件中的信息
        /// </summary>
        /// <param name="strName">oracleTNS实例名</param>
        /// <param name="strIP">数据库IP地址</param>
        /// <param name="strPort">数据库端口</param>
        /// <param name="strSID">数据库SID</param>
        /// <returns></returns>
        public static string GetTNSSever(string strTnsName)
        {
            try
            {
                if (dtcTns == null)
                {
                    dtcTns = GetDictionaryByOracleHomePath();
                }
                return dtcTns[strTnsName.ToLower()];
            }
            catch
            {
                return string.Empty;
            }
        }




        [System.Runtime.InteropServices.DllImport("user32.dll ")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int wndproc);
        [System.Runtime.InteropServices.DllImport("user32.dll ")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        public const int GWL_STYLE = -16;
        public const int WS_DISABLED = 0x8000000;

        /// <summary>
        /// 设置控件不可用且不变灰
        /// </summary>
        public static void SetEnabled(Control c, bool enabled)
        {
            if (enabled)
            { SetWindowLong(c.Handle, GWL_STYLE, (~WS_DISABLED) & GetWindowLong(c.Handle, GWL_STYLE)); }
            else
            { SetWindowLong(c.Handle, GWL_STYLE, WS_DISABLED + GetWindowLong(c.Handle, GWL_STYLE)); }
        }

    }
}
