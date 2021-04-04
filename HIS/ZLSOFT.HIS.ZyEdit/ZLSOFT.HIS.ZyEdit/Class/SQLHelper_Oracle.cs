using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.Configuration;
using System.Reflection;
using System.IO;
using System.Collections;
using ZLSOFT.HIS.ZyEdit.Model;

namespace ZLSOFT.HIS.ZyEdit
{
    class SQLHelper_Oracle
    {
        /// <summary>
        /// Oracle数据库帮助类
        /// </summary>
        static string m_Host;
        static string m_Port;
        static string m_OracleSID;
        static string m_Server;//服务地址，格式：Host:Port/OracleSID

        static string m_userName;
        static string m_password;
        static string m_TnsName;
        static bool mblnTrans = false;
        public static OracleConnection OracleCnn { set; get; }//当前程序Oracle连接

        /// <summary>
        /// TNS初始化Oralce连接信息
        /// </summary>
        /// <param name="userName">用户名</param>
        /// <param name="password">密码</param>
        /// <param name="TnsName">TNS服务名</param>
        public static void IntData(string userName, string password, string TnsName)
        {
            m_userName = userName;
            m_password = password;
            m_TnsName = TnsName;
            if (m_TnsName != "")
            {
                m_Server = ComLib.GetTNSSever(m_TnsName);
            }
            mblnTrans = true;
        }

        /// <summary>
        /// 实例名初始化Oralce连接信息
        /// </summary>
        /// <param name="userName">用户名</param>
        /// <param name="password">密码</param>
        /// <param name="Host">IP地址</param>
        /// <param name="Port">端口</param>
        /// <param name="OracleSID">实例名</param>
        public static void IntData(string userName, string password, string Host, string Port, string OracleSID)
        {
            m_userName = userName;
            m_password = password;
            m_Host = Host;
            m_Port = Port;
            m_OracleSID = OracleSID;
            m_Server = string.Format("{0}:{1}/{2}", m_Host, m_Port, m_OracleSID);
            mblnTrans = false;
        }



        /// <summary>
        /// 执行Oracle过程，返回DataTable
        /// </summary>
        /// <param name="userName">用户名</param>
        /// <param name="password">密码</param>
        /// <param name="procedureName">过程名</param>
        /// <param name="tranPasswd">是否根据ZLHIS密码规则转换密码</param>
        /// <param name="paraValue">入参Input_In得值，JSON字符串</param>
        /// <returns>返回DataTable</returns>
        public DataTable ExcuteProcedureQueryData(string userName, string password,
                string procedureName, string paraValue, bool tranPasswd)
        {
            try
            {
                DataTable dtbl = ExcuteProcedureQueryData(OracleCnn, procedureName, paraValue);
                return dtbl;
            }
            catch (OracleException ex)
            {
                if (ex.Number == 1017) //ORA-01017: invalid username/password; logon denied
                {
                    throw new Exception("用户名或密码错误！");
                }
                else
                {
                    throw;
                }
            }
        }

        //传入ODP连接执行Oracle过程
        private DataTable ExcuteProcedureQueryData(OracleConnection conn, string methodName, string paraValue)
        {

            using (OracleCommand command = new OracleCommand(methodName, conn))
            {
                command.CommandType = CommandType.StoredProcedure;
                OracleParameter[] parameters = new OracleParameter[2];
                parameters[0] = new OracleParameter("Input_In", OracleDbType.Clob);
                parameters[0].Value = paraValue;
                parameters[1] = new OracleParameter("Output_Out", OracleDbType.RefCursor);
                parameters[1].Direction = ParameterDirection.Output;
                command.Parameters.AddRange(parameters);

                using (OracleDataAdapter dataAdapter = new OracleDataAdapter(command))
                {
                    DataTable dtbl = new DataTable();
                    dataAdapter.Fill(dtbl);
                    return dtbl;
                }
            }
        }


        /// <summary>
        /// 执行Oracle服务，返回Json
        /// </summary>
        /// <param name="userName">用户名</param>
        /// <param name="password">密码</param>
        /// <param name="procedureName">过程名</param>
        /// <param name="tranPasswd">是否根据ZLHIS密码规则转换密码</param>
        /// <param name="paraValue">入参Input_In得值，JSON字符串</param>
        /// <returns>返回DataTable</returns>
        public string ExcuteProcedureJson(string userName, string password,
                string procedureName, string paraValue, bool tranPasswd)
        {
            try
            {
                    string dtbl = ExcuteProcedureJson(OracleCnn, procedureName, paraValue);
                    return dtbl;
            }
            catch (OracleException ex)
            {
                if (ex.Number == 1017) //ORA-01017: invalid username/password; logon denied
                {
                    throw new Exception("用户名或密码错误！");
                }
                else
                {
                    throw;
                }
            }
        }

        //传入ODP连接执行Oracle过程
        private string  ExcuteProcedureJson(OracleConnection conn, string methodName, string paraValue)
        {
            string strOut = "";
            using (OracleCommand command = new OracleCommand(methodName, conn))
            {
                command.CommandType = CommandType.StoredProcedure;
                OracleParameter[] parameters = new OracleParameter[2];
                parameters[0] = new OracleParameter("Json_In", OracleDbType.Clob);
                parameters[0].Value = paraValue;
                parameters[1] = new OracleParameter("Json_Out", OracleDbType.Char);
                parameters[1].Direction = ParameterDirection.Output;
                parameters[1].Size = 9999;
                command.Parameters.AddRange(parameters);
                command.ExecuteReader();

                strOut = System.Convert.ToString(command.Parameters[1].Value);
                return strOut;
            }
        }



        //获取ODP数据库连接
        public static OracleConnection GetOdpConnection()
        {
            if (mblnTrans)
            {
                return ZLSoft.HIS.Library.Helper.GetOdpConnection(m_Server, m_password, m_userName, mblnTrans);
            }
            else
            {
                string connStr = string.Format("Data Source=(" +
                                                    "DESCRIPTION=" +
                                                    " (ADDRESS=(PROTOCOL=TCP)(HOST={0})(PORT={1}))" +
                                                    " (CONNECT_DATA=(SERVICE_NAME={2}))" +
                                                    ");" +
                                                    "User Id={3};" +
                                                    "Password={4};",
                                                    m_Host, m_Port, m_OracleSID, m_userName, m_password);
                return new OracleConnection(connStr);
            }
        }


        public static string[] GetDatabases()
        {
            string output = "";
            string fileLine;
            Stack parens = new Stack();

            //	open tnsnames.ora 
            StreamReader sr;
            try
            {
                sr = new StreamReader(@"C:\oracle\ora92\network\ADMIN\TNSNAMES.ORA");
            }
            catch (System.IO.FileNotFoundException ex)
            {
                throw ex;
            }
            //	Read the first line of the file 
            fileLine = sr.ReadLine();
            //	loop through, reading each line of the file 
            while (fileLine != null)
            {
                //	if the first non whitespace character is a #, ignore the line 
                //	and go to the next line in the file 
                if (fileLine.Length > 0 && fileLine.Trim().Substring(0, 1) != "#")
                {
                    //	Read through the input line character by character 
                    char lineChar;
                    for (int i = 0; i < fileLine.Length; i++)
                    {
                        lineChar = fileLine[i];
                        if (lineChar == '(')
                        {
                            //	if the char is a ( push it onto the stack 
                            parens.Push(lineChar);
                        }
                        else if (lineChar == ')')
                        {
                            //	if the char is a ), pop the stack 
                            parens.Pop();
                        }
                        else
                        {
                            if (parens.Count == 0)
                            {
                                output += lineChar;
                            }
                        }
                    }
                }
                //	Read the next line of the file 
                fileLine = sr.ReadLine();
            }

            //	Close the stream reader 
            sr.Close();

            //	Split the output string into a string[] 
            string[] split = output.Split('=');
            //	trim each string in the array 
            for (int i = 0; i < split.Length; i++)
            {
                split[i] = split[i].Trim();
            }
            Array.Sort(split);
            return split;
        }
    }
}

