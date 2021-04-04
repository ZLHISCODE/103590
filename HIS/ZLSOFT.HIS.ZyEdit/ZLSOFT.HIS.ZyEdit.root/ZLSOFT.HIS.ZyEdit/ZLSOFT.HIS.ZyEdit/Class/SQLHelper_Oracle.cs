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

namespace ZLSOFT.HIS.ZyEdit
{
    class SQLHelper_Oracle
    {
        /// <summary>
        /// Oracle数据库帮助类
        /// </summary>
        string m_Host;
        string m_Port;
        string m_OracleSID;
        string m_Server;//服务地址，格式：Host:Port/OracleSID

        public SQLHelper_Oracle()
        {
            if (ComLib.strTnsName != "") {
                m_Server= ComLib.GetTNSSever();
            }
            else
            {
                m_Host = ComLib.GetAppSettingValue("HOST", "").ToString();
                m_Port = ComLib.GetAppSettingValue("PORT", "").ToString();
                m_OracleSID = ComLib.GetAppSettingValue("ORACLE_SID", "").ToString();
                m_Server = string.Format("{0}:{1}/{2}", m_Host, m_Port, m_OracleSID);
            }
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
                using (OracleConnection conn = GetOdpConnection(userName, password, tranPasswd))
                {
                    DataTable dtbl = ExcuteProcedureQueryData(conn, procedureName, paraValue);
                    return dtbl;
                }
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

        //获取ODP数据库连接
        private OracleConnection GetOdpConnection(string userName, string password, bool tranPasswd)
        {
            if (tranPasswd)
            {
                return ZLSoft.HIS.Library.Helper.GetOdpConnection(m_Server, password, userName, true);
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
                                                    m_Host, m_Port, m_OracleSID, userName, password);
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

