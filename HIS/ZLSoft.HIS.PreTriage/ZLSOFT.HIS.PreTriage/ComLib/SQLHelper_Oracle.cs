using System;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.IO;
using System.Collections;
using ZLSOFT.HIS.PreTriage.Models;

namespace ZLSOFT.HIS.PreTriage.ComLib
{
    public static class SQLHelper_Oracle
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
        /// <param name="procedureName">过程名</param>
        /// <param name="paraValue">入参Input_In得值，JSON字符串</param>
        /// <returns>返回DataTable</returns>
        public static DataTable GetDataTable(
                string procedureName, string paraValue)
        {
            paraValue = @"{ ""input"":" + paraValue + "}";

            ComLib.SaveErrorlog(Environment.NewLine + "接口：" + procedureName + Environment.NewLine + "参数：" + paraValue);
            try
            {
                DataTable dtbl = ExcuteProcedureQueryData(BaseData.OracleCnn, procedureName, paraValue);
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

        /// <summary>
        /// 执行Oracle过程，返回String
        /// </summary>
        /// <param name="procedureName">过程名</param>
        /// <param name="paraValue">入参Input_In得值，JSON字符串</param>
        /// <returns>返回DataTable</returns>
        public static string GetDataStr(
                string procedureName, string paraValue)
        {
            paraValue = @"{ ""input"":" + paraValue + "}";

            ComLib.SaveErrorlog(Environment.NewLine + "接口：" + procedureName + Environment.NewLine + "参数：" + paraValue);
            try
            {
                string strOut = ExcuteOutStr(BaseData.OracleCnn, procedureName, paraValue);
                return strOut.Trim();
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
        static DataTable ExcuteProcedureQueryData(OracleConnection conn, string methodName, string paraValue)
        {
            try
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
        static string ExcuteOutStr(OracleConnection conn, string methodName, string paraValue)
        {
            try
            {
                string strOut = "";
                
                OracleCommand command = conn.CreateCommand();// new OracleCommand(methodName, conn);
                command.CommandText = methodName;
                command.CommandType = CommandType.StoredProcedure;
                OracleParameter[] parameters = new OracleParameter[2];
                parameters[0] = new OracleParameter("Json_In", OracleDbType.Clob);
                parameters[0].Value = paraValue;
                parameters[1] = new OracleParameter("Json_Out", OracleDbType.Char);
                parameters[1].Direction = ParameterDirection.Output;
                parameters[1].Size = 9999;
                command.Parameters.AddRange(parameters);
                if (conn.State == ConnectionState.Closed) conn.Open();
                command.ExecuteReader();

                strOut = command.Parameters[1].Value == null ? "" : command.Parameters[1].Value.ToString() == "null" ? "" : System.Convert.ToString( command.Parameters[1].Value);
                return strOut;

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

    }
}
