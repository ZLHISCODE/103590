using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Text;

namespace zlMsgBLL
{
    public static class OracleConnect
    {
        public static Boolean GetConn(string strName, string strPwd, string strServerName,string strRole)
        {
            try
            {
                OracleConnectionStringBuilder conStrB = new OracleConnectionStringBuilder
                {
                    DataSource = strServerName,
                    UserID = strName,
                    Password = strPwd
                };
                if (!strRole.Equals("Normal"))
                {
                    conStrB.DBAPrivilege = strRole;
                }

                OracleConnection connObject = new OracleConnection(conStrB.ConnectionString);
                connObject.Open();
                if (connObject.State == System.Data.ConnectionState.Open)
                {
                    zlMsgDAL.SqlHelper.SetConnection(connObject);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
