using Oracle.ManagedDataAccess.Client;
using System;
using System.ComponentModel;
using System.Data;

namespace zlMsgDAL
{
    public  static class SqlHelper
    {
        private static OracleConnection connObject;
        private static OracleTransaction transaction;
        private static OracleDependency dependency;

        public static void BeginTrans()
        {
            transaction = connObject.BeginTransaction(IsolationLevel.ReadCommitted);
        }

        public static void CommitTrans()
        {
            transaction.Commit();
        }

        public static void RollBackTrans()
        {
            transaction.Rollback();
        }

        public static void SetConnection(OracleConnection oracleConnection)
        {
            connObject = oracleConnection;
        }

        /// <summary>
        /// 执行带有绑定变量的SQL过程,并返回一个记录集
        /// </summary>
        public static DataTable ExecuteDataTable(string query, OracleParameter[] parameters)
        {
            try
            {
                if (connObject.State == ConnectionState.Broken || connObject.State == ConnectionState.Closed)
                {
                    connObject.Open();
                }

                DataTable dt = new DataTable();
                using (OracleCommand cmd = new OracleCommand(query, connObject))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.AddToStatementCache = true;
                    cmd.BindByName = true;

                    if (parameters != null) cmd.Parameters.AddRange(parameters);
                    using (OracleDataAdapter adapter = new OracleDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        ///执行一个不需要返回值过程
        /// </summary>
        /// <param name="cmdText">存储过程的名字或者 T-SQL 语句</param>
        /// <param name="parameters">以数组形式提供SqlCommand命令中用到的参数列表</param>
        public static void ExecteNonQuery(string cmdText, OracleParameter[] parameters)
        {
            try
            {
                if (connObject.State == ConnectionState.Broken || connObject.State == ConnectionState.Closed)
                {
                    connObject.Open();
                }
                using (OracleCommand cmd = new OracleCommand(cmdText, connObject))
                {
                    cmd.BindByName = true;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = cmdText;
                    if (parameters != null) cmd.Parameters.AddRange(parameters);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        public static bool RegistDcn(string strSql,OnChangeEventHandler changeEventHandler)
        {
            using (OracleCommand cmd = new OracleCommand(strSql, connObject))
            {
                cmd.AddRowid = true;
                dependency = new OracleDependency(cmd);
                cmd.Notification.IsNotifiedOnce = false;
                dependency.RowidInfo = OracleRowidInfo.Include;
                dependency.OnChange += changeEventHandler;
                cmd.ExecuteNonQuery();
            }
            return true;
        }

        public static bool UnRegistDcn()
        {
            dependency.RemoveRegistration(connObject);
            return true;
        }


    }
}
