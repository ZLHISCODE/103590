using Oracle.ManagedDataAccess.Client;
using System;
using System.Data;
using zlMsgModel;

namespace zlMsgDAL
{
    public class MsgProviderDal
    {
        /// <summary>
        /// 通过编码或名称获取provider
        /// </summary>
        /// <param name="ProviderCode">编码</param>
        /// <param name="ProviderName">名称</param>
        /// <returns></returns>
        public MsgProvider GetMsgProvider(long ProviderCode =0 ,string ProviderName = "")
        {
            string sql = "Zltools.b_Base_Msg.p_Getprovider";

            //初始化参数
            OracleParameter[] paras = new OracleParameter[3];
            paras[0] = new OracleParameter("Val", OracleDbType.RefCursor);
            paras[0].Direction = ParameterDirection.Output;
            paras[1] = new OracleParameter("Providercode_In", OracleDbType.Int16);
            paras[1].Value = ProviderCode;
            paras[2] = new OracleParameter("Providername_In", OracleDbType.Varchar2);
            paras[2].Value = ProviderName;


            DataTable dt  = SqlHelper.ExecuteDataTable(sql,paras);
            if(dt.Rows.Count>0)
            {
                MsgProvider msg = new MsgProvider((long)dt.Rows[0]["Providercode"].ToString().Val(),
                                                                        dt.Rows[0]["Providername"].ToString(),
                                                                        dt.Rows[0]["Appkey"].ToString(),
                                                                        dt.Rows[0]["Appsecret"].ToString());
                return msg;
            }

            return null;
        }

        /// <summary>
        /// 获取平台信息的主键
        /// </summary>
        /// <returns></returns>
        public long GetProviderCode()
        {
            string sql = "Zltools.b_Base_Msg.p_Getprovidercode";
            OracleParameter[] paras = new OracleParameter[1];
            paras[0] = new OracleParameter("Val", OracleDbType.RefCursor);
            paras[0].Direction = ParameterDirection.Output;

            DataTable dt = SqlHelper.ExecuteDataTable(sql, paras);

            if (dt.Rows.Count > 0) return long.Parse(dt.Rows[0]["Providercode"].ToString());
            return 0;
        }

        /// <summary>
        /// 向数据库中添加Provider
        /// </summary>
        /// <param name="msgProvider"></param>
        /// <returns></returns>
        public bool AddProvider(MsgProvider msgProvider)
        {
            string sql = "Zltools.b_Base_Msg.p_Providerinsert";
            //初始化参数
            OracleParameter[] paras = new OracleParameter[4];
            paras[0] = new OracleParameter("Providercode_In", OracleDbType.Int16);
            paras[1] = new OracleParameter("Providername_In", OracleDbType.Varchar2);
            paras[2] = new OracleParameter("Appkey_In", OracleDbType.Varchar2);
            paras[3] = new OracleParameter("Appsecret_In", OracleDbType.Varchar2);

            paras[0].Value = msgProvider.ProviderCode;
            paras[1].Value = msgProvider.PrividerName;
            paras[2].Value = msgProvider.AppKey;
            paras[3].Value = msgProvider.AppSecret;

            try
            {
                SqlHelper.ExecteNonQuery(sql, paras);
                return true;
            }
            catch
            {
                return false;
            }
           
        }

        /// <summary>
        /// 修改数据库中的Provider
        /// </summary>
        /// <param name="msgProvider"></param>
        /// <returns></returns>
        public bool UpdateProvider(MsgProvider msgProvider)
        {

            string sql = "Zltools.b_Base_Msg.p_Providerupdate";
            //初始化参数
            OracleParameter[] paras = new OracleParameter[4];
            paras[0] = new OracleParameter("Providercode_In", OracleDbType.Int16);
            paras[1] = new OracleParameter("Providername_In", OracleDbType.Varchar2);
            paras[2] = new OracleParameter("Appkey_In", OracleDbType.Varchar2);
            paras[3] = new OracleParameter("Appsecret_In", OracleDbType.Varchar2);

            paras[0].Value = msgProvider.ProviderCode;
            paras[1].Value = msgProvider.PrividerName;
            paras[2].Value = msgProvider.AppKey;
            paras[3].Value = msgProvider.AppSecret;

            try
            {
                SqlHelper.ExecteNonQuery(sql, paras);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 删除数据库中的Provider
        /// </summary>
        /// <param name="msgProvider"></param>
        /// <returns></returns>
        public bool DeleteProvider(MsgProvider msgProvider)
        {
            string sql = "Zltools.b_Base_Msg.p_Providerdelete";
            OracleParameter[] paras = new OracleParameter[1];
            paras[0] = new OracleParameter("Providercode_In", OracleDbType.Int16);
            paras[0].Value = msgProvider.ProviderCode;

            try
            {
                SqlHelper.ExecteNonQuery(sql, paras);
                return true;
            }
            catch
            {
                return false;
            }
        }

    }
}
