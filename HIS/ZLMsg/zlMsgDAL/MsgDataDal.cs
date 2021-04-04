using System;
using System.Collections.Generic;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using zlMsgModel;

namespace zlMsgDAL
{
    public class MsgDataDal
    {
        private MsgData InitDataFromTbl(DataTable t,int intIndex)
        {
            MsgTemplate template;
            MsgProvider provider;
            MsgData msgData = null;

            //初始化一个ProvidermsgTemplate
            provider = new MsgProvider(
                (long)t.Rows[intIndex]["Providercode"].ToString().Val(),
                t.Rows[intIndex]["Providername"].ToString(),
                t.Rows[intIndex]["Appkey"].ToString(),
                t.Rows[intIndex]["Appsecret"].ToString()
            );

            //初始化模版msgTemplate
            template = new MsgTemplate(
                (long)t.Rows[intIndex]["Templateid"].ToString().Val(),
                provider,
                t.Rows[intIndex]["Templatecode"].ToString(),
                t.Rows[intIndex]["Signcode"].ToString(),
                t.Rows[intIndex]["SignNumber"].ToString(),
                t.Rows[intIndex]["Templatepara"].ToString(),
                t.Rows[intIndex]["Templatekind"].ToString(),
                t.Rows[intIndex]["Templatetext"].ToString()
            );

            //初始化msgData
            msgData = new MsgData(long.Parse(t.Rows[intIndex]["Id"].ToString()),
                template,
                (long)t.Rows[intIndex]["Receiver"].ToString().Val(),
                t.Rows[intIndex]["Para"].ToString(),
                t.Rows[intIndex]["Sender"].ToString(),
                t.Rows[intIndex]["Terminal"].ToString(),
                (int)t.Rows[intIndex]["State"].ToString().Val(),
                (int)t.Rows[intIndex]["Kind"].ToString().Val(),
                t.Rows[intIndex]["Msgtext"].ToString(),
                t.Rows[intIndex]["Extend"].ToString(),
                t.Rows[intIndex]["Senddate"].ToString() == "" ? DateTime.Now : Convert.ToDateTime(t.Rows[intIndex]["Senddate"].ToString())
            );
            return msgData;
        }


        public List<MsgData> GetErrorMsg()
        {
            string sql;
            OracleParameter[] paras;
            sql = "Zltools.b_Base_Msg.p_Geterrmsgdata";
            paras = new OracleParameter[1];
            paras[0] = new OracleParameter("Val", OracleDbType.RefCursor);
            paras[0].Direction = ParameterDirection.Output;

            try
            {
                DataTable t = SqlHelper.ExecuteDataTable(sql, paras);
                List<MsgData> lists = new List<MsgData>();
                for(int i=0;i<t.Rows.Count;i++)
                {
                    lists.Add(InitDataFromTbl(t, i)); 
                }
                return lists;
            }
            catch (Exception exception)
            {
                return null;
                throw exception;
            }
        }

        public MsgData GetShortMsg(string strRowID)
        {
            string sql;
            OracleParameter[] paras;
            sql = "Zltools.b_Base_Msg.p_Getmsgdatabyrowid";
            paras = new OracleParameter[2];
            paras[0] = new OracleParameter("Val", OracleDbType.RefCursor);
            paras[0].Direction = ParameterDirection.Output;
            paras[1] = new OracleParameter("Rowid_In", OracleDbType.Varchar2);
            paras[1].Value = strRowID;

            try
            {
                DataTable t = SqlHelper.ExecuteDataTable(sql, paras);
                MsgData msgData = null;
                if (t.Rows.Count > 0)
                {
                    msgData = InitDataFromTbl(t, 0);
                }
                return msgData;
            }
            catch(Exception exception)
            {
                return null;
                throw exception;
            }
        }

        public bool AddShortMsg(MsgData shortMsg)
        {
            string sql = "Zltools.b_Base_Msg.p_Msgdatainsert";
            OracleParameter[] paras = new OracleParameter[10];
            paras[0] = new OracleParameter("Templateid_In", OracleDbType.Int16);
            paras[0].Value = shortMsg.Template.TemplateID;
            paras[1] = new OracleParameter("Receiver_In", OracleDbType.Int16);
            paras[1].Value = shortMsg.Receiver;
            paras[2] = new OracleParameter("Para_In", OracleDbType.Varchar2);
            paras[2].Value = shortMsg.Para;
            paras[3] = new OracleParameter("Sender_In", OracleDbType.Varchar2);
            paras[3].Value = shortMsg.Sender;
            paras[4] = new OracleParameter("Terminal_In", OracleDbType.Varchar2);
            paras[4].Value = shortMsg.Terminal;
            paras[5] = new OracleParameter("State_In", OracleDbType.Int16);
            paras[5].Value = shortMsg.State;
            paras[6] = new OracleParameter("Kind_In", OracleDbType.Int16);
            paras[6].Value = shortMsg.Kind;
            paras[7] = new OracleParameter("Msgtext_In", OracleDbType.Varchar2);
            paras[7].Value = shortMsg.MsgText;
            paras[8] = new OracleParameter("Senddate_In", OracleDbType.Date);
            paras[8].Value = shortMsg.SendDate;
            paras[9] = new OracleParameter("Extend_In", OracleDbType.Varchar2);
            paras[9].Value = shortMsg.Extend;

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

        public bool UpdateShortMsg(MsgData shortMsg)
        {
            string sql = "Zltools.b_Base_Msg.p_Msgdataupdate";
            OracleParameter[] paras = new OracleParameter[11];
            paras[0] = new OracleParameter("Id_In", OracleDbType.Int16);
            paras[0].Value = shortMsg.ID;
            paras[1] = new OracleParameter("Templateid_In", OracleDbType.Int16);
            paras[1].Value = shortMsg.Template.TemplateID;
            paras[2] = new OracleParameter("Receiver_In", OracleDbType.Int16);
            paras[2].Value = shortMsg.Receiver;
            paras[3] = new OracleParameter("Para_In", OracleDbType.Varchar2);
            paras[3].Value = shortMsg.Para;
            paras[4] = new OracleParameter("Sender_In", OracleDbType.Varchar2);
            paras[4].Value = shortMsg.Sender;
            paras[5] = new OracleParameter("Terminal_In", OracleDbType.Varchar2);
            paras[5].Value = shortMsg.Terminal;
            paras[6] = new OracleParameter("State_In", OracleDbType.Int16);
            paras[6].Value = shortMsg.State;
            paras[7] = new OracleParameter("Kind_In", OracleDbType.Int16);
            paras[7].Value = shortMsg.Kind;
            paras[8] = new OracleParameter("Msgtext_In", OracleDbType.Varchar2);
            paras[8].Value = shortMsg.MsgText;
            paras[9] = new OracleParameter("Senddate_In", OracleDbType.Date);
            paras[9].Value = shortMsg.SendDate;
            paras[10] = new OracleParameter("Extend_In", OracleDbType.Varchar2);
            paras[10].Value = shortMsg.Extend;

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

        public bool DeleteShortMsg(List<long> MsgIDs)
        {
            string sql = "Zltools.b_Base_Msg.p_Msgdatadelete";
            OracleParameter[] paras = new OracleParameter[1];
            paras[0] = new OracleParameter("Id_In", OracleDbType.Int16);

            try
            {
                SqlHelper.BeginTrans();
                foreach (long id in MsgIDs)
                {
                    paras[0].Value = id;
                    SqlHelper.ExecteNonQuery(sql, paras);
                }
                SqlHelper.CommitTrans();
                return true;
            }
            catch (Exception exception)
            {
                SqlHelper.RollBackTrans();
                return false;
                throw exception;
            }
        }

        public void ShortMsgRegist(OnChangeEventHandler changeEventHandler)
        {
            string strSql = "Select * from zltools.ZlMsgData";
            SqlHelper.RegistDcn(strSql, changeEventHandler);
        }

        public void ShortMsgUnRegist()
        {
            SqlHelper.UnRegistDcn();
        }

    }
}
