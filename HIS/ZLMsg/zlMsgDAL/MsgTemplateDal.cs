using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using zlMsgModel;

namespace zlMsgDAL
{
    public class MsgTemplateDal
    {

        public List<MsgTemplate> GetTemplateLists()
        {
            string sql = "Zltools.b_Base_Msg.p_Gettemplate";
            OracleParameter[] paras = new OracleParameter[3];
            paras[0] = new OracleParameter("Val", OracleDbType.RefCursor);
            paras[0].Direction = ParameterDirection.Output;
            paras[1] = new OracleParameter("Type_In", OracleDbType.Int16);
            paras[2] = new OracleParameter("Templateid_In", OracleDbType.Int16);

            paras[1].Value = 1; paras[2].Value = 0;

            DataTable t = SqlHelper.ExecuteDataTable(sql, paras);

            List<MsgTemplate> templateList = new List<MsgTemplate>();
            MsgTemplate template;
            MsgProvider provider;
            for (int i = 0; i < t.Rows.Count; i++)
            {
                //初始化一个ProvidermsgTemplate
                provider = new MsgProvider(
                                    (long)t.Rows[i]["Providercode"].ToString().Val(),
                                    t.Rows[i]["Providername"].ToString(),
                                    t.Rows[i]["Appkey"].ToString(),
                                    t.Rows[i]["Appsecret"].ToString()
                                    );

                //初始化模版msgTemplate
                template = new MsgTemplate(
                        (long)t.Rows[i]["Templateid"].ToString().Val(),
                        provider,
                        t.Rows[i]["Templatecode"].ToString(),
                        t.Rows[i]["Signcode"].ToString(),
                        t.Rows[i]["SignNumber"].ToString(),
                        t.Rows[i]["Templatepara"].ToString(),
                        t.Rows[i]["Templatekind"].ToString(),
                        t.Rows[i]["Templatetext"].ToString()
                    );
                templateList.Add(template);
            }
            return templateList;
        }

        /// <summary>
        /// 获取模版信息的主键
        /// </summary>
        /// <returns></returns>
        public long GetTemplateID()
        {
            string sql = "Zltools.b_Base_Msg.p_Gettemplateid";
            OracleParameter[] paras = new OracleParameter[1];
            paras[0] = new OracleParameter("Val", OracleDbType.RefCursor);
            paras[0].Direction = ParameterDirection.Output;

            DataTable dt = SqlHelper.ExecuteDataTable(sql, paras);

            if (dt.Rows.Count > 0) return long.Parse(dt.Rows[0]["Templateid"].ToString());
            return 0;
        }

        public bool AddTemplate(MsgTemplate msgTemplate)
        {
            string sql = "Zltools.b_Base_Msg.p_Templateinsert";
            OracleParameter[] paras = new OracleParameter[8];
            paras[0] = new OracleParameter("Templateid_In", OracleDbType.Int16);
            paras[0].Value = msgTemplate.TemplateID;
            paras[1] = new OracleParameter("Providercode_In", OracleDbType.Int16);
            paras[1].Value = msgTemplate.Provider.ProviderCode;
            paras[2] = new OracleParameter("Templatecode_In", OracleDbType.Varchar2);
            paras[2].Value = msgTemplate.TemplateCode;
            paras[3] = new OracleParameter("Signcode_In", OracleDbType.Varchar2);
            paras[3].Value = msgTemplate.SignCode;
            paras[4] = new OracleParameter("Signnumber_In", OracleDbType.Varchar2);
            paras[4].Value = msgTemplate.SignNumber;
            paras[5] = new OracleParameter("Templatepara_In", OracleDbType.Varchar2);
            paras[5].Value = msgTemplate.TemplatePara;
            paras[6] = new OracleParameter("Templatekind_In", OracleDbType.Int16);
            paras[6].Value = msgTemplate.TemplateKind;
            paras[7] = new OracleParameter("Templatetext_In", OracleDbType.Varchar2);
            paras[7].Value = msgTemplate.TemplateText;
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

        public bool UpdateTemplate(MsgTemplate msgTemplate)
        {
            string sql = "Zltools.b_Base_Msg.p_Templateupdate";
            OracleParameter[] paras = new OracleParameter[8];
            paras[0] = new OracleParameter("Templateid_In", OracleDbType.Int16);
            paras[0].Value = msgTemplate.TemplateID;
            paras[1] = new OracleParameter("Providercode_In", OracleDbType.Int16);
            paras[1].Value = msgTemplate.Provider.ProviderCode;
            paras[2] = new OracleParameter("Templatecode_In", OracleDbType.Varchar2);
            paras[2].Value = msgTemplate.TemplateCode;
            paras[3] = new OracleParameter("Signcode_In", OracleDbType.Varchar2);
            paras[3].Value = msgTemplate.SignCode;
            paras[4] = new OracleParameter("Signnumber_In", OracleDbType.Varchar2);
            paras[4].Value = msgTemplate.SignNumber;
            paras[5] = new OracleParameter("Templatepara_In", OracleDbType.Varchar2);
            paras[5].Value = msgTemplate.TemplatePara;
            paras[6] = new OracleParameter("Templatekind_In", OracleDbType.Varchar2);
            paras[6].Value = msgTemplate.TemplateKind;
            paras[7] = new OracleParameter("Templatetext_In", OracleDbType.Varchar2);
            paras[7].Value = msgTemplate.TemplateText;
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

        public bool DeleteTemplate(List<long> TemplateIDs)
        {
            string sql = "Zltools.b_Base_Msg.p_Templatedelete";
            OracleParameter[] paras = new OracleParameter[1];
            paras[0] = new OracleParameter("Templateid_In", OracleDbType.Int16);
            
            try
            {
                SqlHelper.BeginTrans();
                foreach (long id in TemplateIDs)
                {
                    paras[0].Value = id;
                    SqlHelper.ExecteNonQuery(sql, paras);
                }
                SqlHelper.CommitTrans();                
                return true;
            }
            catch(Exception exception)
            {
                SqlHelper.RollBackTrans();
                return false;
                throw exception;
            }
        }
    }
}
