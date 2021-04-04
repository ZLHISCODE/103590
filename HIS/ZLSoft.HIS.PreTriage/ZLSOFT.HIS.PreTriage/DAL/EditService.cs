using System;
using ZLSOFT.HIS.PreTriage.ComLib;
using System.Data;
using Newtonsoft.Json;

namespace ZLSOFT.HIS.PreTriage.DAL
{
    class EditService
    {
        public DataTable Get病人来源()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Patifrom", "");
        }
        public DataTable Get意识状态()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Patistate", "");
        }
        public DataTable Get陪同人员()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Entourage", "");
        }
        public DataTable Get常见既往史()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Dishistory", "");
        }

        public DataTable Get评分指标()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Scorebase", "");
        }

        public DataTable Get性别()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Sexbase", "");
        }

        public DataTable Get民族()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Nationbase", "");
        }

        public DataTable Get国籍()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Nationality", "");
        }

        public DataTable Get急诊科室()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Dept", "");
        }

        public DataTable Get急诊等级()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Level", "");
        }

        public DataTable Get常用主诉()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Paticc", "");
        }

        public DataTable Get人工评估规则()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Rules", "");
        }


        public string Get儿童年龄上限()
        {
            return SQLHelper_Oracle.GetDataStr("Pkg_Pretriage_Dql.Get_Childmaxage", "");
        }

        public DataTable Get病人列表(string dtbegin, string dtend, string strstate, string strtimeout)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                begin = dtbegin,
                end = dtend,
                state = strstate,
                timeout = strtimeout
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Patlist", strJson);
        }

        public DataTable Get病人分诊信息(string strid)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                id = strid
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Patidetail", strJson);
        }

        /// <summary>
        /// 根据出生日期返回年龄
        /// </summary>
        /// <returns>年龄字符串</returns>
        public string GetDateToAge(DateTime dtBirthDay)
        {
            if (dtBirthDay == null || dtBirthDay <= DateTime.MinValue)
            {
                return "";
            }
            string str出生日期 = JsonConvert.SerializeObject(new
            {
                birthday = dtBirthDay.ToString("yyyy-MM-dd HH:mm")
            });
            return SQLHelper_Oracle.GetDataStr("Pkg_Pretriage_Dql.Get_Datetoage", str出生日期);
        }


        /// <summary>
        /// 根据疼痛选择返回病人等级和分数

        /// </summary>
        public DataTable Get疼痛评分(string str疼痛)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                pain = str疼痛
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Pain_Num_Rating", strJson);
        }


        /// <summary>
        /// 根据昏迷评分分级方法
        /// </summary>
        public DataTable Get昏迷评分(string str睁眼反应, string str语言反应, string str活动反应)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                open_reaction = str睁眼反应,
                language_reaction = str语言反应,
                activity_reaction = str活动反应
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Coma_Rating", strJson);
        }

        /// <summary>
        /// 客观评价分级方法成人和儿童方法

        /// </summary>
        public DataTable Get客观评分(string str年龄, string str年龄单位, string str指标信息)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                agenum = str年龄,
                ageunit = str年龄单位,
                indexid_describe = str指标信息
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Objective_Rating", strJson);
        }

        /// <summary>
        /// 获取当前系统时间
        /// </summary>
        /// <returns></returns>
        public DateTime GetNowTime()
        {
            string strTime = SQLHelper_Oracle.GetDataStr("Pkg_Pretriage_Dql.Get_Now_Time", "");
            return Convert.ToDateTime(strTime);
        }

        /// <summary>
        /// 保存病人分诊信息
        /// </summary>
        /// <returns></returns>
        public DataTable SavePretriage(string strJson)
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dml.Save_Pretriage", strJson);
        }

        /// <summary>
        /// 读取病人分诊信息
        /// </summary>
        /// <returns></returns>
        public DataTable LoadPretriage(string str分诊ID)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                id = str分诊ID
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Load_Pretriage", strJson);
        }


        /// <summary>
        /// 删除病人就诊信息
        /// </summary>
        /// <returns></returns>
        public string DelPretriage(string str就诊ID)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                id = str就诊ID
            });
            return SQLHelper_Oracle.GetDataStr("Pkg_Pretriage_Dml.Del_Pretriage", strJson);
        }

        /// <summary>
        /// 获取病人分诊评分信息
        /// </summary>
        /// <returns></returns>
        public DataTable Load评分信息(string str分诊ID)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                id = str分诊ID
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Load_Levelinfo", strJson);
        }

        /// <summary>
        /// 获取病人分诊指标信息
        /// </summary>
        /// <returns></returns>
        public DataTable Load指标信息(string str分诊ID)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                id = str分诊ID
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Load_Rulesinfo", strJson);
        }

        /// <summary>
        /// 更新最新的挂号安排
        /// </summary>
        /// <returns></returns>
        public void Update挂号安排()
        {
            SQLHelper_Oracle.GetDataStr("Pkg_Pretriage_Dml.Register_Update", "");
            return;
        }

        /// <summary>
        /// 清除挂号事务锁定
        /// </summary>
        /// <returns></returns>
        public void RegisterUnlock()
        {
            SQLHelper_Oracle.GetDataStr("Pkg_Pretriage_Dml.Register_Unlock", "");
            return;
        }

        /// <summary>
        /// 获取最新的就诊状态

        /// </summary>
        /// <returns></returns>
        public DataTable Get就诊状态(string str分诊ID)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                id = str分诊ID
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Getvisitstate", strJson);
        }

        public bool Check就诊状态(string str分诊ID)
        {
            if (str分诊ID == "") return false;
            DataTable data就诊状态 = Get就诊状态(str分诊ID);
            if (data就诊状态 == null) return false;
            if (data就诊状态.Rows.Count == 0) return false;

            if (data就诊状态.Rows[0]["就诊状态"].ToString() != "0")
            {
                return false;
            }
            else
            {
                return true;
            }
        }


        /// <summary>
        /// 变更病人就诊记录的绿色通道状态

        /// </summary>
        /// <returns></returns>
        public string ChangeGreenchannel(string str就诊ID, string str状态)
        {
            string strJson = JsonConvert.SerializeObject(new
            {
                id = str就诊ID,
                是否绿色通道 = str状态

            });
            return SQLHelper_Oracle.GetDataStr("Pkg_Pretriage_Dml.Change_Greenchannel", strJson);
        }


        /// <summary>
        /// 通过输入姓名匹配病人信息列表清单
        /// </summary>
        public DataTable GetPatlistbyName(string str姓名)
        {
            if (str姓名 == "") return null;
            string strJson = JsonConvert.SerializeObject(new
            {
                姓名输入 = str姓名
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_PatlistbyName", strJson);
        }


        /// <summary>
        /// 通过Zl_Age_Check检查年龄录入是否正确

        /// </summary>
        /// <returns>年龄字符串</returns>
        public string CheckAge(string strAge)
        {
            if (strAge == "") return "";
            string str年龄 = JsonConvert.SerializeObject(new
            {
                age = strAge
            });
            return SQLHelper_Oracle.GetDataStr("Pkg_Pretriage_Dql.Checkage", str年龄);
        }

        /// <summary>
        /// 通过Zl_Fun_Checkidcard检查身份证录入是否正确
        /// </summary>
        /// <returns>年龄字符串</returns>
        public string CheckIDCard(string strID)
        {
            if (strID == "") return "";
            string str身份证 = JsonConvert.SerializeObject(new
            {
                idcard = strID
            });
            return SQLHelper_Oracle.GetDataStr("Pkg_Pretriage_Dql.Checkidcard", str身份证);
        }

        /// <summary>
        /// 通过输入身份证号匹配病人信息
        /// </summary>
        public DataTable GetPatlistbyIdCard(string str身份证号)
        {
            if (str身份证号 == "") return null;
            string strJson = JsonConvert.SerializeObject(new
            {
                身份证号 = str身份证号
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Patlistbyidcard", strJson);
        }

        /// <summary>
        /// 通过医保号读取病人信息列表清单
        /// </summary>
        public DataTable Get_Patlistbymedical(string str医保号, string str医保类型)
        {
            if (str医保号 == "" || str医保类型 == "") return null;
            string strJson = JsonConvert.SerializeObject(new
            {
                医保号 = str医保号,
                医保类型 = str医保类型
            });
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Patlistbymedical", strJson);
        }

        public DataTable Get报表列表()
        {
            return SQLHelper_Oracle.GetDataTable("Pkg_Pretriage_Dql.Get_Reportlist", "");
        }

    }
}
