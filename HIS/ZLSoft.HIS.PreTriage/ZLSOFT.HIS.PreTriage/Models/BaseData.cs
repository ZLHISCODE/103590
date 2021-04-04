using Oracle.ManagedDataAccess.Client;

namespace ZLSOFT.HIS.PreTriage.Models
{

    /// <summary>
    /// 基础数据类


    /// </summary>
    public static class BaseData
    {
        public static string 站点 { set; get; }//站点
        public static string  SYS { set; get; }//系统号

        public static string OperatorID { set; get; }//操作员ID
        public static string OperatorName { set; get; }//操作员姓名

        public static string OperatorCode { set; get; }//操作员编码

        public static ADODB.Connection gcnOracle { set; get; }//ZLHIS数据库连接

        public static OracleConnection OracleCnn { set; get; }//当前程序Oracle连接
    }
}
