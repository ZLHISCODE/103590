using System;
using System.Windows.Forms;
using DevExpress.UserSkins;
using ZLSOFT.HIS.PreTriage.ComLib;
using ZLSOFT.HIS.PreTriage.Models;

namespace ZLSOFT.HIS.PreTriage
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            BonusSkins.Register();

            //正式用的连接方法
            //SQLHelper_Oracle.IntData("zlhis", "his", "dyey");

            DevExpress.XtraEditors.Controls.Localizer.Active = new MessboxClass();


            BaseData.OperatorID = "1158";
            BaseData.OperatorName = "管理员";
            BaseData.OperatorCode = "9999";

            //测试用的连接方法
            SQLHelper_Oracle.IntData("zlhis", "his", "192.168.33.60", "1521", "TESTBASE");



            //BaseData.OperatorID = "43423";
            //BaseData.OperatorName = "管理员";
            //BaseData.OperatorCode = "nb024H";

            ////测试用的连接方法
            //SQLHelper_Oracle.IntData("zlhis", "his", "192.168.0.60", "1524", "dyey");

            //初始化Oracle连接
            BaseData.OracleCnn = SQLHelper_Oracle.GetOdpConnection();

            Application.Run(new frmMain());
        }
    }
}
