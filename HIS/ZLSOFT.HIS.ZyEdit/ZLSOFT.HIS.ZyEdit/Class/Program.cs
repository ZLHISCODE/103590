using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ZLSOFT.HIS.ZyEdit.From;
using System.IO;
using System.Collections;
using System.Reflection;

namespace ZLSOFT.HIS.ZyEdit
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]


        /// <summary>
        /// 中医辩证论治
        /// </summary>
        /// <param name="message">JSON字符串:
        /// {
        /// 来源系统,站点, 使用类型, 病人ID, 挂号单,
        /// 门诊号, 处方ID,
        /// 当前科室ID, 当前科室名, 操作员ID,操作员姓名,
        /// 用户名, 用户密码
        /// }
        /// 说明：传参均为字符串类型
        /// 来源系统(0-ZLHIS/1-新门诊)
        /// 使用类型(0-新增/1-修改)
        /// </param>
        /// <returns>是否进行了新增或修改中医处方</returns>

        static void Main()
        {
        string message = JsonConvert.SerializeObject(new
            {
                来源系统 = "0",
                站点 = "-",
                使用类型 = "1",
                病人ID = "1201",
                挂号单 = "S0000063",
                门诊号 = "1602050002",
                病人姓名 = "武大浪",
                病人性别 = "男",
                病人年龄 = "22岁",
                病人民族 = "汉族",
                出生日期 = "1996-08-17",
                诊断ID = "1",
                当前科室ID = "23",
                当前科室名 = "门诊内科",
                操作员ID = "281",
                操作员姓名 = "张永康",
                用户名 = "ZLHIS",
                用户密码 = "aqa",
                TNSNAME = "33.116TestBase"
            });

            //message = "{ "来源系统":"1","站点":"-","使用类型":"0","病人ID":"2167","挂号单":"S0000056","门诊号":"1806250007","病人姓名":"田七","病人性别":"男","病人年龄":"66岁","病人民族":"汉族","出生日期":"1952/6/25","诊断ID":"0","当前科室ID":"23","当前科室名":"门诊内科","操作员ID":"281","操作员姓名":"张永康","用户名":"ZLHIS","用户密码":"AQA"}


            //string message = JsonConvert.SerializeObject(new
            //{
            //    来源系统 = "0",
            //    站点 = "-",
            //    使用类型 = "2",
            //    病人ID = "4416",
            //    挂号单 = "S0000012",
            //    门诊号 = "20096203",
            //    处方ID = "0",
            //    当前科室ID = "56",
            //    当前科室名 = "门诊内科",
            //    操作员ID = "3607",
            //    操作员姓名 = "张恒恒",
            //    用户名 = "ZLHIS",
            //    用户密码 = "his"
            //});

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //frmZyEdit frm = new frmZyEdit(message);frmBase
            frmBase frm = new frmBase(message);
            Application.Run(frm);

        }
    }
}
