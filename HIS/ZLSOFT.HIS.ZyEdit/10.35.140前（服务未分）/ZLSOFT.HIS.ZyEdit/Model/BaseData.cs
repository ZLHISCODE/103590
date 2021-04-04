using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ZLSOFT.HIS.ZyEdit.Model
{
    /// <summary>
    /// 基础数据类
    /// </summary>
    public class BaseData
    {
        public string System { set; get; }//来源系统(0-ZLHIS/1-新门诊)
        public string 站点 { set; get; }//站点
        public string UseType { set; get; }//使用类型(0-新增/1-修改/2-查看)
        public string 病人ID { set; get; }//病人ID
        public string 挂号单 { set; get; }//挂号单
        public string 门诊号 { set; get; }//门诊号
        public string Name { set; get; }//姓名
        public string Sex { set; get; }//性别
        public string Age { set; get; }//年龄
        public string 民族 { set; get; }//民族
        public string Date { set; get; }//出生日期,yyyy-MM-dd HH:mm:ss
        public string 诊断ID { set; get; }//诊断ID
        public string DeptID { set; get; }//当前科室ID
        public string DeptName { set; get; }//当前科室名
        public string OperatorID { set; get; }//操作员ID
        public string OperatorName { set; get; }//操作员姓名
        public string UserName { set; get; }//用户名
        public string UserPassword { set; get; }//用户密码
        public string TNSNAME { set; get; }//Oracle实例名
    }
}
