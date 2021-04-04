using System;
using System.IO;


namespace ZLSOFT.HIS.ZLIDCard
{
   public class PersonInfor
    {
        //中国人的属性

        public string Name { get; set; }//姓名
        public string Nation { get; set; }//民族 
        public string Address { get; set; }//地址
        public string Birthday { get; set; }//出生日期-公共属性

        public string Identity { get; set; }//身份证号-公共属性 
        public string Sex { get; set; }//性别-公共属性

        public string Signdate { get; set; }//签发机关-公共属性

        public string ValidtermOfStart { get; set; }//身份证号发售日期-公共属性

        public string ValidtermOfEnd { get; set; }//身份证结束日期 -公共属性

        public string Samid { get; set; }//安全模块号-公共属性

        public string PeopleNation { get; set; }//国籍-公共属性

        public MemoryStream Picture { get; set; }//照片-公共属性


        //foreigner-外国人的属性

        public string ForeigNername { get; set; }//外国人本身姓名

        public string CnName { get; set; }//中文姓名
        public string PeopleNationCode { get; set; }//国籍代码
    }

}
