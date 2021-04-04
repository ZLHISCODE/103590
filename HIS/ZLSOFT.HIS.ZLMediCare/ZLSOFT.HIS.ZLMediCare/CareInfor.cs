using System;

namespace ZLSOFT.HIS.ZLMediCare
{
    /// <summary>
    /// 封装医保卡信息

    /// </summary>
    /// <returns>医保卡信息</returns> 
    public class CareInfor
    {
        public string cardinfoName { get; set; }//姓名
        public string Card_No { get; set; }//社保卡卡号  
        public string Ic_No { get; set; }//医保应用号
        public string Birthday { get; set; }//出生日期
        public string Id_No { get; set; }//身份证号 
        public string Sex { get; set; }//性别  
        public string Card_Type { get; set; }//险类    
    }

}
