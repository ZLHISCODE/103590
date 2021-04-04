using System;
using System.IO;


namespace ZLSOFT.HIS.ZLVitalSignsCapture
{
    public class VitalSignInfo
    {
        //病人的属性
        public string 体温 { get; set; }
        public string 心率 { get; set; } 
        public string 收缩压 { get; set; }
        public string 舒张压 { get; set; }
        public string 呼吸频率 { get; set; }
        public string 指氧饱和度 { get; set; }
        public string 血糖 { get; set; }
        public string 血钾 { get; set; }
    }
}
