using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ZLSOFT.HIS.ZLVitalSignsCapture
{
    /// <summary>
    /// 界面病人信息
    /// 由于是界面录入，病人信息一些属性节点可能为空，请自行判断
    /// </summary>
    public class VitalPatiInfo
    {
        public int 病人ID { get; set; }
        public string 姓名 { get; set; }
        public string 性别 { get; set; }

        public string 身份证号 { get; set; }
        public string 医保卡号 { get; set; }

    }
}
