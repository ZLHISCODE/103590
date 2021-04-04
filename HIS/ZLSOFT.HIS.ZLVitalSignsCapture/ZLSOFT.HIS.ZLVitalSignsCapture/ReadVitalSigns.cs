using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZLSOFT.HIS.ZLVitalSignsCapture
{
    public class ReadVitalSigns
    {
        public string ErrMeassge = "";//用于返回错误信息 每个过程执行前清空

        /// <summary>
        /// 初始化读取生命体征的部件
        /// </summary>
        /// <returns>是否初始化成功</returns>
        public bool IntMain()
        {
            ErrMeassge = "";//清空错误信息
            return false;
        }

        /// <summary>
        /// 读取生命体征信息
        /// </summary>
        /// <returns>生命体征</returns>
        public VitalSignInfo ReadInfo()
        {
            ErrMeassge = "";//清空错误信息
            try
            {
                VitalSignInfo vsInfo = new VitalSignInfo();
                vsInfo.体温 = "";
                vsInfo.心率 = "";
                vsInfo.收缩压 = "";
                vsInfo.舒张压 = "";
                vsInfo.呼吸频率 = "";
                vsInfo.指氧饱和度 = "";
                vsInfo.血糖 = "";
                vsInfo.血钾 = "";
                return vsInfo;
            }
            catch (Exception ex)
            {
                ErrMeassge = ex.Message;//返回错误信息
                return null;
            }
        }

        /// <summary>
        /// 卸载
        /// </summary>
        /// <returns>是否初始化成功</returns>
        public bool UnloadMain()
        {
            ErrMeassge = "";//清空错误信息
            return true;
        }


    }
}
