using System;
using System.IO;
using System.Xml;


namespace ZLSOFT.HIS.ZLMediCare

{

    /// <summary>
    /// 读取医保卡信息类。

    /// </summary>
    public class ReadMediCare
    {
        public string ErrMessage = "";//用于返回错误信息

        private CardBeiJing  ReadCard=null;//医保类
        /// <summary>
        /// 初始化设备
        /// 
        public ReadMediCare() {
            ReadCard = new CardBeiJing();
        }

        /// </summary>
        public bool IntMain()
        {
            bool blnOut = false;
            ErrMessage = "";//方法执行前清空错误信息

            blnOut = ReadCard.OpenDevice();//打开接口
            ErrMessage = ReadCard.ErrMessage;//返回错误信息
            return blnOut;
        }

        /// <summary>
        /// 获取读卡信息方法。

        /// </summary>
        public CareInfor GetCardInfo()
        {
            CareInfor careinfo = null;
            ErrMessage = "";//方法执行前清空错误信息

            careinfo = ReadCard.GetCardInfo();//打开接口
            ErrMessage = ReadCard.ErrMessage;//返回错误信息
            return careinfo;
        }

        /// <summary>
        /// 关闭设备。

        /// </summary>
        public bool UnloadMain()
        {
            bool blnOut = false;
            ErrMessage = "";//方法执行前清空错误信息

            blnOut = ReadCard.CloseDevice();//关闭接口
            ErrMessage = ReadCard.ErrMessage;//返回错误信息
            return blnOut;
        }

    }
}
