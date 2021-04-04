using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MedicareComLib;
using System.Xml;

namespace ZLSOFT.HIS.ZLMediCare
{
    /// <summary>
    /// 北京医保卡读取类
    /// </summary>
    public class CardBeiJing
    {
        OutpatientClass patiRead = null;
        public string ErrMessage = "";//用于返回错误信息
        public CardBeiJing()
        {
            patiRead = new OutpatientClass();
        }
        private XmlDocument GetXmlDoc(string sXML)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(sXML);
            return xmlDoc;
        }

        private XmlNode GetNodeFromPath(XmlNode oParentNode, string sPath)
        {
            XmlNode tmpNode = oParentNode.SelectNodes(sPath)[0];
            return tmpNode;
        }

        private bool CheckOutputState(XmlDocument xmlDoc)
        {
            try
            {
                string sState = GetNodeFromPath(xmlDoc.DocumentElement, "state").Attributes["success"].InnerText;
                if (sState.Equals("true"))
                {
                    return true;
                }
                string sErrMsg = "";


                //读取错误信息
                XmlNodeList errNodes = GetNodeFromPath(xmlDoc.DocumentElement, "state").SelectNodes("error");
                for (int i = 0; i < errNodes.Count; i++)
                {
                    if (errNodes[i].Attributes.Count > 0)
                    {
                        sErrMsg = sErrMsg + (sErrMsg == "" ? "" : Environment.NewLine) + "[" + errNodes[i].Attributes["no"].InnerText + "]" + errNodes[i].Attributes["info"].InnerText;
                    }
                }

                //读取警告信息
                XmlNodeList warNodes = GetNodeFromPath(xmlDoc.DocumentElement, "state").SelectNodes("warning");
                for (int i = 0; i < warNodes.Count; i++)
                {
                    if (warNodes[i].Attributes.Count > 0)
                    {
                        sErrMsg = sErrMsg + (sErrMsg == "" ? "" : Environment.NewLine) + "[" + warNodes[i].Attributes["no"].InnerText + "]" + warNodes[i].Attributes["info"].InnerText;
                    }
                }
                ErrMessage = sErrMsg;
                return false;
            }
            catch (Exception ex)
            {
                ErrMessage = ex.Message;
                return false;
            }
        }

        public bool OpenDevice()
        {
            ErrMessage = "";///调用前清空错误信息
            try
            {
                string sOut = "";
                patiRead.Open(out sOut);
                XmlDocument xmlDoc = GetXmlDoc(sOut);
                bool bRet = CheckOutputState(xmlDoc);
                xmlDoc = null;
                return bRet;
            }
            catch 
            {
                ErrMessage = "医保接口初始化失败,请检查接口文件!";
                return false;
            }
        }

        public bool CloseDevice()
        {
            ErrMessage = "";///调用前清空错误信息
            try
            {
                string sOut = "";

                patiRead.Close(out sOut);

                XmlDocument xmlDoc = GetXmlDoc(sOut);

                bool bRet = CheckOutputState(xmlDoc);
                xmlDoc = null;

                return bRet;
            }
            catch 
            {
                ErrMessage = "医保接口卸载失败,请检查接口文件!";
                return false;
            }
        }

        public CareInfor GetCardInfo()
        {
            ErrMessage = "";///调用前清空错误信息

            try
            {
                string sOut;

                patiRead.GetCardInfo(out sOut);

                XmlDocument xmlDoc = GetXmlDoc(sOut);

                bool bRet = CheckOutputState(xmlDoc);
                if (bRet)
                {
                    CareInfor careinfo = new CareInfor();
                    XmlNode dataNode = GetNodeFromPath(xmlDoc.DocumentElement, "output/ic");
                    careinfo.Ic_No = dataNode.SelectNodes("ic_no")[0].InnerText.Trim();
                    careinfo.cardinfoName = dataNode.SelectNodes("personname")[0].InnerText.Trim();//姓名
                    careinfo.Card_No = dataNode.SelectNodes("card_no")[0].InnerText.Trim();
                    careinfo.Birthday = dataNode.SelectNodes("birthday")[0].InnerText.Trim();//出生日期 19960929
                    careinfo.Sex = dataNode.SelectNodes("sex")[0].InnerText.Trim() == "1" ? "男" : "女";
                    careinfo.Card_Type = "北京市医疗保险";//险类    
                    careinfo.Id_No = dataNode.SelectNodes("id_no")[0].InnerText.Trim();//身份证号 

                    return careinfo;
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                ErrMessage = "医保卡信息读取失败,请检查接口文件!";
                return null;
            }
        }

    }
}
