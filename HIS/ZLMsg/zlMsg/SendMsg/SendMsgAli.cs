using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Configuration;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Configuration;
using Oracle.ManagedDataAccess.Client;

namespace zlShortMsg
{
    /// <summary>
    /// 实现Icomper接口,忽略大小写进行排序
    /// </summary>
    public class AsciiCompare : IComparer<String>
    {
        public int Compare(String x, String y)
        {
            return string.CompareOrdinal(x, y);
        }
    }

    public class SendMsgAli:SendMsg
    {
        private string OutID;   //外部流水字段
        private string RegionId = "cn-hangzhou";    //API支持的RegionID，如短信API的值为：cn-hangzhou
        private string Action = "SendSms";  //	API的命名，固定值，如发送短信API的值为：SendSms
        private string Version = "2017-05-25";  //	API的版本，固定值，如短信API的值为：2017-05-25
        private string Format = "XML";  //没传默认为JSON，可选填值：XML
        private string SignatureMethod = "HMAC-SHA1";   // 加密方法,建议固定值：HMAC-SHA1
        private string SignatureVersion = "1.0";    //建议固定值：1.0
        private string SignatureNonce = ""; //用于请求的防重放攻击，每次请求唯一	
        private string Signature = "";  //最终生成的签名结果值
        private string url = "http://dysmsapi.aliyuncs.com/?Signature=";    //API调用地址

        private Dictionary<string, string> paras = new Dictionary<string, string>();

        public void Clone(SendMsg s)
        {
           PropertyInfo[] propertyInfos = typeof(SendMsg).GetProperties();

           foreach (PropertyInfo  p in  propertyInfos)
           {
               p.SetValue(this,p.GetValue(s,null),null);
           }
        }

        /// <summary>
        /// 将业务参数填充到模版参数中
        /// </summary>
        /// <param name="TemplatePara">模版参数</param>
        /// <param name="para">业务参数</param>
        /// <returns></returns>
        private string CombinePara(string TemplatePara, string para)
        {
            if (TemplatePara=="" || TemplatePara.Length == 0)
            {
                return "";
            }
            //阿里云业务参数为Json格式,如: {"para1":"v1","para2":"v2"}
            string[] arrTem = TemplatePara.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            string[] arrPara = para.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            string strResult = "{" + "\"" + arrTem[0] + "\":\"" + arrPara[0] + "\"";
            for (int i = 1; i > arrTem.Length; i++)
            {
                strResult = strResult+",\"" + arrTem[0] + "\":\"" + arrPara[0] + "\"";
            }

            strResult = strResult + "}";
            return strResult;
        }

        /// <summary>
        /// 获取调用API的Http地址
        /// </summary>
        /// <returns></returns>
        public override string GetMessageUrl()
        {
            //1.添加系统参数
            paras.Add("SignatureMethod", SignatureMethod);
            paras.Add("SignatureNonce", Guid.NewGuid().ToString("B"));  //使用Guid作为唯一标识
            paras.Add("AccessKeyId", base.AppKey);
            paras.Add("SignatureVersion", SignatureVersion);
            paras.Add("Timestamp", DateTime.Now.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"));    //当前日期,需要转换为GMT格式
            paras.Add("Format",Format);

            //2.添加业务参数
            paras.Add("Action", Action);
            paras.Add("Version", Version);
            paras.Add("RegionId", RegionId);
            paras.Add("PhoneNumbers", Tel);
            paras.Add("SignName", SignName);
            paras.Add("TemplateParam", CombinePara(TemplateParas, Paras));
            paras.Add("TemplateCode", TemplateCode);

            //3.按照Key进行排序
            Dictionary<string, string> ascDictionary = paras.OrderBy(x => x.Key, new AsciiCompare()).ToDictionary(x => x.Key, y => y.Value);

            //4.遍历每一个参数 ,拼接出参数列表
            StringBuilder sortQueryStringTmp = new StringBuilder();
            foreach (KeyValuePair<string, string> kv in ascDictionary)
            {
                sortQueryStringTmp.Append("&").Append(base.UrlEncode(kv.Key)).Append("=").Append(base.UrlEncode(kv.Value));
            }
            String sortedQueryString = sortQueryStringTmp.ToString().Substring(1);// 去除第一个多余的&符号

            //5.参数列表前拼接请求头
            StringBuilder stringToSign = new StringBuilder();
            stringToSign.Append("GET").Append("&");
            stringToSign.Append(base.UrlEncode("/")).Append("&");
            stringToSign.Append(base.UrlEncode(sortedQueryString));

            //6.在请求串前添加Appsecret并签名加密
            String sign = base.ToHMACSHA1(base.AppSecret + "&", stringToSign.ToString());
            sign = base.UrlEncode(sign);    //加密之后,再进行一次URL格式转换

            return url + sign + sortQueryStringTmp;
        }

        /// <summary>
        /// 解析调用API后返回的结果
        /// </summary>
        /// <param name="strRsponse"></param>
        /// <param name="strErrMessage"></param>
        /// <param name="strErrCode"></param>
        /// <returns></returns>
        public bool ResolveResponse(string strRsponse,ref string strErrMessage ,ref string strErrCode)
        {
            if (strRsponse.Like("*<Code>*</Code>*"))
            {
                if (strRsponse.Like("*<Code>Ok</Code>*"))
                {
                    return true;
                }
                else
                {
                    int i = 0;
                    int j = 0;

                    i = strRsponse.IndexOf("<Code>") + "<Code>".Length;
                    j = strRsponse.IndexOf("</Code>");
                    strErrCode = strRsponse.Substring(i, j - i);

                    i = strRsponse.IndexOf("<Message>") + "<Message>".Length;
                    j = strRsponse.IndexOf("</Message>");
                    strErrMessage = strRsponse.Substring(i, j - i);
                    return false;
                }
                    
            }
            else
            {
                strErrMessage = strRsponse;
                strErrCode = "0";
                return false;
            }

        }

        /// <summary>
        /// 解析模版文字和业务参数,返回实际发送的短信文本
        /// </summary>
        /// <param name="templateText"></param>
        /// <param name="parta"></param>
        /// <returns></returns>
        public string ResolveText(string templateText,string paraString)
        {
            //阿里云短信模版中的参数,都用${}进行分隔,如 验证码为：${code}: 

            //没有参数的情况
            if (paraString == "" || paraString.Length == 0)
            {
                return templateText;
            }

            string strMatch = @"\$\{.*?\}";
            string[] arrParaSplit = paraString.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);

            Regex regex = new Regex(strMatch);
            foreach (string s in arrParaSplit)
            {
                templateText = regex.Replace(templateText, s, 1);
            }
            return templateText;
        }
    }
}