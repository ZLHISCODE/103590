using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace zlShortMsg
{
    public class SendMsg
    {
        public string AppKey { get; set; }
        public string AppSecret { get; set; }
        public string Tel { get; set; }
        public string TemplateCode { get; set; }
        public string TemplateParas { get; set; } //参数模版
        public string SignName { get; set; }
        public string SignNumber{ get; set; }
        public string Paras { get; set; }  //业务参数

        public SendMsg()
        {
        }

        public SendMsg(string appKey, string appSecret, string tel, string templateCode, string templateParas, string signName, string paras)
        {
            AppKey = appKey;
            AppSecret = appSecret;
            Tel = tel;
            TemplateCode = templateCode;
            TemplateParas = templateParas;
            SignName = signName;
            Paras = paras;
        }


        /// <summary>
        /// 获取短信发送的http地址
        /// </summary>
        public virtual string GetMessageUrl()
        {
            return null;
        }

        /// <summary>
        /// 对参数进行Url格式的转换
        /// </summary>
        /// <param name="strValue"></param>
        /// <returns></returns>
        public string UrlEncode(string strValue)
        {
            return UrlEncodeUpper(strValue).Replace("+", "%20").Replace("*", "%2A").Replace("%7E", "~");
        }

        private  string UrlEncodeUpper(string str)
        {
            StringBuilder builder = new StringBuilder();
            string strTmp;
            foreach (char c in str)
            {
                //由于一些服务器要求传入的特殊字符和汉字转码后是大写,就在这里进行特殊处理
                //汉字和特殊字符转码后,长度会大于2 ,英文字母转码不变
                strTmp = HttpUtility.UrlEncode(c.ToString(), Encoding.UTF8);
                if (strTmp.Length > 1)
                {
                    builder.Append(strTmp.ToUpper());
                }
                else
                {
                    builder.Append(c);
                }
            }
            return builder.ToString();
        }


        /// <summary>
        /// 使用HmacSha1进行加密,并返回base64编码的结果
        /// </summary>
        /// <param name="encryptText"></param>
        /// <param name="encryptKey"></param>
        /// <returns></returns>
        public string ToHMACSHA1(string encryptKey, string encryptText)
        {
            //HMACSHA1加密
            HMACSHA1 hmacsha1 = new HMACSHA1();
            hmacsha1.Key = System.Text.Encoding.UTF8.GetBytes(encryptKey);
            byte[] dataBuffer = System.Text.Encoding.UTF8.GetBytes(encryptText);
            byte[] hashBytes = hmacsha1.ComputeHash(dataBuffer);
            return Convert.ToBase64String(hashBytes);
        }
    }
}