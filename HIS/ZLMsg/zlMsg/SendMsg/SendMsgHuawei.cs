using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace zlShortMsg
{
    public class SendMsgHuawei:SendMsg
    {
        private string url = "https://api.rtc.huaweicloud.com:10443/sms/batchSendSms/v1"; //API调用地址
        private string sender = "csms12345678"; //国内短信签名通道号或国际/港澳台短信通道号
        string statusCallBack = "";//选填,短信状态报告接收地址,推荐使用域名,为空或者不填表示不接收状态报告

        public void Clone(SendMsg s)
        {
            PropertyInfo[] propertyInfos = typeof(SendMsg).GetProperties();

            foreach (PropertyInfo p in propertyInfos)
            {
                p.SetValue(this, p.GetValue(s, null), null);
            }
        }

        public override string GetMessageUrl()
        {
            return url;
        }

        public string GetHeader()
        {
            string Header = "Authorization||WSSE realm=\"SDP\",profile=\"UsernameToken\",type=\"Appkey\"";
            Header = Header + ";" + "X-WSSE||" + BuildWSSEHeader(AppKey, AppSecret);
            return Header;
        }   

        public JObject GetRequestBody()
        {
            JSONObjectBuilder body = new JSONObjectBuilder()
                .Put("from", SignNumber)
                .Put("to", Tel)
                .Put("templateId", TemplateCode)
                .Put("templateParas","[\"" +  Paras.Replace("||","\",'\"") + "\"]")
                .Put("statusCallback", statusCallBack);
            return body.Build();
        }

        /// <summary>
        /// 构造X-WSSE参数值
        /// </summary>
        /// <param name="appKey"></param>
        /// <param name="appSecret"></param>
        /// <returns></returns>
        public string BuildWSSEHeader(string appKey, string appSecret)
        {
            string now = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"); //Created
            string nonce = Guid.NewGuid().ToString().Replace("-", ""); //Nonce

            byte[] material = Encoding.UTF8.GetBytes(nonce + now + appSecret);
            byte[] hashed = SHA256Managed.Create().ComputeHash(material);
            string hexdigest = BitConverter.ToString(hashed).Replace("-", "");
            string base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(hexdigest)); //PasswordDigest

            return String.Format("UsernameToken Username=\"{0}\",PasswordDigest=\"{1}\",Nonce=\"{2}\",Created=\"{3}\"",
                appKey, base64, nonce, now);
        }

        internal bool ResolveResponse(string strResponse, ref string strErrMessage, ref string strErrCode)
        {
            return true;
        }

        internal string ResolveText(string templateText, string para)
        {
            return "";
        }
    }
}
