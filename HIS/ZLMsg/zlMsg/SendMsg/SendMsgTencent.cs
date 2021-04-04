using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace zlShortMsg
{
    public class SendMsgTencent : SendMsg
    {
        private string sig; //App 凭证,通过计算得出
        private int nationcode = 86;    //国家码,国内为86
        private string ext; //用户的 session 内容，腾讯 server 回包中会原样返回，可选字段，不需要就是设置为空
        private string extend;  //短信码号扩展号
        private string time;    //请求发起时间，UNIX 时间戳（单位：秒），如果和系统时间相差超过 10 分钟则会返回失败
        private string random;  //随机码
        private string url = "https://yun.tim.qq.com/v5/tlssmssvr/sendsms"; //API调用地址

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
            if (random == "")
            {
                random = MainHelper.getRandom().ToString();
                time = MainHelper.getCurrentTime().ToString();
            }
            return url + "?sdkappid=" + AppKey + "&random=" + random;
        }

        public JObject GetRequestBody()
        {
            if (random == "")
            {
                random = MainHelper.getRandom().ToString();
                time = MainHelper.getCurrentTime().ToString();
            }

            JSONObjectBuilder body = new JSONObjectBuilder()
                .Put("tel", (new JSONObjectBuilder()).Put("nationcode", nationcode).Put("mobile", Tel).Build())
                .Put("sig", calculateSignature(AppKey, (long)random.Val(), (long)time.Val(), Tel))
                .Put("tpl_id", TemplateCode)
                .PutArray("params", Paras.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries))
                .Put("sign", !String.IsNullOrEmpty(SignName) ? SignName : "")
                .Put("time", time)
                .Put("extend", !String.IsNullOrEmpty(extend) ? extend : "")
                .Put("ext", !String.IsNullOrEmpty(ext) ? ext : "");

            return body.Build();
        }



        public static string calculateSignature(string appkey, long random, long time, string phoneNumber)
        {
            StringBuilder builder = new StringBuilder("appkey=")
                .Append(appkey)
                .Append("&random=")
                .Append(random)
                .Append("&time=")
                .Append(time)
                .Append("&mobile=")
                .Append(phoneNumber);

            return sha256(builder.ToString());
        }

        private static string sha256(string rawString)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(rawString);
            byte[] hash = SHA256Managed.Create().ComputeHash(bytes);

            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < hash.Length; i++)
            {
                builder.Append(hash[i].ToString("X2"));
            }
            return builder.ToString().ToLower();
        }

        internal bool ResolveResponse(string strResponse, ref string strErrMessage, ref string strErrCode)
        {
            if (strResponse.Like("*OK*"))
            {
                return true;
            }
            else
            {
                // 错误信息形如: {"result":1004,"errmsg":"package format error,cannot get /tel/mobile"}
                int i, j;

                i = strResponse.IndexOf(":") + ":".Length;
                j = strResponse.IndexOf(",");
                strErrCode = strResponse.Substring(i, j - i);

                strResponse.Substring(j + 1);

                i = strResponse.IndexOf(":") + ":".Length;
                j = strResponse.IndexOf("}");
                strErrMessage = strResponse.Substring(i + 1, j - i - 2);
                return false;
            }

        }

        internal string ResolveText(string templateText, string para)
        {

            //腾讯云短信模版中的参数,用大括号包住参数位置表示如:  {1} {2}
            //没有参数的情况
            if (string.IsNullOrEmpty(para))
            {
                return templateText;
            }
            string[] arrParaSplit = para.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 1; i <= arrParaSplit.Length; i++)
            {
                templateText.Replace("{" + i + "}", arrParaSplit[i - 1]);
            }

            return templateText;
        }
    }
}
