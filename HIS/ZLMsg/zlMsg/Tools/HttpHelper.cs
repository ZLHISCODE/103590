using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Newtonsoft.Json.Linq;

namespace zlShortMsg
{
    public static class HttpHelper
    {

        /// <summary>
        /// 发送Get请求
        /// </summary>
        /// <param name="getUrl"></param>
        /// <returns></returns>
        public static string HttpGet(String getUrl)
        {
            string result = string.Empty;
            try
            {
                HttpWebRequest wbRequest = (HttpWebRequest)WebRequest.Create(getUrl);
                wbRequest.Method = "GET";
                HttpWebResponse wbResponse = (HttpWebResponse)wbRequest.GetResponse();
                using (Stream responseStream = wbResponse.GetResponseStream())
                {
                    using (StreamReader sReader = new StreamReader(responseStream))
                    {
                        result = sReader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }
            return result;
        }

        /// <summary>
        /// 发送Post请求
        /// </summary>
        /// <param name="postUrl"></param>
        /// <returns></returns>
        public static string HttpPost(string postUrl, string paramData, Dictionary<string, string> headerDic = null)
        {
            string result = string.Empty;
            try
            {
                HttpWebRequest wbRequest = (HttpWebRequest)WebRequest.Create(postUrl);
                wbRequest.Method = "POST";
                wbRequest.ContentType = "application/x-www-form-urlencoded";
                wbRequest.ContentLength = Encoding.UTF8.GetByteCount(paramData);
                if (headerDic != null && headerDic.Count > 0)
                {
                    foreach (var item in headerDic)
                    {
                        wbRequest.Headers.Add(item.Key, item.Value);
                    }
                }
                using (Stream requestStream = wbRequest.GetRequestStream())
                {
                    using (StreamWriter swrite = new StreamWriter(requestStream))
                    {
                        swrite.Write(paramData);
                    }
                }
                HttpWebResponse wbResponse = (HttpWebResponse)wbRequest.GetResponse();
                using (Stream responseStream = wbResponse.GetResponseStream())
                {
                    using (StreamReader sread = new StreamReader(responseStream))
                    {
                        result = sread.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            { }

            return result;
        }

        /// <summary>
        /// Post方式发送请求,并传入Json参数
        /// </summary>
        /// <returns></returns>
        public static string HttpPost(string strUrl,JObject data,string Header ="")
        {
            //发送请求
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strUrl);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";

            //添加头部信息
            //传入字符串格式:  name1||value1;name2||value2
            if (Header != "")
            {
                string[] arrHeaders = Header.Split(';');
                string[] arrHeaderValue;
                foreach (var arr in arrHeaders)
                {
                    arrHeaderValue  = arr.Split(new string[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
                    request.Headers.Add(arrHeaderValue[0], arrHeaderValue[1]);
                }
            }


            byte[] requestData = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(data));
            request.ContentLength = requestData.Length;
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(requestData, 0, requestData.Length);
            requestStream.Close();

            //接收结果
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream responseStream = response.GetResponseStream();
            StreamReader streamReader = new StreamReader(responseStream, Encoding.GetEncoding("utf-8"));
            string responseStr = streamReader.ReadToEnd();
            streamReader.Close();
            responseStream.Close();

            return responseStr;
        }
    }
}
