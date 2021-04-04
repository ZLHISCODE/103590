using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace zlShortMsg
{
    public static class LogWriter
    {

        /// <summary>
        /// 采用线程安全的方式书写单行日志文件
        /// </summary>
        /// <param name="strFile">日志文件</param>
        /// <param name="strLog">日志信息</param>
        public static void WriteLog(string strFile, string strLog)
        {
            byte[] encodeingBytes = Encoding.UTF8.GetBytes(strLog + Environment.NewLine );
            using (FileStream logFile =new FileStream(strFile, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Write))
            {
                logFile.Seek(0, SeekOrigin.End);
                logFile.Write(encodeingBytes, 0, encodeingBytes.Length);
            }
        }

    }
}
