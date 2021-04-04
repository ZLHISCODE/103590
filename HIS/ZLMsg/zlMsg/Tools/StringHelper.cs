using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace zlShortMsg
{
    public static class StringHelper
    {
        /// <summary>
        /// 将字符串转换为数字
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static double Val(this string s, IFormatProvider fmt = null)
        {
            try
            {
                return Convert.ToDouble(s, fmt);
            }
            catch
            {
                return Convert.ToDouble("0", fmt);
            }
        }

        /// <summary>
        /// 对String类进行拓展,添加Like方法
        /// ? 任何单一字符。 
        /// * 零个或多个字符。 
        /// # 任何一个数字 (0–9)。 
        /// </summary>
        /// <param name="strMatch"></param>
        /// <returns></returns>
        public static bool Like(this String s, string strMatch)
        {
            strMatch = strMatch.Replace("*", ".*");  //替换*
            strMatch = strMatch.Replace("?", ".");  //替换?
            strMatch = strMatch.Replace("#", @"\d");  //替换#
            return Regex.IsMatch(s, strMatch, RegexOptions.IgnoreCase);
        }

    }
}
