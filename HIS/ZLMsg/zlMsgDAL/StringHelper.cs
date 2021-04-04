using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace zlMsgDAL
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
    }
}
