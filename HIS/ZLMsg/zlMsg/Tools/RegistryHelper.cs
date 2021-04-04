using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace zlShortMsg
{
    public static class RegistryHelper
    {
        //注册表信息全部保存在HKEY_CURRENT_USER\Software\zlMsg下
        private static RegistryKey Key = Registry.CurrentUser;
        private static string path = @"Software\zlMsg";

        /// <summary>
        /// 查看注册表指定项的值
        /// </summary>
        /// <param name="path">路经</param>
        /// <param name="name">项名称</param>
        /// <returns>项值</returns>
        public static string GetValue(string name)
        {
            string value;
            try
            {
                RegistryKey objItem = Key.CreateSubKey(path);
                value = objItem.GetValue(name).ToString();
            }
            catch (Exception)
            {
                return "";
            }
            return value;
        }

        /// <summary>
        /// 设置注册表中的值
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public static void SetValue(string name,string value)
        {
            RegistryKey objItem = Key.CreateSubKey(path);
            objItem.SetValue(name, value);
        }
    }
}
