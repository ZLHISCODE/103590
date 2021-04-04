using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;

namespace ZLSOFT.HIS.PreTriage.ComLib
{
    public enum RegDomain
    {
        /// <summary>          
        /// 对应于HKEY_CLASSES_ROOT主键19         
        /// </summary>
        ClassesRoot = 0,
            
        /// <summary>
        /// 对应于HKEY_CURRENT_USER主键
        /// </summary>
        CurrentUser = 1,
        
        /// <summary>
        /// 对应于 HKEY_LOCAL_MACHINE主键
        /// </summary>
        LocalMachine = 2,
        
        /// <summary>
        /// 对应于 HKEY_USER主键
        /// </summary>
        User = 3,
        
        /// <summary>
        /// 对应于HEKY_CURRENT_CONFIG主键
        /// </summary>
        CurrentConfig = 4,

        /// <summary>
        /// 对应于HKEY_DYN_DATA主键
        /// </summary>
        DynDa = 5,
        
        /// <summary>
        /// 对应于HKEY_PERFORMANCE_DATA主键
        /// </summary>
        PerformanceData = 6
    }

    /// <summary> 
    /// 指定在注册表中存储值时所用的数据类型，或标识注册表中某个值的数据类型 
    /// 主要包括：
    /// 1.RegistryValueKind.Unknown
    /// 2.RegistryValueKind.String 
    /// 3.RegistryValueKind.ExpandString 
    /// 4.RegistryValueKind.Binary
    /// 5.RegistryValueKind.DWord10
    /// 6.RegistryValueKind.MultiString
    /// 7.RegistryValueKind.QWord
    
    /// 版本:1.0 
    /// </summary>
    public enum RegValueKind
    {
        /// <summary>
        /// 指示一个不受支持的注册表数据类型。例如，不支持 Microsoft Win32 API 注册表数据类型 REG_RESOURCE_LIST。使用此值指定
        /// </summary>
        Unknown = 0,
        /// <summary>
        /// 指定一个以 Null 结尾的字符串。此值与 Win32 API 注册表数据类型 REG_SZ 等效。
        /// </summary>
        String = 1,
        /// <summary>
        /// 指定一个以 NULL 结尾的字符串，该字符串中包含对环境变量（如 %PATH%，当值被检索时，就会展开）的未展开的引用。
        /// 此值与 Win32 API注册表数据类型 REG_EXPAND_SZ 等效。
        /// </summary>
        ExpandString = 2,
        /// <summary>
        /// 指定任意格式的二进制数据。此值与 Win32 API 注册表数据类型 REG_BINARY 等效。
        /// </summary>
        Binary = 3,
        /// <summary>
        /// 指定一个 32 位二进制数。此值与 Win32 API 注册表数据类型 REG_DWORD 等效。
        /// </summary>
        DWord = 4,
        /// <summary>
        /// 指定一个以 NULL 结尾的字符串数组，以两个空字符结束。此值与 Win32 API 注册表数据类型 REG_MULTI_SZ 等效。
        /// </summary>
        MultiString = 5,
        /// <summary>
        /// 指定一个 64 位二进制数。此值与 Win32 API 注册表数据类型 REG_QWORD 等效。
        /// </summary>
        QWord = 6
    }

   
    /// <summary>
    /// 注册表操作类
    /// 主要包括以下操作：
    /// 1.创建注册表项 
    /// 2.读取注册表项
    /// 3.判断注册表项是否存在 
    /// 4.删除注册表项
    /// 5.创建注册表键值
    /// 6.读取注册表键值
    /// 7.判断注册表键值是否存在
    /// 8.删除注册表键值
    /// 版本:1.0     
    /// </summary> 
    public class RegisterEx 
    {
        private const string REG_DEFAULT_NODE = "Software\\";
        private const string REG_CUR_ROOT = "/";

        #region 字段定义

        /// <summary> 
        /// 注册表项名称
        /// </summary>
        private string _regPath;

        /// <summary>
        /// 注册表基项域
        /// </summary>  
        private RegDomain _regDomain; 

        #endregion 


        #region 属性

        /// <summary>
        /// 设置注册表项名称
        /// </summary> 
        public string RegPath
        { 
            get { return _regPath; }
            set { _regPath = value; } 
        } 
        
        /// <summary>
        /// 注册表基项域
        /// </summary> 
        public RegDomain Domain 
        { 
            get { return _regDomain; }
            set { _regDomain = value; }
        }

        #endregion


        #region 构造函数

        public RegisterEx()
            : this(REG_DEFAULT_NODE, RegDomain.CurrentUser)
        { 
        }

        public RegisterEx(string subKey)
            :this(subKey, RegDomain.CurrentUser)
        {
        }
        
        /// <summary> 
        /// 构造函数
        /// </summary>
        /// <param name="subKey">注册表项名称</param> 
        /// <param name="regDomain">注册表基项域</param> 
        public RegisterEx(string subKey, RegDomain regDomain)
        { 
            ///设置注册表项名称Software\\ZLSoft或者Software\\\ZLSoft替换为Software\ZLSoft
            _regPath = (subKey + @"\").Replace(@"\\", @"\").Replace(@"\\", @"\");

            ///设置注册表基项域 
            _regDomain = regDomain;
        } 

        #endregion

        /// <summary>
        /// 获取注册表的全路径
        /// </summary>
        /// <param name="regSubPath"></param>
        /// <returns></returns>
        private string GetFullPath(string regSubPath)
        {
            if (regSubPath == REG_CUR_ROOT) return _regPath;

            return _regPath + regSubPath + @"\";
        }


        #region 公有方法
            
        #region 创建注册表项 

        /// <summary>
        /// 创建注册表项，默认创建在注册表基项 HKEY_LOCAL_MACHINE下面（请先设置SubKey属性）
        /// 虚方法，子类可进行重写 
        /// </summary>
        public virtual RegisterEx CreateSubKey()
        {
            return CreateSubKey(REG_CUR_ROOT);
        }

        /// <summary>
        /// 创建注册表项（请先设置SubKey属性） 170   
        /// 虚方法，子类可进行重写 171     
        /// 例子：如regDomain是HKEY_LOCAL_MACHINE，subkey是software\\higame\\，则将创建HKEY_LOCAL_MACHINE\\software\\higame\\注册表项 172     
        /// </summary> 173      
        /// <param name="regSubPath">注册表项名称</param> 174      
        /// <param name="regDomain">注册表基项域</param> 175    
        public virtual RegisterEx CreateSubKey(string regSubPath) 
        { 
            //判断注册表项名称是否为空，如果为空，返回false
            if (regSubPath == string.Empty || regSubPath == null)
            { 
                return null; 
            } 
            
            //创建基于注册表基项的节点 
            RegistryKey key = GetRegDomain(_regDomain);
            try
            {
                //要创建的注册表项的节点 
                if (!IsSubKeyExist(regSubPath))
                {
                    using (RegistryKey item = key.CreateSubKey(GetFullPath(regSubPath)))
                    {
                        //item为空，则注册表节点创建失败
                        if (item == null) return null;

                        item.Close();
                    }

                    //返回子节点的操作对象
                    return new RegisterEx(GetFullPath(regSubPath));
                }
                else
                {
                    //节点存在，则直接返回注册表节点对象
                    return new RegisterEx(GetFullPath(regSubPath));
                }
            }
            finally
            {
                //关闭对注册表项的更改 
                key.Close();
            }
        }

        #endregion
        

        #region 判断注册表项是否存在 

        /// <summary> 
        /// 判断注册表项是否存在，默认是在注册表基项HKEY_LOCAL_MACHINE下判断（请先设置SubKey属性） 
        /// 虚方法，子类可进行重写 
        /// 例子：如果设置了Domain和SubKey属性，则判断Domain\\SubKey，否则默认判断HKEY_LOCAL_MACHINE\\software\\ 
        /// </summary> 204         
        /// <returns>返回注册表项是否存在，存在返回true，否则返回false</returns> 
        public virtual bool IsSubKeyExist() 
        {
            return IsSubKeyExist(REG_CUR_ROOT);
        }       
        
        /// <summary> 
        /// 判断注册表项是否存在
        /// 虚方法，子类可进行重写
        /// 例子：如regDomain是HKEY_CLASSES_ROOT，subkey是software\\higame\\，则将判断HKEY_CLASSES_ROOT\\software\\higame\\注册表项是否存在
        /// </summary> 
        /// <param name="regSubPath">注册表项名称</param> 
        /// <param name="regDomain">注册表基项域</param> 
        /// <returns>返回注册表项是否存在，存在返回true，否则返回false</returns> 
        public virtual bool IsSubKeyExist(string regSubPath) 
        { 
            //判断注册表项名称是否为空，如果为空，返回false 
            if (regSubPath == string.Empty || regSubPath == null)
            { 
                return false; 
            }
            //检索注册表子项 
            //如果sKey为null,说明没有该注册表项不存在，否则存在
            using(RegistryKey item = OpenSubPathWithMS(regSubPath))
            {
                if (item == null)
                {
                    return false;
                }

                item.Close();
            }

            return true; 
        } 

        #endregion 

        #region 删除注册表项 

        /// <summary> 
        /// 删除注册表项（请先设置SubKey属性）
        /// 虚方法，子类可进行重写
        /// </summary> 
        /// <returns>如果删除成功，则返回true，否则为false</returns>
        public virtual bool DeleteSubKey()
        {
            return DeleteSubKey(REG_CUR_ROOT);
        }        

        /// <summary>
        /// 删除注册表项
        /// 虚方法，子类可进行重写 
        /// </summary>
        /// <param name="regSubPath"></param>
        /// <param name="regDomain"></param>
        /// <returns></returns>
        public virtual bool DeleteSubKey(string regSubPath)
        { 
            //判断注册表项名称是否为空，如果为空，返回false 
            if (regSubPath == string.Empty || regSubPath == null) 
            { 
                return false; 
            } 

            //创建基于注册表基项的节点 
            using(RegistryKey item = GetRegDomain(_regDomain))
            {
                if (IsSubKeyExist(regSubPath))
                {
                    //删除注册表项 
                    item.DeleteSubKey(GetFullPath(regSubPath));
                }

                //关闭对注册表项的更改              
                item.Close();
            }
            return true;        
        }        

        #endregion          

        #region 判断键值是否存在         
                
        /// <summary>
        /// 判断键值是否存在（请先设置SubKey属性）          
        /// 虚方法，子类可进行重写          
        /// 如果SubKey为空、null或者SubKey指定的注册表项不存在，返回false      
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public virtual bool IsRegKeyExist(string key)    
        {
            return IsRegKeyExist(key, REG_CUR_ROOT); 
        }

        /// <summary> 
        /// 判断键值是否存在 
        /// 虚方法，子类可进行重写 
        /// </summary> 
        /// <param name="key">键值名称</param> 
        /// <param name="regSubPath">注册表项名称</param>
        /// <param name="regDomain">注册表基项域</param> 
        /// <returns>返回键值是否存在，存在返回true，否则返回false</returns> 
        public virtual bool IsRegKeyExist(string key, string regSubPath) 
        { 
            //返回结果 
            bool result = false; 

            //判断是否设置键值属性 
            if (key == string.Empty || key == null) 
            { 
                return false; 
            } 
            //判断注册表项是否存在
            if (IsSubKeyExist())
            {
                //打开注册表项 
                using (RegistryKey item = OpenSubPathWithMS(regSubPath))
                {
                    //键值集合 
                    string[] regeditKeyNames;
                    //获取键值集合 
                    regeditKeyNames = item.GetValueNames();
                    //遍历键值集合，如果存在键值，则退出遍历 
                    foreach (string regeditKey in regeditKeyNames)
                    {
                        if (string.Compare(regeditKey, key, true) == 0)
                        {
                            result = true;
                            break;
                        }
                    }

                    //关闭对注册表项的更改 
                    item.Close();
                }
            }
            return result; 
        } 

        #endregion

        #region 设置键值内容 

        /// <summary>
        /// 设置指定的键值内容，不指定内容数据类型（请先设置SubKey属性）
        /// 存在改键值则修改键值内容，不存在键值则先创建键值，再设置键值内容
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public virtual bool WriteRegKey(string key, object value) 
        {
            return WriteRegKey(key, value, RegValueKind.String);
        }
        
        /// <summary>
        /// 设置指定的键值内容，不指定内容数据类型（请先设置SubKey属性）
        /// 存在改键值则修改键值内容，不存在键值则先创建键值，再设置键值内容
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <param name="regValueKind"></param>
        /// <returns></returns>
        public virtual bool WriteRegKey(string key, object value, RegValueKind regValueKind) 
        { 
            //返回结果 
            bool result = false;
            //判断键值是否存在 
            if (key == string.Empty || key == null) 
            { 
                return false; 
            } 
            //判断注册表项是否存在，如果不存在，则直接创建 
            if (!IsSubKeyExist(REG_CUR_ROOT)) 
            {
                CreateSubKey(REG_CUR_ROOT); 
            }
            //以可写方式打开注册表项 
            using (RegistryKey item = OpenSubPathWithMS(REG_CUR_ROOT))
            {
                //如果注册表项打开失败，则返回false 
                if (key == null)
                {
                    return false;
                }

                item.SetValue(key, value, GetRegValueKind(regValueKind));
                result = true;

                //关闭对注册表项的更改 
                item.Close();
            }

            return result; 
        } 

        #endregion

        #region 读取键值内容
        
        /// <summary>
        /// 读取键值内容（请先设置SubKey属性）
        /// 1.如果SubKey为空、null或者SubKey指示的注册表项不存在，返回null
        /// 2.反之，则返回键值内容
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public virtual object ReadRegKey(string key) 
        {
            try
            {
                return ReadRegKey(key, REG_CUR_ROOT);
            }
            catch 
            {
                return null;
            }
        } 
        
        /// <summary>
        /// 读取键值内容
        /// </summary>
        /// <param name="key"></param>
        /// <param name="regSubPath"></param>
        /// <param name="regDomain"></param>
        /// <returns></returns>
        public virtual object ReadRegKey(string key, string regSubPath) 
        {
            try
            {
                //键值内容结果 
                object obj = null;
                //判断是否设置键值属性 
                if (key == string.Empty || key == null)
                {
                    return null;
                }
                //判断键值是否存在 
                if (IsRegKeyExist(key, regSubPath))
                {
                    //打开注册表项 
                    using (RegistryKey item = OpenSubPathWithMS(regSubPath))
                    {
                        if (key != null)
                        {
                            obj = item.GetValue(key);
                        }

                        //关闭对注册表项的更改 
                        item.Close();
                    }
                }
                return obj;
            }
            catch
            {
                return null;
            }
        } 

        #endregion
 
        #region 删除键值      
        
        /// <summary>
        /// 删除键值（请先设置SubKey属性）
        /// 如果SubKey为空、null或者SubKey指示的注册表项不存在，返回false 
        /// </summary> 
        /// <param name="key">键值名称</param>
        /// <returns>如果删除成功，返回true，否则返回false</returns>
        public virtual bool DeleteRegeditKey(string key) 
        {
            return DeleteRegeditKey(key, REG_CUR_ROOT);
        } 
        
        /// <summary>
        /// 删除键值
        /// </summary>
        /// <param name="key"></param>
        /// <param name="regSubPath"></param>
        /// <param name="regDomain"></param>
        /// <returns></returns>
        public virtual bool DeleteRegeditKey(string key, string regSubPath)
        {
            //判断键值名称和注册表项名称是否为空，如果为空，则返回false
            if (key == string.Empty || key == null || regSubPath == string.Empty || regSubPath == null)
            {
                return false;
            }
            //判断键值是否存在
            if (IsRegKeyExist(key))
            {
                //以可写方式打开注册表项
                using (RegistryKey item = OpenSubPathWithMS(regSubPath))
                {
                    if (key != null)
                    {
                        //删除键值
                        item.DeleteValue(key);
                        
                        //关闭对注册表项的更改
                        item.Close();
                    }
                }
            }
            return true;
        }

        #endregion

        #endregion

        #region 受保护方法

        /// <summary>
        /// 获取注册表基项域对应顶级节点
        /// 例子：如regDomain是ClassesRoot，则返回Registry.ClassesRoot
        /// </summary>
        /// <param name="regDomain">注册表基项域</param>
        /// <returns>注册表基项域对应顶级节点</returns>
        protected RegistryKey GetRegDomain(RegDomain regDomain)
        {
            //创建基于注册表基项的节点
            RegistryKey key;
            #region 判断注册表基项域
            switch (regDomain)
            {
                case RegDomain.ClassesRoot:
                    key = Registry.ClassesRoot; break;
                case RegDomain.CurrentUser:
                    key = Registry.CurrentUser; break;
                case RegDomain.LocalMachine:
                    key = Registry.LocalMachine; break;
                case RegDomain.User:
                    key = Registry.Users; break;
                case RegDomain.CurrentConfig:
                    key = Registry.CurrentConfig; break;
                case RegDomain.DynDa:
                    key = Registry.DynData; break;
                case RegDomain.PerformanceData:
                    key = Registry.PerformanceData; break;
                default:
                    key = Registry.LocalMachine; break;
            }
            
            #endregion
            return key;
        }
        
        /// <summary>
        /// 获取在注册表中对应的值数据类型
        /// 例子：如regValueKind是DWord，则返回RegistryValueKind.DWord
        /// </summary>
        /// <param name="regValueKind">注册表数据类型</param>
        /// <returns>注册表中对应的数据类型</returns>
        protected RegistryValueKind GetRegValueKind(RegValueKind regValueKind)
        {
            RegistryValueKind regValueK;
            #region 判断注册表数据类型
            
            switch (regValueKind)
            {
                case RegValueKind.Unknown:
                    regValueK = RegistryValueKind.Unknown; break;
                case RegValueKind.String:
                    regValueK = RegistryValueKind.String; break;
                case RegValueKind.ExpandString:
                    regValueK = RegistryValueKind.ExpandString; break;
                case RegValueKind.Binary:
                    regValueK = RegistryValueKind.Binary; break;
                case RegValueKind.DWord:
                    regValueK = RegistryValueKind.DWord; break;
                case RegValueKind.MultiString:
                    regValueK = RegistryValueKind.MultiString; break;
                case RegValueKind.QWord:
                    regValueK = RegistryValueKind.QWord; break;
                default:
                    regValueK = RegistryValueKind.String; break;
            }
            #endregion

            return regValueK;
        }

        #region 打开注册表项

        /// <summary>
        /// 打开注册表项节点
        /// 虚方法，子类可进行重写
        /// </summary>
        /// <returns></returns>
        public virtual RegistryKey OpenSubPathWithMS()
        {
            return OpenSubPathWithMS(REG_CUR_ROOT);
        }      
        
        /// <summary>
        /// 打开注册表项节点
        /// 虚方法，子类可进行重写
        /// </summary>
        /// <param name="regSubPath"></param>
        /// <param name="regDomain"></param>
        /// <returns></returns>
        public virtual RegistryKey OpenSubPathWithMS(string regSubPath)
        {
            //判断注册表项名称是否为空
            if (regSubPath == string.Empty || regSubPath == null)
            {
                return null;
            }
            //要打开的注册表项的节点
            RegistryKey sKey = null;

            //创建基于注册表基项的节点
            using (RegistryKey key = GetRegDomain(_regDomain))
            {
                    //打开注册表项
                    sKey = key.OpenSubKey(GetFullPath(regSubPath), true);

                    //关闭对注册表项的更改
                    key.Close();
            }

            //返回注册表节点
            return sKey;
        }
        
        #endregion

        /// <summary>
        /// 获取键值名称
        /// </summary>
        /// <param name="regSubPath"></param>
        /// <returns></returns>
        public virtual string[] GetKeyNames()
        {
            RegistryKey regKey = OpenSubPathWithMS(REG_CUR_ROOT);

            if (regKey == null) return null;

            return regKey.GetValueNames();
        }

        /// <summary>
        /// 返回子项信息
        /// </summary>
        /// <returns></returns>
        public virtual string[] GetSubItemName()
        {
            RegistryKey regKey = OpenSubPathWithMS(REG_CUR_ROOT);

            if (regKey == null) return null;

            return regKey.GetSubKeyNames();
        }


        #region

        /// <summary>
        /// 返回子节点的注册表操作对象
        /// </summary>
        /// <param name="regSubPath"></param>
        /// <returns></returns>
        public virtual RegisterEx OpenSubPath(string regSubPath)
        {
            return OpenSubPath(regSubPath, true);
        }

        /// <summary>
        /// 返回子节点的注册表操作对象
        /// </summary>
        /// <param name="regSubPath"></param>
        /// <returns></returns>
        public virtual RegisterEx OpenSubPath(string regSubPath, bool isAutoCreate)
        {
            if (IsSubKeyExist(regSubPath))
            {
                return new RegisterEx(GetFullPath(regSubPath), _regDomain);
            }
            else
            {
                if (isAutoCreate == false) return null;

                return CreateSubKey(regSubPath);
            }
        }



        #endregion

        #endregion
    }
}
