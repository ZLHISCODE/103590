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
        /// ��Ӧ��HKEY_CLASSES_ROOT����19         
        /// </summary>
        ClassesRoot = 0,
            
        /// <summary>
        /// ��Ӧ��HKEY_CURRENT_USER����
        /// </summary>
        CurrentUser = 1,
        
        /// <summary>
        /// ��Ӧ�� HKEY_LOCAL_MACHINE����
        /// </summary>
        LocalMachine = 2,
        
        /// <summary>
        /// ��Ӧ�� HKEY_USER����
        /// </summary>
        User = 3,
        
        /// <summary>
        /// ��Ӧ��HEKY_CURRENT_CONFIG����
        /// </summary>
        CurrentConfig = 4,

        /// <summary>
        /// ��Ӧ��HKEY_DYN_DATA����
        /// </summary>
        DynDa = 5,
        
        /// <summary>
        /// ��Ӧ��HKEY_PERFORMANCE_DATA����
        /// </summary>
        PerformanceData = 6
    }

    /// <summary> 
    /// ָ����ע����д洢ֵʱ���õ��������ͣ����ʶע�����ĳ��ֵ���������� 
    /// ��Ҫ������
    /// 1.RegistryValueKind.Unknown
    /// 2.RegistryValueKind.String 
    /// 3.RegistryValueKind.ExpandString 
    /// 4.RegistryValueKind.Binary
    /// 5.RegistryValueKind.DWord10
    /// 6.RegistryValueKind.MultiString
    /// 7.RegistryValueKind.QWord
    
    /// �汾:1.0 
    /// </summary>
    public enum RegValueKind
    {
        /// <summary>
        /// ָʾһ������֧�ֵ�ע����������͡����磬��֧�� Microsoft Win32 API ע����������� REG_RESOURCE_LIST��ʹ�ô�ֵָ��
        /// </summary>
        Unknown = 0,
        /// <summary>
        /// ָ��һ���� Null ��β���ַ�������ֵ�� Win32 API ע����������� REG_SZ ��Ч��
        /// </summary>
        String = 1,
        /// <summary>
        /// ָ��һ���� NULL ��β���ַ��������ַ����а����Ի����������� %PATH%����ֵ������ʱ���ͻ�չ������δչ�������á�
        /// ��ֵ�� Win32 APIע����������� REG_EXPAND_SZ ��Ч��
        /// </summary>
        ExpandString = 2,
        /// <summary>
        /// ָ�������ʽ�Ķ��������ݡ���ֵ�� Win32 API ע����������� REG_BINARY ��Ч��
        /// </summary>
        Binary = 3,
        /// <summary>
        /// ָ��һ�� 32 λ������������ֵ�� Win32 API ע����������� REG_DWORD ��Ч��
        /// </summary>
        DWord = 4,
        /// <summary>
        /// ָ��һ���� NULL ��β���ַ������飬���������ַ���������ֵ�� Win32 API ע����������� REG_MULTI_SZ ��Ч��
        /// </summary>
        MultiString = 5,
        /// <summary>
        /// ָ��һ�� 64 λ������������ֵ�� Win32 API ע����������� REG_QWORD ��Ч��
        /// </summary>
        QWord = 6
    }

   
    /// <summary>
    /// ע��������
    /// ��Ҫ�������²�����
    /// 1.����ע����� 
    /// 2.��ȡע�����
    /// 3.�ж�ע������Ƿ���� 
    /// 4.ɾ��ע�����
    /// 5.����ע����ֵ
    /// 6.��ȡע����ֵ
    /// 7.�ж�ע����ֵ�Ƿ����
    /// 8.ɾ��ע����ֵ
    /// �汾:1.0     
    /// </summary> 
    public class RegisterEx 
    {
        private const string REG_DEFAULT_NODE = "Software\\";
        private const string REG_CUR_ROOT = "/";

        #region �ֶζ���

        /// <summary> 
        /// ע���������
        /// </summary>
        private string _regPath;

        /// <summary>
        /// ע��������
        /// </summary>  
        private RegDomain _regDomain; 

        #endregion 


        #region ����

        /// <summary>
        /// ����ע���������
        /// </summary> 
        public string RegPath
        { 
            get { return _regPath; }
            set { _regPath = value; } 
        } 
        
        /// <summary>
        /// ע��������
        /// </summary> 
        public RegDomain Domain 
        { 
            get { return _regDomain; }
            set { _regDomain = value; }
        }

        #endregion


        #region ���캯��

        public RegisterEx()
            : this(REG_DEFAULT_NODE, RegDomain.CurrentUser)
        { 
        }

        public RegisterEx(string subKey)
            :this(subKey, RegDomain.CurrentUser)
        {
        }
        
        /// <summary> 
        /// ���캯��
        /// </summary>
        /// <param name="subKey">ע���������</param> 
        /// <param name="regDomain">ע��������</param> 
        public RegisterEx(string subKey, RegDomain regDomain)
        { 
            ///����ע���������Software\\ZLSoft����Software\\\ZLSoft�滻ΪSoftware\ZLSoft
            _regPath = (subKey + @"\").Replace(@"\\", @"\").Replace(@"\\", @"\");

            ///����ע�������� 
            _regDomain = regDomain;
        } 

        #endregion

        /// <summary>
        /// ��ȡע����ȫ·��
        /// </summary>
        /// <param name="regSubPath"></param>
        /// <returns></returns>
        private string GetFullPath(string regSubPath)
        {
            if (regSubPath == REG_CUR_ROOT) return _regPath;

            return _regPath + regSubPath + @"\";
        }


        #region ���з���
            
        #region ����ע����� 

        /// <summary>
        /// ����ע����Ĭ�ϴ�����ע������ HKEY_LOCAL_MACHINE���棨��������SubKey���ԣ�
        /// �鷽��������ɽ�����д 
        /// </summary>
        public virtual RegisterEx CreateSubKey()
        {
            return CreateSubKey(REG_CUR_ROOT);
        }

        /// <summary>
        /// ����ע������������SubKey���ԣ� 170   
        /// �鷽��������ɽ�����д 171     
        /// ���ӣ���regDomain��HKEY_LOCAL_MACHINE��subkey��software\\higame\\���򽫴���HKEY_LOCAL_MACHINE\\software\\higame\\ע����� 172     
        /// </summary> 173      
        /// <param name="regSubPath">ע���������</param> 174      
        /// <param name="regDomain">ע��������</param> 175    
        public virtual RegisterEx CreateSubKey(string regSubPath) 
        { 
            //�ж�ע����������Ƿ�Ϊ�գ����Ϊ�գ�����false
            if (regSubPath == string.Empty || regSubPath == null)
            { 
                return null; 
            } 
            
            //��������ע������Ľڵ� 
            RegistryKey key = GetRegDomain(_regDomain);
            try
            {
                //Ҫ������ע�����Ľڵ� 
                if (!IsSubKeyExist(regSubPath))
                {
                    using (RegistryKey item = key.CreateSubKey(GetFullPath(regSubPath)))
                    {
                        //itemΪ�գ���ע���ڵ㴴��ʧ��
                        if (item == null) return null;

                        item.Close();
                    }

                    //�����ӽڵ�Ĳ�������
                    return new RegisterEx(GetFullPath(regSubPath));
                }
                else
                {
                    //�ڵ���ڣ���ֱ�ӷ���ע���ڵ����
                    return new RegisterEx(GetFullPath(regSubPath));
                }
            }
            finally
            {
                //�رն�ע�����ĸ��� 
                key.Close();
            }
        }

        #endregion
        

        #region �ж�ע������Ƿ���� 

        /// <summary> 
        /// �ж�ע������Ƿ���ڣ�Ĭ������ע������HKEY_LOCAL_MACHINE���жϣ���������SubKey���ԣ� 
        /// �鷽��������ɽ�����д 
        /// ���ӣ����������Domain��SubKey���ԣ����ж�Domain\\SubKey������Ĭ���ж�HKEY_LOCAL_MACHINE\\software\\ 
        /// </summary> 204         
        /// <returns>����ע������Ƿ���ڣ����ڷ���true�����򷵻�false</returns> 
        public virtual bool IsSubKeyExist() 
        {
            return IsSubKeyExist(REG_CUR_ROOT);
        }       
        
        /// <summary> 
        /// �ж�ע������Ƿ����
        /// �鷽��������ɽ�����д
        /// ���ӣ���regDomain��HKEY_CLASSES_ROOT��subkey��software\\higame\\�����ж�HKEY_CLASSES_ROOT\\software\\higame\\ע������Ƿ����
        /// </summary> 
        /// <param name="regSubPath">ע���������</param> 
        /// <param name="regDomain">ע��������</param> 
        /// <returns>����ע������Ƿ���ڣ����ڷ���true�����򷵻�false</returns> 
        public virtual bool IsSubKeyExist(string regSubPath) 
        { 
            //�ж�ע����������Ƿ�Ϊ�գ����Ϊ�գ�����false 
            if (regSubPath == string.Empty || regSubPath == null)
            { 
                return false; 
            }
            //����ע������� 
            //���sKeyΪnull,˵��û�и�ע�������ڣ��������
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

        #region ɾ��ע����� 

        /// <summary> 
        /// ɾ��ע������������SubKey���ԣ�
        /// �鷽��������ɽ�����д
        /// </summary> 
        /// <returns>���ɾ���ɹ����򷵻�true������Ϊfalse</returns>
        public virtual bool DeleteSubKey()
        {
            return DeleteSubKey(REG_CUR_ROOT);
        }        

        /// <summary>
        /// ɾ��ע�����
        /// �鷽��������ɽ�����д 
        /// </summary>
        /// <param name="regSubPath"></param>
        /// <param name="regDomain"></param>
        /// <returns></returns>
        public virtual bool DeleteSubKey(string regSubPath)
        { 
            //�ж�ע����������Ƿ�Ϊ�գ����Ϊ�գ�����false 
            if (regSubPath == string.Empty || regSubPath == null) 
            { 
                return false; 
            } 

            //��������ע������Ľڵ� 
            using(RegistryKey item = GetRegDomain(_regDomain))
            {
                if (IsSubKeyExist(regSubPath))
                {
                    //ɾ��ע����� 
                    item.DeleteSubKey(GetFullPath(regSubPath));
                }

                //�رն�ע�����ĸ���              
                item.Close();
            }
            return true;        
        }        

        #endregion          

        #region �жϼ�ֵ�Ƿ����         
                
        /// <summary>
        /// �жϼ�ֵ�Ƿ���ڣ���������SubKey���ԣ�          
        /// �鷽��������ɽ�����д          
        /// ���SubKeyΪ�ա�null����SubKeyָ����ע�������ڣ�����false      
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public virtual bool IsRegKeyExist(string key)    
        {
            return IsRegKeyExist(key, REG_CUR_ROOT); 
        }

        /// <summary> 
        /// �жϼ�ֵ�Ƿ���� 
        /// �鷽��������ɽ�����д 
        /// </summary> 
        /// <param name="key">��ֵ����</param> 
        /// <param name="regSubPath">ע���������</param>
        /// <param name="regDomain">ע��������</param> 
        /// <returns>���ؼ�ֵ�Ƿ���ڣ����ڷ���true�����򷵻�false</returns> 
        public virtual bool IsRegKeyExist(string key, string regSubPath) 
        { 
            //���ؽ�� 
            bool result = false; 

            //�ж��Ƿ����ü�ֵ���� 
            if (key == string.Empty || key == null) 
            { 
                return false; 
            } 
            //�ж�ע������Ƿ����
            if (IsSubKeyExist())
            {
                //��ע����� 
                using (RegistryKey item = OpenSubPathWithMS(regSubPath))
                {
                    //��ֵ���� 
                    string[] regeditKeyNames;
                    //��ȡ��ֵ���� 
                    regeditKeyNames = item.GetValueNames();
                    //������ֵ���ϣ�������ڼ�ֵ�����˳����� 
                    foreach (string regeditKey in regeditKeyNames)
                    {
                        if (string.Compare(regeditKey, key, true) == 0)
                        {
                            result = true;
                            break;
                        }
                    }

                    //�رն�ע�����ĸ��� 
                    item.Close();
                }
            }
            return result; 
        } 

        #endregion

        #region ���ü�ֵ���� 

        /// <summary>
        /// ����ָ���ļ�ֵ���ݣ���ָ�������������ͣ���������SubKey���ԣ�
        /// ���ڸļ�ֵ���޸ļ�ֵ���ݣ������ڼ�ֵ���ȴ�����ֵ�������ü�ֵ����
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public virtual bool WriteRegKey(string key, object value) 
        {
            return WriteRegKey(key, value, RegValueKind.String);
        }
        
        /// <summary>
        /// ����ָ���ļ�ֵ���ݣ���ָ�������������ͣ���������SubKey���ԣ�
        /// ���ڸļ�ֵ���޸ļ�ֵ���ݣ������ڼ�ֵ���ȴ�����ֵ�������ü�ֵ����
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <param name="regValueKind"></param>
        /// <returns></returns>
        public virtual bool WriteRegKey(string key, object value, RegValueKind regValueKind) 
        { 
            //���ؽ�� 
            bool result = false;
            //�жϼ�ֵ�Ƿ���� 
            if (key == string.Empty || key == null) 
            { 
                return false; 
            } 
            //�ж�ע������Ƿ���ڣ���������ڣ���ֱ�Ӵ��� 
            if (!IsSubKeyExist(REG_CUR_ROOT)) 
            {
                CreateSubKey(REG_CUR_ROOT); 
            }
            //�Կ�д��ʽ��ע����� 
            using (RegistryKey item = OpenSubPathWithMS(REG_CUR_ROOT))
            {
                //���ע������ʧ�ܣ��򷵻�false 
                if (key == null)
                {
                    return false;
                }

                item.SetValue(key, value, GetRegValueKind(regValueKind));
                result = true;

                //�رն�ע�����ĸ��� 
                item.Close();
            }

            return result; 
        } 

        #endregion

        #region ��ȡ��ֵ����
        
        /// <summary>
        /// ��ȡ��ֵ���ݣ���������SubKey���ԣ�
        /// 1.���SubKeyΪ�ա�null����SubKeyָʾ��ע�������ڣ�����null
        /// 2.��֮���򷵻ؼ�ֵ����
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
        /// ��ȡ��ֵ����
        /// </summary>
        /// <param name="key"></param>
        /// <param name="regSubPath"></param>
        /// <param name="regDomain"></param>
        /// <returns></returns>
        public virtual object ReadRegKey(string key, string regSubPath) 
        {
            try
            {
                //��ֵ���ݽ�� 
                object obj = null;
                //�ж��Ƿ����ü�ֵ���� 
                if (key == string.Empty || key == null)
                {
                    return null;
                }
                //�жϼ�ֵ�Ƿ���� 
                if (IsRegKeyExist(key, regSubPath))
                {
                    //��ע����� 
                    using (RegistryKey item = OpenSubPathWithMS(regSubPath))
                    {
                        if (key != null)
                        {
                            obj = item.GetValue(key);
                        }

                        //�رն�ע�����ĸ��� 
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
 
        #region ɾ����ֵ      
        
        /// <summary>
        /// ɾ����ֵ����������SubKey���ԣ�
        /// ���SubKeyΪ�ա�null����SubKeyָʾ��ע�������ڣ�����false 
        /// </summary> 
        /// <param name="key">��ֵ����</param>
        /// <returns>���ɾ���ɹ�������true�����򷵻�false</returns>
        public virtual bool DeleteRegeditKey(string key) 
        {
            return DeleteRegeditKey(key, REG_CUR_ROOT);
        } 
        
        /// <summary>
        /// ɾ����ֵ
        /// </summary>
        /// <param name="key"></param>
        /// <param name="regSubPath"></param>
        /// <param name="regDomain"></param>
        /// <returns></returns>
        public virtual bool DeleteRegeditKey(string key, string regSubPath)
        {
            //�жϼ�ֵ���ƺ�ע����������Ƿ�Ϊ�գ����Ϊ�գ��򷵻�false
            if (key == string.Empty || key == null || regSubPath == string.Empty || regSubPath == null)
            {
                return false;
            }
            //�жϼ�ֵ�Ƿ����
            if (IsRegKeyExist(key))
            {
                //�Կ�д��ʽ��ע�����
                using (RegistryKey item = OpenSubPathWithMS(regSubPath))
                {
                    if (key != null)
                    {
                        //ɾ����ֵ
                        item.DeleteValue(key);
                        
                        //�رն�ע�����ĸ���
                        item.Close();
                    }
                }
            }
            return true;
        }

        #endregion

        #endregion

        #region �ܱ�������

        /// <summary>
        /// ��ȡע���������Ӧ�����ڵ�
        /// ���ӣ���regDomain��ClassesRoot���򷵻�Registry.ClassesRoot
        /// </summary>
        /// <param name="regDomain">ע��������</param>
        /// <returns>ע���������Ӧ�����ڵ�</returns>
        protected RegistryKey GetRegDomain(RegDomain regDomain)
        {
            //��������ע������Ľڵ�
            RegistryKey key;
            #region �ж�ע��������
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
        /// ��ȡ��ע����ж�Ӧ��ֵ��������
        /// ���ӣ���regValueKind��DWord���򷵻�RegistryValueKind.DWord
        /// </summary>
        /// <param name="regValueKind">ע�����������</param>
        /// <returns>ע����ж�Ӧ����������</returns>
        protected RegistryValueKind GetRegValueKind(RegValueKind regValueKind)
        {
            RegistryValueKind regValueK;
            #region �ж�ע�����������
            
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

        #region ��ע�����

        /// <summary>
        /// ��ע�����ڵ�
        /// �鷽��������ɽ�����д
        /// </summary>
        /// <returns></returns>
        public virtual RegistryKey OpenSubPathWithMS()
        {
            return OpenSubPathWithMS(REG_CUR_ROOT);
        }      
        
        /// <summary>
        /// ��ע�����ڵ�
        /// �鷽��������ɽ�����д
        /// </summary>
        /// <param name="regSubPath"></param>
        /// <param name="regDomain"></param>
        /// <returns></returns>
        public virtual RegistryKey OpenSubPathWithMS(string regSubPath)
        {
            //�ж�ע����������Ƿ�Ϊ��
            if (regSubPath == string.Empty || regSubPath == null)
            {
                return null;
            }
            //Ҫ�򿪵�ע�����Ľڵ�
            RegistryKey sKey = null;

            //��������ע������Ľڵ�
            using (RegistryKey key = GetRegDomain(_regDomain))
            {
                    //��ע�����
                    sKey = key.OpenSubKey(GetFullPath(regSubPath), true);

                    //�رն�ע�����ĸ���
                    key.Close();
            }

            //����ע���ڵ�
            return sKey;
        }
        
        #endregion

        /// <summary>
        /// ��ȡ��ֵ����
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
        /// ����������Ϣ
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
        /// �����ӽڵ��ע����������
        /// </summary>
        /// <param name="regSubPath"></param>
        /// <returns></returns>
        public virtual RegisterEx OpenSubPath(string regSubPath)
        {
            return OpenSubPath(regSubPath, true);
        }

        /// <summary>
        /// �����ӽڵ��ע����������
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
