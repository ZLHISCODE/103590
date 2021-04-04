using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.VisualBasic;
using System.Runtime.InteropServices;

namespace ZLSOFT.HIS.PreTriage.ComLib
{
    public class DynamicCall : DisposeImp
    {
        private object _objProxy = null;
        private bool _isNewCreate = true;

        private DynamicCall()
        {
        }


        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="objInstance"></param>
        public DynamicCall(object objInstance)
        {
            _objProxy = objInstance;
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="progId">程序ID</param>
        /// <param name="isNewInstance">是否每次创建新的实例</param>
        public DynamicCall(string progId, bool isNewInstance = true )
        {
            _isNewCreate = false;
            if (isNewInstance)
            {
                _objProxy = Interaction.CreateObject(progId);
                _isNewCreate = true;
            }
            else
            {
                _objProxy = Interaction.GetObject("", progId);
                if (_objProxy == null)
                {
                    _objProxy = Interaction.CreateObject(progId);
                    _isNewCreate = true;
                }
            }
        }

        public object Tag
        {
            get;
            set;
        }

        /// <summary>
        /// 判断ProgId是否存在
        /// </summary>
        /// <param name="progId"></param>
        /// <returns></returns>
        static public bool HasProgId(string progId)
        {
            try
            {
                RegisterEx re = new RegisterEx("\\", RegDomain.ClassesRoot);

                return re.IsSubKeyExist(progId);
            }
            catch
            {
                return false;
            }
        }

        public object Call(string methodName, params object[] pars)
        {
            try
            {
                return Interaction.CallByName(_objProxy, methodName, CallType.Method, pars);
            }
            catch (Exception ex)
            {
                throw new Exception("动态调用方法[" + methodName + "]产生异常。" , ex);
            }
        }

        public object Get(string propertyName)
        {
            try
            {
                return Interaction.CallByName(_objProxy, propertyName, CallType.Get);
            }
            catch (Exception ex)
            {
                throw new Exception("动态获取属性[" + propertyName + "]产生异常。", ex);
            }
        }

        public void Set(string propertyName, object value)
        {
            try
            {
                Interaction.CallByName(_objProxy, propertyName, CallType.Set, new object[] { value });
            }
            catch (Exception ex)
            {
                throw new Exception("动态设置属性[" + propertyName + "]产生异常。", ex);
            }            
        }

        public void Let(string propertyName, object value)
        {
            try
            {
                Interaction.CallByName(_objProxy, propertyName, CallType.Let, new object[] { value });
            }
            catch (Exception ex)
            {
                throw new Exception("动态设置属性[" + propertyName + "]产生异常。", ex);
            }
        }

        /// <summary>
        /// 释放com对象，需谨慎考虑
        /// </summary>
        public void ReleaseCom()
        {
            if (_objProxy != null) Marshal.ReleaseComObject(_objProxy);
            _objProxy = null;
        }

        /// <summary>
        /// 释放托管资源
        /// </summary>
        protected override void DisposeHostedRes()
        {
        }

        /// <summary>
        /// 释放非托管资源
        /// </summary>
        protected override void DisposeNotHostedRes()
        {
            if (_isNewCreate)
            {
                //使用CreateObject创建的对象，才需要进行释放
                _objProxy = null;
            }
        }
    }
}
