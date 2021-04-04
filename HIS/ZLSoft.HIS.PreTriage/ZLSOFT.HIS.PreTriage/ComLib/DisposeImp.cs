using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ZLSOFT.HIS.PreTriage.ComLib
{
    /// <summary>
    /// 使用IDisposable方式释放对象资源
    /// </summary>
    public abstract class DisposeImp : IDisposable
    {
        private bool _isDisposed;

        /// <summary>
        /// 释放状态
        /// </summary>
        protected bool Disposed
        {
            get
            {
                return _isDisposed;
            }
        }

        #region 构造方法

        public DisposeImp()
        {
            _isDisposed = false;
        }

        #endregion

        #region 析构方法

        ~DisposeImp()
        {
            Dispose(false);
        }

        #endregion

        public virtual void Dispose()
        {
            //当超出对象的作用域后，系统会自动调用实现了IDisposable接口的Dispose方法
            Dispose(true);

            //将对象从垃圾回收器链表中移除
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (!this._isDisposed)
            {
                if (disposing)
                {
                    //释放托管资源
                    DisposeHostedRes();
                }

                //释放非托管资源，如果有
                DisposeNotHostedRes();
            }

            this._isDisposed = true;
        }

        /// <summary>
        /// 释放托管资源
        /// </summary>
        protected abstract void DisposeHostedRes();

        /// <summary>
        /// 释放非托管资源
        /// </summary>
        protected abstract void DisposeNotHostedRes();

    }
}
