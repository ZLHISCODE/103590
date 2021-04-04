using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ZLSOFT.HIS.PreTriage.ComLib
{
    // 首先建立接口，这个是COM必须使用的   
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]

    public interface Ishowfrm
    {

        /// <summary>
        /// 急诊预检分诊工作站
        /// </summary>
        /// <param name="message">JSON字符串:
        /// {
        /// tnsname,用户名, 用户密码,站点
        /// 操作员id,操作员姓名,操作员编码
        /// }
        /// 说明：传参均为字符串类型
        /// </param>
        /// <returns></returns>
        bool EditPreTriage(string message, ADODB.Connection cn);
    }
}
