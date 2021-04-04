using System.Runtime.InteropServices;

namespace ZLSOFT.HIS.ZyEdit
{
    // 首先建立接口，这个是COM必须使用的   
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]


    public interface Ishowfrm
    {

        /// <summary>
        /// 中医辩证论治
        /// </summary>
        /// <param name="message">JSON字符串:
        /// {
        /// 来源系统, 使用类型, 病人ID, 挂号单,
        /// 门诊号,病人姓名,病人性别,病人年龄,病人民族,出生日期, 诊断ID,
        /// 当前科室ID, 当前科室名, 操作员ID,操作员姓名,
        /// 用户名, 用户密码
        /// }
        /// 说明：传参均为字符串类型
        /// 来源系统(0-ZLHIS/1-新门诊)
        /// 使用类型(0-新增/1-修改)
        /// </param>
        /// <returns>是否进行了新增或修改中医处方</returns>
        /// <returns>strOut JSON字符串:HIS医嘱ID,HIS诊断ID,诊断ID,处方ID</returns>
         bool EditZyInfo(string message, out string strOut);



        /// <summary>
        /// 中医基础数据维护
        /// </summary>
        /// <param name="message">JSON字符串:
        /// {
        /// 来源系统, 站点, 操作员ID,操作员姓名,
        /// 用户名, 用户密码
        /// }
        /// 说明：传参均为字符串类型
        /// 来源系统(0-ZLHIS/1-新门诊)
        /// 使用类型(0-新增/1-修改)
        /// </param>
        /// <returns></returns>
        bool EditZyBase(string message);


        /// <summary>
        /// 检查是否为中医辩证论治下达的医嘱或者诊断
        /// </summary>
        /// <param name="message">JSON字符串:
        /// {
        /// 来源系统, 病人ID, 挂号单,
        /// 门诊号,HIS医嘱ID,HIS诊断ID,用户名, 用户密码
        /// }
        /// 说明：传参均为字符串类型
        /// 来源系统(0-ZLHIS/1-新门诊)
        /// </param>
        /// <returns>CheckDiag:是否为中医辩证论治下达的医嘱或者诊断</returns>
        /// <returns>strOut:诊断ID|处方ID</returns>
        bool CheckDiag(string message, out string strOut);
    }
}
